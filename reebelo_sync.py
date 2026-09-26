"""
Reebelo Stock Sync — core engine
Mister Mobile Singapore

Reads:
  * POS Masterlist  : stock report_used device*.xlsx  (Column F = Total -> Available Quantity)
  * Reebelo export  : reebelo-export-2836-*.csv       (29 columns, Cobalt -> Inventory -> Export)
  * SKU Registry    : Reebelo_Match_Review_DD-MM-YYYY.xlsx  (Locked Matches / New Masterlist SKUs /
                                                             Match Review / Not Selling in Reebelo /
                                                             Not on Reebelo Yet)

Writes:
  * Reebelo_Match_Review_DD-MM-YYYY.xlsx   (the SKU registry for the next run)
  * Reebelo_Stock_Update_DD-MM-YYYY.csv    ->  sku,price,stock,minprice,market
    price / minprice / market are ALWAYS left blank so Cobalt keeps the existing prices.

LOCKED RULES
  - Only column J "stock (you)" is ever changed. Price, min price, competitor stock: never.
  - Used sets only. Brand New / MMACC accessories are skipped (manual).
  - Oversell buffer: POS Available Qty 1-2  ->  0   (configurable)
  - Anything that is not a 100% match is NOT uploaded; it is flagged for review.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import date
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ----------------------------------------------------------------------------------
# Branding (copied from the TikTok Stock Sync tool)
# ----------------------------------------------------------------------------------
HEADER_BG = "FF111111"
HEADER_FG = "FFFFEB00"
TITLE_BG = "FFFFEB00"

SHEET_LOCKED = "Locked Matches"
SHEET_NEW_ML = "New Masterlist SKUs"
SHEET_REVIEW = "Match Review"
SHEET_NOT_SELLING = "Not Selling in Reebelo"
SHEET_NOT_YET = "Not on Reebelo Yet"
SHEET_ERRORS = "Validation Errors"
SHEET_SKUS = "_ReebeloSKUs"
SHEET_SUMMARY = "Summary"

REQUIRED_REGISTRY_SHEETS = [
    SHEET_LOCKED,
    SHEET_NEW_ML,
    SHEET_REVIEW,
    SHEET_NOT_SELLING,
    SHEET_NOT_YET,
]

DECISION_LINKED = "Linked"
DECISION_NOT_SELLING = "Not Selling in Reebelo"
DECISION_NOT_YET = "Not on Reebelo yet"

DV_REVIEW = '"Linked (fill col K),Not Selling in Reebelo,Not on Reebelo yet"'
DV_NEWML = '"Linked (fill col I),Not Selling in Reebelo,Not on Reebelo yet"'

REEBELO_REQUIRED_COLS = [
    "sku",
    "psku",
    "price",
    "min price",
    "stock (you)",
    "stock (all)",
    "reebelo model name",
    "reebelo brand",
    "reebelo category",
    "reebelo color",
    "reebelo condition",
    "reebelo storage",
]

UPLOAD_HEADERS = ["sku", "price", "stock", "minprice", "market"]

# Export-set / regional tokens that must never be sold on Reebelo SG
EXPORT_TOKENS = {
    "JP", "TH", "TW", "HK", "CN", "KR", "MY", "VN", "US", "IND", "INDO", "INDIA",
    "PY", "CH", "UK", "EU", "UAE", "AUS", "GER", "LL", "ZA", "PH", "CA", "MOR",
    "FR", "SA",
}

BRAND_MAP = {
    "IPHONE": "Apple",
    "IPAD": "Apple",
    "APPLE": "Apple",
    "APPLE WATCH": "Apple",
    "MACBOOK": "Apple",
    "SAMSUNG": "Samsung",
    "SAMSUNG WATCH": "Samsung",
    "SAMSUNG TAB": "Samsung",
    "GOOGLE": "Google",
    "GOOGLE WATCH": "Google",
    "PIXEL": "Google",
    "XIAOMI": "Xiaomi",
    "REDMI": "Xiaomi",
    "POCO": "Xiaomi",
    "HUAWEI": "Huawei",
    "HONOR": "Honor",
    "HONOR TABLET": "Honor",
    "ONE PLUS": "OnePlus",
    "ONEPLUS": "OnePlus",
    "OPPO": "OPPO",
    "OPPO WATCH": "OPPO",
    "NOTHING": "Nothing",
    "CMF BY NOTHING": "Nothing",
    "CMF": "Nothing",
    "SONY": "Sony",
    "LENOVO": "Lenovo",
    "VIVO": "Vivo",
    "ASUS": "Asus",
    "MOTOROLA": "Motorola",
    "REALME": "Realme",
}

# Tokens dropped from the model token-set on BOTH sides before comparing
NOISE_TOKENS = {
    "APPLE", "SAMSUNG", "GALAXY", "GOOGLE", "XIAOMI", "HUAWEI", "HONOR", "ONEPLUS",
    "ONE", "OPPO", "NOTHING", "SONY", "LENOVO", "VIVO", "ASUS", "MOTOROLA", "REALME",
    "BY", "STANDARD", "BATTERY", "PHYSICAL", "SIM", "ESIM", "SIMS", "INCH", "GEN",
    "GENERATION", "AND", "THE", "WITH", "NEW", "USED", "REFURBISHED",
}

NETWORK_TOKENS = {"5G", "4G", "LTE", "3G"}
CONNECT_TOKENS = {"WIFI", "CELL", "CELLULAR", "BLUETOOTH", "GPS"}

COLOR_SYNONYMS = {
    "GREY": "GRAY",
    "SPACE GREY": "SPACE GRAY",
    "SPACEGRAY": "SPACE GRAY",
}


# ----------------------------------------------------------------------------------
# Output file names  (same convention as the iShopChangi tool: DD-MM-YYYY)
# ----------------------------------------------------------------------------------
def date_tag(run_date: date) -> str:
    return run_date.strftime("%d-%m-%Y")


def match_review_filename(run_date: date) -> str:
    return f"Reebelo_Match_Review_{date_tag(run_date)}.xlsx"


def stock_update_filename(run_date: date) -> str:
    return f"Reebelo_Stock_Update_{date_tag(run_date)}.csv"


# ----------------------------------------------------------------------------------
# Small helpers
# ----------------------------------------------------------------------------------
def _s(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    return str(v).strip()


def _int(v: Any, default: int = 0) -> int:
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return default
        return int(float(str(v).replace(",", "").strip()))
    except Exception:
        return default


def _id(v: Any) -> str:
    """Normalise a Masterlist Stock Type ID to a clean string (30244.0 -> 30244)."""
    t = _s(v)
    if not t:
        return ""
    try:
        f = float(t)
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    return t


def split_ids(v: Any) -> List[str]:
    t = _s(v)
    if not t:
        return []
    parts = re.split(r"[,;/|\s]+", t)
    return [_id(p) for p in parts if _id(p)]


def norm_color(v: Any) -> str:
    t = re.sub(r"[^A-Z0-9 ]+", " ", _s(v).upper())
    t = re.sub(r"\s+", " ", t).strip()
    t = COLOR_SYNONYMS.get(t, t)
    t = re.sub(r"\bGREY\b", "GRAY", t)
    return t.strip()


def norm_storage(v: Any) -> str:
    """'64GBGB/4' -> 64GB, 'M1 512GB/8' -> 512GB, '1TB' -> 1TB."""
    t = _s(v).upper()
    m = re.search(r"(?<![A-Z0-9])(\d+)\s*(TB|GB|MB)", t)
    if not m:
        return ""
    return f"{int(m.group(1))}{m.group(2)}"


def _split_alpha_digit(text: str) -> str:
    text = re.sub(r"(?<=[A-Z])(?=\d)", " ", text)
    text = re.sub(r"(?<=\d)(?=[A-Z])", " ", text)
    return text


def model_tokens(text: str) -> frozenset:
    """Canonical token set for a model string (POS or Reebelo)."""
    t = _s(text).upper()
    t = t.replace("+", " PLUS ")
    # strip capacity (and the /RAM that may follow it) BEFORE splitting letters from digits
    t = re.sub(r"\d+\s*(?:GB|TB|MB)\s*(?:GB)?\s*/\s*\d+", " ", t)
    t = re.sub(r"\d+\s*(?:GB|TB|MB)\s*(?:GB)?\b", " ", t)
    t = re.sub(r"\b\d+\s*GB\s*RAM\b", " ", t)
    t = re.sub(r"\b(\d+(?:\.\d+)?)\s*(?:MM|INCH|\")", r" \1MM ", t)   # watch / tablet sizes
    t = re.sub(r"(?<=\d)(?:ST|ND|RD|TH)\b", " ", t)                   # 8th Gen -> 8
    # network / connectivity words must go before letters are split from digits (5G -> 5 + G)
    t = re.sub(r"(?<![A-Z0-9])(?:5G|4G|3G|LTE|WI-?FI|CELLULAR|CELL|BLUETOOTH|GPS)(?![A-Z0-9])", " ", t)
    t = re.sub(r"[^A-Z0-9]+", " ", t)
    t = _split_alpha_digit(t)
    toks = [x for x in t.split() if x]
    out = []
    for x in toks:
        if x in NOISE_TOKENS or x in NETWORK_TOKENS or x in CONNECT_TOKENS:
            continue
        if x in ("GB", "TB", "MB", "RAM"):
            continue
        out.append(x)
    return frozenset(out)


def network_of(text: str) -> str:
    t = _s(text).upper()
    for n in ("5G", "4G", "LTE", "3G"):
        if re.search(rf"\b{n}\b", t):
            return "4G" if n == "LTE" else n
    return ""


def conn_of(text: str) -> str:
    """Cellular vs WiFi-only (iPads, watches). WiFi/Bluetooth/GPS are treated as the
    same 'no mobile data' pool; Cellular/LTE as the mobile-data pool."""
    t = _s(text).upper()
    t = re.sub(r"\bE-?SIM\b", " ", t)   # "1 Physical SIM + eSIM" is not a WiFi/Cellular marker
    if re.search(r"\b(CELL|CELLULAR|LTE)\b", t):
        return "CELL"
    if re.search(r"\b(WI-?FI|BLUETOOTH|GPS)\b", t):
        return "WIFI"
    return ""


def ram_of(text: str) -> str:
    t = _s(text).upper()
    m = re.search(r"(\d+)\s*GB\s*RAM", t)
    if m:
        return m.group(1)
    m = re.search(r"\d+\s*(?:GB|TB)\s*/\s*(\d+)", t)
    if m:
        return m.group(1)
    return ""


# ----------------------------------------------------------------------------------
# POS Masterlist
# ----------------------------------------------------------------------------------
@dataclass
class PosRow:
    stock_id: str
    category: str
    brand: str
    model: str
    color: str
    qty: int
    excluded: str = ""          # reason if the row must never be sold
    brand_canon: str = ""
    storage: str = ""
    ram: str = ""
    network: str = ""
    conn: str = ""
    tokens: frozenset = field(default_factory=frozenset)
    color_norm: str = ""

    @property
    def label(self) -> str:
        return f"{self.stock_id}: {self.model} | {self.color}"


def _strip_model_code(model: str) -> str:
    """S26 ULTRA 256GB/12 5G-S948B  ->  S26 ULTRA 256GB/12 5G"""
    return re.sub(r"-\s*[A-Z0-9]{3,8}\s*$", "", _s(model).upper()).strip()


def _rewind(f):
    """Streamlit UploadedFile objects keep their read position between reruns."""
    try:
        f.seek(0)
    except Exception:
        pass
    return f


def load_pos(file_like) -> Tuple[List[PosRow], List[str]]:
    """Return (rows, errors). Only Category == Used rows are returned."""
    errors: List[str] = []
    file_like = _rewind(file_like)
    wb = load_workbook(file_like, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    grid = [list(r) for r in ws.iter_rows(values_only=True)]
    if len(grid) < 3:
        return [], ["POS Masterlist: file has no data rows."]

    head1 = [_s(x) for x in grid[0]]
    head2 = [_s(x) for x in grid[1]]

    qty_idx = None
    for i, h in enumerate(head2):
        if h.lower() == "available quantity":
            qty_idx = i
            break
    if qty_idx is None:
        for i, h in enumerate(head1):
            if h.lower() in ("available quantity", "total available"):
                qty_idx = i
                break
    if qty_idx is None:
        return [], ["POS Masterlist: could not find the 'Total -> Available Quantity' column."]

    def col_of(name: str, default: int) -> int:
        for i, h in enumerate(head1):
            if h.lower() == name:
                return i
        return default

    c_id = col_of("stock type id", 0)
    c_cat = col_of("category", 1)
    c_brand = col_of("brand", 2)
    c_model = col_of("model", 3)
    c_color = col_of("color", 4)

    rows: List[PosRow] = []
    for raw in grid[2:]:
        if raw is None or all(x is None for x in raw):
            continue
        sid = _id(raw[c_id] if c_id < len(raw) else None)
        if not sid:
            continue
        cat = _s(raw[c_cat] if c_cat < len(raw) else "")
        if cat.upper() != "USED":
            continue  # Reebelo sells used sets only
        brand = _s(raw[c_brand] if c_brand < len(raw) else "")
        model = _s(raw[c_model] if c_model < len(raw) else "")
        color = _s(raw[c_color] if c_color < len(raw) else "")
        qty = _int(raw[qty_idx] if qty_idx < len(raw) else 0)

        r = PosRow(sid, cat, brand, model, color, qty)
        mu = _s(model).upper()
        base = _strip_model_code(mu)

        excluded = ""
        if "FREEBIE" in mu:
            excluded = "FREEBIE row"
        else:
            words = set(re.split(r"[^A-Z0-9]+", base))
            hit = words & EXPORT_TOKENS
            if "SRI" in words and "LANKA" in words:
                hit = hit | {"SRI LANKA"}
            if hit:
                excluded = "Export set (" + ", ".join(sorted(hit)) + ")"
        r.excluded = excluded

        r.brand_canon = BRAND_MAP.get(brand.upper().strip(), brand.title())
        r.storage = norm_storage(base)
        r.ram = ram_of(base)
        r.network = network_of(base)
        r.conn = conn_of(base)
        r.color_norm = norm_color(color)
        family = ""
        bu = brand.upper().strip()
        if bu.startswith("IPHONE"):
            family = "IPHONE"
        elif bu.startswith("IPAD"):
            family = "IPAD"
        elif "WATCH" in bu:
            family = "WATCH"
        r.tokens = model_tokens(f"{family} {base}")
        rows.append(r)

    if not rows:
        errors.append("POS Masterlist: no rows with Category = 'Used' were found.")
    return rows, errors


# ----------------------------------------------------------------------------------
# Reebelo export
# ----------------------------------------------------------------------------------
@dataclass
class RebRow:
    idx: int
    sku: str
    psku: str
    model_name: str
    brand: str
    category: str
    color: str
    condition: str
    storage: str
    stock_you: int
    stock_all: int
    status: str
    is_brand_new: bool = False
    brand_canon: str = ""
    ram: str = ""
    network: str = ""
    conn: str = ""
    tokens: frozenset = field(default_factory=frozenset)
    color_norm: str = ""


def _reb_model_base(model_name: str, brand: str, storage: str) -> str:
    """'Apple iPhone 17 Pro - 256GB - 1 Physical SIM + eSIM - Silver - Like New - ...'
       -> 'iPhone 17 Pro'"""
    t = _s(model_name)
    parts = [p.strip() for p in t.split(" - ")]
    base = parts[0] if parts else t
    b = _s(brand)
    if b and base.upper().startswith(b.upper()):
        base = base[len(b):].strip()
    st = _s(storage)
    if st and base.upper().endswith(st.upper()):
        base = base[: -len(st)].strip(" -")
    return base


def load_reebelo(file_like, filename: str = "") -> Tuple[List[RebRow], pd.DataFrame, List[str]]:
    errors: List[str] = []
    file_like = _rewind(file_like)
    name = (filename or getattr(file_like, "name", "") or "").lower()
    if name.endswith((".xlsx", ".xlsm", ".xls")):
        df = pd.read_excel(file_like, dtype=str)
    else:
        df = pd.read_csv(file_like, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REEBELO_REQUIRED_COLS if c not in df.columns]
    if missing:
        errors.append("Reebelo inventory file is missing required column(s): " + ", ".join(missing))
        return [], df, errors

    rows: List[RebRow] = []
    for i, rec in df.iterrows():
        sku = _s(rec.get("sku"))
        if not sku:
            continue
        cond = _s(rec.get("reebelo condition"))
        r = RebRow(
            idx=int(i),
            sku=sku,
            psku=_s(rec.get("psku")),
            model_name=_s(rec.get("reebelo model name")),
            brand=_s(rec.get("reebelo brand")),
            category=_s(rec.get("reebelo category")),
            color=_s(rec.get("reebelo color")),
            condition=cond,
            storage=_s(rec.get("reebelo storage")),
            stock_you=_int(rec.get("stock (you)")),
            stock_all=_int(rec.get("stock (all)")),
            status=_s(rec.get("status")),
        )
        r.is_brand_new = (cond.upper() == "BRAND NEW") or sku.upper().startswith("MMACC")
        r.brand_canon = BRAND_MAP.get(r.brand.upper().strip(), r.brand)
        base = _reb_model_base(r.model_name, r.brand, r.storage)
        r.storage = norm_storage(r.storage) or norm_storage(r.model_name)
        r.ram = ram_of(r.model_name)
        r.network = network_of(r.model_name)
        r.conn = conn_of(r.model_name)
        r.color_norm = norm_color(r.color)
        r.tokens = model_tokens(base)
        rows.append(r)
    return rows, df, errors


# ----------------------------------------------------------------------------------
# Matching
# ----------------------------------------------------------------------------------
def _hard_key(brand: str, storage: str, tokens: frozenset, color: str, conn: str = "") -> Tuple:
    return (brand.upper(), storage, tokens, color, conn)


def compatible(p: PosRow, r: RebRow) -> bool:
    """RAM / network are only compared when BOTH sides state them."""
    if p.ram and r.ram and p.ram != r.ram:
        return False
    if p.network and r.network and p.network != r.network:
        return False
    return True


def build_auto_matches(pos_rows: List[PosRow], reb_rows: List[RebRow]):
    """Return (locked_map, suggestions, pos_hints)
       locked_map  : sku -> pos_id   (strict 1:1, brand+storage+model+colour all equal)
       suggestions : sku -> (pos_id, note)   for the Match Review sheet
       pos_hints   : pos_id -> (sku, note)   for the New Masterlist SKUs sheet
    """
    pos_by_key: Dict[Tuple, List[PosRow]] = {}
    for p in pos_rows:
        if p.excluded or not p.storage and not p.tokens:
            continue
        pos_by_key.setdefault(_hard_key(p.brand_canon, p.storage, p.tokens, p.color_norm, p.conn), []).append(p)

    reb_by_key: Dict[Tuple, List[RebRow]] = {}
    for r in reb_rows:
        if r.is_brand_new:
            continue
        reb_by_key.setdefault(_hard_key(r.brand_canon, r.storage, r.tokens, r.color_norm, r.conn), []).append(r)

    locked: Dict[str, str] = {}
    notes: Dict[str, str] = {}
    for key, rlist in reb_by_key.items():
        plist = pos_by_key.get(key)
        if not plist:
            continue
        plist = [p for p in plist if compatible(p, rlist[0])]
        if len(rlist) == 1 and len(plist) == 1:
            locked[rlist[0].sku] = plist[0].stock_id
        elif plist:
            for r in rlist:
                notes[r.sku] = (
                    f"{len(plist)} POS row(s) and {len(rlist)} Reebelo listing(s) share this "
                    f"model / colour / storage - confirm which listing is the live one"
                )

    # near matches: same brand + storage + model tokens, colour different
    pos_soft: Dict[Tuple, List[PosRow]] = {}
    for p in pos_rows:
        if p.excluded:
            continue
        pos_soft.setdefault((p.brand_canon.upper(), p.storage, p.tokens, p.conn), []).append(p)

    suggestions: Dict[str, Tuple[str, str]] = {}
    for r in reb_rows:
        if r.is_brand_new or r.sku in locked:
            continue
        cands = pos_soft.get((r.brand_canon.upper(), r.storage, r.tokens, r.conn), [])
        cands = [p for p in cands if compatible(p, r)]
        if len(cands) == 1:
            p = cands[0]
            suggestions[r.sku] = (
                p.stock_id,
                f"Colour differs: POS '{p.color}' vs Reebelo '{r.color}' - confirm before linking",
            )
        elif len(cands) > 1:
            suggestions[r.sku] = (
                "",
                "Several POS colours match this model / storage: "
                + ", ".join(sorted({c.color for c in cands}))[:180],
            )
        elif r.sku in notes:
            suggestions[r.sku] = ("", notes[r.sku])

    # fuzzy hints (never auto-locked): same brand + storage + colour, model wording differs
    taken_pos = set(locked.values())
    free_reb = [r for r in reb_rows if not r.is_brand_new and r.sku not in locked]
    by_bsc: Dict[Tuple, List[RebRow]] = {}
    for r in free_reb:
        by_bsc.setdefault((r.brand_canon.upper(), r.storage, r.color_norm), []).append(r)

    pos_hints: Dict[str, Tuple[str, str]] = {}
    for p in pos_rows:
        if p.excluded or p.stock_id in taken_pos or p.qty <= 0:
            continue
        cands = by_bsc.get((p.brand_canon.upper(), p.storage, p.color_norm), [])
        best, best_score = None, 0.0
        for r in cands:
            if not compatible(p, r):
                continue
            union = p.tokens | r.tokens
            if not union:
                continue
            score = len(p.tokens & r.tokens) / len(union)
            if score > best_score:
                best, best_score = r, score
        if best is not None and best_score >= 0.5:
            pos_hints[p.stock_id] = (
                best.sku,
                f"Wording differs ({int(best_score * 100)}% token match) - check the model before linking",
            )
            suggestions.setdefault(best.sku, (p.stock_id, pos_hints[p.stock_id][1]))
    return locked, suggestions, pos_hints


# ----------------------------------------------------------------------------------
# Registry (SKU Registry workbook) I/O
# ----------------------------------------------------------------------------------
@dataclass
class Registry:
    locked: Dict[str, List[str]] = field(default_factory=dict)      # reebelo sku -> [pos ids]
    review_decisions: Dict[str, Tuple[str, str, str]] = field(default_factory=dict)
    #   reebelo sku -> (decision, corrected_pos_id, notes)
    newml_decisions: Dict[str, Tuple[str, str, str]] = field(default_factory=dict)
    #   pos id -> (decision, linked sku, notes)
    not_selling: List[str] = field(default_factory=list)            # pos ids
    not_yet: List[str] = field(default_factory=list)                # pos ids
    present: bool = False


def _sheet_records(ws) -> List[Dict[str, Any]]:
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [_s(h) for h in rows[0]]
    out = []
    for r in rows[1:]:
        if r is None or all(x is None for x in r):
            continue
        out.append({header[i]: r[i] for i in range(min(len(header), len(r)))})
    return out


def _pick(rec: Dict[str, Any], *names: str) -> Any:
    for n in names:
        for k in rec:
            if k.lower().strip() == n.lower().strip():
                return rec[k]
    return None


def _decision_of(raw: Any) -> str:
    t = _s(raw).lower()
    if not t:
        return ""
    if t.startswith("linked"):
        return DECISION_LINKED
    if "not selling" in t:
        return DECISION_NOT_SELLING
    if "not on reebelo" in t:
        return DECISION_NOT_YET
    return ""


def load_registry(file_like) -> Tuple[Registry, List[str]]:
    reg = Registry()
    errors: List[str] = []
    if file_like is None:
        return reg, errors
    file_like = _rewind(file_like)
    wb = load_workbook(file_like, data_only=True, read_only=True)
    missing = [s for s in REQUIRED_REGISTRY_SHEETS if s not in wb.sheetnames]
    if missing:
        errors.append("SKU Registry is missing worksheet(s): " + ", ".join(missing))
        return reg, errors
    reg.present = True

    for rec in _sheet_records(wb[SHEET_LOCKED]):
        sku = _s(_pick(rec, "Reebelo SKU", "sku"))
        ids = split_ids(_pick(rec, "LOCKED Masterlist ID(s)", "Masterlist Stock Type ID"))
        if sku and ids:
            reg.locked.setdefault(sku, [])
            for i in ids:
                if i not in reg.locked[sku]:
                    reg.locked[sku].append(i)

    for rec in _sheet_records(wb[SHEET_REVIEW]):
        sku = _s(_pick(rec, "Reebelo SKU", "sku"))
        if not sku:
            continue
        dec = _decision_of(_pick(rec, "Reviewer Decision"))
        corrected = _id(_pick(rec, "Corrected Masterlist ID", "Link to Masterlist ID"))
        note = _s(_pick(rec, "Notes"))
        if dec or corrected:
            reg.review_decisions[sku] = (dec, corrected, note)

    for rec in _sheet_records(wb[SHEET_NEW_ML]):
        pid = _id(_pick(rec, "Masterlist Stock Type ID"))
        if not pid:
            continue
        dec = _decision_of(_pick(rec, "Reviewer Decision"))
        sku = _s(_pick(rec, "Link to Reebelo SKU", "Link to Reebelo SKU ID"))
        note = _s(_pick(rec, "Notes"))
        if dec or sku:
            reg.newml_decisions[pid] = (dec, sku, note)

    reg.not_selling = [
        _id(_pick(rec, "Masterlist Stock Type ID"))
        for rec in _sheet_records(wb[SHEET_NOT_SELLING])
        if _id(_pick(rec, "Masterlist Stock Type ID"))
    ]
    reg.not_yet = [
        _id(_pick(rec, "Masterlist Stock Type ID"))
        for rec in _sheet_records(wb[SHEET_NOT_YET])
        if _id(_pick(rec, "Masterlist Stock Type ID"))
    ]
    return reg, errors


# ----------------------------------------------------------------------------------
# The sync itself
# ----------------------------------------------------------------------------------
@dataclass
class Options:
    buffer_max: int = 2            # POS qty <= buffer_max becomes 0 (0 disables the buffer)
    relink_returning: bool = True  # move Not Selling / Not on Reebelo Yet back when stock returns
    changed_rows_only: bool = True
    seed_auto_lock: bool = True    # first run: auto-lock 100% exact matches


@dataclass
class SyncResult:
    locked_rows: List[Dict[str, Any]] = field(default_factory=list)
    review_rows: List[Dict[str, Any]] = field(default_factory=list)
    newml_rows: List[Dict[str, Any]] = field(default_factory=list)
    not_selling_rows: List[Dict[str, Any]] = field(default_factory=list)
    not_yet_rows: List[Dict[str, Any]] = field(default_factory=list)
    error_rows: List[Dict[str, Any]] = field(default_factory=list)
    upload_rows: List[Dict[str, Any]] = field(default_factory=list)
    stock_no_match: List[Dict[str, Any]] = field(default_factory=list)
    counts: Dict[str, int] = field(default_factory=dict)
    blocking: List[str] = field(default_factory=list)
    reebelo_skus: List[str] = field(default_factory=list)


def apply_buffer(qty: int, buffer_max: int) -> int:
    if qty <= 0:
        return 0
    if buffer_max > 0 and qty <= buffer_max:
        return 0
    return qty


def run_sync(
    pos_rows: List[PosRow],
    reb_rows: List[RebRow],
    registry: Registry,
    opts: Options,
) -> SyncResult:
    res = SyncResult()
    pos_by_id = {p.stock_id: p for p in pos_rows}
    reb_by_sku = {r.sku: r for r in reb_rows}
    res.reebelo_skus = [r.sku for r in reb_rows if not r.is_brand_new]

    auto_locked, suggestions, pos_hints = build_auto_matches(pos_rows, reb_rows)

    # ---------------- 1. build the locked map -----------------------------------
    locked: Dict[str, List[str]] = {}

    def add_link(sku: str, pid: str, source: str):
        if not sku or not pid:
            return
        r = reb_by_sku.get(sku)
        if r is None:
            res.error_rows.append({
                "Severity": "Warning",
                "Issue": f"{source}: Reebelo SKU '{sku}' is not in today's export - link kept but not uploaded.",
            })
            return
        if r.is_brand_new:
            res.error_rows.append({
                "Severity": "Warning",
                "Issue": f"{source}: '{sku}' is a Brand New / accessory listing - skipped (handled manually).",
            })
            return
        locked.setdefault(sku, [])
        if pid not in locked[sku]:
            locked[sku].append(pid)

    # a) carried over from the registry - locked links NEVER drop
    for sku, ids in registry.locked.items():
        for pid in ids:
            add_link(sku, pid, "Locked Matches")

    # b) reviewer decisions from Match Review
    for sku, (dec, corrected, _note) in registry.review_decisions.items():
        if dec == DECISION_LINKED and corrected:
            add_link(sku, corrected, "Match Review decision")

    # c) reviewer decisions from New Masterlist SKUs
    for pid, (dec, sku, _note) in registry.newml_decisions.items():
        if dec == DECISION_LINKED and sku:
            add_link(sku, pid, "New Masterlist decision")

    # d) first-run seeding / new exact matches
    already_linked_pos = {pid for ids in locked.values() for pid in ids}
    parked = set(registry.not_selling) | set(registry.not_yet)
    for pid, (dec, _sku, _n) in registry.newml_decisions.items():
        if dec in (DECISION_NOT_SELLING, DECISION_NOT_YET):
            parked.add(pid)

    if opts.seed_auto_lock:
        for sku, pid in auto_locked.items():
            if sku in locked or pid in already_linked_pos:
                continue
            if pid in parked and not opts.relink_returning:
                continue
            if pid in parked:
                p = pos_by_id.get(pid)
                if not p or p.qty <= 0:
                    continue
                res.error_rows.append({
                    "Severity": "Warning",
                    "Issue": (
                        f"Re-linked: Masterlist {pid} ({p.model} | {p.color}) was parked in "
                        f"Not Selling / Not on Reebelo Yet but has {p.qty} in POS and an exact "
                        f"Reebelo listing '{sku}' - moved back to Locked Matches."
                    ),
                })
            add_link(sku, pid, "Auto-match")
            already_linked_pos.add(pid)

    linked_pos_ids = {pid for ids in locked.values() for pid in ids}

    # ---------------- 2. locked rows + target stock ------------------------------
    n_buffered = 0
    for sku in sorted(locked):
        r = reb_by_sku.get(sku)
        ids = locked[sku]
        if r is None:
            continue
        qty = 0
        labels, missing = [], []
        for pid in ids:
            p = pos_by_id.get(pid)
            if p is None:
                missing.append(pid)
                labels.append(f"{pid}: (not in today's POS)")
                continue
            if p.excluded:
                res.error_rows.append({
                    "Severity": "Warning",
                    "Issue": f"{sku}: Masterlist {pid} is {p.excluded} - counted as 0.",
                })
                labels.append(f"{pid}: {p.model} | {p.color} (EXCLUDED)")
                continue
            qty += max(0, p.qty)
            labels.append(p.label)
        target = apply_buffer(qty, opts.buffer_max)
        if qty > 0 and target == 0:
            n_buffered += 1
        if missing:
            res.error_rows.append({
                "Severity": "Info",
                "Issue": (f"{sku}: Masterlist ID(s) {', '.join(missing)} are not in today's POS "
                          f"export (sold out) - counted as 0. Link kept."),
            })
        status = "Unchanged"
        if target != r.stock_you:
            status = "UPDATE"
        if qty > 0 and target == 0:
            status = "BUFFER → 0"
        res.locked_rows.append({
            "Reebelo SKU": sku,
            "Reebelo Model Name": r.model_name,
            "Brand": r.brand,
            "Category": r.category,
            "Color": r.color,
            "Condition": r.condition,
            "Storage": r.storage,
            "LOCKED Masterlist ID(s)": ", ".join(ids),
            "Masterlist Model | Color": " ; ".join(labels),
            "POS Available Qty": qty,
            "Current Stock (you)": r.stock_you,
            "Target Stock": target,
            "In Upload CSV": status,
            "# IDs": len(ids),
        })
        if (not opts.changed_rows_only) or target != r.stock_you:
            res.upload_rows.append({"sku": sku, "price": "", "stock": target, "minprice": "", "market": ""})

    # ---------------- 3. Match Review (Reebelo listings not linked) ---------------
    for r in reb_rows:
        if r.is_brand_new or r.sku in locked:
            continue
        prev = registry.review_decisions.get(r.sku, ("", "", ""))
        sugg_id, sugg_note = suggestions.get(r.sku, ("", ""))
        note = prev[2] or sugg_note
        p = pos_by_id.get(sugg_id) if sugg_id else None
        res.review_rows.append({
            "Reebelo SKU": r.sku,
            "Reebelo Model Name": r.model_name,
            "Brand": r.brand,
            "Category": r.category,
            "Color": r.color,
            "Condition": r.condition,
            "Storage": r.storage,
            "Current Seller Stock": r.stock_you,
            "Suggested Masterlist ID": (f"{sugg_id}: {p.model} | {p.color}" if p else sugg_id),
            "Corrected Masterlist ID": prev[1],
            "Reviewer Decision": prev[0],
            "Notes": note,
        })
        if r.stock_you > 0:
            res.stock_no_match.append({
                "Reebelo SKU": r.sku,
                "Reebelo Model Name": r.model_name,
                "Color": r.color,
                "Storage": r.storage,
                "Current Stock (you)": r.stock_you,
            })
            res.error_rows.append({
                "Severity": "Alert",
                "Issue": (f"{r.sku} shows stock {r.stock_you} on Reebelo but has no confirmed POS "
                          f"match - oversell risk. Left untouched; zero it manually or link it."),
            })

    # ---------------- 4. New Masterlist SKUs / parked tabs ------------------------
    not_selling_ids, not_yet_ids = [], []
    for pid in registry.not_selling:
        if pid not in linked_pos_ids:
            not_selling_ids.append(pid)
    for pid in registry.not_yet:
        if pid not in linked_pos_ids:
            not_yet_ids.append(pid)
    for pid, (dec, _sku, _n) in registry.newml_decisions.items():
        if pid in linked_pos_ids:
            continue
        if dec == DECISION_NOT_SELLING and pid not in not_selling_ids:
            not_selling_ids.append(pid)
        elif dec == DECISION_NOT_YET and pid not in not_yet_ids:
            not_yet_ids.append(pid)

    def pos_block(pid: str) -> Dict[str, Any]:
        p = pos_by_id.get(pid)
        return {
            "Masterlist Stock Type ID": pid,
            "Category": p.category if p else "",
            "Brand": p.brand if p else "",
            "Model": p.model if p else "(not in today's POS export)",
            "Color": p.color if p else "",
            "Available Qty": p.qty if p else 0,
        }

    res.not_selling_rows = [pos_block(p) for p in dict.fromkeys(not_selling_ids)]
    res.not_yet_rows = [pos_block(p) for p in dict.fromkeys(not_yet_ids)]

    parked_now = set(not_selling_ids) | set(not_yet_ids)
    for p in pos_rows:
        if p.stock_id in linked_pos_ids or p.stock_id in parked_now:
            continue
        if p.excluded:
            continue
        if p.qty <= 0:
            continue  # only POS SKUs that actually have stock need a decision
        prev = registry.newml_decisions.get(p.stock_id, ("", "", ""))
        cand = [sku for sku, pid in auto_locked.items() if pid == p.stock_id]
        hint_sku, hint_note = pos_hints.get(p.stock_id, ("", ""))
        res.newml_rows.append({
            "Masterlist Stock Type ID": p.stock_id,
            "Category": p.category,
            "Brand": p.brand,
            "Model": p.model,
            "Color": p.color,
            "Available Qty": p.qty,
            "Suggested Reebelo SKU": (cand[0] if cand else hint_sku),
            "Link to Reebelo SKU": prev[1],
            "Reviewer Decision": prev[0],
            "Notes": prev[2] or hint_note or "No Reebelo listing found for this model / colour",
        })

    # ---------------- 5. counts --------------------------------------------------
    excluded_pos = [p for p in pos_rows if p.excluded]
    for p in excluded_pos[:200]:
        res.error_rows.append({
            "Severity": "Info",
            "Issue": f"POS {p.stock_id} {p.model} | {p.color}: skipped - {p.excluded}.",
        })

    res.counts = {
        "locked_total": len(res.locked_rows),
        "locked_updated": len(res.upload_rows),
        "locked_buffered": n_buffered,
        "new_masterlist": len(res.newml_rows),
        "review": len(res.review_rows),
        "not_selling": len(res.not_selling_rows),
        "not_yet": len(res.not_yet_rows),
        "errors": len([e for e in res.error_rows if e["Severity"] in ("Alert", "Error")]),
        "warnings": len([e for e in res.error_rows if e["Severity"] == "Warning"]),
        "unmatched_with_stock": len(res.stock_no_match),
        "brand_new_skipped": len([r for r in reb_rows if r.is_brand_new]),
        "pos_used_rows": len(pos_rows),
        "pos_excluded": len(excluded_pos),
        "reebelo_used_listings": len([r for r in reb_rows if not r.is_brand_new]),
    }
    return res


# ----------------------------------------------------------------------------------
# Output builders
# ----------------------------------------------------------------------------------
def build_upload_csv(res: SyncResult) -> bytes:
    df = pd.DataFrame(res.upload_rows, columns=UPLOAD_HEADERS)
    if df.empty:
        df = pd.DataFrame(columns=UPLOAD_HEADERS)
    buf = io.StringIO()
    df.to_csv(buf, index=False)
    return buf.getvalue().encode("utf-8")


MIN_W, MAX_W = 9.0, 95.0


def _autofit(ws, headers: List[str], rows: List[List[Any]]):
    """Size every column to its widest cell so no text is cut off."""
    ncols = len(headers)
    widths = []
    for i in range(ncols):
        longest = len(str(headers[i]))
        for r in rows:
            if i < len(r):
                v = r[i]
                if v is None:
                    continue
                for line in str(v).split("\n"):
                    longest = max(longest, len(line))
        widths.append(min(max(longest + 3, MIN_W), MAX_W))
    widths[0] = 6.0  # the "#" column
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return widths


def _write_sheet(wb, title: str, headers: List[str], rows: List[Dict[str, Any]],
                 dv: Optional[Tuple[str, str]] = None, dv_rows: int = 700):
    ws = wb.create_sheet(title)
    all_headers = ["#"] + headers
    ws.append(all_headers)
    table = []
    for n, rec in enumerate(rows, start=1):
        line = [n] + [rec.get(h, "") for h in headers]
        table.append(line)
        ws.append(line)

    fill = PatternFill("solid", fgColor=HEADER_BG)
    font = Font(bold=True, color=HEADER_FG)
    for c in range(1, len(all_headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(vertical="center", horizontal="left")
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    _autofit(ws, all_headers, table)

    last = max(1, len(rows) + 1)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(all_headers))}{last}"

    if dv:
        col_letter, formula = dv
        validator = DataValidation(type="list", formula1=formula, allow_blank=True, showDropDown=False)
        ws.add_data_validation(validator)
        end = max(len(rows) + 1, dv_rows) + 200
        validator.add(f"{col_letter}2:{col_letter}{end}")
    return ws


def build_registry_workbook(res: SyncResult, run_date: date, opts: Options) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    # ---- Summary -----------------------------------------------------------
    ws = wb.create_sheet(SHEET_SUMMARY)
    ws["B2"] = "Reebelo Stock Bulk Update — Match Review"
    ws["B2"].fill = PatternFill("solid", fgColor=TITLE_BG)
    ws["B2"].font = Font(bold=True, size=15, color="FF111111")
    ws["B3"] = (f"Mister Mobile Singapore · generated {date_tag(run_date)} · "
                f"POS Masterlist → Reebelo (Cobalt)")
    ws["B3"].font = Font(size=9, color="FF333333")
    c = res.counts
    buf_txt = "off (exact POS qty)" if opts.buffer_max == 0 else f"1–{opts.buffer_max} units → 0"
    lines = [
        ("Locked Matches (total)", c.get("locked_total", 0)),
        ("Locked Matches in upload CSV", c.get("locked_updated", 0)),
        ("Zeroed by oversell buffer", c.get("locked_buffered", 0)),
        ("New Masterlist SKUs to review", c.get("new_masterlist", 0)),
        ("Reebelo listings awaiting review", c.get("review", 0)),
        ("Not Selling in Reebelo", c.get("not_selling", 0)),
        ("Not on Reebelo Yet", c.get("not_yet", 0)),
        ("Alerts (stock > 0, no POS match)", c.get("unmatched_with_stock", 0)),
        ("Warnings", c.get("warnings", 0)),
        ("Brand New / accessory listings skipped", c.get("brand_new_skipped", 0)),
        ("POS used rows read", c.get("pos_used_rows", 0)),
        ("POS rows excluded (export sets / freebies)", c.get("pos_excluded", 0)),
        ("Reebelo used listings in export", c.get("reebelo_used_listings", 0)),
        ("Oversell buffer used", buf_txt),
    ]
    r = 5
    for label, val in lines:
        ws.cell(row=r, column=2, value=label).font = Font(bold=True)
        ws.cell(row=r, column=3, value=val)
        r += 1
    ws.cell(row=r + 1, column=2,
            value=("Reminder: price, minprice and market stay BLANK in the upload CSV "
                   "(columns kept, values empty) — stock only.")).font = Font(color="FFB00020", bold=True)
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["C"].width = 22

    _write_sheet(
        wb, SHEET_LOCKED,
        ["Reebelo SKU", "Reebelo Model Name", "Brand", "Category", "Color", "Condition",
         "Storage", "LOCKED Masterlist ID(s)", "Masterlist Model | Color", "POS Available Qty",
         "Current Stock (you)", "Target Stock", "In Upload CSV", "# IDs"],
        res.locked_rows,
    )

    _write_sheet(
        wb, SHEET_NEW_ML,
        ["Masterlist Stock Type ID", "Category", "Brand", "Model", "Color", "Available Qty",
         "Suggested Reebelo SKU", "Link to Reebelo SKU", "Reviewer Decision", "Notes"],
        res.newml_rows,
        dv=("J", DV_NEWML),
    )

    _write_sheet(
        wb, SHEET_REVIEW,
        ["Reebelo SKU", "Reebelo Model Name", "Brand", "Category", "Color", "Condition",
         "Storage", "Current Seller Stock", "Suggested Masterlist ID", "Corrected Masterlist ID",
         "Reviewer Decision", "Notes"],
        res.review_rows,
        dv=("L", DV_REVIEW),
    )

    for title, rows in ((SHEET_NOT_SELLING, res.not_selling_rows), (SHEET_NOT_YET, res.not_yet_rows)):
        _write_sheet(
            wb, title,
            ["Masterlist Stock Type ID", "Category", "Brand", "Model", "Color", "Available Qty"],
            rows,
        )

    _write_sheet(wb, SHEET_ERRORS, ["Severity", "Issue"], res.error_rows)

    ws = wb.create_sheet(SHEET_SKUS)
    ws.append(["Reebelo SKU (used listings only)"])
    ws["A1"].fill = PatternFill("solid", fgColor=HEADER_BG)
    ws["A1"].font = Font(bold=True, color=HEADER_FG)
    for s in res.reebelo_skus:
        ws.append([s])
    ws.column_dimensions["A"].width = 62
    ws.sheet_state = "hidden"

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

"""
Reebelo Stock Sync — accessories engine
Mister Mobile Singapore

Handles the Brand New MMACC accessory listings, which follow completely different
naming rules from used devices:

  POS accessories export (stock_report_DD-MM-YYYY_accessories_new.xlsx)
      Brand  column holds the CATEGORY   e.g. "TEMPERED GLASS APPLE", "CASE APPLE IPHONE 14"
      Model  column holds the VARIANT    e.g. "IPHONE 14 PRO MAX FULL COVER", "BACK MERCURY JELLY"
      Color  column holds the finish     e.g. CLEAR / MATT / PRIVACY / BLACK

  Reebelo listing
      MMACC - Tempered Glass - Privacy - Apple iPhone 13
      MMACC - Phone Case - Clear - Apple iPhone 14

RULES (confirmed with Carmen, 30 Sep 2026)
  * Grade A / Grade B items (cables, adapters, earpieces) are NEVER auto-matched —
    the retail team still has to confirm whether POS "APPLE ORI" is Grade A or Grade B.
  * Accessories use their own oversell buffer (default: none).
  * Shared pools (one POS row fits several phones, e.g. one glass for 11 Pro / X / XS):
    pool >= share_threshold  -> every listing gets the full quantity
    pool <  share_threshold  -> the pool is split evenly between the listings
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

# Reebelo finish  ->  POS Color
GLASS_FINISHES = {"CLEAR": "CLEAR", "MATT": "MATT", "MATTE": "MATT", "PRIVACY": "PRIVACY",
                  "BLACK": "BLACK"}

# Reebelo case variant -> (POS Model, POS Color)
CASE_VARIANTS = {
    "CLEAR": ("BACK MERCURY JELLY", "CLEAR"),
    "BLACK": ("BACK MERCURY SOFT", "BLACK"),
    "MATTE BLACK": ("BACK MERCURY SOFT", "BLACK"),
    "MATT BLACK": ("BACK MERCURY SOFT", "BLACK"),
}

# Reebelo device wording -> POS device wording
DEVICE_ALIASES = {
    "SE 2020": "SE 2",
    "SE 2022": "SE 3",
    "SE 2ND GEN": "SE 2",
    "SE 3RD GEN": "SE 3",
}

# POS model decorations that are not part of the device name
GLASS_DECOR = ("FULL COVER 360", "FULL COVER", "BACK CAMERA", "360")

BRAND_WORDS = {"APPLE", "SAMSUNG", "OPPO", "XIAOMI", "HONOR", "HUAWEI", "ONEPLUS",
               "ONE PLUS", "GOOGLE", "NOTHING", "VIVO", "REALME", "SONY", "LENOVO"}


def _u(v) -> str:
    return "" if v is None else str(v).replace("\xa0", " ").strip().upper()


def device_tokens(text: str) -> frozenset:
    """'Apple iPhone 14 Pro Max' / 'IPHONE 14 PRO MAX FULL COVER' -> {IPHONE,14,PRO,MAX}"""
    t = _u(text)
    for dec in GLASS_DECOR:
        t = t.replace(dec, " ")
    for a, b in DEVICE_ALIASES.items():
        t = re.sub(rf"\b{re.escape(a)}\b", b, t)
    t = re.sub(r"[^A-Z0-9. ]+", " ", t)
    toks = [x for x in t.split() if x and x not in BRAND_WORDS]
    return frozenset(toks)


# ----------------------------------------------------------------------------------
# POS accessories rows
# ----------------------------------------------------------------------------------
@dataclass
class AccRow:
    stock_id: str
    category: str
    brand: str            # POS "Brand" column = the category string
    model: str
    color: str
    qty: int
    kind: str = "OTHER"                      # GLASS / CASE / OTHER
    device_brand: str = ""                   # APPLE / SAMSUNG / ...
    devices: List[frozenset] = field(default_factory=list)
    finish: str = ""
    decorated: bool = False                  # FULL COVER / 360 / BACK CAMERA
    camera: bool = False

    @property
    def label(self) -> str:
        return f"{self.stock_id}: {self.brand} | {self.model} | {self.color}"


def classify_pos_accessory(stock_id: str, category: str, brand: str, model: str,
                           color: str, qty: int) -> AccRow:
    r = AccRow(stock_id, category, brand.strip(), model.strip(), color.strip(), qty)
    b, m = _u(brand), _u(model)
    r.finish = _u(color)

    if b.startswith("TEMPERED GLASS") and "ALIGNMENT" not in b:
        r.kind = "GLASS"
        r.device_brand = b.replace("TEMPERED GLASS", "").strip()
        r.camera = "BACK CAMERA" in m
        r.decorated = ("FULL COVER" in m) or ("360" in m)
        r.devices = [device_tokens(seg) for seg in m.split("-") if device_tokens(seg)]

    elif b.startswith("CASE "):
        r.kind = "CASE"
        segs = [s.strip() for s in b.split("-") if s.strip()]
        first = segs[0][len("CASE "):].strip()
        words = first.split()
        r.device_brand = words[0] if words else ""
        devices = [first]
        for s in segs[1:]:
            s = s[len("CASE "):].strip() if s.startswith("CASE ") else s
            if not s.startswith(r.device_brand):
                s = f"{r.device_brand} {s}"
            devices.append(s)
        r.devices = [device_tokens(d) for d in devices if device_tokens(d)]

    return r


# ----------------------------------------------------------------------------------
# Reebelo accessory listings
# ----------------------------------------------------------------------------------
@dataclass
class RebAcc:
    sku: str
    kind: str = "OTHER"        # GLASS / CASE / GRADE / OTHER
    variant: str = ""
    brand: str = ""
    devices: frozenset = field(default_factory=frozenset)


def parse_reebelo_accessory(sku: str) -> RebAcc:
    """'MMACC - Tempered Glass - Privacy - Apple iPhone 13' -> GLASS / PRIVACY / APPLE"""
    out = RebAcc(sku)
    parts = [p.strip() for p in str(sku).split(" - ")]
    if len(parts) < 4 or not parts[0].upper().startswith("MMACC"):
        return out
    head, variant, device = _u(parts[1]), _u(parts[2]), " - ".join(parts[3:])
    if head == "TEMPERED GLASS":
        out.kind = "GLASS"
    elif head == "PHONE CASE":
        out.kind = "CASE"
    else:
        out.kind = "GRADE"        # 'Apple A', 'Samsung B', ... cables / adapters / earpieces
        return out
    out.variant = variant
    dev_u = _u(device)
    words = dev_u.split()
    out.brand = words[0] if words and words[0] in BRAND_WORDS else ""
    out.devices = device_tokens(device)
    return out


# ----------------------------------------------------------------------------------
# Matching
# ----------------------------------------------------------------------------------
def build_accessory_matches(acc_rows: List[AccRow], reb_accs: List[Tuple[str, RebAcc]]):
    """reb_accs: [(sku, RebAcc)] for the Brand New / MMACC listings.
       Returns (locked, notes):
         locked : sku -> pos stock_id   (exactly one candidate)
         notes  : sku -> explanation for the Match Review sheet
    """
    glass = [a for a in acc_rows if a.kind == "GLASS" and not a.camera]
    cases = [a for a in acc_rows if a.kind == "CASE"]

    locked: Dict[str, str] = {}
    notes: Dict[str, str] = {}

    for sku, ra in reb_accs:
        if ra.kind == "GRADE":
            notes[sku] = ("OEM Grade A / Grade B item - not auto-matched (retail team still to "
                          "confirm whether POS 'ORI' = Grade A or Grade B)")
            continue
        if ra.kind == "GLASS":
            want = GLASS_FINISHES.get(ra.variant, ra.variant)
            cands = [a for a in glass
                     if a.device_brand == ra.brand
                     and a.finish == want
                     and any(d == ra.devices for d in a.devices)]
        elif ra.kind == "CASE":
            mapping = CASE_VARIANTS.get(ra.variant)
            if not mapping:
                notes[sku] = f"Unknown case variant '{ra.variant}' - link it by hand"
                continue
            want_model, want_color = mapping
            cands = [a for a in cases
                     if a.device_brand == ra.brand
                     and _u(a.model) == want_model
                     and a.finish == want_color
                     and any(d == ra.devices for d in a.devices)]
        else:
            continue

        if len(cands) == 1:
            locked[sku] = cands[0].stock_id
        elif not cands:
            notes[sku] = "No POS accessory row matches this model / finish"
        else:
            notes[sku] = ("Several POS rows match: "
                          + "; ".join(f"{c.stock_id} ({c.model} | {c.color})" for c in cands[:4]))
    return locked, notes


def share_pool(qty: int, n_listings: int, threshold: int = 30) -> int:
    """Hybrid sharing rule for one POS pool feeding several Reebelo listings."""
    if n_listings <= 1:
        return max(0, qty)
    if qty >= threshold:
        return max(0, qty)          # plenty in the pool - show it on every listing
    return max(0, qty // n_listings)  # thin pool - split so Reebelo can never oversell

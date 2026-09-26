# Reebelo Stock Sync Tool

Mister Mobile Singapore — generates the Reebelo (Cobalt) bulk stock update from the POS
used-device stock report, and keeps a Match Review / SKU Registry workbook between runs.
Same workflow and design as the TikTok Stock Sync Tool.

---

## What it does

| Input | Where it comes from |
|---|---|
| **POS Masterlist** | pos.mistermobile.sg → `stock report_used device*.xlsx` (Column F = Total → Available Quantity) |
| **Reebelo Bulk Stock Inventory** | Cobalt → Inventory → **Export** → `reebelo-export-2836-*.csv` (29 columns) |
| **SKU Registry** | the `Reebelo_Match_Review_*.xlsx` this tool produced last run |

| Output | Use |
|---|---|
| `Reebelo_Stock_Upload_YYYY-MM-DD.csv` | Cobalt → Inventory → **Update offers** → Upload CSV |
| `Reebelo_Match_Review_YYYY-MM-DD.xlsx` | review the new/unmatched SKUs, then feed it back next run |

---

## Locked rules

- **Only stock changes.** The upload CSV is `sku,price,stock,minprice,market` with
  **price, minprice and market left blank** — the columns stay, the values are empty, so Cobalt
  keeps the existing price and min price. Column K `stock (all)` (competitor stock) is never touched.
- **Used sets only.** `Brand New` / `MMACC` accessory listings are skipped and handled manually.
- **Oversell buffer** (button on the page): POS Available Qty **1–2 → 0** (default), or 1 → 0, or off.
  Qty 0 / missing → 0.
- **Export sets excluded:** JP, TH, TW, HK, CN, KR, MY, VN, US, IND, INDO, INDIA, PY, CH, UK, EU,
  UAE, AUS, GER, LL, ZA, PH, CA, MOR, FR, SA, SRI LANKA, plus FREEBIE rows.
- **Golden rule:** anything that is not a 100% match is **not** uploaded — it is listed for review.

---

## The registry worksheets

| Sheet | Meaning | You do |
|---|---|---|
| **Locked Matches** | confirmed Reebelo SKU ↔ Masterlist Stock Type ID links | nothing — these never drop |
| **New Masterlist SKUs** | POS SKUs with stock that aren't linked or parked | pick a decision in the dropdown |
| **Match Review** | Reebelo listings with no confirmed POS match | pick a decision in the dropdown |
| **Not Selling in Reebelo** | parked from a decision | nothing |
| **Not on Reebelo Yet** | parked from a decision | nothing |
| **Validation Errors** | alerts, warnings, skipped rows | read |

**Reviewer Decision** dropdown (both review sheets):

1. `Linked (fill col I)` → also type the ID/SKU in column I. Next run it moves into Locked Matches.
2. `Not Selling in Reebelo`
3. `Not on Reebelo yet`

Carry-over behaviour:

- A locked link **never drops**, even if the POS row disappears (that run just pushes 0).
- A parked SKU that gets POS stock again **and** has an exact Reebelo listing is **re-linked
  automatically** and removed from the parked sheets (logged as a warning so you can see it).
- A linked listing disappears from Match Review and from the parked sheets.

---

## Run it locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Create a GitHub repo, e.g. `mistermobile/reebelo-stock-sync`, and push these files:

   ```
   app.py
   reebelo_sync.py
   requirements.txt
   README.md
   .streamlit/config.toml
   ```

   ```bash
   git init
   git add .
   git commit -m "Reebelo Stock Sync Tool"
   git branch -M main
   git remote add origin https://github.com/<your-account>/reebelo-stock-sync.git
   git push -u origin main
   ```

2. Go to **share.streamlit.io** → **New app** → pick the repo, branch `main`, main file `app.py` → Deploy.
3. No secrets or API keys are needed — every file is uploaded by hand in the browser.

---

## Uploading to Cobalt

1. Cobalt → **Inventory → Update offers → Upload CSV**.
2. Upload the CSV exactly as downloaded. **Do not open or re-save it in Google Sheets** — that
   produces `line N: sku is a required field`.
3. Processing takes several minutes. Wait, then click **refresh**. Don't re-upload because the
   change isn't visible yet.
4. `Catalogue → Upload Inventory` is for publishing NEW offers only — never use it for stock.

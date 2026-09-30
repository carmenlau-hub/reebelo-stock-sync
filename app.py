"""
Reebelo Stock Sync Tool — Streamlit app
Mister Mobile Singapore

Run locally:   streamlit run app.py
Deploy:        push to GitHub -> share.streamlit.io -> point at app.py
"""

from datetime import date

import streamlit as st

import reebelo_sync as R

st.set_page_config(page_title="Reebelo Stock Sync Tool", page_icon="📦", layout="wide")

# ----------------------------------------------------------------------------------
# Styling — same black / yellow identity as the TikTok Stock Sync Tool
# ----------------------------------------------------------------------------------
st.markdown(
    """
    <style>
      .mm-banner {background:#FFEB00;padding:14px 18px;border-radius:6px 6px 0 0;}
      .mm-banner h1 {margin:0;font-size:26px;font-weight:800;color:#111;}
      .mm-sub {background:#111;color:#EEE;padding:8px 18px;border-radius:0 0 6px 6px;
               font-size:13px;margin-bottom:18px;}
      .mm-sub b {color:#FFEB00;}
      .mm-warn {background:#FFF8CC;border-left:6px solid #FFEB00;padding:12px 16px;
                border-radius:4px;font-size:14px;margin:6px 0 18px 0;}
      .mm-alert {background:#FDECEA;border-left:6px solid #B00020;padding:12px 16px;
                 border-radius:4px;font-size:14px;margin:6px 0 18px 0;}
      .mm-ok {background:#E9F7EC;border-left:6px solid #1E8E3E;padding:12px 16px;
              border-radius:4px;font-size:14px;margin:6px 0 18px 0;}
      div[data-testid="stMetricValue"] {font-size:24px;}
      .stDownloadButton button {width:100%;font-weight:700;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="mm-banner"><h1>📦 Reebelo Stock Sync Tool</h1></div>', unsafe_allow_html=True)
st.markdown(
    '<div class="mm-sub"><b>Mister Mobile Singapore</b> · POS Masterlist → Reebelo Cobalt '
    '"Update offers" bulk stock update</div>',
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------------
# Sidebar — 1 · Upload files
# ----------------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 1 · Upload files")
    pos_file = st.file_uploader("POS Masterlist (used device stock report)", type=["xlsx", "xlsm"])
    reb_files = st.file_uploader(
        "Reebelo Bulk Stock Inventory (export CSV)",
        type=["csv", "xlsx", "xlsm"],
        accept_multiple_files=True,
        help="You can drop more than one export here — e.g. the full export plus an "
             "accessories-only export. Duplicate SKUs: the newest file wins.",
    )
    acc_file = st.file_uploader(
        "POS Accessories stock report (optional)",
        type=["xlsx", "xlsm"],
        help="Leave empty and Brand New / MMACC accessory listings stay manual, exactly as before.",
    )
    reg_file = st.file_uploader("SKU Registry (5 worksheets)", type=["xlsx", "xlsm"])

    st.caption(
        "Required registry worksheets:\n\n"
        "· Locked Matches · New Masterlist SKUs · Match Review · Not Selling in Reebelo "
        "· Not on Reebelo Yet"
    )

    st.markdown("### 2 · Options")
    seed = st.checkbox(
        "Auto-lock 100% exact matches", value=True,
        help="Brand + model + storage + colour must all match exactly, and the match must be 1-to-1. "
             "Anything less goes to Match Review.",
    )
    relink = st.checkbox(
        "Re-link SKUs that come back into POS", value=True,
        help="A Masterlist SKU parked under Not Selling / Not on Reebelo Yet is moved back into "
             "Locked Matches when it has POS stock again and an exact Reebelo listing exists.",
    )
    changed_only = st.checkbox(
        "Upload CSV = changed rows only", value=True,
        help="Only listings whose stock actually changes are written to the CSV.",
    )

# ----------------------------------------------------------------------------------
# Gate
# ----------------------------------------------------------------------------------
if not pos_file or not reb_files:
    st.info("⬅️ Upload the **POS Masterlist** and the **Reebelo Bulk Stock Inventory** file to begin.")
    st.markdown(
        '<div class="mm-warn">🔒 <b>Locked rules.</b> Only <b>column J “stock (you)”</b> is ever '
        'changed. <b>price, minprice and market stay blank</b> in the upload CSV — the columns are '
        'kept, the values are empty — so Cobalt keeps your existing price and min price. '
        'Brand New / MMACC accessories are only synced when you also upload the accessories POS '
        'report; Grade A / Grade B cables and adapters are never auto-matched.</div>',
        unsafe_allow_html=True,
    )
    st.stop()

# ----------------------------------------------------------------------------------
# Load + validate
# ----------------------------------------------------------------------------------
blocking = []
pos_rows, pos_err = R.load_pos(pos_file)
blocking += pos_err
reb_rows, reb_df, reb_err = R.load_reebelo(reb_files)
blocking += reb_err
acc_rows, acc_err = R.load_pos_accessories(acc_file)
blocking += acc_err
registry, reg_err = R.load_registry(reg_file) if reg_file else (R.Registry(), [])
blocking += reg_err

if blocking:
    st.markdown(
        '<div class="mm-alert"><b>Export stopped.</b> A required file, worksheet or column is '
        "missing:<br>• " + "<br>• ".join(blocking) + "</div>",
        unsafe_allow_html=True,
    )
    st.stop()

if not registry.present:
    st.warning(
        "No SKU Registry uploaded — this run **creates** one. "
        "Download the Match Review workbook below, fill in every New Masterlist SKU decision, "
        "and upload it on the next run."
    )

if not acc_rows:
    st.info(
        "No accessories POS report uploaded — the Brand New / MMACC listings are skipped and "
        "stay manual. Upload `stock_report_..._accessories_new.xlsx` in the sidebar to include them."
    )

# ----------------------------------------------------------------------------------
# 3 · Oversell buffer
# ----------------------------------------------------------------------------------
st.markdown("#### 3 · Oversell buffer")
b1, b2 = st.columns(2)
with b1:
    dev_choice = st.radio(
        "Used devices",
        ["1–2 units → 0  (locked Reebelo rule)", "1 unit → 0", "No buffer (exact POS qty)"],
        index=0,
    )
with b2:
    acc_choice = st.radio(
        "Accessories",
        ["No buffer (exact POS qty)", "1–5 units → 0", "1–10 units → 0", "1–2 units → 0"],
        index=0,
        disabled=not acc_rows,
    )
buffer_max = {"1–2 units → 0  (locked Reebelo rule)": 2, "1 unit → 0": 1,
              "No buffer (exact POS qty)": 0}[dev_choice]
acc_buffer_max = {"No buffer (exact POS qty)": 0, "1–5 units → 0": 5,
                  "1–10 units → 0": 10, "1–2 units → 0": 2}[acc_choice]

opts = R.Options(
    buffer_max=buffer_max,
    acc_buffer_max=acc_buffer_max,
    relink_returning=relink,
    changed_rows_only=changed_only,
    seed_auto_lock=seed,
)
res = R.run_sync(pos_rows, reb_rows, registry, opts, acc_rows=acc_rows)
c = res.counts

# ----------------------------------------------------------------------------------
# 4 · Validation summary
# ----------------------------------------------------------------------------------
st.markdown(
    '<div class="mm-warn">💰 <b>Price is never touched.</b> The upload CSV keeps all five headers '
    '<code>sku,price,stock,minprice,market</code> with <b>price, minprice and market blank</b>. '
    'Stock only.</div>',
    unsafe_allow_html=True,
)

st.markdown("#### 4 · Validation summary")
r1 = st.columns(4)
r1[0].metric("Locked Matches updated", c["locked_updated"], help="Rows written to the upload CSV")
r1[1].metric("New Masterlist SKUs", c["new_masterlist"])
r1[2].metric("SKUs requiring review", c["review"])
r1[3].metric("Zeroed by buffer", c["locked_buffered"])

r2 = st.columns(4)
r2[0].metric("Locked — used devices", c["locked_devices"])
r2[1].metric("Locked — accessories", c["locked_accessories"])
r2[2].metric("Validation errors", c["errors"])
r2[3].metric("Unmatched with stock", c["unmatched_with_stock"])

r3 = st.columns(4)
r3[0].metric("Not Selling in Reebelo", c["not_selling"])
r3[1].metric("Not on Reebelo Yet", c["not_yet"])
r3[2].metric("Accessories on shared pools", c["shared_pools"],
             help=f"One POS pool feeding several listings. Full qty at {opts.share_threshold}+ "
                  "pieces, split evenly below that.")
r3[3].metric("Accessory listings to review", c["review_accessories"])

st.caption(
    f"Locked Matches total {c['locked_total']} · Reebelo used listings {c['reebelo_used_listings']} · "
    f"Reebelo accessory listings {c['reebelo_acc_listings']} · "
    f"POS used rows {c['pos_used_rows']} · POS accessory rows {c['pos_acc_rows']} · "
    f"POS rows excluded (export sets / freebies) {c['pos_excluded']} · Warnings {c['warnings']}"
)

if res.stock_no_match:
    st.markdown(
        f'<div class="mm-alert">⚠️ <b>{len(res.stock_no_match)} Reebelo listing(s) show stock but '
        "have no confirmed POS match.</b> They are left untouched (not in the CSV) — link them in "
        "Match Review or zero them manually in Cobalt.</div>",
        unsafe_allow_html=True,
    )

# ----------------------------------------------------------------------------------
# 5 · Download
#     The Reebelo bulk upload CSV unlocks only when EVERY New Masterlist SKU
#     has a Reviewer Decision. The Match Review workbook is always available.
# ----------------------------------------------------------------------------------
st.markdown("#### 5 · Download")

pending = [r for r in res.newml_rows if not str(r.get("Reviewer Decision", "")).strip()]
today = date.today()
csv_name = R.stock_update_filename(today)
xlsx_name = R.match_review_filename(today)
xlsx_bytes = R.build_registry_workbook(res, today, opts)

if pending:
    st.markdown(
        f'<div class="mm-alert">🔒 <b>Bulk upload CSV locked.</b> '
        f'<b>{len(pending)}</b> of {len(res.newml_rows)} New Masterlist SKUs still have no '
        "Reviewer Decision. Download the Match Review workbook, fill the Reviewer Decision column "
        "on the <b>New Masterlist SKUs</b> tab for every row (Linked / Not Selling in Reebelo / "
        "Not on Reebelo yet), then upload it again as the SKU Registry.</div>",
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        '<div class="mm-ok">✅ <b>All New Masterlist SKUs reviewed.</b> The Reebelo bulk upload '
        "CSV is unlocked.</div>",
        unsafe_allow_html=True,
    )

d1, d2 = st.columns(2)
if pending:
    d1.button(f"🔒 {csv_name} — locked ({len(pending)} to review)", disabled=True)
else:
    d1.download_button(
        f"⬇️ {csv_name}  ({len(res.upload_rows)} rows)",
        data=R.build_upload_csv(res), file_name=csv_name, mime="text/csv",
    )
d2.download_button(
    f"⬇️ {xlsx_name}", data=xlsx_bytes, file_name=xlsx_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption(
    "Cobalt → Inventory → **Update offers** → Upload CSV. Do not open or re-save the CSV in Google "
    "Sheets first. Processing can take several minutes — wait, then hit refresh."
)

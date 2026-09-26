"""
Reebelo Stock Sync Tool — Streamlit app
Mister Mobile Singapore

Run locally:   streamlit run app.py
Deploy:        push to GitHub -> share.streamlit.io -> point at app.py
"""

from datetime import date

import pandas as pd
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
    reb_file = st.file_uploader("Reebelo Bulk Stock Inventory (export CSV)", type=["csv", "xlsx", "xlsm"])
    reg_file = st.file_uploader("SKU Registry (5 worksheets)", type=["xlsx", "xlsm"])

    st.caption(
        "Required registry worksheets:\n\n"
        "· Locked Matches · New Masterlist SKUs · Match Review · Not Selling in Reebelo "
        "· Not on Reebelo Yet"
    )

    st.markdown("### 2 · Options")
    seed = st.checkbox(
        "Auto-lock 100% exact matches",
        value=True,
        help="Brand + model + storage + colour must all match exactly, and the match must be 1-to-1. "
             "Anything less goes to Match Review.",
    )
    relink = st.checkbox(
        "Re-link SKUs that come back into POS",
        value=True,
        help="A Masterlist SKU parked under Not Selling / Not on Reebelo Yet is moved back into "
             "Locked Matches when it has POS stock again and an exact Reebelo listing exists.",
    )
    changed_only = st.checkbox(
        "Upload CSV = changed rows only",
        value=True,
        help="Only listings whose stock actually changes are written to the CSV.",
    )

# ----------------------------------------------------------------------------------
# Gate
# ----------------------------------------------------------------------------------
if not pos_file or not reb_file:
    st.info("⬅️ Upload the **POS Masterlist** and the **Reebelo Bulk Stock Inventory** file to begin.")
    st.markdown(
        '<div class="mm-warn">🔒 <b>Locked rules.</b> Only <b>column J “stock (you)”</b> is ever '
        'changed. <b>price, minprice and market stay blank</b> in the upload CSV — the columns are '
        'kept, the values are empty — so Cobalt keeps your existing price and min price. Used sets '
        'only: Brand New / MMACC accessory listings are skipped and stay manual.</div>',
        unsafe_allow_html=True,
    )
    st.stop()

# ----------------------------------------------------------------------------------
# Load + validate
# ----------------------------------------------------------------------------------
blocking = []
pos_rows, pos_err = R.load_pos(pos_file)
blocking += pos_err
reb_rows, reb_df, reb_err = R.load_reebelo(reb_file, getattr(reb_file, "name", ""))
blocking += reb_err
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
        "Download the Match Review workbook at the bottom, review it, and upload it on the next run."
    )

# ----------------------------------------------------------------------------------
# Oversell buffer
# ----------------------------------------------------------------------------------
st.markdown("#### 3 · Oversell buffer")
buf_choice = st.radio(
    "Zero out low POS stock so the last unit is never oversold:",
    ["1–2 units → 0  (locked Reebelo rule)", "1 unit → 0", "No buffer (use exact POS qty)"],
    index=0,
    horizontal=True,
)
buffer_max = {"1–2 units → 0  (locked Reebelo rule)": 2, "1 unit → 0": 1,
              "No buffer (use exact POS qty)": 0}[buf_choice]

opts = R.Options(
    buffer_max=buffer_max,
    relink_returning=relink,
    changed_rows_only=changed_only,
    seed_auto_lock=seed,
)
res = R.run_sync(pos_rows, reb_rows, registry, opts)
c = res.counts

# ----------------------------------------------------------------------------------
# Validation summary
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
r2[0].metric("Not Selling in Reebelo", c["not_selling"])
r2[1].metric("Not on Reebelo Yet", c["not_yet"])
r2[2].metric("Validation errors", c["errors"])
r2[3].metric("Unmatched with stock", c["unmatched_with_stock"])

st.caption(
    f"Locked Matches total {c['locked_total']} · Reebelo used listings {c['reebelo_used_listings']} · "
    f"Brand New / accessory listings skipped {c['brand_new_skipped']} · "
    f"POS used rows {c['pos_used_rows']} · POS rows excluded (export sets / freebies) {c['pos_excluded']} · "
    f"Warnings {c['warnings']}"
)

if res.stock_no_match:
    st.markdown(
        f'<div class="mm-alert">⚠️ <b>{len(res.stock_no_match)} Reebelo listing(s) show stock but '
        "have no confirmed POS match.</b> They are left untouched (not in the CSV) — link them in "
        "Match Review or zero them manually in Cobalt.</div>",
        unsafe_allow_html=True,
    )
    st.dataframe(pd.DataFrame(res.stock_no_match), use_container_width=True, hide_index=True)

# ----------------------------------------------------------------------------------
# Tabs
# ----------------------------------------------------------------------------------
tabs = st.tabs(
    ["Upload CSV preview", R.SHEET_LOCKED, R.SHEET_REVIEW, R.SHEET_NEW_ML,
     R.SHEET_NOT_SELLING, R.SHEET_NOT_YET, R.SHEET_ERRORS]
)
frames = [
    pd.DataFrame(res.upload_rows, columns=R.UPLOAD_HEADERS),
    pd.DataFrame(res.locked_rows),
    pd.DataFrame(res.review_rows),
    pd.DataFrame(res.newml_rows),
    pd.DataFrame(res.not_selling_rows),
    pd.DataFrame(res.not_yet_rows),
    pd.DataFrame(res.error_rows),
]
for tab, frame in zip(tabs, frames):
    with tab:
        if frame.empty:
            st.caption("Nothing here for this run.")
        else:
            st.dataframe(frame.head(3000), use_container_width=True, hide_index=True)
            if len(frame) > 3000:
                st.caption(f"Showing the first 3,000 of {len(frame):,} rows — the file has them all.")

# ----------------------------------------------------------------------------------
# Downloads
# ----------------------------------------------------------------------------------
st.markdown("#### 5 · Download")
today = date.today()
csv_bytes = R.build_upload_csv(res)
xlsx_bytes = R.build_registry_workbook(res, today, opts)
csv_name = R.stock_update_filename(today)
xlsx_name = R.match_review_filename(today)

d1, d2 = st.columns(2)
d1.download_button(
    f"⬇️ {csv_name}  ({len(res.upload_rows)} rows)",
    data=csv_bytes,
    file_name=csv_name,
    mime="text/csv",
)
d2.download_button(
    f"⬇️ {xlsx_name}",
    data=xlsx_bytes,
    file_name=xlsx_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.caption(
    "Cobalt → Inventory → **Update offers** → Upload CSV. Do not open or re-save the CSV in Google "
    "Sheets first. Processing can take several minutes — wait, then hit refresh."
)

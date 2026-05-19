import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import io

st.set_page_config(
    page_title="Fleet Summary Analyser",
    page_icon="🚢",
    layout="wide",
)

st.title("🚢 Fleet Summary Analyser")
st.markdown("Upload your Fleet Summary Excel file to analyse LO report status across vessels.")

# ── File upload ───────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Upload Fleet Summary File",
    type=["xlsx", "xls", "csv"],
    help="Supports Excel (.xlsx / .xls) and CSV files",
)

if not uploaded:
    st.info("👆 Upload a Fleet Summary file to get started.")

    # ── Legend shown before upload ──────────────────────────────────────────
    st.markdown("---")
    st.markdown("### Status Logic & Colour Guide")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.success("🟢 **Up-to-date**\nLatest Report Date is within the last **89 days**.\nRow highlighted **green**.")
    with col2:
        st.error("🔴 **Overdue**\nLatest Report Date is **more than 89 days** ago.\nRow highlighted **red**.")
    with col3:
        st.warning("⚪ **No Date**\nNo Latest Report Date recorded.\nRow left **un-highlighted**.")

    st.markdown("#### Overdue Severity Buckets")
    b1, b2, b3 = st.columns(3)
    b1.info("🟡 **90–180 days** — Attention needed")
    b2.warning("🟠 **181–210 days** — Escalate")
    b3.error("🔴 **210+ days** — Critical")

    st.markdown("#### Suggested Extra Analysis (available after upload)")
    st.markdown("""
| Analysis | What it shows |
|---|---|
| **Fleet Health Score** | % of vessels with reports submitted within 89 days |
| **Overdue Severity Breakdown** | Buckets: 90–180 / 181–210 / 210+ days overdue |
| **VesselCheck Compliance** | How many vessels have a valid VesselCheck date |
| **Reporting Activity by Month** | Count of reports submitted per calendar month |
| **Vessels Never Reported** | Vessels with no date at all |
| **Remarks Frequency** | Most common words/phrases in the Remarks column |
""")
    st.stop()

# ── Parse ─────────────────────────────────────────────────────────────────────
TARGET_SHEET = "Fleet Summary"

@st.cache_data(show_spinner="Reading file…")
def load_file(file_bytes, filename):
    if filename.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(file_bytes))
        df.columns = df.columns.str.strip()
        return df, None
    xf = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = xf.sheet_names
    sheet = next((s for s in sheets if s.strip().lower() == TARGET_SHEET.lower()), None)
    if sheet is None:
        return None, sheets
    df = xf.parse(sheet)
    df.columns = df.columns.str.strip()
    df = df.dropna(how="all")
    return df, sheet

try:
    raw_bytes = uploaded.read()
    df, sheet_info = load_file(raw_bytes, uploaded.name)
except Exception as e:
    st.error(f"Could not read file: {e}")
    st.stop()

if df is None:
    available_sheets = sheet_info
    chosen = st.selectbox(
        f'Sheet **"{TARGET_SHEET}"** not found. Select the correct sheet:',
        available_sheets,
    )
    @st.cache_data(show_spinner="Reading sheet…")
    def load_sheet(file_bytes, sheet):
        xf = pd.ExcelFile(io.BytesIO(file_bytes))
        df = xf.parse(sheet)
        df.columns = df.columns.str.strip()
        df = df.dropna(how="all")
        return df
    df = load_sheet(raw_bytes, chosen)
else:
    if isinstance(sheet_info, str):
        st.caption(f"Sheet loaded: **{sheet_info}**")

df = df.dropna(how="all").reset_index(drop=True)

# ── Column mapping ────────────────────────────────────────────────────────────
def find_col(df, exact_hints, partial_hints=None):
    lower = {c.strip().lower(): c for c in df.columns}
    for h in exact_hints:
        if h.lower() in lower:
            return lower[h.lower()]
    if partial_hints:
        for h in partial_hints:
            for k, v in lower.items():
                if h.lower() in k:
                    return v
    return None

col_vessel       = find_col(df, ["vessel"],                      partial_hints=["vessel"])
col_report_date  = find_col(df, ["latest report date", "report date", "last report date"],
                                partial_hints=["latest report", "report date", "last report", "date"])
col_remarks      = find_col(df, ["remarks", "remark"],           partial_hints=["remark"])
col_vessel_check = find_col(df, ["vesselcheck", "vessel check"], partial_hints=["vesselcheck", "vessel check", "check"])

missing = [name for name, c in [
    ("Vessel", col_vessel), ("Latest Report Date", col_report_date),
    ("Remarks", col_remarks), ("VesselCheck", col_vessel_check),
] if c is None]

if missing:
    st.warning(f"Could not auto-detect columns: **{', '.join(missing)}**. Please map them below.")
    all_cols = list(df.columns)
    with st.expander("Map columns manually", expanded=True):
        c1, c2, c3, c4 = st.columns(4)
        col_vessel       = c1.selectbox("Vessel column",             all_cols, index=0)
        col_report_date  = c2.selectbox("Latest Report Date column", all_cols, index=min(1, len(all_cols)-1))
        col_remarks      = c3.selectbox("Remarks column",            all_cols, index=min(2, len(all_cols)-1))
        col_vessel_check = c4.selectbox("VesselCheck column",        all_cols, index=min(3, len(all_cols)-1))

# ── Normalise ─────────────────────────────────────────────────────────────────
df = pd.DataFrame({
    "Vessel":             df[col_vessel],
    "Latest Report Date": df[col_report_date],
    "Remarks":            df[col_remarks],
    "VesselCheck":        df[col_vessel_check],
})

report_date_dt  = pd.to_datetime(df["Latest Report Date"], errors="coerce").dt.normalize()
vc_date_dt      = pd.to_datetime(df["VesselCheck"],        errors="coerce").dt.normalize()
today           = pd.Timestamp(date.today())

df["Days Since Report"] = (today - report_date_dt).dt.days.astype("Int64")

def classify_status(d):
    if pd.isna(d):
        return "No Date"
    return "Up-to-date" if d <= 89 else "Overdue"

df["Status"] = df["Days Since Report"].apply(classify_status)

def overdue_bucket(d):
    if pd.isna(d) or d <= 89:
        return None
    if d <= 180:
        return "90–180 days"
    if d <= 210:
        return "181–210 days"
    return "210+ days"

df["Overdue Bucket"] = df["Days Since Report"].apply(overdue_bucket)

# Format dates as plain strings
df["Latest Report Date"] = report_date_dt.dt.strftime("%Y-%m-%d").fillna("")
df["VesselCheck"]        = vc_date_dt.dt.strftime("%Y-%m-%d").fillna(
    df["VesselCheck"].astype(str).replace("nan", ""))

# ── Derived metrics ───────────────────────────────────────────────────────────
total      = len(df)
n_uptodate = (df["Status"] == "Up-to-date").sum()
n_overdue  = (df["Status"] == "Overdue").sum()
n_nodate   = (df["Status"] == "No Date").sum()
health_pct = round(n_uptodate / total * 100, 1) if total else 0

n_31_60 = (df["Overdue Bucket"] == "90–180 days").sum()
n_61_90 = (df["Overdue Bucket"] == "181–210 days").sum()
n_90p   = (df["Overdue Bucket"] == "210+ days").sum()

vc_has_date  = (df["VesselCheck"].str.match(r"\d{4}-\d{2}-\d{2}")).sum()
vc_no_date   = total - vc_has_date

# ── Sidebar filters ───────────────────────────────────────────────────────────
st.sidebar.header("Filters")
status_opts = df["Status"].dropna().unique().tolist()
sel_status  = st.sidebar.multiselect("Status", status_opts, default=status_opts)
vc_vals     = df["VesselCheck"].dropna().unique().tolist()
sel_vc      = st.sidebar.multiselect("VesselCheck", vc_vals, default=vc_vals)
search      = st.sidebar.text_input("Search vessel", "")

mask = df["Status"].isin(sel_status) & df["VesselCheck"].isin(sel_vc)
if search:
    mask &= df["Vessel"].astype(str).str.contains(search, case=False, na=False)
filtered = df[mask].copy()

st.sidebar.divider()
st.sidebar.markdown("**Export**")
display_cols = ["Vessel", "Latest Report Date", "Days Since Report",
                "Status", "Overdue Bucket", "VesselCheck", "Remarks"]
display_cols = [c for c in display_cols if c in filtered.columns]
csv_buf = filtered[display_cols].to_csv(index=False).encode()
st.sidebar.download_button("Download filtered CSV", csv_buf,
                           file_name="fleet_summary_filtered.csv", mime="text/csv")

# ── Legend ────────────────────────────────────────────────────────────────────
with st.expander("📖 Status Logic & Colour Guide", expanded=False):
    lg1, lg2, lg3 = st.columns(3)
    with lg1:
        st.success("🟢 **Up-to-date**\nReport within **≤ 89 days**.\nRow highlighted **green**.")
    with lg2:
        st.error("🔴 **Overdue**\nReport **> 89 days** ago.\nRow highlighted **red**.")
    with lg3:
        st.warning("⚪ **No Date**\nNo report date recorded.\nRow un-highlighted.")

    st.markdown("**Overdue Severity Buckets:**")
    b1, b2, b3 = st.columns(3)
    b1.info("🟡 **90–180 days** — Attention needed")
    b2.warning("🟠 **181–210 days** — Escalate")
    b3.error("🔴 **210+ days** — Critical")

st.divider()

# ── KPI cards ─────────────────────────────────────────────────────────────────
st.markdown("### Fleet Overview")
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Total Vessels",    total)
k2.metric("Fleet Health",     f"{health_pct}%", help="% vessels with report ≤ 89 days")
k3.metric("Up-to-date",       n_uptodate)
k4.metric("Overdue",          n_overdue)
k5.metric("No Report Date",   n_nodate)

st.markdown("#### Overdue Severity")
s1, s2, s3 = st.columns(3)
s1.metric("90–180 days (Attention)", n_31_60)
s2.metric("181–210 days (Escalate)", n_61_90)
s3.metric("210+ days (Critical)",    n_90p)

st.divider()

# ── Charts row 1 ──────────────────────────────────────────────────────────────
st.markdown("### Analytics")
ch1, ch2 = st.columns(2)

with ch1:
    color_map = {"Up-to-date": "#2ecc71", "Overdue": "#e74c3c", "No Date": "#95a5a6"}
    status_counts = df["Status"].value_counts().reset_index()
    status_counts.columns = ["Status", "Count"]
    fig_pie = px.pie(
        status_counts, names="Status", values="Count",
        title="Report Status Distribution",
        color="Status", color_discrete_map=color_map, hole=0.4,
    )
    fig_pie.update_traces(textinfo="percent+label")
    st.plotly_chart(fig_pie, use_container_width=True)

with ch2:
    bucket_order = ["90–180 days", "181–210 days", "210+ days"]
    bucket_colors = {"90–180 days": "#f39c12", "181–210 days": "#e67e22", "210+ days": "#e74c3c"}
    overdue_buckets = (
        df[df["Overdue Bucket"].notna()]["Overdue Bucket"]
        .value_counts()
        .reindex(bucket_order, fill_value=0)
        .reset_index()
    )
    overdue_buckets.columns = ["Bucket", "Count"]
    fig_bucket = px.bar(
        overdue_buckets, x="Bucket", y="Count",
        title="Overdue Severity Breakdown",
        color="Bucket", color_discrete_map=bucket_colors,
        text_auto=True,
    )
    fig_bucket.update_layout(showlegend=False, xaxis_title="", yaxis_title="Vessels")
    st.plotly_chart(fig_bucket, use_container_width=True)

# ── Charts row 2 ──────────────────────────────────────────────────────────────
ch3, ch4 = st.columns(2)

with ch3:
    # Reporting activity by month
    rpt_dates = pd.to_datetime(df["Latest Report Date"], errors="coerce")
    monthly = (
        rpt_dates.dropna()
        .dt.to_period("M")
        .value_counts()
        .sort_index()
        .reset_index()
    )
    monthly.columns = ["Month", "Count"]
    monthly["Month"] = monthly["Month"].astype(str)
    fig_monthly = px.bar(
        monthly, x="Month", y="Count",
        title="Reporting Activity by Month",
        text_auto=True,
        color_discrete_sequence=["#3498db"],
    )
    fig_monthly.update_layout(xaxis_tickangle=-35, yaxis_title="Reports Submitted")
    st.plotly_chart(fig_monthly, use_container_width=True)

with ch4:
    # VesselCheck compliance
    vc_data = pd.DataFrame({
        "Category": ["Has VesselCheck Date", "No VesselCheck Date"],
        "Count": [vc_has_date, vc_no_date],
    })
    fig_vc = px.pie(
        vc_data, names="Category", values="Count",
        title="VesselCheck Date Compliance",
        color="Category",
        color_discrete_map={"Has VesselCheck Date": "#2ecc71", "No VesselCheck Date": "#e74c3c"},
        hole=0.4,
    )
    fig_vc.update_traces(textinfo="percent+label")
    st.plotly_chart(fig_vc, use_container_width=True)

# ── Top overdue bar ───────────────────────────────────────────────────────────
overdue_df = df[df["Status"] == "Overdue"].sort_values("Days Since Report", ascending=False)
if not overdue_df.empty:
    st.markdown("### Top Overdue Vessels")
    fig_over = px.bar(
        overdue_df.head(20),
        x="Vessel", y="Days Since Report",
        color="Overdue Bucket",
        color_discrete_map={"90–180 days": "#f39c12", "181–210 days": "#e67e22", "210+ days": "#e74c3c"},
        text_auto=True,
        title="Top 20 Overdue Vessels (coloured by severity)",
    )
    fig_over.update_layout(xaxis_tickangle=-35, yaxis_title="Days Since Last Report",
                           legend_title="Severity")
    st.plotly_chart(fig_over, use_container_width=True)

st.divider()

# ── Data table ────────────────────────────────────────────────────────────────
st.markdown(f"### Vessel Table ({len(filtered)} of {total})")

def highlight_status(row):
    s = row.get("Status", "")
    if s == "Overdue":
        return ["background-color: #fdecea"] * len(row)
    if s == "Up-to-date":
        return ["background-color: #eafaf1"] * len(row)
    return [""] * len(row)

styled = filtered[display_cols].style.apply(highlight_status, axis=1)
st.dataframe(styled, use_container_width=True, height=450)

# ── Vessels never reported ────────────────────────────────────────────────────
with st.expander("🚫 Vessels with No Report Date"):
    no_date_df = df[df["Status"] == "No Date"][["Vessel", "VesselCheck", "Remarks"]]
    if no_date_df.empty:
        st.success("All vessels have a report date recorded.")
    else:
        st.warning(f"{len(no_date_df)} vessel(s) have no Latest Report Date.")
        st.dataframe(no_date_df, use_container_width=True)

# ── Remarks summary ───────────────────────────────────────────────────────────
with st.expander("📝 Remarks Summary"):
    remarks_df = filtered[["Vessel", "Status", "Remarks"]].dropna(subset=["Remarks"])
    remarks_df = remarks_df[remarks_df["Remarks"].astype(str).str.strip().ne("")]
    if remarks_df.empty:
        st.write("No remarks to display for current filter.")
    else:
        st.dataframe(remarks_df, use_container_width=True)

        # Top recurring words in remarks
        all_remarks = " ".join(remarks_df["Remarks"].astype(str).str.lower())
        stop = {"the","a","an","and","or","of","to","in","is","it","for","on",
                "with","at","by","from","as","was","are","has","have","been",
                "this","that","not","no","n/a","nan","be","will"}
        words = [w.strip(".,;:()") for w in all_remarks.split() if len(w) > 3 and w not in stop]
        if words:
            word_freq = pd.Series(words).value_counts().head(15).reset_index()
            word_freq.columns = ["Word", "Frequency"]
            fig_words = px.bar(
                word_freq, x="Frequency", y="Word", orientation="h",
                title="Top 15 Words in Remarks",
                color_discrete_sequence=["#8e44ad"],
            )
            fig_words.update_layout(yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig_words, use_container_width=True)

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
        st.success("🟢 **Up-to-date**\nLatest Report Date is within the last **120 days**.\nRow highlighted **green**.")
    with col2:
        st.error("🔴 **Overdue**\nLatest Report Date is **more than 120 days** ago.\nRow highlighted **red**.")
    with col3:
        st.warning("⚪ **No Date**\nNo Latest Report Date recorded.\nRow left **un-highlighted**.")

    st.markdown("#### Overdue Severity Buckets")
    b1, b2, b3 = st.columns(3)
    b1.info("🟡 **120–180 days** — Attention needed")
    b2.warning("🟠 **181–210 days** — Escalate")
    b3.error("🔴 **210+ days** — Critical")

    st.markdown("#### Suggested Extra Analysis (available after upload)")
    st.markdown("""
| Analysis | What it shows |
|---|---|
| **Fleet Health Score** | % of vessels with reports submitted within 120 days |
| **Overdue Severity Breakdown** | Buckets: 120–180 / 181–210 / 210+ days overdue |
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
    df = xf.parse(sheet, keep_default_na=False, na_values=[""])
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

# Drop ghost rows: no vessel name (NaN, blank, or whitespace-only)
df = df[df["Vessel"].notna() & (df["Vessel"].astype(str).str.strip() != "")].reset_index(drop=True)

report_date_dt  = pd.to_datetime(df["Latest Report Date"], errors="coerce").dt.normalize()
vc_date_dt      = pd.to_datetime(df["VesselCheck"],        errors="coerce").dt.normalize()
today           = pd.Timestamp(date.today())

df["Days Since Report"] = (today - report_date_dt).dt.days.astype("Int64")

def classify_status(d):
    if pd.isna(d):
        return "No Date"
    return "Up-to-date" if d <= 120 else "Overdue"

df["Status"] = df["Days Since Report"].apply(classify_status)

def overdue_bucket(d):
    if pd.isna(d) or d <= 120:
        return None
    if d <= 180:
        return "120–180 days"
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

n_31_60 = (df["Overdue Bucket"] == "120–180 days").sum()
n_61_90 = (df["Overdue Bucket"] == "181–210 days").sum()
n_90p   = (df["Overdue Bucket"] == "210+ days").sum()

vc_has_date  = (df["VesselCheck"].str.match(r"\d{4}-\d{2}-\d{2}")).sum()
vc_no_date   = total - vc_has_date

# ── Normalise Remarks → clean strings, then classify ─────────────────────────
# "Blank" = empty/NaN cells   "NA" = literal "NA"/"N/A"/dash/null text
_NA_TEXT_VALUES = {
    "na", "n/a", "n.a.", "n.a", "nan", "none", "null",
    "-", "–", "—", ".", "0", "nil", "not applicable", "not available",
}
# 1. Strip whitespace; NaN → "" (empty string)
_remarks_clean  = df["Remarks"].fillna("").astype(str).str.strip()
# 2. Classify
is_blank_mask   = (_remarks_clean == "")
is_na_text_mask = _remarks_clean.str.lower().isin(_NA_TEXT_VALUES) & ~is_blank_mask
is_na_mask      = is_blank_mask | is_na_text_mask   # combined for fleet-health
# 3. Write display values: blank→"Blank", NA text→"NA", everything else unchanged
df["Remarks"] = _remarks_clean.copy()
df.loc[is_blank_mask,   "Remarks"] = "Blank"
df.loc[is_na_text_mask, "Remarks"] = "NA"

# 4. Jibe mask — operates on clean column (Blank/NA rows never contain jibe text)
is_jibe_mask = df["Remarks"].str.contains(
    r"hasn.t been rolled out|not rolled out|jibe", case=False, na=False, regex=True
)

n_remarks_blank = int(is_blank_mask.sum())
n_remarks_na    = int(is_na_text_mask.sum())
n_not_jibe      = int(is_jibe_mask.sum())
n_other         = int((~is_na_mask & ~is_jibe_mask).sum())
# Sanity: n_remarks_blank + n_remarks_na + n_not_jibe + n_other == total

# Fleet Health — denominator excludes Jibe and Blank/NA-remarks vessels
active_mask       = ~is_jibe_mask & ~is_na_mask
active_fleet      = int(active_mask.sum())
n_uptodate_active = int(((df["Status"] == "Up-to-date") & active_mask).sum())
health_pct        = round(n_uptodate_active / active_fleet * 100, 1) if active_fleet else 0

# Unique Remarks values for sidebar filter — exclude Blank/NA (handled by checkbox)
remarks_opts = sorted(df.loc[~is_na_mask, "Remarks"].unique().tolist())

# VesselCheck date series and month-year labels for filter/chart
vc_date_dt_series = pd.to_datetime(df["VesselCheck"], errors="coerce")
vc_month_series   = vc_date_dt_series.dt.to_period("M")
# Format as "Jan 2026" for display
vc_month_labels_sorted = sorted(vc_month_series.dropna().unique().tolist(), key=lambda p: p.ordinal)
vc_month_display  = [p.strftime("%b %Y") for p in vc_month_labels_sorted]
# Map "Jan 2026" -> Period for reverse lookup
vc_display_to_period = {p.strftime("%b %Y"): p for p in vc_month_labels_sorted}

# ── Sidebar filters ───────────────────────────────────────────────────────────
st.sidebar.header("Filters")
status_opts = df["Status"].dropna().unique().tolist()
sel_status  = st.sidebar.multiselect("Status", status_opts, default=status_opts)
vessel_opts = sorted(df["Vessel"].dropna().astype(str).unique().tolist())
sel_vessels = st.sidebar.multiselect("Vessel", options=vessel_opts, default=[],
                                     help="Leave blank to include all vessels")

st.sidebar.markdown("**Remarks**")
include_blank_remarks = st.sidebar.checkbox("Include Blank remarks", value=True)
include_na_remarks    = st.sidebar.checkbox("Include NA remarks", value=True)
sel_remarks = st.sidebar.multiselect(
    "Filter by specific Remarks",
    options=remarks_opts,
    default=[],
    help="Leave blank to include all; select specific remarks to filter to those only",
)

st.sidebar.markdown("**VesselCheck Date Range**")
vc_valid_dates = vc_date_dt_series.dropna()
if not vc_valid_dates.empty:
    vc_min_date = vc_valid_dates.min().date()
    vc_max_date = vc_valid_dates.max().date()
    import datetime as _dt
    vc_date_from = st.sidebar.date_input(
        "From", value=vc_min_date,
        min_value=vc_min_date, max_value=vc_max_date, key="vc_from"
    )
    vc_date_to = st.sidebar.date_input(
        "To", value=vc_max_date,
        min_value=vc_min_date, max_value=vc_max_date, key="vc_to"
    )
    sel_vc_months = [
        lbl for lbl, p in vc_display_to_period.items()
        if _dt.date(p.year, p.month, 1) >= _dt.date(vc_date_from.year, vc_date_from.month, 1)
        and _dt.date(p.year, p.month, 1) <= _dt.date(vc_date_to.year, vc_date_to.month, 1)
    ]
else:
    vc_date_from = vc_date_to = None
    sel_vc_months = []

mask = df["Status"].isin(sel_status)
if sel_vessels:
    mask &= df["Vessel"].astype(str).isin(sel_vessels)
if sel_remarks or not include_blank_remarks or not include_na_remarks:
    if sel_remarks:
        remark_match = df["Remarks"].isin(sel_remarks)
        if include_blank_remarks:
            remark_match |= (df["Remarks"] == "Blank")
        if include_na_remarks:
            remark_match |= (df["Remarks"] == "NA")
        mask &= remark_match
    else:
        if not include_blank_remarks:
            mask &= (df["Remarks"] != "Blank")
        if not include_na_remarks:
            mask &= (df["Remarks"] != "NA")
if sel_vc_months:
    sel_periods = set(vc_display_to_period[lbl] for lbl in sel_vc_months if lbl in vc_display_to_period)
    vc_month_mask = vc_month_series.isin(sel_periods) | vc_month_series.isna()
    mask &= vc_month_mask
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
        st.success("🟢 **Up-to-date**\nReport within **≤ 120 days**.\nRow highlighted **green**.")
    with lg2:
        st.error("🔴 **Overdue**\nReport **> 120 days** ago.\nRow highlighted **red**.")
    with lg3:
        st.warning("⚪ **No Date**\nNo report date recorded.\nRow un-highlighted.")

    st.markdown("**Overdue Severity Buckets:**")
    b1, b2, b3 = st.columns(3)
    b1.info("🟡 **120–180 days** — Attention needed")
    b2.warning("🟠 **181–210 days** — Escalate")
    b3.error("🔴 **210+ days** — Critical")

st.divider()

# ── KPI cards ─────────────────────────────────────────────────────────────────
st.markdown("### Fleet Overview")

# Row 1 — Vessel count breakdown
r1c1, r1c2, r1c3, r1c4, r1c5, r1c6 = st.columns(6)
r1c1.metric("Total Vessels",                       total)
r1c2.metric("Hasn't been rolled out with Jibe",    n_not_jibe,
            help="Vessels whose Remarks contain 'hasn't been rolled out', 'not rolled out', or 'jibe'")
r1c3.metric("Remarks — NA",                        n_remarks_na,
            help="Vessels with literal 'NA', 'N/A', 'none', '-' etc. in the Remarks cell")
r1c4.metric("Remarks — Blank",                     n_remarks_blank,
            help="Vessels where the Remarks cell is empty")
r1c5.metric("Fleet Health",                        f"{health_pct}%",
            help="% active vessels with report ≤ 120 days")
r1c6.metric("No Report Date",                      n_nodate)

st.info(
    f"**Fleet Health formula:** "
    f"{n_uptodate_active} up-to-date (active) ÷ {active_fleet} active vessels × 100 = **{health_pct}%**  \n"
    f"Active fleet = {total} total − {n_not_jibe} Jibe − {n_remarks_na} NA − {n_remarks_blank} Blank = **{active_fleet}**  \n"
    f"ℹ️ The *Up-to-date* KPI above shows **{n_uptodate}** (all vessels reported within 120 days, including Jibe / NA / Blank). "
    f"Fleet Health uses only the **{n_uptodate_active}** within the active fleet — difference of {n_uptodate - n_uptodate_active}."
)

# Row 2 — Report status
r2c1, r2c2, r2c3 = st.columns(3)
r2c1.metric("Up-to-date",  n_uptodate)
r2c2.metric("Overdue",     n_overdue)
r2c3.metric("No Date",     n_nodate)

with st.expander("🔍 Remarks Breakdown — all remarks with counts (total = vessel count)"):
    # Full frequency table: every unique Remark value → count, sorted desc, with Total row
    remarks_freq = (
        df["Remarks"]
        .value_counts()
        .reset_index()
    )
    remarks_freq.columns = ["Remark", "Count"]
    total_row = pd.DataFrame([{"Remark": "TOTAL", "Count": remarks_freq["Count"].sum()}])
    remarks_table = pd.concat([remarks_freq, total_row], ignore_index=True)
    st.dataframe(
        remarks_table.style.apply(
            lambda row: ["font-weight: bold; background-color: #f0f0f0"] * len(row)
            if row["Remark"] == "TOTAL" else [""] * len(row),
            axis=1,
        ),
        use_container_width=True,
        height=min(400, 35 * (len(remarks_table) + 1) + 40),
    )
    st.caption(
        "**Blank** = empty cell in the sheet &nbsp;|&nbsp; "
        "**NA** = cell contains 'NA', 'N/A', 'none', '-' or similar &nbsp;|&nbsp; "
        "All other values are shown as-is."
    )

st.markdown("#### Overdue Severity")
s1, s2, s3 = st.columns(3)
s1.metric("120–180 days (Attention)", n_31_60)
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
    bucket_order = ["120–180 days", "181–210 days", "210+ days"]
    bucket_colors = {"120–180 days": "#f39c12", "181–210 days": "#e67e22", "210+ days": "#e74c3c"}
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
    # VesselCheck compliance pie
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

# ── VesselCheck by Month ───────────────────────────────────────────────────────
st.markdown("### VesselCheck by Month")

# Build month series using only the selected period range
if sel_vc_months:
    sel_periods_chart = set(vc_display_to_period[lbl] for lbl in sel_vc_months if lbl in vc_display_to_period)
    vc_monthly_filtered = vc_month_series[vc_month_series.isin(sel_periods_chart)]
else:
    vc_monthly_filtered = vc_month_series.dropna()

# Aggregate: count vessels per month, format label as "Mon YYYY"
vc_by_month = (
    vc_monthly_filtered
    .value_counts()
    .sort_index()
    .reset_index()
)
vc_by_month.columns = ["Period", "Vessels"]
vc_by_month["Month"] = vc_by_month["Period"].apply(lambda p: p.strftime("%b %Y"))

if vc_by_month.empty:
    st.info("No VesselCheck dates available for the selected date range.")
else:
    fig_vc_month = px.bar(
        vc_by_month, x="Month", y="Vessels",
        title="Number of Vessels with VesselCheck by Month",
        text_auto=True,
        color_discrete_sequence=["#1abc9c"],
        category_orders={"Month": vc_by_month["Month"].tolist()},
    )
    fig_vc_month.update_layout(
        xaxis_type="category",
        xaxis_tickangle=-35,
        yaxis_title="Vessel Count",
        xaxis_title="VesselCheck Month",
    )
    st.plotly_chart(fig_vc_month, use_container_width=True)

    with st.expander("📅 VesselCheck Month Detail Table"):
        if sel_vc_months:
            vc_detail_mask = vc_month_series.isin(sel_periods_chart)
        else:
            vc_detail_mask = vc_month_series.notna()
        vc_detail = df[vc_detail_mask][
            ["Vessel", "VesselCheck", "Status", "Remarks"]
        ].sort_values("VesselCheck")
        st.dataframe(vc_detail, use_container_width=True)

# ── Vessel report status bar (overdue + up-to-date, respects filters) ─────────
overdue_df   = filtered[filtered["Status"] == "Overdue"].copy()
uptodate_df  = filtered[filtered["Status"] == "Up-to-date"].copy()

overdue_df["Bar Value"]  = overdue_df["Days Since Report"].astype(float)
overdue_df["Label"]      = overdue_df["Overdue Bucket"].fillna("Overdue")
overdue_df["Bar Type"]   = "Overdue"

uptodate_df["Days Left"] = 120 - uptodate_df["Days Since Report"].astype(float)
uptodate_df["Bar Value"] = uptodate_df["Days Left"]
uptodate_df["Label"]     = "Up-to-date"
uptodate_df["Bar Type"]  = "Up-to-date"

# Top 20 overdue (worst first) + top 20 soonest-expiring up-to-date
top_overdue   = overdue_df.sort_values("Bar Value", ascending=False).head(20)
top_uptodate  = uptodate_df.sort_values("Days Left", ascending=True).head(20)
bar_df = pd.concat([top_overdue, top_uptodate], ignore_index=True)

if not bar_df.empty:
    st.markdown("### Vessel Report Status")
    color_map = {
        "120–180 days":  "#f39c12",
        "181–210 days":  "#e67e22",
        "210+ days":     "#e74c3c",
        "Overdue":       "#e74c3c",
        "Up-to-date":    "#2ecc71",
    }
    fig_over = px.bar(
        bar_df.sort_values("Bar Value", ascending=False),
        x="Vessel", y="Bar Value",
        color="Label",
        color_discrete_map=color_map,
        text_auto=True,
        title=(
            f"Overdue vessels (days since report) & "
            f"Up-to-date vessels (days left before overdue) — "
            f"{len(overdue_df)} overdue / {len(uptodate_df)} up-to-date in current filter"
        ),
        labels={"Bar Value": "Days", "Label": "Status"},
        custom_data=["Status", "Days Since Report", "Days Left"] if "Days Left" in bar_df.columns else ["Status"],
    )
    fig_over.update_traces(
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Status: %{customdata[0]}<br>"
            "Days since report: %{customdata[1]}<br>"
            "Days left: %{customdata[2]}<extra></extra>"
        )
    )
    fig_over.update_layout(
        xaxis_tickangle=-35,
        yaxis_title="Days",
        legend_title="Status",
        bargap=0.15,
    )
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


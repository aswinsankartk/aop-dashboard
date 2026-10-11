import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Target Dashboard FY26-27", layout="wide")

# ----------------------------------------------------------------------------
# Config
# ----------------------------------------------------------------------------
METRICS = [
    "Gross Profit",
    "Sales Revenue",
    "Accessory Revenue",
    "Expenses, Depreciation & Finance Cost",
    "Net Profit",
]
MONTHS = ["Apr-26", "May-26", "Jun-26", "Jul-26", "Aug-26", "Sep-26",
          "Oct-26", "Nov-26", "Dec-26", "Jan-27", "Feb-27", "Mar-27"]
QUARTER_OF = {m: q for ms, q in [
    (MONTHS[0:3], "Q1 (Apr-Jun)"), (MONTHS[3:6], "Q2 (Jul-Sep)"),
    (MONTHS[6:9], "Q3 (Oct-Dec)"), (MONTHS[9:12], "Q4 (Jan-Mar)")] for m in ms}
QUARTERS = list(dict.fromkeys(QUARTER_OF.values()))
CLUSTER_ORDER = ["KA1", "KA3", "KA4", "KL"]
ID_COLS = {"location": "Location", "cluster": "Cluster", "cm": "CM"}


# ----------------------------------------------------------------------------
# Data loading
# ----------------------------------------------------------------------------
def normalise_metric(text: str):
    t = str(text).lower()
    if "gross" in t:
        return "Gross Profit"
    if "accessor" in t:
        return "Accessory Revenue"
    if "sale" in t:
        return "Sales Revenue"
    if "expense" in t or "deprec" in t or "finance" in t:
        return "Expenses, Depreciation & Finance Cost"
    if "net" in t:
        return "Net Profit"
    return None


def normalise_month(x) -> str:
    if isinstance(x, (pd.Timestamp,)) or hasattr(x, "strftime"):
        return x.strftime("%b-%y")
    s = str(x).strip()
    if "total" in s.lower():
        return "Total"
    # handles strings like "2026-04-01 00:00:00"
    try:
        if len(s) >= 10 and s[4] == "-":
            return pd.to_datetime(s).strftime("%b-%y")
    except Exception:
        pass
    return s


def region_of(cluster: str) -> str:
    c = str(cluster).upper().strip()
    if c.startswith("KA"):
        return "KA"
    if c.startswith("KL"):
        return "KL"
    return c


@st.cache_data(show_spinner="Reading Excel...")
def load_data(file, sheet) -> pd.DataFrame:
    raw = pd.read_excel(file, sheet_name=sheet, header=[0, 1])

    # Build flat column names, forward-filling the merged metric header
    flat, current = [], None
    for top, sub in raw.columns:
        top_s, sub_n = str(top).strip(), normalise_month(sub)
        if sub_n.lower() in ID_COLS:
            flat.append(ID_COLS[sub_n.lower()])
            continue
        if not top_s.startswith("Unnamed") and top_s.lower() != "nan":
            current = normalise_metric(top_s) or current
        flat.append(f"{current}||{sub_n}")
    raw.columns = flat

    # Merged Location/Cluster cells -> forward fill
    raw[["Location", "Cluster"]] = raw[["Location", "Cluster"]].ffill()
    raw = raw.dropna(subset=["Location", "Cluster"], how="all")
    # Drop any grand-total / sub-total rows present in the sheet
    mask = (raw["Location"].astype(str).str.contains("total", case=False)
            | raw["Cluster"].astype(str).str.contains("total", case=False))
    raw = raw[~mask]

    long = raw.melt(id_vars=["Location", "Cluster", "CM"],
                    var_name="key", value_name="Value")
    long[["Metric", "Month"]] = long["key"].str.split(r"\|\|", expand=True)
    long = long.drop(columns="key")
    long = long[long["Month"].isin(MONTHS)]      # use months; Total is recomputed
    long["Value"] = pd.to_numeric(long["Value"], errors="coerce").fillna(0)
    long["Region"] = long["Cluster"].map(region_of)
    long["Quarter"] = long["Month"].map(QUARTER_OF)
    return long


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def fmt(v: float) -> str:
    return f"{v:,.0f}"


def pivot(df, index, columns="Metric"):
    p = df.pivot_table(index=index, columns=columns, values="Value",
                       aggfunc="sum", fill_value=0)
    if columns == "Metric":
        p = p.reindex(columns=[m for m in METRICS if m in p.columns])
    return p


def show_table(p: pd.DataFrame):
    st.dataframe(p.style.format("{:,.0f}"), use_container_width=True)


def kpi_row(df):
    totals = df.groupby("Metric")["Value"].sum()
    cols = st.columns(len(METRICS))
    for col, m in zip(cols, METRICS):
        col.metric(m, fmt(totals.get(m, 0)))


def metric_bar(df, x, metric_choice, color=None, title=""):
    d = df[df["Metric"] == metric_choice]
    d = d.groupby([x] + ([color] if color else []), as_index=False)["Value"].sum()
    fig = px.bar(d, x=x, y="Value", color=color, barmode="group",
                 text_auto=",.0f", title=title)
    fig.update_layout(yaxis_tickformat=",")
    return fig


def all_metrics_chart(df, x, title=""):
    d = df.groupby([x, "Metric"], as_index=False)["Value"].sum()
    fig = px.bar(d, x=x, y="Value", color="Metric", barmode="group",
                 category_orders={"Metric": METRICS}, title=title)
    fig.update_layout(yaxis_tickformat=",", legend_orientation="h",
                      legend_y=-0.25)
    return fig


# ----------------------------------------------------------------------------
# Sidebar - input & filters
# ----------------------------------------------------------------------------
st.title("📊 AOP Dashboard - FY 2026-27")

with st.sidebar:
    st.header("Data")
    uploaded = st.file_uploader("Upload Excel file", type=["xlsx", "xlsm", "xls"])
    if uploaded is None:
        st.info("Upload the Excel file to begin.")
        st.stop()
    sheets = pd.ExcelFile(uploaded).sheet_names
    sheet = st.selectbox("Sheet", sheets)

data = load_data(uploaded, sheet)
if data.empty:
    st.error("No data found. Please check the sheet's header layout.")
    st.stop()

with st.sidebar:
    st.header("Filters")
    regions = sorted(data["Region"].unique())
    sel_regions = st.multiselect("Region", regions, default=regions)
    cl_pool = [c for c in CLUSTER_ORDER if c in data["Cluster"].unique()] + \
              sorted(set(data["Cluster"].unique()) - set(CLUSTER_ORDER))
    cl_pool = [c for c in cl_pool if region_of(c) in sel_regions]
    sel_clusters = st.multiselect("Cluster", cl_pool, default=cl_pool)
    loc_pool = sorted(data[data["Cluster"].isin(sel_clusters)]["Location"].unique())
    sel_locs = st.multiselect("Location", loc_pool, default=loc_pool)

df = data[data["Region"].isin(sel_regions)
          & data["Cluster"].isin(sel_clusters)
          & data["Location"].isin(sel_locs)]
if df.empty:
    st.warning("No data for the selected filters.")
    st.stop()

# ----------------------------------------------------------------------------
# Tabs
# ----------------------------------------------------------------------------
tabs = st.tabs(["Total Yearly", "Region Yearly", "Cluster Yearly",
                "Quarterly (Cluster)", "Monthly (Cluster)", "Location"])

# 1. Total yearly ------------------------------------------------------------
with tabs[0]:
    st.subheader("Total Target - Yearly")
    kpi_row(df)
    c1, c2 = st.columns(2)
    totals = df.groupby("Metric", as_index=False)["Value"].sum()
    totals["Metric"] = pd.Categorical(totals["Metric"], METRICS, ordered=True)
    totals = totals.sort_values("Metric")
    with c1:
        fig = px.bar(totals, x="Metric", y="Value", text_auto=",.0f",
                     color="Metric", title="Yearly totals by metric")
        fig.update_layout(showlegend=False, yaxis_tickformat=",")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        mt = pivot(df, "Month", "Metric").reindex(MONTHS)
        fig = px.line(mt.reset_index().melt("Month", var_name="Metric",
                                            value_name="Value"),
                      x="Month", y="Value", color="Metric", markers=True,
                      title="Monthly trend")
        fig.update_layout(yaxis_tickformat=",")
        st.plotly_chart(fig, use_container_width=True)
    st.markdown("**Monthly breakdown**")
    show_table(mt.assign(**{}).T.assign(Total=lambda t: t.sum(axis=1)))

# 2. Region yearly -----------------------------------------------------------
with tabs[1]:
    st.subheader("Region-wise Yearly (KA & KL)")
    rp = pivot(df, "Region")
    for reg in rp.index:
        st.markdown(f"##### {reg}")
        cols = st.columns(len(METRICS))
        for col, m in zip(cols, rp.columns):
            col.metric(m, fmt(rp.loc[reg, m]))
    st.plotly_chart(all_metrics_chart(df, "Region", "Region comparison"),
                    use_container_width=True)
    show_table(rp)

# 3. Cluster yearly ----------------------------------------------------------
with tabs[2]:
    st.subheader("Cluster-wise Yearly (KA1, KA3, KA4, KL)")
    cp = pivot(df, "Cluster")
    cp = cp.reindex([c for c in CLUSTER_ORDER if c in cp.index]
                    + [c for c in cp.index if c not in CLUSTER_ORDER])
    st.plotly_chart(all_metrics_chart(df, "Cluster", "Cluster comparison"),
                    use_container_width=True)
    show_table(cp)
    pick = st.selectbox("Drill into metric", METRICS, key="cl_metric")
    st.plotly_chart(metric_bar(df, "Cluster", pick, title=f"{pick} by cluster"),
                    use_container_width=True)

# 4. Quarterly cluster-wise --------------------------------------------------
with tabs[3]:
    st.subheader("Quarterly - Cluster-wise")
    pick = st.selectbox("Metric", METRICS, key="q_metric")
    st.plotly_chart(
        metric_bar(df, "Quarter", pick, color="Cluster",
                   title=f"{pick} - quarterly by cluster")
        .update_xaxes(categoryorder="array", categoryarray=QUARTERS),
        use_container_width=True)
    st.markdown("**All metrics (Cluster x Quarter)**")
    qp = pivot(df, ["Cluster", "Quarter"])
    qp = qp.reindex(sorted(qp.index, key=lambda i: (
        CLUSTER_ORDER.index(i[0]) if i[0] in CLUSTER_ORDER else 99,
        QUARTERS.index(i[1]))))
    show_table(qp)

# 5. Monthly cluster-wise ----------------------------------------------------
with tabs[4]:
    st.subheader("Monthly - Cluster-wise")
    pick = st.selectbox("Metric", METRICS, key="m_metric")
    d = (df[df["Metric"] == pick].groupby(["Month", "Cluster"], as_index=False)
         ["Value"].sum())
    fig = px.line(d, x="Month", y="Value", color="Cluster", markers=True,
                  category_orders={"Month": MONTHS},
                  title=f"{pick} - monthly by cluster")
    fig.update_layout(yaxis_tickformat=",")
    st.plotly_chart(fig, use_container_width=True)
    st.markdown(f"**{pick}: Cluster x Month**")
    mp = df[df["Metric"] == pick].pivot_table(index="Cluster", columns="Month",
                                              values="Value", aggfunc="sum",
                                              fill_value=0)
    mp = mp.reindex(columns=MONTHS)
    mp["Total"] = mp.sum(axis=1)
    show_table(mp)
    with st.expander("All metrics (Cluster x Month)"):
        show_table(pivot(df, ["Cluster", "Month"]))

# 6. Location wise -----------------------------------------------------------
with tabs[5]:
    st.subheader("Location-wise")
    lp = pivot(df, "Location")
    pick = st.selectbox("Metric", METRICS, key="loc_metric")
    ld = (df[df["Metric"] == pick].groupby(["Location", "Cluster"],
                                           as_index=False)["Value"].sum()
          .sort_values("Value", ascending=False))
    fig = px.bar(ld, x="Location", y="Value", color="Cluster",
                 text_auto=",.0f", title=f"{pick} by location")
    fig.update_layout(yaxis_tickformat=",")
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("**All metrics by location**")
    show_table(lp)

    st.markdown("**Single-location view**")
    one = st.selectbox("Location", sorted(df["Location"].unique()))
    one_df = df[df["Location"] == one]
    kpi_row(one_df)
    mt1 = pivot(one_df, "Month", "Metric").reindex(MONTHS)
    fig = px.line(mt1.reset_index().melt("Month", var_name="Metric",
                                         value_name="Value"),
                  x="Month", y="Value", color="Metric", markers=True,
                  title=f"{one} - monthly trend")
    fig.update_layout(yaxis_tickformat=",")
    st.plotly_chart(fig, use_container_width=True)

# Download of the cleaned data ------------------------------------------------
with st.sidebar:
    st.download_button("⬇ Download filtered data (CSV)",
                       df.to_csv(index=False).encode(), "filtered_data.csv",
                       "text/csv")


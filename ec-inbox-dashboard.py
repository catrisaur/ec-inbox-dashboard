import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime

# =========================================================
# PAGE CONFIGURATION
# =========================================================
st.set_page_config(
    page_title="E&C Inbox Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📊 E&C Inbox Dashboard")
st.caption(
    "Comprehensive operational intelligence for email volumes, workload patterns, "
    "automation suitability, and efficiency forecasting."
)

# =========================================================
# FILE UPLOAD
# =========================================================
uploaded_file = st.file_uploader(
    "Upload your dataset (Excel or CSV)", type=["xlsx", "xls", "csv"]
)

# Cache loading for performance
@st.cache_data
def load_file(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file)
    else:
        return pd.read_excel(file)

if uploaded_file:

    # -----------------------------------------------------
    # Load Data
    # -----------------------------------------------------
    try:
        df = load_file(uploaded_file)
        st.success("✅ Data loaded successfully")
    except Exception as e:
        st.error(f"Error loading file: {e}")
        st.stop()

    # -----------------------------------------------------
    # Validate Required Schema
    # -----------------------------------------------------
    required_cols = ["DateTimeReceived", "Category", "Sub-Category", "Chatbot_Addressable"]
    missing_cols = [col for col in required_cols if col not in df.columns]

    with st.expander("🔍 View Detected Columns"):
        st.write(df.columns.tolist())

    if missing_cols:
        st.error(f"❌ Missing required columns: {missing_cols}")
        st.stop()

    # -----------------------------------------------------
    # CLEAN & FORMAT
    # -----------------------------------------------------
    df["DateTimeReceived"] = pd.to_datetime(df["DateTimeReceived"], errors="coerce")
    df.dropna(subset=["DateTimeReceived"], inplace=True)

    df.sort_values("DateTimeReceived", inplace=True)
    df["Date"] = df["DateTimeReceived"].dt.date
    df["Month"] = df["DateTimeReceived"].dt.to_period("M").astype(str)
    df["Hour"] = df["DateTimeReceived"].dt.hour
    df["Weekday"] = df["DateTimeReceived"].dt.day_name()

    # -----------------------------------------------------
    # Clean category columns to prevent plotting issues
    # -----------------------------------------------------
    df["Category"] = df["Category"].fillna("Unknown").astype(str).str.strip()
    df["Sub-Category"] = df["Sub-Category"].fillna("Unknown").astype(str).str.strip()
    df["Chatbot_Addressable"] = df["Chatbot_Addressable"].fillna("No").astype(str).str.strip()

    # =========================================================
    # SIDEBAR FILTERS
    # =========================================================
    st.sidebar.header("🔎 Filters")

    category_list = sorted(df["Category"].unique())
    selected_categories = st.sidebar.multiselect("Filter by Category", category_list)

    subcat_list = sorted(df["Sub-Category"].unique())
    selected_subcats = st.sidebar.multiselect("Filter by Sub-Category", subcat_list)

    date_min, date_max = df["DateTimeReceived"].min(), df["DateTimeReceived"].max()
    date_range = st.sidebar.date_input("Date Range", [date_min, date_max])

    filtered_df = df[
        (df["DateTimeReceived"] >= pd.to_datetime(date_range[0])) &
        (df["DateTimeReceived"] <= pd.to_datetime(date_range[1]))
    ]

    if selected_categories:
        filtered_df = filtered_df[filtered_df["Category"].isin(selected_categories)]

    if selected_subcats:
        filtered_df = filtered_df[filtered_df["Sub-Category"].isin(selected_subcats)]

    if filtered_df.empty:
        st.warning("No data matches your selected filters.")
        st.stop()

    # =========================================================
    # KPI CALCULATIONS
    # =========================================================
    today = pd.Timestamp.today()

    total_volume = len(filtered_df)
    chatbot_count = filtered_df[filtered_df["Chatbot_Addressable"] == "Yes"].shape[0]
    pct_chatbot = (chatbot_count / total_volume * 100) if total_volume > 0 else 0

    minutes_per_manual = 4
    minutes_chatbot = 0.1
    manual_min_total = total_volume * minutes_per_manual
    chatbot_min_total = chatbot_count * minutes_chatbot
    hours_saved = (manual_min_total - chatbot_min_total) / 60
    fte_saved = hours_saved / 160

    peak_hour = filtered_df["Hour"].mode()[0]
    peak_day = filtered_df["Date"].mode()[0]
    busiest_weekday = filtered_df["Weekday"].mode()[0]

    # =========================================================
    # KPI DASHBOARD
    # =========================================================
    st.subheader("📈 Executive KPIs")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Volume (Filtered)", total_volume)
    k2.metric("Automation-Addressable %", f"{pct_chatbot:.1f}%")
    k3.metric("Estimated Hours Saved", f"{hours_saved:.1f}")
    k4.metric("Estimated FTE Savings", f"{fte_saved:.2f}")

    k5, k6, k7 = st.columns(3)
    k5.metric("Peak Hour", f"{peak_hour}:00")
    k6.metric("Peak Day", str(peak_day))
    k7.metric("Busiest Weekday", busiest_weekday)

    # =========================================================
    # TREND ANALYSIS
    # =========================================================
    st.subheader("📉 Volume Trends & Workload Distribution")

    # Monthly Trend
    monthly = filtered_df.groupby("Month", as_index=False).size().rename(columns={"size": "Count"})
    fig_month = px.line(monthly, x="Month", y="Count", markers=True, title="Monthly Email Volume")
    st.plotly_chart(fig_month, use_container_width=True)

    # Daily Trend
    daily = filtered_df.groupby("Date", as_index=False).size().rename(columns={"size": "Count"})
    fig_daily = px.line(daily, x="Date", y="Count", markers=True, title="Daily Email Volume")
    st.plotly_chart(fig_daily, use_container_width=True)

    # Hourly Distribution
    hourly = filtered_df.groupby("Hour", as_index=False).size().rename(columns={"size": "Count"})
    fig_hour = px.bar(hourly, x="Hour", y="Count", title="Hourly Distribution")
    st.plotly_chart(fig_hour, use_container_width=True)

    # Weekday Heatmap
    weekday_hour = filtered_df.groupby(["Weekday", "Hour"], as_index=False).size().rename(columns={"size": "Count"})
    fig_heat = px.density_heatmap(
        weekday_hour,
        x="Hour", y="Weekday", z="Count",
        title="Workload Heatmap by Hour & Weekday",
        histfunc="avg",
        color_continuous_scale="Viridis"
    )
    st.plotly_chart(fig_heat, use_container_width=True)

    # =========================================================
    # CATEGORY INSIGHTS
    # =========================================================
    st.subheader("📂 Category & Sub-Category Insights")

    category_counts = (
        filtered_df.groupby("Category", as_index=False).size().rename(columns={"size": "Count"})
    )
    category_counts["% of Total"] = category_counts["Count"] / total_volume * 100

    fig_cat = px.bar(
        category_counts,
        x="Count",
        y="Category",
        orientation="h",
        title="Email Volume by Category",
        text="Count"
    )
    st.plotly_chart(fig_cat, use_container_width=True)

    subcat_counts = (
        filtered_df.groupby(["Category", "Sub-Category"], as_index=False).size().rename(columns={"size": "Count"})
    )
    fig_subcat = px.treemap(
        subcat_counts,
        path=["Category", "Sub-Category"],
        values="Count",
        title="Sub-Category Structure & Volume"
    )
    st.plotly_chart(fig_subcat, use_container_width=True)

    # =========================================================
    # AUTOMATION OPPORTUNITY
    # =========================================================
    st.subheader("🤖 Automation Opportunity Analysis")
    auto_summary = (
        filtered_df.groupby("Category", as_index=False)
        .agg(Total=("Category", "count"), Auto=("Chatbot_Addressable", lambda s: (s == "Yes").sum()))
    )
    auto_summary["Auto %"] = (auto_summary["Auto"] / auto_summary["Total"] * 100).round(1)
    auto_summary["Manual Hrs"] = auto_summary["Total"] * 4 / 60
    auto_summary["Potential Savings (hrs)"] = (auto_summary["Auto"] * 4 - auto_summary["Auto"] * 0.1) / 60

    st.dataframe(auto_summary)

    fig_auto = px.scatter(
        auto_summary,
        x="Auto %",
        y="Total",
        size="Total",
        color="Auto %",
        text="Category",
        title="Automation Potential vs Workload Impact"
    )
    fig_auto.update_traces(textposition="top center")
    st.plotly_chart(fig_auto, use_container_width=True)

    # =========================================================
    # RISK & ANOMALY FLAGS
    # =========================================================
    st.subheader("⚠ Operational Alerts & Risk Flags")
    if "Data Protection" in filtered_df["Category"].values:
        st.error("🚨 Data Protection emails detected — human review recommended")

    threshold = daily["Count"].mean() + 2 * daily["Count"].std()
    anomalies = daily[daily["Count"] > threshold]
    if not anomalies.empty:
        st.warning("📈 High-volume anomaly detected (above 2σ threshold).")
        st.dataframe(anomalies)

    # =========================================================
    # EXECUTIVE SUMMARY
    # =========================================================
    st.subheader("📌 Executive Summary & Recommendations")
    peak_month = monthly.loc[monthly["Count"].idxmax()]["Month"]
    top_category = category_counts.iloc[0]["Category"]

    st.markdown(f"""
### **Operational Insight Summary**

- **Highest-Volume Month:** {peak_month}  
- **Top Category:** {top_category}  
- **Automation Potential:** {pct_chatbot:.1f}%  
- **Estimated Hours Saved:** {hours_saved:.1f} hrs  
- **FTE Equivalence:** {fte_saved:.2f} FTEs  
- **Busiest Operating Window:** {busiest_weekday}, {peak_hour}:00  
- **Peak Day Volume:** {peak_day}  

### **Strategic Recommendations**
1. **Prioritize Automation for High-Volume & High-Suitability Categories**  
2. **Re-engineer High-Volume, Low-Automation Categories**  
3. **Align Staffing With Workload Peaks**  
4. **Monitor Sensitive Categories (e.g., Data Protection)**  
5. **Use Sub-Category Patterns to Inform Chatbot Training**
""")

else:
    st.info("Upload a dataset to enable the dashboard.")

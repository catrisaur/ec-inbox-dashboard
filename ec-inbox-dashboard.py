import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# =========================================================
# Application Header
# =========================================================
st.set_page_config(page_title="Email Operations Dashboard", layout="wide")

st.title("📊 Email Operations & Automation Intelligence Dashboard")
st.caption("Operational insights, automation potential, and volume trends for strategic decision-making.")

# =========================================================
# File Upload
# =========================================================
uploaded_file = st.file_uploader(
    "Upload your dataset (Excel or CSV)", type=["xlsx", "xls", "csv"]
)

if uploaded_file:

    # -----------------------------------------------------
    # Load File
    # -----------------------------------------------------
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        st.success("✅ Data loaded successfully")
    except Exception as e:
        st.error(f"Error loading file: {e}")
        st.stop()

    # -----------------------------------------------------
    # Validate Required Columns
    # -----------------------------------------------------
    required_cols = ["DateTimeReceived", "Category", "Sub-Category", "Chatbot_Addressable"]
    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        st.error(f"❌ Missing required columns: {missing_cols}")
        st.stop()

    # -----------------------------------------------------
    # Clean & Prepare Data
    # -----------------------------------------------------
    df["DateTimeReceived"] = pd.to_datetime(df["DateTimeReceived"], errors="coerce")
    df.dropna(subset=["DateTimeReceived"], inplace=True)
    df.sort_values("DateTimeReceived", inplace=True)

    # =========================================================
    # Sidebar Filters
    # =========================================================
    st.sidebar.header("🔎 Filters")

    category_list = sorted(df["Category"].unique())
    selected_categories = st.sidebar.multiselect("Category", category_list)

    date_min, date_max = df["DateTimeReceived"].min(), df["DateTimeReceived"].max()
    date_range = st.sidebar.date_input("Date Range", [date_min, date_max])

    filtered_df = df[
        (df["DateTimeReceived"] >= pd.to_datetime(date_range[0])) &
        (df["DateTimeReceived"] <= pd.to_datetime(date_range[1]))
    ]

    if selected_categories:
        filtered_df = filtered_df[filtered_df["Category"].isin(selected_categories)]

    if filtered_df.empty:
        st.warning("No data available for the selected filters.")
        st.stop()

    # =========================================================
    # KPI CALCULATIONS
    # =========================================================
    today = pd.Timestamp.today()

    total_ytd = len(filtered_df)
    total_mtd = filtered_df[filtered_df["DateTimeReceived"].dt.month == today.month].shape[0]
    total_wtd = filtered_df[filtered_df["DateTimeReceived"].dt.isocalendar().week == today.isocalendar().week].shape[0]
    total_today = filtered_df[filtered_df["DateTimeReceived"].dt.date == today.date()].shape[0]

    chatbot_count = filtered_df[filtered_df["Chatbot_Addressable"] == "Yes"].shape[0]
    pct_chatbot = (chatbot_count / len(filtered_df)) * 100

    # Operational efficiency assumptions
    manual_time_min = len(filtered_df) * 4           # 4 min per manual handling
    chatbot_time_min = chatbot_count * 0.1           # near-zero handling time
    time_saved_hours = (manual_time_min - chatbot_time_min) / 60
    fte_saved = time_saved_hours / 160               # 160 hours per FTE month

    # =========================================================
    # KPI Dashboard (Executive View)
    # =========================================================
    st.subheader("📈 Key Operational Metrics")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("YTD Volume", total_ytd)
    col2.metric("MTD Volume", total_mtd)
    col3.metric("WTD Volume", total_wtd)
    col4.metric("Today", total_today)

    col5, col6, col7 = st.columns(3)
    col5.metric("Automation-Addressable", f"{pct_chatbot:.1f}%")
    col6.metric("Hours Saved (Est.)", f"{time_saved_hours:.1f}")
    col7.metric("FTE Savings (Est.)", f"{fte_saved:.2f}")

    # =========================================================
    # Trend Analysis (Volumes)
    # =========================================================
    st.subheader("📉 Volume Trends")

    # Monthly
    df_monthly = (
        filtered_df.groupby(filtered_df["DateTimeReceived"].dt.to_period("M"))
        .size()
        .reset_index(name="Email Count")
    )
    df_monthly["Month"] = df_monthly["DateTimeReceived"].astype(str)

    fig_monthly = px.line(
        df_monthly,
        x="Month",
        y="Email Count",
        title="Monthly Email Volume",
        markers=True,
    )
    st.plotly_chart(fig_monthly, use_container_width=True)

    # Daily
    df_daily = (
        filtered_df.groupby(filtered_df["DateTimeReceived"].dt.date)
        .size()
        .reset_index(name="Email Count")
    )

    fig_daily = px.line(
        df_daily,
        x="DateTimeReceived",
        y="Email Count",
        title="Daily Email Volume",
        markers=True,
    )
    st.plotly_chart(fig_daily, use_container_width=True)

    # Hourly
    df_hourly = (
        filtered_df.groupby(filtered_df["DateTimeReceived"].dt.hour)
        .size()
        .reset_index(name="Email Count")
    )

    fig_hourly = px.bar(
        df_hourly,
        x="DateTimeReceived",
        y="Email Count",
        title="Hourly Distribution",
    )
    st.plotly_chart(fig_hourly, use_container_width=True)

    # =========================================================
    # CATEGORY INSIGHTS (Enhanced)
    # =========================================================
    st.subheader("📂 Category Insights")

    # Category Breakdown
    df_category = (
        filtered_df["Category"]
        .value_counts()
        .reset_index()
    )
    df_category.columns = ["Category", "Email Count"]
    df_category["% of Total"] = (df_category["Email Count"] / df_category["Email Count"].sum() * 100).round(1)

    fig_category = px.bar(
        df_category,
        x="Email Count",
        y="Category",
        orientation="h",
        title="Email Volume by Category",
        text="Email Count"
    )
    fig_category.update_traces(textposition="outside")
    st.plotly_chart(fig_category, use_container_width=True)

    # Category KPI Table
    st.markdown("### 📊 Category Performance Overview")

    df_cat_kpi = df_category.copy()
    df_cat_kpi["Chatbot-Addressable"] = df_cat_kpi["Category"].apply(
        lambda x: filtered_df[(filtered_df["Category"] == x) & (filtered_df["Chatbot_Addressable"] == "Yes")].shape[0]
    )
    df_cat_kpi["Automation Potential (%)"] = (
        df_cat_kpi["Chatbot-Addressable"] / df_cat_kpi["Email Count"] * 100
    ).round(1)
    df_cat_kpi["Manual Hours (Est.)"] = df_cat_kpi["Email Count"] * 4 / 60
    df_cat_kpi["Savings (Hrs)"] = (df_cat_kpi["Chatbot-Addressable"] * 4 - df_cat_kpi["Chatbot-Addressable"] * 0.1) / 60

    st.dataframe(df_cat_kpi)

    # Sub-Category Deep Dive
    st.markdown("### 📁 Sub-Category Breakdown")

    df_subcat = (
        filtered_df.groupby(["Category", "Sub-Category"])
        .size()
        .reset_index(name="Email Count")
    )

    fig_subcat = px.treemap(
        df_subcat,
        path=["Category", "Sub-Category"],
        values="Email Count",
        title="Sub-Category Distribution"
    )
    st.plotly_chart(fig_subcat, use_container_width=True)

    # =========================================================
    # AUTOMATION OPPORTUNITY (Enhanced)
    # =========================================================
    st.subheader("🤖 Automation Opportunity")

    df_chatbot = (
        filtered_df["Chatbot_Addressable"]
        .value_counts()
        .reset_index()
    )
    df_chatbot.columns = ["Chatbot", "Count"]

    fig_chatbot = px.pie(
        df_chatbot,
        names="Chatbot",
        values="Count",
        title="Chatbot-Addressable vs Manual Handling",
    )
    st.plotly_chart(fig_chatbot, use_container_width=True)

    # Automation Heatmap
    st.markdown("### 🔥 Automation Priority Heatmap")

    heatmap_data = df_cat_kpi.copy()
    heatmap_data["Volume Tier"] = pd.qcut(heatmap_data["Email Count"], 2, labels=["Low", "High"])
    heatmap_data["Automation Tier"] = pd.qcut(heatmap_data["Automation Potential (%)"], 2, labels=["Low", "High"])

    fig_heatmap = px.scatter(
        heatmap_data,
        x="Automation Potential (%)",
        y="Email Count",
        size="Email Count",
        color="Automation Tier",
        text="Category",
        title="Automation Impact vs Workload",
    )
    fig_heatmap.update_traces(textposition="top center")
    st.plotly_chart(fig_heatmap, use_container_width=True)

    # =========================================================
    # Alerts & Risk Flags
    # =========================================================
    st.subheader("⚠ Operational Alerts")

    peak_day = df_daily.loc[df_daily["Email Count"].idxmax()]
    st.warning(f"📌 **Peak Day:** {peak_day['DateTimeReceived']} with **{peak_day['Email Count']}** emails.")

    if "Data Protection" in filtered_df["Category"].values:
        st.error("🚨 Data Protection emails detected — immediate review recommended")

    # =========================================================
    # Executive Summary (Enhanced)
    # =========================================================
    st.subheader("📌 Executive Summary")

    peak_month = df_monthly.loc[df_monthly["Email Count"].idxmax()]["Month"]
    top_category = df_category.iloc[0]["Category"]

    st.write(f"""
### **Operational Insight Summary**

- **Highest-Volume Month:** {peak_month}  
- **Top Category:** {top_category}  
- **Automation Potential:** {pct_chatbot:.1f}% of emails are suitable for chatbot handling  
- **Productivity Impact:** Estimated **{time_saved_hours:.1f} hours saved** = **{fte_saved:.2f} FTEs**  
- **Risk Signals:** {"Data Protection activity detected" if "Data Protection" in filtered_df["Category"].values else "No critical risk categories detected"}  
- **Operational Load Concentration:** {top_category} represents **{df_category.iloc[0]['% of Total']:.1f}%** of all email traffic  

### **Strategic Recommendations**
- **Automate High-Volume, High-Suitability Paths** (fastest ROI).  
- **Re-engineer High-Volume, Low-Suitability Categories** to reduce repeated queries.  
- **Maintain Manual Review** for low-volume but sensitive categories such as Whistleblowing or Data Protection.  

These insights highlight clear opportunities to improve operational capacity, reduce SLA delays, 
and enhance compliance assurance through targeted automation.
""")

else:
    st.info("Upload a dataset to enable the dashboard.")

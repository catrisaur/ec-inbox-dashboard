import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# =========================================================
# PAGE CONFIGURATION
# =========================================================
st.set_page_config(page_title="E&C Inbox Dashboard", layout="wide")
st.title("📊 E&C Inbox Dashboard")
st.caption("Operational intelligence for email volumes, automation potential, and efficiency gains.")

# =========================================================
# FILE UPLOAD
# =========================================================
uploaded_file = st.file_uploader("Upload your dataset (Excel or CSV)", type=["xlsx", "xls", "csv"])

@st.cache_data
def load_file(file):
    return pd.read_csv(file) if file.name.endswith(".csv") else pd.read_excel(file)

if uploaded_file:
    df = load_file(uploaded_file)
    st.success("✅ Data loaded successfully")

    # Validate schema
    required_cols = ["DateTimeReceived", "Category", "Sub-Category", "Sub-Sub-Category", "Chatbot_Addressable"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"❌ Missing required columns: {missing}")
        st.stop()

    # =========================================================
    # CLEAN DATETIME
    # =========================================================
    df["DateTimeReceived"] = pd.to_datetime(
        df["DateTimeReceived"],
        format="%m/%d/%Y %H:%M",
        errors="coerce"
    )
    df.dropna(subset=["DateTimeReceived"], inplace=True)

    df["Date"] = df["DateTimeReceived"].dt.date
    df["Month"] = df["DateTimeReceived"].dt.to_period("M").astype(str)
    df["Hour"] = df["DateTimeReceived"].dt.hour
    df["Weekday"] = df["DateTimeReceived"].dt.day_name()

    min_date = df["DateTimeReceived"].min().date()
    max_date = df["DateTimeReceived"].max().date()

    # =========================================================
    # SIDEBAR FILTERS
    # =========================================================
    st.sidebar.header("🔎 Filters")
    selected_categories = st.sidebar.multiselect("Filter by Category", sorted(df["Category"].unique()))
    selected_subcats = st.sidebar.multiselect("Filter by Sub-Category", sorted(df["Sub-Category"].unique()))
    date_range = st.sidebar.date_input(
        "Date Range",
        value=[min_date, max_date],
        min_value=min_date,
        max_value=max_date
    )

    # Apply filters
    filtered_df = df[
        (df["DateTimeReceived"].dt.date >= date_range[0]) &
        (df["DateTimeReceived"].dt.date <= date_range[1])
    ]
    if selected_categories:
        filtered_df = filtered_df[filtered_df["Category"].isin(selected_categories)]
    if selected_subcats:
        filtered_df = filtered_df[filtered_df["Sub-Category"].isin(selected_subcats)]

    if filtered_df.empty:
        st.warning("No data matches your filters.")
        st.stop()

    # =========================================================
    # KPI DASHBOARD
    # =========================================================
    total_volume = len(filtered_df)
    chatbot_count = filtered_df[filtered_df["Chatbot_Addressable"] == "Yes"].shape[0]
    pct_chatbot = (chatbot_count / total_volume * 100) if total_volume > 0 else 0
    hours_saved = ((total_volume * 4) - (chatbot_count * 0.1)) / 60
    fte_saved = hours_saved / 160

    st.subheader("📈 Executive KPIs")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Emails", total_volume)
    k2.metric("Automation %", f"{pct_chatbot:.1f}%")
    k3.metric("Hours Saved", f"{hours_saved:.1f}")
    k4.metric("FTE Savings", f"{fte_saved:.2f}")

    # =========================================================
    # TREND ANALYSIS
    # =========================================================
    st.subheader("📉 Volume Trends")
    monthly = filtered_df.groupby("Month", as_index=False).size().rename(columns={"size": "Count"})
    fig_month = px.line(
        monthly, x="Month", y="Count", markers=True, title="Monthly Email Volume",
        line_shape='linear'
    )
    st.plotly_chart(fig_month, use_container_width=True)

    # Hourly Heatmap
    weekday_hour = filtered_df.groupby(["Weekday", "Hour"], as_index=False).size().rename(columns={"size": "Count"})
    fig_heat = px.density_heatmap(
        weekday_hour, x="Hour", y="Weekday", z="Count", title="Email Volume Heatmap by Hour and Weekday",
        color_continuous_scale="reds"
    )
    st.plotly_chart(fig_heat, use_container_width=True)

    # =========================================================
    # CATEGORY INSIGHTS
    # =========================================================
    st.subheader("📂 Category Insights")
    category_counts = filtered_df.groupby("Category", as_index=False).size().rename(columns={"size": "Count"})
    fig_cat = px.bar(
        category_counts, x="Count", y="Category", orientation="h",
        title="Volume by Category",
        color="Count", color_continuous_scale=px.colors.sequential.OrRd
    )
    st.plotly_chart(fig_cat, use_container_width=True)

    # =========================================================
    # EXECUTIVE SUMMARY & INSIGHTS
    # =========================================================
    st.subheader("📌 Strategic Recommendations & Insights")

    # Top category & peak month
    top_category = category_counts.iloc[0]['Category'] if not category_counts.empty else "N/A"
    peak_month = monthly.loc[monthly['Count'].idxmax()]['Month'] if len(monthly) else "N/A"

    # Additional insights
    # 1. Automation insights
    automation_text = f"""
**Automation Coverage:** {pct_chatbot:.1f}% of emails are handled by chatbot.  
*Insight:* High automation reduces manual effort, improves response time, and lowers operational costs.  
*Action:* Focus on categories with lower automation for potential efficiency gains.
"""
    st.markdown(automation_text)

    # 2. Peak hour & weekday
    hourly_counts = filtered_df.groupby("Hour").size()
    peak_hour = hourly_counts.idxmax()
    weekday_counts = filtered_df.groupby("Weekday").size()
    peak_weekday = weekday_counts.idxmax()
    st.markdown(f"""
**Peak Times:**  
- Hour: {peak_hour}:00  
- Weekday: {peak_weekday}  
*Insight:* Most emails arrive at this time; adjust staffing to match peak workload.
""")

    # 3. Category automation details
    top_category_data = filtered_df.groupby("Category")["Chatbot_Addressable"].value_counts(normalize=True).unstack().fillna(0)
    top_category_data = top_category_data.sort_values("Yes", ascending=False)
    st.markdown("**Category Automation Insights:**")
    for cat in top_category_data.index[:3]:
        yes_pct = top_category_data.loc[cat, "Yes"]*100 if "Yes" in top_category_data.columns else 0
        no_pct = top_category_data.loc[cat, "No"]*100 if "No" in top_category_data.columns else 0
        st.markdown(f"- {cat}: {yes_pct:.1f}% automated, {no_pct:.1f}% manual. Consider further automation for manual tasks.")

    # 4. Month-over-month trend
    monthly_counts = filtered_df.groupby("Month").size()
    if len(monthly_counts) > 1:
        change = (monthly_counts.iloc[-1] - monthly_counts.iloc[-2]) / monthly_counts.iloc[-2] * 100
        st.markdown(f"- **Month-over-Month Change:** {change:.1f}%")
        st.markdown("  *Insight:* Increase/decrease in volume may indicate seasonal trends or campaigns.")

    # 5. Hours & FTE saved
    st.markdown(f"- **Estimated Hours Saved:** {hours_saved:.1f} hrs, equivalent to {fte_saved:.2f} FTEs.")
    st.markdown("  *Insight:* Reducing manual effort allows staff to focus on higher-value tasks, improving efficiency.")

    # Final summary
    st.markdown(f"""
**Top Category:** {top_category}  
**Peak Month:** {peak_month}  

**Recommendations:**  
1. Prioritize automation for high-volume categories.  
2. Align staffing with peak workload hours.  
3. Monitor sensitive categories for compliance risk.  
""")

else:
    st.info("Upload a dataset to enable the dashboard.")

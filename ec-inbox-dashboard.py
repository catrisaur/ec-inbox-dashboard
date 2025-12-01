
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
    required_cols = ["DateTimeReceived", "Category", "Sub-Category", "Chatbot_Addressable"]
    if any(col not in df.columns for col in required_cols):
        st.error("❌ Missing required columns.")
        st.stop()

    # =========================================================
    # DATA PREP
    # =========================================================
    df["DateTimeReceived"] = pd.to_datetime(df["DateTimeReceived"], errors="coerce")
    df.dropna(subset=["DateTimeReceived"], inplace=True)
    df["Date"] = df["DateTimeReceived"].dt.date
    df["Month"] = df["DateTimeReceived"].dt.to_period("M").astype(str)
    df["Hour"] = df["DateTimeReceived"].dt.hour
    df["Weekday"] = df["DateTimeReceived"].dt.day_name()

    # =========================================================
    # SIDEBAR FILTERS
    # =========================================================
    st.sidebar.header("🔎 Filters")
    selected_categories = st.sidebar.multiselect("Filter by Category", sorted(df["Category"].unique()))
    selected_subcats = st.sidebar.multiselect("Filter by Sub-Category", sorted(df["Sub-Category"].unique()))
    date_range = st.sidebar.date_input("Date Range", [df["DateTimeReceived"].min(), df["DateTimeReceived"].max()])

    filtered_df = df[
        (df["DateTimeReceived"] >= pd.to_datetime(date_range[0])) &
        (df["DateTimeReceived"] <= pd.to_datetime(date_range[1]))
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
    fig_month = px.line(monthly, x="Month", y="Count", markers=True, title="Monthly Email Volume")
    st.plotly_chart(fig_month, use_container_width=True)

    # Hourly Heatmap
    weekday_hour = filtered_df.groupby(["Weekday", "Hour"], as_index=False).size().rename(columns={"size": "Count"})
    fig_heat = px.density_heatmap(weekday_hour, x="Hour", y="Weekday", z="Count", color_continuous_scale="Viridis")
    st.plotly_chart(fig_heat, use_container_width=True)

    # =========================================================
    # CATEGORY INSIGHTS
    # =========================================================
    st.subheader("📂 Category Insights")
    category_counts = filtered_df.groupby("Category", as_index=False).size().rename(columns={"size": "Count"})
    fig_cat = px.bar(category_counts, x="Count", y="Category", orientation="h", title="Volume by Category")
    st.plotly_chart(fig_cat, use_container_width=True)

    # =========================================================
    # EXECUTIVE SUMMARY
    # =========================================================
    st.subheader("📌 Strategic Recommendations")
    st.markdown(f"""
    - **Automation Potential:** {pct_chatbot:.1f}%  
    - **Estimated Hours Saved:** {hours_saved:.1f} hrs  
    - **Top Category:** {category_counts.iloc[0]['Category']}  
    - **Peak Month:** {monthly.loc[monthly['Count'].idxmax()]['Month']}  

    **Recommendations:**
    1. Prioritize automation for high-volume categories.
    2. Align staffing with peak workload hours.
    3. Monitor sensitive categories for compliance risk.
    """)

else:
    st.info("Upload a dataset to enable the dashboard.")

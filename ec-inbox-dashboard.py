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
    if file.name.endswith(".csv"):
        return pd.read_csv(file)
    else:
        return pd.read_excel(file)

if uploaded_file:
    df = load_file(uploaded_file)
    st.success("✅ Data loaded successfully")

    # -----------------------
    # Validate schema
    # -----------------------
    required_cols = ["DateTimeReceived", "Category", "Sub-Category", "Sub-Sub-Category", "Chatbot_Addressable", "Body.TextBody"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"❌ Missing required columns: {missing}")
        st.stop()

    # =========================================================
    # CLEAN DATETIME
    # =========================================================
    df["DateTimeReceived"] = pd.to_datetime(df["DateTimeReceived"], errors="coerce")
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
    chatbot_count = filtered_df["Chatbot_Addressable"].eq("Yes").sum()
    pct_chatbot = (chatbot_count / total_volume * 100) if total_volume else 0
    # Assume 4 minutes per email without automation, 0.1 min if automated
    hours_saved = ((total_volume * 4) - (chatbot_count * 0.1)) / 60
    fte_saved = hours_saved / 160

    st.subheader("📈 Executive KPIs")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Emails", total_volume)
    k2.metric("Automation Potential", f"{pct_chatbot:.1f}%")
    k3.metric("Potential Hours Saved", f"{hours_saved:.1f}")
    k4.metric("Potential FTE Savings", f"{fte_saved:.2f}")

    # =========================================================
    # TREND ANALYSIS
    # =========================================================
    st.subheader("📉 Volume Trends")
    monthly = filtered_df.groupby("Month").size().reset_index(name="Count")
    fig_month = px.line(monthly, x="Month", y="Count", markers=True, title="Monthly Email Volume")
    st.plotly_chart(fig_month, use_container_width=True)

    # Hourly Heatmap
    weekday_hour = filtered_df.groupby(["Weekday", "Hour"]).size().reset_index(name="Count")
    fig_heat = px.density_heatmap(
        weekday_hour, x="Hour", y="Weekday", z="Count", title="Email Volume by Hour & Weekday",
        color_continuous_scale="reds"
    )
    st.plotly_chart(fig_heat, use_container_width=True)

    # =========================================================
    # CATEGORY INSIGHTS
    # =========================================================
    st.subheader("📂 Category Insights")
    category_counts = filtered_df.groupby("Category").size().reset_index(name="Count").sort_values("Count", ascending=False)
    fig_cat = px.bar(
        category_counts, x="Count", y="Category", orientation="h",
        color="Count", color_continuous_scale=px.colors.sequential.OrRd,
        title="Volume by Category"
    )
    st.plotly_chart(fig_cat, use_container_width=True)

    # =========================================================
    # CATEGORY & SUB-CATEGORY BREAKDOWN — FLAT
    # =========================================================
    st.subheader("📝 Category & Sub-Category Overview")

for cat, cat_df in filtered_df.groupby("Category"):
    st.markdown(f"### 📂 {cat} — {len(cat_df)} emails | Automation: {cat_df['Chatbot_Addressable'].eq('Yes').mean()*100:.1f}%")
    
    subcat_cols = st.columns(1)  # You can make 2-3 per row if wide enough
    for subcat, subcat_df in cat_df.groupby("Sub-Category"):
        total_sub = len(subcat_df)
        pct_chatbot_sub = subcat_df['Chatbot_Addressable'].eq("Yes").mean() * 100
        peak_hour = subcat_df["Hour"].mode().iloc[0] if not subcat_df.empty else "N/A"
        peak_weekday = subcat_df["Weekday"].mode().iloc[0] if not subcat_df.empty else "N/A"
        
        with st.expander(f"▶ {subcat} — {total_sub} emails | Automation: {pct_chatbot_sub:.1f}%"):
            st.markdown(f"- **Peak Hour:** {peak_hour}:00")
            st.markdown(f"- **Peak Weekday:** {peak_weekday}")
            st.markdown("- **Sample Emails:**")
            for sample in subcat_df["Body.TextBody"].dropna().head(3):
                st.markdown(f"    - {sample[:200]}{'...' if len(sample) > 200 else ''}")

    # =========================================================
    # EXECUTIVE SUMMARY & STRATEGIC INSIGHTS
    # =========================================================
    st.subheader("📌 Strategic Recommendations & Insights")
    top_category = category_counts.iloc[0]['Category'] if not category_counts.empty else "N/A"
    peak_month = monthly.loc[monthly['Count'].idxmax()]['Month'] if len(monthly) else "N/A"

    st.markdown(f"""
**Top Category:** {top_category}  
**Peak Month:** {peak_month}  

**Recommendations:**  
1. Prioritize automation for high-volume categories.  
2. Align staffing with peak workload hours.  
3. Monitor sensitive categories for compliance risk.  
4. Review sample emails per category to identify patterns for process improvement.
""")

else:
    st.info("Upload a dataset to enable the dashboard.")

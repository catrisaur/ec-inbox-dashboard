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
# FILE LOAD (default file or upload)
# =========================================================
DEFAULT_FILE = "ECInbox_Analysis_20251201.xlsx"  # <-- Replace with your actual file path

@st.cache_data
def load_default_data(file_path):
    df = pd.read_excel(file_path)
    return df

df = load_default_data(DEFAULT_FILE)
st.success(f"✅ Dataset loaded: {DEFAULT_FILE}")


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
# SUB-CATEGORY TABLE — EXECUTIVE VIEW
# =========================================================
st.subheader("📝 Sub-Category Insights (Top Categories)")

overview_data = []
for cat in filtered_df["Category"].unique():
    cat_df = filtered_df[filtered_df["Category"] == cat]
    for subcat in cat_df["Sub-Category"].unique():
        subcat_df = cat_df[cat_df["Sub-Category"] == subcat]
        total_emails = len(subcat_df)
        chatbot_yes = subcat_df[subcat_df["Chatbot_Addressable"] == "Yes"].shape[0]
        pct_auto = round((chatbot_yes / total_emails * 100) if total_emails else 0, 1)
        peak_hour = subcat_df.groupby("Hour").size().idxmax()
        peak_weekday = subcat_df.groupby("Weekday").size().idxmax()
        sample_email = subcat_df["Body.TextBody"].dropna().iloc[0][:150] + ("..." if len(subcat_df["Body.TextBody"].iloc[0]) > 150 else "")

        overview_data.append({
            "Category": cat,
            "Sub-Category": subcat,
            "Total Emails": total_emails,
            "Automation Potential (%)": pct_auto,
            "Peak Hour": peak_hour,
            "Peak Weekday": peak_weekday,
            "Sample Email": sample_email
        })

overview_df = pd.DataFrame(overview_data)
st.dataframe(
    overview_df.sort_values(["Automation Potential (%)", "Total Emails"], ascending=[False, False]),
    use_container_width=True
)

# Optional: Expand sample emails per row
st.info("🔹 Sample emails truncated to 150 characters.")


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

st.success("✅ Dashboard generation complete.")
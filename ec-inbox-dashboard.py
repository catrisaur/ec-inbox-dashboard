
import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from wordcloud import WordCloud
import matplotlib.pyplot as plt

# =========================================================
# PAGE CONFIGURATION
# =========================================================
st.set_page_config(page_title="E&C Inbox Dashboard", layout="wide")
st.title("📊 **E&C Inbox Executive Dashboard**")
st.caption("Operational intelligence for email volumes, automation potential, and efficiency gains.")

# =========================================================
# FILE UPLOAD
# =========================================================
uploaded_file = st.file_uploader("📂 Upload your dataset (Excel or CSV)", type=["xlsx", "xls", "csv"])

@st.cache_data
def load_file(file):
    return pd.read_csv(file) if file.name.endswith(".csv") else pd.read_excel(file)

if uploaded_file:
    df = load_file(uploaded_file)
    st.success("✅ Data loaded successfully")

    # Validate schema
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

    min_date, max_date = df["DateTimeReceived"].min().date(), df["DateTimeReceived"].max().date()

    # =========================================================
    # SIDEBAR FILTERS
    # =========================================================
    st.sidebar.header("🔎 **Filters**")
    selected_categories = st.sidebar.multiselect("Filter by Category", sorted(df["Category"].unique()))
    selected_subcats = st.sidebar.multiselect("Filter by Sub-Category", sorted(df["Sub-Category"].unique()))
    date_range = st.sidebar.date_input("Date Range", value=[min_date, max_date], min_value=min_date, max_value=max_date)

    filtered_df = df[(df["DateTimeReceived"].dt.date >= date_range[0]) & (df["DateTimeReceived"].dt.date <= date_range[1])]
    if selected_categories:
        filtered_df = filtered_df[filtered_df["Category"].isin(selected_categories)]
    if selected_subcats:
        filtered_df = filtered_df[filtered_df["Sub-Category"].isin(selected_subcats)]

    if filtered_df.empty:
        st.warning("⚠ No data matches your filters.")
        st.stop()

    # =========================================================
    # KPI DASHBOARD
    # =========================================================
    st.markdown("### 📈 **Executive KPIs**")

    total_volume = len(filtered_df)
    chatbot_count = filtered_df["Chatbot_Addressable"].eq("Yes").sum()
    pct_chatbot = (chatbot_count / total_volume * 100) if total_volume else 0
    hours_saved = ((total_volume * 4) - (chatbot_count * 0.1)) / 60
    fte_saved = hours_saved / 160

    # Calculate averages
    days_range = (filtered_df["DateTimeReceived"].max() - filtered_df["DateTimeReceived"].min()).days + 1
    avg_per_day = round(total_volume / days_range, 2)

    months_range = len(filtered_df["Month"].unique())
    avg_per_month = round(total_volume / months_range, 2)

    # ================= Titles for Clarity =================
    st.markdown("###### **Volume Metrics**")
    k1, k5, k6 = st.columns(3)
    k1.metric("📧 Total Emails", f"{total_volume:,}")
    k5.metric("📅 Avg Emails per Day", f"{avg_per_day}")
    k6.metric("🗓 Avg Emails per Month", f"{avg_per_month}")

    st.markdown("###### **Automation Efficiency**")
    k2, k3, k4 = st.columns(3)
    k2.metric("⚙️ Automation Potential", f"{pct_chatbot:.1f}%")
    k3.metric("⏳ Estimated Hours Saved", f"{hours_saved:.1f}")
    k4.metric("👥 Estimated FTE Savings", f"{fte_saved:.2f}")

    st.divider()


    # =========================================================
    # TREND ANALYSIS
    # =========================================================
    st.markdown("### 📉 **Volume Trends**")
    monthly = filtered_df.groupby("Month").size().reset_index(name="Count")
    fig_month = px.line(monthly, x="Month", y="Count", markers=True, title="Monthly Email Volume", color_discrete_sequence=["#1f77b4"])
    st.plotly_chart(fig_month, use_container_width=True)

    # Weekly Heatmap
    weekday_hour = filtered_df.groupby(["Weekday", "Hour"]).size().reset_index(name="Count")
    fig_heat = px.density_heatmap(weekday_hour, x="Hour", y="Weekday", z="Count", title="Email Volume by Hour & Weekday", color_continuous_scale="Blues")
    st.plotly_chart(fig_heat, use_container_width=True)

    st.divider()

    # =========================================================
    # CATEGORY INSIGHTS
    # =========================================================
    st.markdown("### 📂 **Category Insights**")
    category_counts = filtered_df.groupby("Category").size().reset_index(name="Count").sort_values("Count", ascending=False)
    fig_cat = px.bar(category_counts, x="Count", y="Category", orientation="h", color="Count", color_continuous_scale=px.colors.sequential.Blues, title="Volume by Category")
    st.plotly_chart(fig_cat, use_container_width=True)

    # Treemap Visualization
    st.markdown("#### 📦 Treemap: Category & Sub-Category Distribution")
    treemap_df = filtered_df.groupby(["Category", "Sub-Category"]).size().reset_index(name="Count")
    fig_treemap = px.treemap(treemap_df, path=["Category", "Sub-Category"], values="Count", color="Category", color_discrete_sequence=px.colors.sequential.Blues, title="Category and Sub-Category Distribution")
    fig_treemap.update_traces(root_color="white")
    st.plotly_chart(fig_treemap, use_container_width=True)

    st.divider()

    # =========================================================
    # NEW VISUAL: Automation Potential Bubble Chart
    # =========================================================
    st.markdown("### 🔍 **Automation Opportunity Map**")
    bubble_df = filtered_df.groupby("Category").agg({"Chatbot_Addressable": lambda x: (x == "Yes").sum(), "Category": "count"}).rename(columns={"Category": "Total"}).reset_index()
    bubble_df["Automation %"] = (bubble_df["Chatbot_Addressable"] / bubble_df["Total"]) * 100
    fig_bubble = px.scatter(bubble_df, x="Total", y="Automation %", size="Total", color="Category", hover_name="Category", title="Automation Potential vs Volume", color_discrete_sequence=px.colors.qualitative.Set2)
    st.plotly_chart(fig_bubble, use_container_width=True)

    st.divider()



    # =========================================================
    # TOP TWO-WORD PHRASES (BIGRAMS)
    # =========================================================
    st.markdown("### 🗂 **Top Two-Word Phrases in Emails**")

    import re
    from collections import Counter

    # Combine all email text
    text_data = " ".join(filtered_df["Subject"].dropna().tolist())
    words = re.findall(r'\b\w+\b', text_data.lower())

    # Remove common stopwords
    stopwords = set(["the", "and", "to", "of", "in", "for", "on", "at", "a", "is", "with", "by", "an", "be", "or"])
    filtered_words = [w for w in words if w not in stopwords and len(w) > 2]

    # Create bigrams (two-word phrases)
    bigrams = zip(filtered_words, filtered_words[1:])
    bigram_phrases = [" ".join(pair) for pair in bigrams]

    # Get top 20 bigrams
    common_bigrams = Counter(bigram_phrases).most_common(20)
    bigrams_df = pd.DataFrame(common_bigrams, columns=["Phrase", "Frequency"])

    # Plot interactive bar chart
    fig_bigrams = px.bar(
        bigrams_df, x="Frequency", y="Phrase", orientation="h",
        color="Frequency", color_continuous_scale="Blues",
        title="Top Two-Word Phrases in Emails"
    )
    st.plotly_chart(fig_bigrams, use_container_width=True)
    st.divider()


    # =========================================================
    # ACTIONABLE INSIGHTS
    # =========================================================
    st.markdown("### 📌 **Strategic Recommendations & Insights**")
    top_category = category_counts.iloc[0]['Category'] if not category_counts.empty else "N/A"
    peak_month = monthly.loc[monthly['Count'].idxmax()]['Month'] if len(monthly) else "N/A"

    st.markdown(f"""
**Top Category:** `{top_category}`  
**Peak Month:** `{peak_month}`  

✅ **Actionable Insights:**  
- **Automate high-volume, low-automation categories** (see bubble chart).  
- **Focus on peak workload days/hours** for resource allocation.  
- **Monitor compliance-sensitive categories** for risk mitigation.  
- **Leverage keyword analysis** to identify recurring requests for chatbot scripts.  
- **Forecast next quarter volumes** to plan staffing and automation investments.
""")

else:
    st.info("📥 Upload a dataset to enable the dashboard.")

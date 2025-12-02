import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from collections import Counter
import re

# =========================================================
# PAGE CONFIGURATION
# =========================================================
st.set_page_config(page_title="E&C Inbox Dashboard", layout="wide")
st.title("📊 **E&C Inbox Dashboard**")
st.caption("Operational intelligence for email volumes, automation potential, and efficiency gains.")

# =========================================================
# LOAD DATASET
# =========================================================
@st.cache_data
def load_file(filepath="ECInbox_Analysis_20251202.xlsx"):
    return pd.read_excel(filepath)

try:
    df = load_file()
    st.success("✅ Data loaded successfully")
except Exception as e:
    st.error(f"❌ Failed to load dataset: {e}")
    st.stop()

# =========================================================
# VALIDATE REQUIRED COLUMNS
# =========================================================
required_cols = ["DateTimeReceived", "Category", "Sub-Category", "Sub-Sub-Category",
                 "Chatbot_Addressable", "Body.TextBody", "Subject"]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    st.error(f"❌ Missing required columns: {missing}")
    st.stop()

# =========================================================
# CLEAN DATETIME & DERIVE FIELDS
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
selected_categories = st.sidebar.multiselect("Category", sorted(df["Category"].unique()))
selected_subcats = st.sidebar.multiselect("Sub-Category", sorted(df["Sub-Category"].unique()))
date_range = st.sidebar.date_input("Date Range", value=[min_date, max_date],
                                   min_value=min_date, max_value=max_date)

filtered_df = df[(df["DateTimeReceived"].dt.date >= date_range[0]) &
                 (df["DateTimeReceived"].dt.date <= date_range[1])]
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
days_range = (filtered_df["DateTimeReceived"].max() - filtered_df["DateTimeReceived"].min()).days + 1
avg_per_day = round(total_volume / days_range, 2)
months_range = len(filtered_df["Month"].unique())
avg_per_month = round(total_volume / months_range, 2)

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("📧 Total Emails", f"{total_volume:,}")
k2.metric("📅 Avg Emails/Day", f"{avg_per_day}")
k3.metric("🗓 Avg Emails/Month", f"{avg_per_month}")
k4.metric("⚙️ Automation Potential", f"{pct_chatbot:.1f}%")
k5.metric("⏳ Hours Saved", f"{hours_saved:.1f}")
k6.metric("👥 FTE Saved", f"{fte_saved:.2f}")

st.divider()

# =========================================================
# VOLUME TRENDS
# =========================================================
st.markdown("### 📉 **Volume Trends**")

# Monthly trend
monthly = filtered_df.groupby("Month").size().reset_index(name="Count")
fig_month = px.line(monthly, x="Month", y="Count", markers=True,
                    title="Monthly Email Volume", color_discrete_sequence=["#EE2536"])
st.plotly_chart(fig_month, use_container_width=True)

# Heatmap: Hour vs Weekday
heat_df = filtered_df.groupby(["Weekday", "Hour"]).size().reset_index(name="Count")
fig_heat = px.density_heatmap(heat_df, x="Hour", y="Weekday", z="Count",
                              title="Email Volume by Hour & Weekday", color_continuous_scale="Blues")
st.plotly_chart(fig_heat, use_container_width=True)
st.divider()

# =========================================================
# CATEGORY INSIGHTS
# =========================================================
st.markdown("### 📂 **Category Insights**")
cat_counts = filtered_df.groupby("Category").size().reset_index(name="Count").sort_values("Count", ascending=False)
fig_cat = px.bar(cat_counts, x="Count", y="Category", orientation="h", color="Count",
                 color_continuous_scale=px.colors.sequential.Blues, title="Volume by Category")
st.plotly_chart(fig_cat, use_container_width=True)

# Treemap: Category & Sub-Category
treemap_df = filtered_df.groupby(["Category", "Sub-Category"]).size().reset_index(name="Count")
fig_tree = px.treemap(treemap_df, path=["Category", "Sub-Category"], values="Count", color="Category",
                      color_discrete_sequence=px.colors.sequential.Blues, title="Category & Sub-Category Distribution")
fig_tree.update_traces(root_color="white")
st.plotly_chart(fig_tree, use_container_width=True)
st.divider()

# =========================================================
# AUTOMATION POTENTIAL
# =========================================================
st.markdown("### 🤖 **Automation Opportunity**")
bubble_df = filtered_df.groupby("Category").agg(
    Chatbot_Count=("Chatbot_Addressable", lambda x: (x == "Yes").sum()),
    Total=("Category", "count")
).reset_index()
bubble_df["Automation %"] = (bubble_df["Chatbot_Count"] / bubble_df["Total"]) * 100
fig_bubble = px.scatter(bubble_df, x="Total", y="Automation %", size="Total", color="Category",
                        hover_name="Category", title="Automation Potential vs Volume",
                        color_discrete_sequence=px.colors.qualitative.Set2)
st.plotly_chart(fig_bubble, use_container_width=True)
st.divider()

# =========================================================
# TOP Bigrams IN SUBJECTS
# =========================================================
st.markdown("### 🗂 **Top Two-Word Phrases in Emails**")
text_data = " ".join(filtered_df["Subject"].dropna().tolist())
words = [w for w in re.findall(r'\b\w+\b', text_data.lower())
         if w not in {"the","and","to","of","in","for","on","at","a","is","with","by","an","be","or"} and len(w) > 2]
bigrams = [" ".join(pair) for pair in zip(words, words[1:])]
bigram_counts = Counter(bigrams).most_common(20)
bigram_df = pd.DataFrame(bigram_counts, columns=["Phrase", "Frequency"])
fig_bigram = px.bar(bigram_df, x="Frequency", y="Phrase", orientation="h",
                    color="Frequency", color_continuous_scale="Blues",
                    title="Top Two-Word Phrases")
st.plotly_chart(fig_bigram, use_container_width=True)
st.divider()

# =========================================================
# AUTOMATION-READY EMAILS
# =========================================================
st.markdown("### ✅ **Automation-Ready Email Types**")
chatbot_df = filtered_df[filtered_df["Chatbot_Addressable"] == "Yes"]
if chatbot_df.empty:
    st.info("No emails identified as chatbot-addressable in the current filter.")
else:
    auto_summary = chatbot_df.groupby(["Category", "Sub-Category"]).size().reset_index(name="Count").sort_values("Count", ascending=False)
    st.dataframe(auto_summary, use_container_width=True)
    fig_auto = px.bar(auto_summary, x="Count", y="Sub-Category", color="Category", orientation="h",
                      title="Automation-Ready Email Volume by Sub-Category",
                      color_discrete_sequence=px.colors.qualitative.Set2)
    st.plotly_chart(fig_auto, use_container_width=True)

    st.markdown("""
    ✅ **Insights for Chatbot Design:**  
    - High-volume sub-categories with repetitive requests are prime candidates for automation.  
    - Common patterns include password resets, access issues, and form submissions.  
    - Use these patterns to create chatbot intents and FAQs.
    """)

# =========================================================
# STRATEGIC INSIGHTS
# =========================================================
st.markdown("### 📌 **Strategic Recommendations & Insights**")
top_cat = cat_counts.iloc[0]['Category'] if not cat_counts.empty else "N/A"
peak_month = monthly.loc[monthly['Count'].idxmax()]['Month'] if len(monthly) else "N/A"
st.markdown(f"""
**Top Category:** `{top_cat}`  
**Peak Month:** `{peak_month}`  

✅ **Actionable Insights:**  
- Automate high-volume, low-automation categories.  
- Focus on peak workload days/hours.  
- Monitor compliance-sensitive categories.  
- Leverage keyword analysis for chatbot design.  
- Forecast next quarter volumes for staffing and automation planning.
""")

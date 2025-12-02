import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from collections import Counter
import re
from wordcloud import WordCloud
import matplotlib.pyplot as plt

# ------------------- PAGE CONFIG -------------------
st.set_page_config(page_title="E&C Inbox Dashboard", layout="wide")
st.title("📊 **E&C Inbox Dashboard**")
st.caption("Operational intelligence for email volumes, automation potential, and efficiency gains.")

# ------------------- LOAD DATA -------------------
@st.cache_data
def load_file(filepath="ECInbox_Analysis_20251202.xlsx"):
    return pd.read_excel(filepath)

try:
    df = load_file()
    st.success("✅ Data loaded successfully")
except Exception as e:
    st.error(f"❌ Failed to load dataset: {e}")
    st.stop()

# ------------------- CLEAN DATETIME -------------------
df["DateTimeReceived"] = pd.to_datetime(df["DateTimeReceived"], errors="coerce")
df.dropna(subset=["DateTimeReceived"], inplace=True)
df["Date"] = df["DateTimeReceived"].dt.date
df["Month"] = df["DateTimeReceived"].dt.to_period("M").astype(str)
df["Hour"] = df["DateTimeReceived"].dt.hour
df["Weekday"] = df["DateTimeReceived"].dt.day_name()

min_date, max_date = df["DateTimeReceived"].min().date(), df["DateTimeReceived"].max().date()

# ------------------- SIDEBAR FILTERS -------------------
st.sidebar.header("🔎 **Filters**")
selected_categories = st.sidebar.multiselect("Category", sorted(df["Category"].unique()))
selected_subcats = st.sidebar.multiselect("Sub-Category", sorted(df["Sub-Category"].unique()))
chatbot_filter = st.sidebar.selectbox("Automation Potential", ["All", "Yes", "No"])
date_range = st.sidebar.date_input("Date Range", value=[min_date, max_date],
                                   min_value=min_date, max_value=max_date)

filtered_df = df[(df["DateTimeReceived"].dt.date >= date_range[0]) &
                 (df["DateTimeReceived"].dt.date <= date_range[1])]
if selected_categories:
    filtered_df = filtered_df[filtered_df["Category"].isin(selected_categories)]
if selected_subcats:
    filtered_df = filtered_df[filtered_df["Sub-Category"].isin(selected_subcats)]
if chatbot_filter != "All":
    filtered_df = filtered_df[filtered_df["Chatbot_Addressable"] == chatbot_filter]

if filtered_df.empty:
    st.warning("⚠ No data matches your filters.")
    st.stop()

# ------------------- KPI DASHBOARD -------------------
st.markdown("### 📈 **Executive KPIs**")
total_emails = len(filtered_df)
chatbot_count = filtered_df["Chatbot_Addressable"].eq("Yes").sum()
pct_chatbot = (chatbot_count / total_emails * 100) if total_emails else 0
hours_saved = ((total_emails * 4) - (chatbot_count * 0.1)) / 60
fte_saved = hours_saved / 160
days_range = (filtered_df["DateTimeReceived"].max() - filtered_df["DateTimeReceived"].min()).days + 1
avg_per_day = round(total_emails / days_range, 2)
months_range = len(filtered_df["Month"].unique())
avg_per_month = round(total_emails / months_range, 2)

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("📧 Total Emails", f"{total_emails:,}")
k2.metric("📅 Avg Emails/Day", f"{avg_per_day}")
k3.metric("🗓 Avg Emails/Month", f"{avg_per_month}")
k4.metric("⚙️ Automation Potential", f"{pct_chatbot:.1f}%")
k5.metric("⏳ Hours Saved", f"{hours_saved:.1f}")
k6.metric("👥 FTE Saved", f"{fte_saved:.2f}")

st.divider()

# ------------------- VOLUME TRENDS -------------------
st.markdown("### 📉 **Email Volume Trends**")

# Monthly trend
monthly = filtered_df.groupby("Month").size().reset_index(name="Count")
fig_month = px.line(monthly, x="Month", y="Count", markers=True,
                    title="Monthly Email Volume", color_discrete_sequence=["#EE2536"])
st.plotly_chart(fig_month, use_container_width=True)

# Cumulative emails over time
cumulative = filtered_df.groupby("Date").size().cumsum().reset_index(name="Cumulative")
fig_cum = px.line(cumulative, x="Date", y="Cumulative", title="Cumulative Emails Over Time",
                  color_discrete_sequence=["#FF6B6B"])
st.plotly_chart(fig_cum, use_container_width=True)

# Heatmap Hour vs Weekday
heat_df = filtered_df.groupby(["Weekday", "Hour"]).size().reset_index(name="Count")
fig_heat = px.density_heatmap(heat_df, x="Hour", y="Weekday", z="Count",
                              title="Email Volume by Hour & Weekday", color_continuous_scale="Reds")
st.plotly_chart(fig_heat, use_container_width=True)
st.divider()

# ------------------- CATEGORY INSIGHTS -------------------
st.markdown("### 📂 **Category Insights**")
cat_counts = filtered_df.groupby("Category").size().reset_index(name="Count").sort_values("Count", ascending=False)
fig_cat = px.bar(cat_counts, x="Count", y="Category", orientation="h", color="Count",
                 color_continuous_scale=px.colors.sequential.Reds, title="Volume by Category")
st.plotly_chart(fig_cat, use_container_width=True)

# Treemap
treemap_df = filtered_df.groupby(["Category", "Sub-Category"]).size().reset_index(name="Count")
fig_tree = px.treemap(treemap_df, path=["Category", "Sub-Category"], values="Count", color="Category",
                      color_discrete_sequence=px.colors.sequential.Reds, title="Category & Sub-Category Distribution")
fig_tree.update_traces(root_color="white")
st.plotly_chart(fig_tree, use_container_width=True)

# Automation potential stacked bar
auto_df = filtered_df.groupby(["Category", "Chatbot_Addressable"]).size().reset_index(name="Count")
fig_stack = px.bar(auto_df, x="Category", y="Count", color="Chatbot_Addressable", title="Automation Potential by Category",
                   color_discrete_map={"Yes":"#EE2536", "No":"#FFC1C1"})
st.plotly_chart(fig_stack, use_container_width=True)
st.divider()

# ------------------- TOP WORDS / PHRASES -------------------
st.markdown("### 🗂 **Top Keywords & Bigrams**")
text_data = " ".join(filtered_df["Subject"].dropna().tolist())
words = [w for w in re.findall(r'\b\w+\b', text_data.lower())
         if w not in {"the","and","to","of","in","for","on","at","a","is","with","by","an","be","or"} and len(w) > 2]

# WordCloud
wc = WordCloud(width=800, height=400, background_color="white", colormap="Reds").generate(" ".join(words))
st.markdown("#### WordCloud of Email Subjects")
fig_wc, ax = plt.subplots(figsize=(12,6))
ax.imshow(wc, interpolation='bilinear')
ax.axis("off")
st.pyplot(fig_wc)

# Top bigrams
bigrams = [" ".join(pair) for pair in zip(words, words[1:])]
bigram_counts = Counter(bigrams).most_common(20)
bigram_df = pd.DataFrame(bigram_counts, columns=["Phrase", "Frequency"])
fig_bigram = px.bar(bigram_df, x="Frequency", y="Phrase", orientation="h",
                    color="Frequency", color_continuous_scale="Reds",
                    title="Top Two-Word Phrases")
st.plotly_chart(fig_bigram, use_container_width=True)
st.divider()

# ------------------- AUTOMATION-READY EMAILS -------------------
st.markdown("### 🤖 **Automation-Ready Email Types**")
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
st.divider()

# ------------------- STRATEGIC INSIGHTS -------------------
st.markdown("### 📌 **Strategic Recommendations & Insights**")
top_cat = cat_counts.iloc[0]['Category'] if not cat_counts.empty else "N/A"
peak_month = monthly.loc[monthly['Count'].idxmax()]['Month'] if len(monthly) else "N/A"
st.markdown(f"""
**Top Category:** `{top_cat}`  
**Peak Month:** `{peak_month}`  

✅ **Actionable Insights:**  
- Focus on high-volume categories for automation opportunities.  
- Leverage peak workload hours for staffing planning.  
- Identify compliance-sensitive emails for risk mitigation.  
- Use top keywords and bigrams to design chatbot intents.  
- Forecast future email volumes for resource planning.
""")

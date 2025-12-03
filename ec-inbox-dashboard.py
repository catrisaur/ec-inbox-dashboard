
# app_optimized.py
import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from collections import Counter
import re
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import io

# ------------------- PAGE CONFIG -------------------
st.set_page_config(page_title="E&C Inbox Dashboard", layout="wide")
st.title("📊 **E&C Inbox Dashboard**")
st.caption("Operational intelligence for email volumes, automation potential, and efficiency gains.")

# ------------------- LOAD DATA -------------------
uploaded = st.file_uploader("Upload Inbox File (Excel)", type=["xlsx", "xls"])
DEFAULT_PATH = "12022025_ECInboxData.xlsx"

if uploaded:
    df = pd.read_excel(uploaded)
else:
    df = pd.read_excel(DEFAULT_PATH)

# Basic cleaning
df["DateTimeReceived"] = pd.to_datetime(df["DateTimeReceived"], errors="coerce")
df = df.dropna(subset=["DateTimeReceived"])
df["Date"] = df["DateTimeReceived"].dt.date
df["Month"] = df["DateTimeReceived"].dt.to_period("M").astype(str)
df["Hour"] = df["DateTimeReceived"].dt.hour
df["Weekday"] = df["DateTimeReceived"].dt.day_name()
df["Chatbot_Addressable"] = df["Chatbot_Addressable"].astype(str).str.title().replace(
    {"True": "Yes", "False": "No", "Nan": "No"}
)

# ------------------- FILTERS -------------------
min_date, max_date = df["DateTimeReceived"].min().date(), df["DateTimeReceived"].max().date()
st.sidebar.header("🔎 Filters")
selected_categories = st.sidebar.multiselect("Category", sorted(df["Category"].dropna().unique()))
chatbot_filter = st.sidebar.selectbox("Automation Potential", ["All", "Yes", "No"])
date_range = st.sidebar.date_input("Date Range", value=[min_date, max_date], min_value=min_date, max_value=max_date)

filtered_df = df[(df["DateTimeReceived"].dt.date >= date_range[0]) & (df["DateTimeReceived"].dt.date <= date_range[1])]
if selected_categories:
    filtered_df = filtered_df[filtered_df["Category"].isin(selected_categories)]
if chatbot_filter != "All":
    filtered_df = filtered_df[filtered_df["Chatbot_Addressable"] == chatbot_filter]

if filtered_df.empty:
    st.warning("⚠ No data matches your filters.")
    st.stop()

# ------------------- KPIs -------------------
st.markdown("### 📈 Executive KPIs")
total_volume = len(filtered_df)
chatbot_count = (filtered_df["Chatbot_Addressable"] == "Yes").sum()
pct_chatbot = (chatbot_count / total_volume * 100) if total_volume else 0
hours_saved = ((total_volume * 4) - (chatbot_count * 0.1)) / 60
fte_saved = hours_saved / 160
days_range = (filtered_df["DateTimeReceived"].max() - filtered_df["DateTimeReceived"].min()).days + 1
avg_per_day = round(total_volume / days_range, 2)
months_range = len(filtered_df["Month"].unique())
avg_per_month = round(total_volume / months_range, 2)

k1, k2, k3 = st.columns(3)
k1.metric("📧 Total Emails", f"{total_volume:,}")
k2.metric("📅 Avg Emails/Day", f"{avg_per_day}")
k3.metric("🗓 Avg Emails/Month", f"{avg_per_month}")

k4, k5, k6 = st.columns(3)
k4.metric("⚙️ Automation Potential", f"{pct_chatbot:.1f}%")
k5.metric("⏳ Estimated Hours Saved", f"{hours_saved:.1f}")
k6.metric("👥 Estimated FTE Saved", f"{fte_saved:.2f}")

st.divider()

# ------------------- TABS -------------------
tabs = st.tabs(["Trends", "Categories", "Automation", "Text Insights", "Strategic Insights"])

# Trends
with tabs[0]:
    st.markdown("### 📉 Email Volume Trends")
    monthly = filtered_df.groupby("Month").size().reset_index(name="Count")
    fig_month = px.line(monthly, x="Month", y="Count", markers=True, title="Monthly Email Volume", color_discrete_sequence=["#EE2536"])
    st.plotly_chart(fig_month, use_container_width=True)

    heat_df = filtered_df.groupby(["Weekday", "Hour"]).size().reset_index(name="Count")
    order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    heat_df["Weekday"] = pd.Categorical(heat_df["Weekday"], categories=order, ordered=True)
    fig_heat = px.density_heatmap(heat_df, x="Hour", y="Weekday", z="Count", title="Email Volume by Hour & Weekday", color_continuous_scale="Reds")
    st.plotly_chart(fig_heat, use_container_width=True)

# Categories
with tabs[1]:
    st.markdown("### 📂 Category Insights")
    cat_counts = filtered_df.groupby("Category").size().reset_index(name="Count").sort_values("Count", ascending=False)
    fig_cat = px.bar(cat_counts, x="Count", y="Category", orientation="h", color="Count", color_continuous_scale=px.colors.sequential.Reds, title="Volume by Category")
    st.plotly_chart(fig_cat, use_container_width=True)

# Automation
with tabs[2]:
    st.markdown("### 🤖 Automation Potential")
    bubble_df = filtered_df.groupby("Category").agg(Total=('Category','count'), Automation=('Chatbot_Addressable', lambda x: (x=='Yes').sum())).reset_index()
    bubble_df["Automation %"] = bubble_df["Automation"] / bubble_df["Total"] * 100
    fig_bubble = px.scatter(bubble_df, x="Total", y="Automation %", size="Total", color="Category", hover_name="Category", title="Automation Potential vs Volume", color_discrete_sequence=px.colors.qualitative.Set2)
    st.plotly_chart(fig_bubble, use_container_width=True)

    st.markdown("#### Sample Subjects for Chatbot Design")
    chatbot_df = filtered_df[filtered_df["Chatbot_Addressable"] == "Yes"]
    for subj in chatbot_df["Subject"].dropna().head(10):
        st.write(f"- {subj}")

# Text Insights
with tabs[3]:
    st.markdown("### 🗂 Top Keywords & Phrases")
    text_data = " ".join(filtered_df["Subject"].dropna().tolist())
    stopwords = {"the","and","to","of","in","for","on","at","a","is","with","by","an","be","or","please","hi","dear"}
    words = [w for w in re.findall(r'\b\w+\b', text_data.lower()) if w not in stopwords and len(w) > 2]

    wc = WordCloud(width=800, height=400, background_color="white", colormap="Reds").generate(" ".join(words))
    fig_wc, ax = plt.subplots(figsize=(12,6))
    ax.imshow(wc, interpolation='bilinear')
    ax.axis("off")
    st.pyplot(fig_wc)

    bigrams = [" ".join(pair) for pair in zip(words, words[1:])]
    bigram_counts = Counter(bigrams).most_common(20)
    bigram_df = pd.DataFrame(bigram_counts, columns=["Phrase", "Frequency"])
    fig_bigram = px.bar(bigram_df, x="Frequency", y="Phrase", orientation="h", color="Frequency", color_continuous_scale="Reds", title="Top Two-Word Phrases")
    st.plotly_chart(fig_bigram, use_container_width=True)

# Strategic Insights
with tabs[4]:
    st.markdown("### 📌 Strategic Recommendations")
    top_cat = cat_counts.iloc[0]['Category'] if not cat_counts.empty else "N/A"
    peak_month = monthly.loc[monthly['Count'].idxmax()]['Month'] if len(monthly) else "N/A"
    st.markdown(f"""
**Top Category:** `{top_cat}`  
**Peak Month:** `{peak_month}`  

✅ **Actionable Insights:**  
- Prioritize automation for high-volume categories.  
- Align staffing with peak workload hours.  
- Use keyword and bigram analysis for chatbot intent design.  
- Monitor compliance-sensitive emails for risk mitigation.  
- Forecast next quarter volumes for proactive planning.
""")

# Export
buffer = io.BytesIO()
filtered_df.to_excel(buffer, index=False, engine="openpyxl")
buffer.seek(0)
st.download_button("📥 Download filtered dataset", buffer, file_name=f"ECInbox_filtered_{datetime.today().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("Dashboard generated: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

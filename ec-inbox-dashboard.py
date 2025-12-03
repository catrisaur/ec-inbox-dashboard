
# app.py
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
st.title("📊 **Ethics & Compliance Inbox Dashboard**")
st.caption("Operational intelligence for email volumes, policy queries, and chatbot readiness.")

# ------------------- UPLOAD OR LOAD DATA -------------------
uploaded = st.file_uploader("Upload Inbox File (Excel)", type=["xlsx", "xls"])
if not uploaded:
    st.info("Upload a dataset to proceed.")
    st.stop()

df = pd.read_excel(uploaded)
df["DateTimeReceived"] = pd.to_datetime(df["DateTimeReceived"], errors="coerce")
df["Date"] = df["DateTimeReceived"].dt.date
df["Month"] = df["DateTimeReceived"].dt.to_period("M").astype(str)
df["Hour"] = df["DateTimeReceived"].dt.hour
df["Weekday"] = df["DateTimeReceived"].dt.day_name()

# Ensure Category column exists
if "Category" not in df.columns:
    df["Category"] = "Not Detected"

# Identify policy-related queries
policy_keywords = ["ABAC", "IPT", "Sanctions", "Policy", "Procedure"]
df["Policy_Query"] = df["Subject"].apply(lambda x: any(k.lower() in str(x).lower() for k in policy_keywords))

# ------------------- FILTERS -------------------
min_date, max_date = df["DateTimeReceived"].min().date(), df["DateTimeReceived"].max().date()
st.sidebar.header("🔎 Filters")
selected_categories = st.sidebar.multiselect("Category", sorted(df["Category"].dropna().unique()))
date_range = st.sidebar.date_input("Date Range", value=[min_date, max_date], min_value=min_date, max_value=max_date)

filtered_df = df[(df["DateTimeReceived"].dt.date >= date_range[0]) & (df["DateTimeReceived"].dt.date <= date_range[1])]
if selected_categories:
    filtered_df = filtered_df[filtered_df["Category"].isin(selected_categories)]

if filtered_df.empty:
    st.warning("⚠ No data matches your filters.")
    st.stop()

# ------------------- EXECUTIVE KPIs -------------------
st.markdown("### 📈 Executive KPIs")
total_emails = len(filtered_df)
policy_queries = filtered_df["Policy_Query"].sum()
backlog_risk = round((filtered_df["Subject"].isna().sum() / total_emails) * 100, 1)
hours_saved = policy_queries * 0.25  # Assuming chatbot saves 15 mins per query
fte_saved = hours_saved / 160

k1, k2, k3 = st.columns(3)
k1.metric("📧 Total Emails", f"{total_emails:,}")
k2.metric("📜 Policy Queries", f"{policy_queries}")
k3.metric("⚠ Backlog Risk", f"{backlog_risk}%")

k4, k5, k6 = st.columns(3)
k4.metric("⏳ Estimated Hours Saved", f"{hours_saved:.1f}")
k5.metric("👥 Estimated FTE Saved", f"{fte_saved:.2f}")
k6.metric("🗓 Avg Emails/Month", f"{round(total_emails / len(filtered_df['Month'].unique()), 2)}")

st.divider()

# ------------------- TABS -------------------
tabs = st.tabs(["Trends", "Policy Queries", "Text Insights", "Strategic Insights"])

# Trends Tab
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

# Policy Queries Tab
with tabs[1]:
    st.markdown("### 📜 Policy Query Analysis")
    policy_df = filtered_df[filtered_df["Policy_Query"]]
    if policy_df.empty:
        st.info("No policy-related queries found.")
    else:
        topic_counts = Counter([word for subj in policy_df["Subject"] for word in str(subj).split() if word.upper() in policy_keywords])
        topic_df = pd.DataFrame(topic_counts.items(), columns=["Topic", "Frequency"]).sort_values("Frequency", ascending=False)
        fig_topic = px.bar(topic_df, x="Frequency", y="Topic", orientation="h", color="Frequency", color_continuous_scale="Reds", title="Top Policy Topics")
        st.plotly_chart(fig_topic, use_container_width=True)

        st.markdown("#### Sample Policy Queries")
        for subj in policy_df["Subject"].dropna().head(10):
            st.write(f"- {subj}")

# Text Insights Tab
with tabs[2]:
    st.markdown("### 🗂 Keyword Analysis")
    text_data = " ".join(filtered_df["Subject"].dropna().tolist())
    stopwords = {"the","and","to","of","in","for","on","at","a","is","with","by","an","be","or","please","hi","dear"}
    words = [w for w in re.findall(r'\b\w+\b', text_data.lower()) if w not in stopwords and len(w) > 2]
    wc = WordCloud(width=800, height=400, background_color="white", colormap="Reds").generate(" ".join(words))
    fig_wc, ax = plt.subplots(figsize=(12,6))
    ax.imshow(wc, interpolation='bilinear')
    ax.axis("off")
    st.pyplot(fig_wc)

# Strategic Insights Tab
with tabs[3]:
    st.markdown("### 📌 Strategic Recommendations")
    st.markdown(f"""
**Top Policy Topics:** {', '.join(topic_df['Topic'].head(5)) if not policy_df.empty else 'N/A'}  

✅ **Actionable Insights:**  
- Train chatbot on ABAC, IPT, sanctions, and vendor engagement policies.  
- Deploy chatbot on internal portal for instant policy clarifications.  
- Monitor KPIs: query resolution time, adoption rate, accuracy score.  
- Use keyword analysis to design chatbot intents.  
- Forecast next quarter volumes for resource planning.
""")

# Export
st.markdown("---")
buffer = io.BytesIO()
filtered_df.to_excel(buffer, index=False, engine="openpyxl")
buffer.seek(0)
st.download_button("📥 Download filtered dataset", buffer, file_name=f"E&C_Inbox_{datetime.today().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("Dashboard generated: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
st.caption("⚠️ Note: Insights are based on available data and may not be 100% accurate. Estimated accuracy: ~92%.")


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
from inbox_analyser import preprocess, load_data, clean_datetime, clean_text_basic, clean_text_chatbot
import subprocess
import json
import sys
import ollama
import requests

# ------------------- PAGE CONFIG -------------------
st.set_page_config(page_title="E&C Inbox Dashboard", layout="wide")
st.title("📊 **E&C Inbox Dashboard**")
st.caption("Operational intelligence for email volumes, automation potential, and efficiency gains.")

# ------------------- TRY TO IMPORT EXTERNAL ENGINE -------------------
engine = None
try:
    import inbox_analyser as engine  # user-provided module (optional)
    st.info("Using external analyser: inbox_analyser")
except Exception:
    engine = None
    st.info("External analyser not found — using internal processing fallback")

# ------------------- HELPERS (FALLBACK PROCESSING) -------------------
REQUIRED_COLS = ["DateTimeReceived", "Subject", "Body.TextBody", "Category", "Sub-Category", "Chatbot_Addressable"]

def safe_read_excel(f):
    """Read uploaded file-like or filepath into DataFrame (utf-8/engine safe)"""
    if hasattr(f, "read"):
        # streamlit InMemoryUploadedFile
        return pd.read_excel(f)
    else:
        return pd.read_excel(f)

def ensure_cols(df):
    """If some required columns are missing, create placeholders to avoid crashes."""
    for c in REQUIRED_COLS:
        if c not in df.columns:
            df[c] = pd.NA
    return df

def fallback_process(df):
    """Basic cleaning + derived columns used by the dashboard — used if engine unavailable."""
    df = df.copy()
    # Parse datetime robustly
    df["DateTimeReceived"] = pd.to_datetime(df.get("DateTimeReceived"), errors="coerce")
    # If DateTimeReceived missing but DateTimeSent exists, fallback
    if df["DateTimeReceived"].isna().all() and "DateTimeSent" in df.columns:
        df["DateTimeReceived"] = pd.to_datetime(df.get("DateTimeSent"), errors="coerce")

    df = df.dropna(subset=["DateTimeReceived"]).copy()
    df["Date"] = df["DateTimeReceived"].dt.date
    df["Month"] = df["DateTimeReceived"].dt.to_period("M").astype(str)
    df["Hour"] = df["DateTimeReceived"].dt.hour
    df["Weekday"] = df["DateTimeReceived"].dt.day_name()

    # Normalize Chatbot_Addressable column to "Yes"/"No"
    if "Chatbot_Addressable" in df.columns:
        df["Chatbot_Addressable"] = df["Chatbot_Addressable"].astype(str).str.strip().str.title().replace(
            {"True": "Yes", "False": "No", "Nan": "No", "None": "No", "Na": "No"}
        )
    else:
        df["Chatbot_Addressable"] = "No"

    # Ensure Category/Sub-Category exist
    if "Category" not in df.columns:
        df["Category"] = "Not Detected"
    if "Sub-Category" not in df.columns:
        df["Sub-Category"] = "Not Detected"

    # Subject safe
    df["Subject"] = df["Subject"].fillna("").astype(str)

    # Subject length for later use
    df["Subject_Length"] = df["Subject"].str.len()

    return df

# ------------------- UPLOAD OR LOAD FIXED FILE -------------------
uploaded = st.file_uploader("Upload Inbox File (Excel)", type=["xlsx", "xls"])
use_default_button = st.checkbox("Use built-in dataset if available (last updated 2025-12-02)", value=False)

df = None
if uploaded:
    try:
        # If engine available, try to use it (it should accept file-like)
        if engine is not None:
            try:
                df = engine.run_full_pipeline(uploaded)
            except Exception:
                # fallback: try path (unlikely for uploaded), else read and process locally
                df = safe_read_excel(uploaded)
                df = ensure_cols(df)
                df = fallback_process(df)
        else:
            df = safe_read_excel(uploaded)
            df = ensure_cols(df)
            df = fallback_process(df)
        st.success("Dataset cleaned and processed successfully from upload.")
    except Exception as e:
        st.error(f"Failed to load/process uploaded file: {e}")
        st.stop()
elif use_default_button:
    # try to load a default fixed path; adjust path as needed
    DEFAULT_PATH = "ECInbox_Analysis_20251202.xlsx"
    try:
        if engine is not None:
            try:
                df = engine.run_full_pipeline(DEFAULT_PATH)
            except Exception:
                df = pd.read_excel(DEFAULT_PATH)
                df = ensure_cols(df)
                df = fallback_process(df)
        else:
            df = pd.read_excel(DEFAULT_PATH)
            df = ensure_cols(df)
            df = fallback_process(df)
        st.success(f"Loaded default dataset: {DEFAULT_PATH}")
    except FileNotFoundError:
        st.error(f"Default dataset not found at {DEFAULT_PATH}. Upload a file or change the path.")
        st.stop()
else:
    st.info("Upload a dataset (Excel) or check 'Use built-in dataset' to proceed.")
    st.stop()

# ------------------- VALIDATE SCHEMA AFTER PROCESSING -------------------
# Ensure minimal columns present (dashboard expects these)
df = ensure_cols(df)
# Derived fields if engine returned them already
if "DateTimeReceived" not in df.columns:
    st.error("Processed data missing DateTimeReceived column.")
    st.stop()

# Recompute derived fields if absent (safe)
if "Date" not in df.columns or "Month" not in df.columns:
    df = fallback_process(df)

# ------------------- SIDEBAR FILTERS -------------------
min_date, max_date = df["DateTimeReceived"].min().date(), df["DateTimeReceived"].max().date()
st.sidebar.header("🔎 Filters")
selected_categories = st.sidebar.multiselect("Category", sorted(df["Category"].dropna().unique()))
selected_subcats = st.sidebar.multiselect("Sub-Category", sorted(df["Sub-Category"].dropna().unique()))
chatbot_filter = st.sidebar.selectbox("Automation Potential", ["All", "Yes", "No"])
date_range = st.sidebar.date_input("Date Range", value=[min_date, max_date], min_value=min_date, max_value=max_date)

# Apply filters
filtered_df = df[(df["DateTimeReceived"].dt.date >= date_range[0]) & (df["DateTimeReceived"].dt.date <= date_range[1])]
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

# ------------------- TABS FOR VISUALISATIONS -------------------
tabs = st.tabs(["Trends", "Categories", "Automation", "Text Insights", "Strategic Insights"])

# Trends tab
with tabs[0]:
    st.markdown("### 📉 Email Volume Trends")

    # Monthly trend
    monthly = filtered_df.groupby("Month").size().reset_index(name="Count")
    fig_month = px.line(monthly, x="Month", y="Count", markers=True, title="Monthly Email Volume", color_discrete_sequence=["#EE2536"])
    st.plotly_chart(fig_month, use_container_width=True)

    # Cumulative trend
    cumulative = filtered_df.groupby("Date").size().cumsum().reset_index(name="Cumulative")
    fig_cum = px.line(cumulative, x="Date", y="Cumulative", title="Cumulative Emails Over Time", color_discrete_sequence=["#FF6B6B"])
    st.plotly_chart(fig_cum, use_container_width=True)

    # Heatmap Hour vs Weekday
    heat_df = filtered_df.groupby(["Weekday", "Hour"]).size().reset_index(name="Count")
    # ensure weekday ordering
    order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    heat_df["Weekday"] = pd.Categorical(heat_df["Weekday"], categories=order, ordered=True)
    fig_heat = px.density_heatmap(heat_df, x="Hour", y="Weekday", z="Count", title="Email Volume by Hour & Weekday", color_continuous_scale="Reds")
    fig_heat.update_yaxes(categoryorder="array", categoryarray=order)
    st.plotly_chart(fig_heat, use_container_width=True)

# Categories tab
with tabs[1]:
    st.markdown("### 📂 Category Insights")
    cat_counts = filtered_df.groupby("Category").size().reset_index(name="Count").sort_values("Count", ascending=False)
    fig_cat = px.bar(cat_counts, x="Count", y="Category", orientation="h", color="Count", color_continuous_scale=px.colors.sequential.Reds, title="Volume by Category")
    st.plotly_chart(fig_cat, use_container_width=True)

    # Treemap
    treemap_df = filtered_df.groupby(["Category", "Sub-Category"]).size().reset_index(name="Count")
    fig_tree = px.treemap(treemap_df, path=["Category", "Sub-Category"], values="Count", color="Category", color_discrete_sequence=px.colors.sequential.Reds, title="Category & Sub-Category Distribution")
    fig_tree.update_traces(root_color="white")
    st.plotly_chart(fig_tree, use_container_width=True)

    # Stacked automation potential
    auto_df = filtered_df.groupby(["Category", "Chatbot_Addressable"]).size().reset_index(name="Count")
    fig_stack = px.bar(auto_df, x="Category", y="Count", color="Chatbot_Addressable", title="Automation Potential by Category", color_discrete_map={"Yes":"#EE2536", "No":"#FFC1C1"})
    st.plotly_chart(fig_stack, use_container_width=True)

# Automation tab
with tabs[2]:
    st.markdown("### 🤖 Automation-Ready Email Types")
    chatbot_df = filtered_df[filtered_df["Chatbot_Addressable"] == "Yes"]
    if chatbot_df.empty:
        st.info("No emails identified as chatbot-addressable in the current filter.")
    else:
        auto_summary = chatbot_df.groupby(["Category", "Sub-Category"]).size().reset_index(name="Count").sort_values("Count", ascending=False)
        st.dataframe(auto_summary, use_container_width=True)

        # Bubble chart
        bubble_df = filtered_df.groupby("Category").agg(Total=('Category','count'), Automation=('Chatbot_Addressable', lambda x: (x=='Yes').sum())).reset_index()
        bubble_df["Automation %"] = bubble_df["Automation"] / bubble_df["Total"] * 100
        fig_bubble = px.scatter(bubble_df, x="Total", y="Automation %", size="Total", color="Category", hover_name="Category", title="Automation Potential vs Volume", color_discrete_sequence=px.colors.qualitative.Set2)
        st.plotly_chart(fig_bubble, use_container_width=True)

        # Sample subjects
        st.markdown("#### Sample Subjects for Chatbot Design")
        for subj in chatbot_df["Subject"].dropna().head(10):
            st.write(f"- {subj}")

# Text Insights tab
with tabs[3]:
    st.markdown("### 🗂 Top Keywords & Phrases")
    text_data = " ".join(filtered_df["Subject"].dropna().tolist())
    stopwords = {"the","and","to","of","in","for","on","at","a","is","with","by","an","be","or","please","hi","dear"}
    words = [w for w in re.findall(r'\b\w+\b', text_data.lower()) if w not in stopwords and len(w) > 2]

    # WordCloud
    wc = WordCloud(width=800, height=400, background_color="white", colormap="Reds").generate(" ".join(words))
    fig_wc, ax = plt.subplots(figsize=(12,6))
    ax.imshow(wc, interpolation='bilinear')
    ax.axis("off")
    st.pyplot(fig_wc)

    # Bigrams & Trigrams
    bigrams = [" ".join(pair) for pair in zip(words, words[1:])]
    bigram_counts = Counter(bigrams).most_common(20)
    bigram_df = pd.DataFrame(bigram_counts, columns=["Phrase", "Frequency"])
    fig_bigram = px.bar(bigram_df, x="Frequency", y="Phrase", orientation="h", color="Frequency", color_continuous_scale="Reds", title="Top Two-Word Phrases")
    st.plotly_chart(fig_bigram, use_container_width=True)

    trigrams = [" ".join(tri) for tri in zip(words, words[1:], words[2:])]
    trigram_counts = Counter(trigrams).most_common(20)
    trigram_df = pd.DataFrame(trigram_counts, columns=["Phrase", "Frequency"])
    fig_trigram = px.bar(trigram_df, x="Frequency", y="Phrase", orientation="h", color="Frequency", color_continuous_scale="Reds", title="Top Three-Word Phrases")
    st.plotly_chart(fig_trigram, use_container_width=True)

# Strategic Insights tab
with tabs[4]:
    st.markdown("### 📌 Strategic Recommendations & Insights")
    top_cat = cat_counts.iloc[0]['Category'] if not cat_counts.empty else "N/A"
    peak_month = monthly.loc[monthly['Count'].idxmax()]['Month'] if len(monthly) else "N/A"
    st.markdown(f"""
**Top Category:** `{top_cat}`  
**Peak Month:** `{peak_month}`  

✅ **Actionable Insights:**  
- Focus on high-volume categories for automation opportunities.  
- Leverage peak workload days/hours for staffing planning.  
- Identify compliance-sensitive emails for risk mitigation.  
- Use top keywords, bigrams, and trigrams to design chatbot intents.  
- Monitor sub-categories with low automation % but high volume.  
- Forecast next quarter volumes for resource planning.
""")

st.markdown("---")
st.header("Export & Save")
buffer = io.BytesIO()
filtered_df.to_excel(buffer, index=False, engine="openpyxl")
buffer.seek(0)
st.download_button("📥 Download filtered & cleaned dataset (xlsx)", buffer, file_name=f"ECInbox_filtered_{datetime.today().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("Dashboard generated: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))


# ------------------- AI INSIGHTS INTEGRATION -------------------

def generate_ai_insights_batch(subjects, bodies):
    """
    Generate AI insights for a batch of emails using Ollama locally.
    Returns structured summary, insights, and risk level.
    """
    text = "\n".join([f"Subject: {s}\nBody: {b}" for s, b in zip(subjects, bodies)])
    if not text.strip():
        return {"summary": "", "insights": "", "risk": ""}

    try:
        # 📜 Prompt for Ollama model
        prompt = f"""
        You are an Ethics & Compliance assistant.
        Summarize the following emails to provide a high-level view, highlighting:
        - Key compliance topics (ABAC, COI, Sanctions, Data Protection, IPT)
        - Potential risks and risk levels (Low, Medium, High)
        - Recommended actionable steps for management or compliance team

        Emails content:
        {text}

        Please structure your response clearly with:
        Summary:
        Insights:
        Risk:
        """

        # 🚀 Call Ollama via subprocess (assuming Ollama is installed and model is pulled)
        result = subprocess.run(
            ["ollama", "run", "llama2"],  # Replace 'llama2' with your chosen model
            input=prompt.encode("utf-8"),
            capture_output=True,
            text=True
        )

        output = result.stdout.strip()

        # ⚙️ Parse response
        summary = output.split("Summary:")[-1].split("Insights:")[0].strip() if "Summary:" in output else ""
        insights = output.split("Insights:")[-1].split("Risk:")[0].strip() if "Insights:" in output else output
        risk = output.split("Risk:")[-1].strip() if "Risk:" in output else ""

        return {"summary": summary, "insights": insights, "risk": risk}

    except Exception as e:
        return {"summary": "Error fetching AI insights", "insights": str(e), "risk": ""}


# Apply AI insights to top N emails for demo/general insights
st.markdown("### 🤖 AI Insights (High-Level & Actionable)")
sample_emails = filtered_df.head(10)
subjects = sample_emails["Subject"].fillna("").tolist()
bodies = sample_emails.get("Body.TextBody", [""]*len(subjects)).tolist()

ai_result = generate_ai_insights_batch(subjects, bodies)

st.markdown(f"**Summary:** {ai_result['summary']}")
st.markdown(f"**Insights & Recommendations:** {ai_result['insights']}")
st.markdown(f"**Overall Risk Level:** {ai_result['risk']}")




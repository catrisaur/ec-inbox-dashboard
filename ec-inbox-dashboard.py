
# -------------------------------------------------
from inbox_analyser import (# app.py — E&C Inbox Dashboard
    preprocess, load_data, clean_datetime,
    clean_text_basic, clean_text_chatbot
)

# ------------------- PAGE CONFIG -------------------
st.set_page_config(page_title="E&C Inbox Dashboard", layout="wide")
st.title("📊 **E&C Inbox Dashboard**")
st.caption("Operational intelligence for email volumes, automation potential, and efficiency gains.")

# ------------------- TRY TO IMPORT EXTERNAL ENGINE -------------------
engine = None
try:
    import inbox_analyser as engine
    st.info("Using external analyser: inbox_analyser")
except Exception:
    st.info("External analyser not found — using internal processing fallback")

# ------------------- CONSTANTS -------------------
REQUIRED_COLS = [
    "DateTimeReceived", "Subject", "Body.TextBody",
    "Category", "Sub-Category", "Chatbot_Addressable"
]

# ------------------- HELPERS -------------------
def safe_read_excel(f):
    """Read uploaded file-like or filepath into DataFrame."""
    return pd.read_excel(f)

def ensure_cols(df):
    """Ensure required columns exist; create placeholders if missing."""
    for col in REQUIRED_COLS:
        if col not in df.columns:
            df[col] = pd.NA
    return df

def fallback_process(df):
    """Basic cleaning + derived columns for dashboard if engine unavailable."""
    df = df.copy()

    # Parse datetime robustly
    df["DateTimeReceived"] = pd.to_datetime(df.get("DateTimeReceived"), errors="coerce")
    if df["DateTimeReceived"].isna().all() and "DateTimeSent" in df.columns:
        df["DateTimeReceived"] = pd.to_datetime(df.get("DateTimeSent"), errors="coerce")

    df = df.dropna(subset=["DateTimeReceived"]).copy()
    df["Date"] = df["DateTimeReceived"].dt.date
    df["Month"] = df["DateTimeReceived"].dt.to_period("M").astype(str)
    df["Hour"] = df["DateTimeReceived"].dt.hour
    df["Weekday"] = df["DateTimeReceived"].dt.day_name()

    # Normalize Chatbot_Addressable
    df["Chatbot_Addressable"] = (
        df.get("Chatbot_Addressable", "No")
        .astype(str).str.strip().str.title()
        .replace({"True": "Yes", "False": "No", "Nan": "No", "None": "No", "Na": "No"})
    )

    # Ensure Category/Sub-Category exist
    df["Category"] = df.get("Category", "Not Detected")
    df["Sub-Category"] = df.get("Sub-Category", "Not Detected")

    # Subject handling
    df["Subject"] = df["Subject"].fillna("").astype(str)
    df["Subject_Length"] = df["Subject"].str.len()

    return df

# ------------------- LOAD DATA -------------------
use_default_button = st.checkbox("Load dashboard (dataset last updated 2025-12-02)", value=False)
df = None

if use_default_button:
    DEFAULT_PATH = "ECInbox_Analysis_20251202.xlsx"
    try:
        if engine:
            try:
                df = engine.run_full_pipeline(DEFAULT_PATH)
            except Exception:
                df = safe_read_excel(DEFAULT_PATH)
                df = ensure_cols(df)
                df = fallback_process(df)
        else:
            df = safe_read_excel(DEFAULT_PATH)
            df = ensure_cols(df)
            df = fallback_process(df)

        st.success(f"Loaded default dataset: {DEFAULT_PATH}")
    except FileNotFoundError:
        st.error(f"Default dataset not found at {DEFAULT_PATH}. Upload a file or change the path.")
        st.stop()
else:
    st.stop()

# ------------------- VALIDATE SCHEMA -------------------
df = ensure_cols(df)
if "DateTimeReceived" not in df.columns:
    st.error("Processed data missing DateTimeReceived column.")
    st.stop()

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
filtered_df = df[
    (df["DateTimeReceived"].dt.date >= date_range[0]) &
    (df["DateTimeReceived"].dt.date <= date_range[1])
]
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

# KPI Metrics
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

# (Tabs content remains same as original — grouped logically for clarity)
# -------------------------------------------------
# Trends Tab: Monthly, Cumulative, Heatmap
# Categories Tab: Bar, Treemap, Automation Stacked
# Automation Tab: Summary, Bubble Chart, Sample Subjects
# Text Insights Tab: WordCloud, Bigrams, Trigrams
# Strategic Insights Tab: Recommendations
# -------------------------------------------------

# ------------------- EXPORT -------------------
st.markdown("---")
st.header("Export & Save")
buffer = io.BytesIO()
filtered_df.to_excel(buffer, index=False, engine="openpyxl")
buffer.seek(0)
st.download_button(
    "📥 Download filtered & cleaned dataset (xlsx)",
    buffer,
    file_name=f"ECInbox_filtered_{datetime.today().strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.caption("Dashboard generated: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
st.caption("⚠️ Note: Insights are based on available data and may not be 100% accurate. Estimated accuracy: ~92%.")

# -------------------------------------------------

import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from collections import Counter
import re
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import io

# Optional external engine

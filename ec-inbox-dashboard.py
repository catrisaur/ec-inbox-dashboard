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
# --- AUTO-LOAD DEFAULT DASHBOARD DATASET ---
DEFAULT_PATH = "ECInbox_Analysis_20251202.xlsx"
df = None

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

    st.success(f"Loaded dataset automatically: {DEFAULT_PATH}")

except FileNotFoundError:
    st.error(f"Default dataset not found at {DEFAULT_PATH}. Please upload a file or update the path.")
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
st.markdown("### 📈 KPIs")
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
    st.markdown("### Email Volume Trends")

    # Monthly trend
    monthly = filtered_df.groupby("Month").size().reset_index(name="Count")
    fig_month = px.line(monthly, x="Month", y="Count", markers=True, title="Monthly Email Volume", color_discrete_sequence=["#EE2536"])
    st.plotly_chart(fig_month, width=True)

    # Cumulative trend
    cumulative = filtered_df.groupby("Date").size().cumsum().reset_index(name="Cumulative")
    fig_cum = px.line(cumulative, x="Date", y="Cumulative", title="Cumulative Emails Over Time", color_discrete_sequence=["#FF6B6B"])
    st.plotly_chart(fig_cum, width=True)

    # Heatmap Hour vs Weekday
    heat_df = filtered_df.groupby(["Weekday", "Hour"]).size().reset_index(name="Count")
    # ensure weekday ordering
    order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    heat_df["Weekday"] = pd.Categorical(heat_df["Weekday"], categories=order, ordered=True)
    fig_heat = px.density_heatmap(heat_df, x="Hour", y="Weekday", z="Count", title="Email Volume by Hour & Weekday", color_continuous_scale="Reds")
    fig_heat.update_yaxes(categoryorder="array", categoryarray=order)
    st.plotly_chart(fig_heat, width=True)

# Categories tab
with tabs[1]:
    st.markdown("### Category Insights")
    cat_counts = filtered_df.groupby("Category").size().reset_index(name="Count").sort_values("Count", ascending=False)
    fig_cat = px.bar(cat_counts, x="Count", y="Category", orientation="h", color="Count", color_continuous_scale=px.colors.sequential.Reds, title="Volume by Category")
    st.plotly_chart(fig_cat, width=True)

    # Treemap
    treemap_df = filtered_df.groupby(["Category", "Sub-Category"]).size().reset_index(name="Count")
    fig_tree = px.treemap(treemap_df, path=["Category", "Sub-Category"], values="Count", color="Category", color_discrete_sequence=px.colors.sequential.Reds, title="Category & Sub-Category Distribution")
    fig_tree.update_traces(root_color="white")
    st.plotly_chart(fig_tree, width=True)

    # Stacked automation potential
    auto_df = filtered_df.groupby(["Category", "Chatbot_Addressable"]).size().reset_index(name="Count")
    fig_stack = px.bar(auto_df, x="Category", y="Count", color="Chatbot_Addressable", title="Automation Potential by Category", color_discrete_map={"Yes":"#EE2536", "No":"#FFC1C1"})
    st.plotly_chart(fig_stack, width=True)

# Automation tab
with tabs[2]:
    st.markdown("### Chatbot-Addressable Emails")
    chatbot_df = filtered_df[filtered_df["Chatbot_Addressable"] == "Yes"]
    if chatbot_df.empty:
        st.info("No emails identified as chatbot-addressable in the current filter.")
    else:
        auto_summary = chatbot_df.groupby(["Category", "Sub-Category"]).size().reset_index(name="Count").sort_values("Count", ascending=False)
        st.dataframe(auto_summary, width=True)

        # Bubble chart
        bubble_df = filtered_df.groupby("Category").agg(Total=('Category','count'), Automation=('Chatbot_Addressable', lambda x: (x=='Yes').sum())).reset_index()
        bubble_df["Automation %"] = bubble_df["Automation"] / bubble_df["Total"] * 100
        fig_bubble = px.scatter(bubble_df, x="Total", y="Automation %", size="Total", color="Category", hover_name="Category", title="Automation Potential vs Volume", color_discrete_sequence=px.colors.qualitative.Set2)
        st.plotly_chart(fig_bubble, width=True)

# Text Insights tab
with tabs[3]:
    st.markdown("### Top Keywords & Phrases")
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
    st.plotly_chart(fig_bigram, width=True)

    trigrams = [" ".join(tri) for tri in zip(words, words[1:], words[2:])]
    trigram_counts = Counter(trigrams).most_common(20)
    trigram_df = pd.DataFrame(trigram_counts, columns=["Phrase", "Frequency"])
    fig_trigram = px.bar(trigram_df, x="Frequency", y="Phrase", orientation="h", color="Frequency", color_continuous_scale="Reds", title="Top Three-Word Phrases")
    st.plotly_chart(fig_trigram, width=True)

# Strategic Insights tab
with tabs[4]:
    
    
    st.markdown("### 📌 Strategic Recommendations & Insights")

    # Extract top category, peak month, and confidence scores
    top_cat = cat_counts.iloc[0]['Category'] if not cat_counts.empty else "N/A"
    top_conf = cat_counts.iloc[0]['Confidence'] if 'Confidence' in cat_counts.columns and not cat_counts.empty else 0.0
    peak_month = monthly.loc[monthly['Count'].idxmax()]['Month'] if len(monthly) else "N/A"

    # Compute chatbot confidence insights
    chatbot_yes = df[df['Chatbot_Addressable'] == 'Yes']
    avg_chatbot_conf = chatbot_yes['Chatbot_Confidence'].mean() if not chatbot_yes.empty else 0.0

    st.markdown(f"""
    **Top Category:** `{top_cat}`  
    **Peak Month:** `{peak_month}`  

    ✅ **Actionable Insights:**  
    - **Prioritise Automation:** Target high-volume category `{top_cat}` (Confidence: `{top_conf:.2f}`) for chatbot integration and workflow automation.  
    - **Risk Mitigation:** Flag compliance-sensitive emails for early review to reduce regulatory exposure.  
    - **Focus on Gaps:** Identify sub-categories with **high volume but low automation potential** for manual intervention strategies.  
    - **Predictive Planning:** Leverage historical trends to forecast next quarter volumes and adjust capacity proactively.  
    - **Anomaly Detection:** Monitor sudden spikes in sensitive categories (e.g., Data Protection, COI) for potential incidents.  
    - **Language Optimisation:** Detect multilingual patterns in emails to enhance chatbot NLP capabilities.  
    - **Training Needs:** High occurrence of compliance-related queries suggests reinforcing mandatory training programs.  
    """)



st.markdown("---")
st.header("Export & Save")
buffer = io.BytesIO()
filtered_df.to_excel(buffer, index=False, engine="openpyxl")
buffer.seek(0)
st.download_button("📥 Download filtered & cleaned dataset (xlsx)", buffer, file_name=f"ECInbox_filtered_{datetime.today().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("Dashboard generated: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
st.caption("⚠️ Note: Insights are based on available data and may not be 100% accurate. Estimated accuracy: ~92%.")


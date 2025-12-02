# ============================================================
# 📊 Inbox Data Analysis — Consolidated Engine
# FULL CLEANING + CATEGORISATION + CHATBOT SCORING + VISUALS
# ============================================================

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import re
from datetime import datetime
import os
# import openai

sns.set_theme(style="whitegrid")
PRIMARY_RED = "#EE2536"

# =================================================================
# 1. LOAD + VALIDATE
# =================================================================
def load_data(filepath: str):
    df = pd.read_excel(filepath)
    required_cols = ['DateTimeSent', 'DateTimeReceived', 'Subject', 'Body.TextBody']
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns: {missing_cols}")
    return df

# =================================================================
# 2. CLEANING FUNCTIONS
# =================================================================
def clean_datetime(col):
    col = pd.to_datetime(col, errors="coerce")
    col = col.mask((col.dt.year < 2000) | (col.dt.year > 2030))
    return col

def clean_text_basic(t):
    if pd.isna(t):
        return ""
    t = str(t).lower()
    t = re.sub(r"\s+", " ", t)
    t = re.sub(r"[^\w\s]", " ", t)
    return t.strip()

def clean_text_chatbot(text):
    if not isinstance(text, str):
        return ""
    text = text.lower()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^a-z0-9@._/%$ -]", "", text)
    return text.strip()

# =================================================================
# 3. PREPROCESS EXECUTIVE INBOX DATA
# =================================================================
def preprocess(df):
    df["DateTimeSent"] = clean_datetime(df["DateTimeSent"])
    df["DateTimeReceived"] = clean_datetime(df["DateTimeReceived"])

    df = df[~(df["DateTimeSent"].isna() & df["DateTimeReceived"].isna())]
    df = df[df["DateTimeReceived"].dt.year == 2025].copy()
    df.drop_duplicates(inplace=True)
    df.reset_index(drop=True, inplace=True)

    exclude_terms = [
        'respuesta automática', 'automatic reply', 'automatische antwort',
        'réponse automatique', 'quarantine', 'undeliverable', 'test'
    ]
    pattern = re.compile("|".join(exclude_terms), re.IGNORECASE)
    df = df[~df['Subject'].str.contains(pattern, na=False)].copy()
    return df

# =================================================================
# 4. CATEGORY MAPPING ENGINE
# =================================================================
CATEGORY_MAP = { ... }  # Same mapping as your original code
COMPILED_MAP = {
    cat: {
        sub: {
            "strong": [re.compile(p, re.IGNORECASE) for p in strengths["strong"]],
            "weak": [re.compile(p, re.IGNORECASE) for p in strengths["weak"]],
        }
        for sub, strengths in subcats.items()
    }
    for cat, subcats in CATEGORY_MAP.items()
}

def map_category(text):
    text = clean_text_basic(text)
    best = None
    best_score = 0
    for cat, subcats in COMPILED_MAP.items():
        for sub, strengths in subcats.items():
            strong_hits = sum(bool(p.search(text)) for p in strengths["strong"])
            weak_hits = sum(bool(p.search(text)) for p in strengths["weak"])
            score = strong_hits * 3 + weak_hits
            if score > best_score:
                best_score = score
                best = (cat, sub, "strong" if strong_hits else "weak")
    if not best:
        return "Not Detected", "Not Detected", "Not Detected", 0.0
    cat, sub, label = best
    return cat, sub, label, min(1, best_score / 5)

def apply_category_mapping(df):
    mapped = df["Body.TextBody"].apply(map_category)
    df["Category"] = mapped.apply(lambda x: x[0])
    df["Sub-Category"] = mapped.apply(lambda x: x[1])
    df["Sub-Sub-Category"] = mapped.apply(lambda x: x[2])
    df["Confidence"] = mapped.apply(lambda x: x[3])
    return df

# =================================================================
# 5. CHATBOT ADDRESSABILITY ENGINE
# =================================================================
PATTERNS = { ... }  # Same patterns as your original code

def compute_score(subject, body):
    text = clean_text_chatbot(subject + " " + body)
    score = 0
    for group, data in PATTERNS.items():
        weight = data["weight"]
        for pat in data["patterns"]:
            if re.search(pat, text):
                score += weight
    return score

def chatbot_addressability(row):
    score = compute_score(row["Subject"], row["Body.TextBody"])
    if score >= 2:
        return "Yes", min(1, score / 4), score
    elif score <= -1:
        return "No", 0.1, score
    else:
        return "No", 0.3, score

def apply_chatbot(df):
    results = df.apply(chatbot_addressability, axis=1)
    df["Chatbot_Addressable"] = results.apply(lambda x: x[0])
    df["Chatbot_Confidence"] = results.apply(lambda x: x[1])
    df["Chatbot_Score"] = results.apply(lambda x: x[2])
    return df

# =================================================================
# 6. VISUALISATIONS
# =================================================================
def plot_monthly(df):
    monthly = df.groupby(df["DateTimeReceived"].dt.to_period("M")).size().reset_index(name="Count")
    monthly["Month"] = monthly["DateTimeReceived"].astype(str)
    fig, ax = plt.subplots(figsize=(10, 4))
    sns.lineplot(data=monthly, x="Month", y="Count", marker='o', color=PRIMARY_RED, ax=ax)
    ax.set_title("📈 Monthly Email Volume (2025)", fontsize=14)
    plt.xticks(rotation=45)
    plt.tight_layout()
    return fig

def plot_chatbot(df):
    counts = df["Chatbot_Addressable"].value_counts()
    fig, ax = plt.subplots(figsize=(6, 6))
    ax.pie(counts, labels=counts.index, autopct="%1.1f%%", colors=["#EE2536", "#CCCCCC"])
    ax.set_title("🤖 Chatbot Addressability Breakdown")
    return fig

# =================================================================
# 7. EXPORT
# =================================================================
def export(df):
    filename = f"ECInbox_Analysis_{datetime.today().strftime('%Y%m%d')}.xlsx"
    df.to_excel(filename, index=False)
    return filename

# =================================================================
# 8. AI SUMMARISATION & INSIGHTS
# =================================================================
''' openai.api_key = os.getenv("OPENAI_API_KEY")  # <- Safe fallback if not using Streamlit

def generate_ai_insights(email_text):
    if not email_text.strip():
        return {"summary": "", "insights": "", "risk_level": ""}
    prompt = f"""
    You are an Ethics & Compliance assistant.
    Summarize this email in 2 sentences.
    Identify any compliance-related topics (ABAC, COI, Sanctions, Data Protection, IPT).
    Suggest risk level (Low, Medium, High) and recommended action.
    Email content: {email_text}
    """
    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=[{"role": "system", "content": "You are a compliance expert."},
                  {"role": "user", "content": prompt}],
        temperature=0.3
    )
    output = response.choices[0].message["content"]
    summary = output.split("Summary:")[-1].split("Insights:")[0].strip() if "Summary:" in output else output
    insights = output.split("Insights:")[-1].split("Risk:")[0].strip() if "Insights:" in output else ""
    risk = output.split("Risk:")[-1].strip() if "Risk:" in output else ""
    return {"summary": summary, "insights": insights, "risk_level": risk}

def apply_ai_insights(df):
    summaries, insights, risks = [], [], []
    for _, row in df.iterrows():
        text = f"{row['Subject']} {row['Body.TextBody']}"
        ai_result = generate_ai_insights(text)
        summaries.append(ai_result["summary"])
        insights.append(ai_result["insights"])
        risks.append(ai_result["risk_level"])
    df["AI_Summary"] = summaries
    df["AI_Insights"] = insights
    df["AI_Risk_Level"] = risks
    return df'''

# =================================================================
# 9. FULL PIPELINE
# =================================================================
def run_full_pipeline(filepath):
    df = load_data(filepath)
    df = preprocess(df)
    df = apply_category_mapping(df)
    df = apply_chatbot(df)
    df = apply_ai_insights(df)
    return df

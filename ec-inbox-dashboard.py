# ============================================================
# 📊 EXECUTIVE INBOX ANALYSIS — FULL PIPELINE STREAMLIT APP
# Cleaning + Categorisation + Chatbot + Dashboard
# ============================================================

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import re
from datetime import datetime
import io

st.set_page_config(page_title="EC Inbox Dashboard", layout="wide")
sns.set_theme(style="whitegrid")

PRIMARY_RED = "#EE2536"


# ============================================================
# 1️⃣ FILE UPLOADER
# ============================================================

st.title("📨 E&C Inbox Dashboard")
uploaded = st.file_uploader("Upload Outlook export (Excel or CSV)", type=["xlsx", "csv"])

if not uploaded:
    st.info("Please upload an Outlook-exported Excel/CSV file to continue.")
    st.stop()

# Load file dynamically
if uploaded.name.endswith(".csv"):
    df_raw = pd.read_csv(uploaded, encoding="ISO-8859-1", on_bad_lines="skip")
else:
    df_raw = pd.read_excel(uploaded)

st.success(f"File uploaded: {uploaded.name}")


# ============================================================
# 2️⃣ CLEANING PIPELINE — YOUR ORIGINAL SCRIPT WRAPPED IN A FUNCTION
# ============================================================

def clean_datetime(col):
    col = pd.to_datetime(col, format="%m/%d/%Y %H:%M", errors="coerce")
    col = col.fillna(pd.to_datetime(col, errors="coerce"))
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


# ------------------------------------------------------------
# ---- CATEGORY DICTIONARY (your full mapping preserved)
# ------------------------------------------------------------

category_map = {
    "Anti-Bribery and Anti-Corruption (ABAC)": {
        "ISO 37001": {
            "strong": [r"\biso\s*37001\b", r"\bsurveillance audit\b", r"\banti[- ]bribery standard\b"],
            "weak": [r"certification", r"compliance standard"]
        },
        "Gifts & Entertainment": {
            "strong": [
                r"\bgift\b", r"\bgifts\b", r"\bge\b", r"\bmoon\s*cake\b", r"\bmooncake\b",
                r"\bendowment\b", r"formulaire de", r"\breceiving\b", r"\boffering\b",
                r"hospitality", r"treat", r"lunch", r"dinner", r"gift card", r"declaration",
                r"received", r"offered", r"employee offering"
            ],
            "weak": [r"entertainment", r"meal", r"token", r"cny"]
        },
        "ABAC eLearning / Training": {
            "strong": [r"\babac\b.*(training|elearning)", r"mandatory training", r"translation", r"translations"],
            "weak": []
        },
        "Third-Party Due Diligence / Screening": {
            "strong": [
                r"dow jones", r"\basam\b", r"screening", r"third[- ]party", r"due diligence",
                r"kyc", r"background check", r"supplier audit", r"screening request"
            ],
            "weak": []
        },
        "Charitable Donations, Sponsorship & Political Contributions": {
            "strong": [r"donation", r"sponsorship", r"csr", r"political contribution", r"charitable giving"],
            "weak": []
        }
    },
    "Conflict of Interests (COI)": {
        "COI Declaration": {
            "strong": [r"\bcoi\b", r"conflict of interest", r"interest declaration"],
            "weak": [r"family relationship", r"related party"]
        },
        "External Appointments": {
            "strong": [r"external appointment", r"outside employment", r"side job"],
            "weak": []
        }
    },
    "Data Protection": {
        "Data Incident / Breach": {
            "strong": [r"data breach", r"phishing", r"cyber incident", r"malware", r"ransomware"],
            "weak": [r"security incident", r"personal data"]
        },
        "Data Governance & Classification": {
            "strong": [r"data classification", r"governance", r"GDPR"],
            "weak": []
        }
    },
    "Interested Person Transactions (IPT)": {
        "IPT Policies & Procedures": {
            "strong": [r"\bipt\b policy", r"ipt procedure"],
            "weak": []
        },
        "IPT Portal / System Issues": {
            "strong": [r"ipt portal", r"ipt system", r"ipt access"],
            "weak": [r"login issue", r"access problem"]
        },
        "IPT Refreshers / Training": {
            "strong": [r"ipt training", r"ipt refresher"],
            "weak": []
        }
    },
    "Sanctions": {
        "Sanction Risk Framework": {
            "strong": [r"sanctions risk", r"risk assessment"],
            "weak": []
        },
        "Sanctions Policies & Procedures": {
            "strong": [r"sanctions procedures", r"sanctions operating"],
            "weak": []
        }
    }
}


# Compile regexes once
compiled_category_map = {}
for category, subcats in category_map.items():
    compiled_category_map[category] = {}
    for subcat, strengths in subcats.items():
        compiled_category_map[category][subcat] = {
            strength: [re.compile(pat, re.IGNORECASE) for pat in pats]
            for strength, pats in strengths.items()
        }


def map_category_with_confidence(text):
    text = clean_text_basic(text)
    results = []

    for category, subcats in compiled_category_map.items():
        for subcat, strengths in subcats.items():
            strong_hits = 0
            weak_hits = 0

            for p in strengths.get("strong", []):
                strong_hits += len(p.findall(text))

            for p in strengths.get("weak", []):
                weak_hits += len(p.findall(text))

            total_hits = strong_hits + weak_hits
            if total_hits == 0:
                continue

            raw = strong_hits * 3 + weak_hits
            max_possible = len(strengths.get("strong", [])) * 3 + len(strengths.get("weak", []))
            confidence = raw / max_possible if max_possible > 0 else 0

            results.append((category, subcat, strong_hits, raw, max_possible, confidence))

    if not results:
        return ("Not Detected", "Not Detected", "Not Detected", 0.0)

    results.sort(key=lambda x: x[5], reverse=True)
    best = results[0]

    category, subcat, strong_hits, raw, max_possible, _ = best
    strength = "strong" if strong_hits > 0 else "weak"
    conf = round(raw / max_possible, 3)

    return category, subcat, strength, conf


# Chatbot patterns
PATTERNS = {
    "high_confidence": {
        "weight": 2,
        "patterns": [
            r"\bhow to\b", r"\breset password\b", r"\blogin issue\b",
            r"\baccess denied\b", r"\bsubmit form\b", r"\bupdate profile\b"
        ]
    },
    "medium_confidence": {
        "weight": 1,
        "patterns": [
            r"\bquestion\b", r"\binquiry\b", r"\bclarification\b"
        ]
    },
    "human_required": {
        "weight": -2,
        "patterns": [r"\bresignation\b", r"\btermination\b", r"\bcomplaint\b"]
    }
}


def compute_chatbot_score(subject, body=""):
    text = f"{clean_text_chatbot(subject)} {clean_text_chatbot(body)}"
    total = 0

    for g, gdata in PATTERNS.items():
        w = gdata["weight"]
        for pat in gdata["patterns"]:
            if re.search(pat, text):
                total += w

    return total


def is_chatbot_addressable(subject, body=""):
    score = compute_chatbot_score(subject, body)

    if score >= 1.5:
        return "Yes", min(1, score / 4), score
    if score <= -1:
        return "No", 0.10, score
    if 0 < score < 1.5:
        return "Yes", 0.4, score
    return "No", 0.25, score


# ------------------------------------------------------------
# 🚀 MAIN PROCESSING FUNCTION
# ------------------------------------------------------------

@st.cache_data
def process_raw_data(df):

    # required columns
    required = ['DateTimeSent', 'DateTimeReceived', 'Subject', 'Body.TextBody']
    for col in required:
        if col not in df.columns:
            st.error(f"Missing required column: {col}")
            st.stop()

    df["DateTimeSent"] = clean_datetime(df["DateTimeSent"])
    df["DateTimeReceived"] = clean_datetime(df["DateTimeReceived"])

    df = df[df["DateTimeReceived"].dt.year == 2025].copy()
    df.drop_duplicates(inplace=True)

    # Clean text for category model
    df["Clean-Text"] = df["Body.TextBody"].fillna("").apply(clean_text_basic)

    # Map category + chatbot
    df[["Category", "Sub-Category", "Sub-Sub-Category", "Confidence"]] = df["Clean-Text"].apply(
        lambda x: pd.Series(map_category_with_confidence(x))
    )

    df["Chatbot_Addressable"], df["Chatbot_Confidence"], df["Chatbot_Score"] = zip(*df.apply(
        lambda r: is_chatbot_addressable(
            r.get("Subject", ""),
            r.get("Body.TextBody", "")
        ),
        axis=1
    ))

    return df


# ============================================================
# 3️⃣ PROCESS USING USER UPLOADED DATA
# ============================================================

df = process_raw_data(df_raw)

st.success("Data processed successfully!")


# ============================================================
# 4️⃣ DASHBOARD — KPIs + Charts
# ============================================================

st.header("📈 Executive KPIs")

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("Total Emails (2025)", len(df))
with col2:
    monthly_avg = df.groupby(df["DateTimeReceived"].dt.month).size().mean().round(1)
    st.metric("Avg Emails per Month", monthly_avg)
with col3:
    peak_month = df["DateTimeReceived"].dt.month.value_counts().idxmax()
    st.metric("Peak Month", peak_month)

# Monthly chart
monthly = (
    df.groupby(df["DateTimeReceived"].dt.to_period('M'))
    .size()
    .reset_index(name='Count')
)
monthly["Month"] = monthly["DateTimeReceived"].astype(str)

fig, ax = plt.subplots(figsize=(10, 4))
sns.lineplot(data=monthly, x="Month", y="Count", marker='o', color=PRIMARY_RED)
plt.xticks(rotation=45)
st.pyplot(fig)

# ============================================================
# 5️⃣ FILTER TABLE
# ============================================================

st.header("🔎 Explore Email Records")

category_filter = st.selectbox("Filter by Category", ["All"] + sorted(df["Category"].unique()))

filtered_df = df if category_filter == "All" else df[df["Category"] == category_filter]

st.dataframe(filtered_df, use_container_width=True)


# ============================================================
# 6️⃣ DOWNLOAD CLEANED DATA
# ============================================================

buffer = io.BytesIO()
df.to_excel(buffer, index=False)
st.download_button(
    label="📥 Download Cleaned Excel",
    data=buffer.getvalue(),
    file_name=f"ECInbox_Analysis_{datetime.today().strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


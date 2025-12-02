# ============================================================
# 📊 Inbox Data Analysis — Consolidated Engine
# FULL CLEANING + CATEGORISATION + CHATBOT SCORING + VISUALS
# Designed for integration with Streamlit
# ============================================================

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import re
from datetime import datetime

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
    """Parse timestamps, fix wrong formats, remove out-of-range years."""
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


# =================================================================
# 3. PREPROCESS EXECUTIVE INBOX DATA
# =================================================================

def preprocess(df):
    df["DateTimeSent"] = clean_datetime(df["DateTimeSent"])
    df["DateTimeReceived"] = clean_datetime(df["DateTimeReceived"])

    df = df[~(df["DateTimeSent"].isna() & df["DateTimeReceived"].isna())]

    # Filter to 2025
    df = df[df["DateTimeReceived"].dt.year == 2025].copy()
    df.drop_duplicates(inplace=True)
    df.reset_index(drop=True, inplace=True)

    # Remove autoreplies and spam
    exclude_terms = [
        'respuesta automática', 'automatic reply', 'automatische antwort',
        'réponse automatique', 'quarantine', 'undeliverable', 'test'
    ]
    pattern = re.compile('|'.join(exclude_terms), flags=re.IGNORECASE)
    df = df[~df['Subject'].str.contains(pattern, na=False)].copy()

    return df


# =================================================================
# 4. CATEGORY MAPPING ENGINE
# =================================================================

CATEGORY_MAP = {
    "Anti-Bribery and Anti-Corruption (ABAC)": {
        "ISO 37001": {
            "strong": [r"\biso\s*37001\b", r"\bsurveillance audit\b", r"\banti[- ]bribery standard\b"],
            "weak": [r"certification", r"compliance standard"]
        },
        "Gifts & Entertainment": {
            "strong": [r"\bgift\b", r"\bgifts\b", r"\bge\b", r"\bmoon\s*cake\b", r"\bmooncake\b",
                       r"\bendowment\b", r"formulaire de", r"\breceiving\b", r"\boffering\b",
                       r"hospitality", r"treat", r"lunch", r"dinner", r"gift card", r"declaration", r"received", r"offered", r"employee offering"],
            "weak": [r"entertainment", r"meal", r"token", r"cny"]
        },
        "ABAC eLearning / Training": {
            "strong": [r"\babac\b.*(training|elearning)", r"mandatory training", r"translation", r"translations"],
            "weak": []
        },
        "Third-Party Due Diligence / Screening": {
            "strong": [r"dow jones", r"\basam\b", r"screening", r"third[- ]party", r"due diligence", r"kyc",
                       r"background check", r"supplier audit", r"screening request", r"due diligence", r"diligence screening"],
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
            "strong": [r"external appointment", r"outside employment", r"side job", r"additional role"],
            "weak": []
        }
    },
    "Data Protection": {
        "Data Incident / Breach": {
            "strong": [r"data breach", r"phishing", r"cyber incident", r"malware", r"ransomware", r"leak of data", r"PII"],
            "weak": [r"security incident", r"personal data"]
        },
        "Data Governance & Classification": {
            "strong": [r"data classification", r"governance", r"GDPR", r"data handling", r"sensitive information"],
            "weak": []
        }
    },
    "Interested Person Transactions (IPT)": {
        "IPT Policies & Procedures": {
            "strong": [r"\bipt\b policy", r"ipt procedure", r"ipt compliance"],
            "weak": []
        },
        "IPT Portal / System Issues": {
            "strong": [r"ipt portal", r"ipt system", r"ipt access", r"cannot login", r"cannot access"],
            "weak": [r"login issue", r"access problem"]
        },
        "IPT Refreshers / Training": {
            "strong": [r"ipt training", r"ipt refresher", r"ipt course"],
            "weak": []
        }
    },
    "Sanctions": {
        "Sanction Risk Framework": {
            "strong": [r"sanctions risk", r"risk assessment", r"compliance check"],
            "weak": []
        },
        "Sanctions Policies & Procedures": {
            "strong": [r"sanctions procedures", r"sanctions operating", r"sanctions policy", r"review sanctions"],
            "weak": []
        }
    }
}

# compile patterns
COMPILED_MAP = {}
for cat, subcats in CATEGORY_MAP.items():
    COMPILED_MAP[cat] = {}
    for sub, strengths in subcats.items():
        COMPILED_MAP[cat][sub] = {
            "strong": [re.compile(p, re.IGNORECASE) for p in strengths["strong"]],
            "weak": [re.compile(p, re.IGNORECASE) for p in strengths["weak"]],
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

PATTERNS = {
    "high_confidence": {
        "weight": 2,
        "patterns": [
            r"\bhow to\b",
            r"\breset password\b",
            r"\blogin issue\b",
            r"\baccess denied\b",
            r"\brequest info\b",
            r"\bsubmit form\b",
            r"\bupdate profile\b",
            r"\bcheck status\b",
            r"\bpassword reset\b",
            r"\bactivate account\b"
        ]
    },
    "medium_confidence": {
        "weight": 1,
        "patterns": [
            r"\bquestion\b",
            r"\binquiry\b",
            r"\bclarification\b",
            r"\bissue\b",
            r"\bhelp\b",
            r"\bsupport\b",
            r"\btroubleshoot\b",
            r"\bproblem\b"
        ]
    },
    "borderline_human_indicators": {
        "weight": -1,
        "patterns": [
            r"\bplease advise\b",
            r"\bneed approval\b",
            r"\bconfirm\b",
            r"\brequest approval\b",
            r"\baction required\b"
        ]
    },
    "human_required": {
        "weight": -2,
        "patterns": [
            r"\bresignation\b",
            r"\btermination\b",
            r"\blegal\b",
            r"\blitigation\b",
            r"\bcomplaint\b",
            r"\bHR sensitive\b",
            r"\bconfidential\b",
            r"\bescalation\b"
        ]
    }
}


def compute_score(subject, body):
    text = clean_text_chatbot(subject + " " + body)

    score = 0
    for group, data in PATTERNS.items():
        for pat in data["p"]:
            if re.search(pat, text):
                score += data["w"]

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
# 6. VISUALISATIONS (RETURN FIGURES FOR STREAMLIT)
# =================================================================

def plot_monthly(df):
    monthly = (
        df.groupby(df["DateTimeReceived"].dt.to_period("M"))
        .size().reset_index(name="Count")
    )
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
    ax.pie(
        counts,
        labels=counts.index,
        autopct="%1.1f%%",
        colors=["#EE2536", "#CCCCCC"]
    )
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
# 8. MAIN PIPELINE FUNCTION (USED IN STREAMLIT)
# =================================================================

def run_full_pipeline(filepath):
    df = load_data(filepath)
    df = preprocess(df)
    df = apply_category_mapping(df)
    df = apply_chatbot(df)
    return df

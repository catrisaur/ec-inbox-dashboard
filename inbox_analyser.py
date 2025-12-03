
# ============================================================
# 📊 Inbox Data Analysis — Consolidated Engine
# FULL CLEANING + CATEGORIZATION + CHATBOT SCORING + VISUALS
# Designed for integration with Streamlit
# ============================================================

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import re
from datetime import datetime
from typing import Tuple

sns.set_theme(style="whitegrid")
PRIMARY_RED = "#EE2536"

# =================================================================
# 1. LOAD + VALIDATE
# =================================================================

def load_data(filepath: str) -> pd.DataFrame:
    """Load Excel file and validate required columns."""
    df = pd.read_excel(filepath)
    required_cols = ['DateTimeSent', 'DateTimeReceived', 'Subject', 'Body.TextBody']
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns: {missing_cols}")
    return df


# =================================================================
# 2. CLEANING FUNCTIONS
# =================================================================

def clean_datetime(series: pd.Series) -> pd.Series:
    """Convert to datetime, remove invalid years."""
    series = pd.to_datetime(series, errors="coerce")
    return series.mask((series.dt.year < 2000) | (series.dt.year > 2030))


def clean_text_basic(text: str) -> str:
    """Basic text cleaning: lowercase, remove special chars."""
    if pd.isna(text):
        return ""
    text = str(text).lower()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^\w\s]", " ", text)
    return text.strip()


def clean_text_chatbot(text: str) -> str:
    """Chatbot-specific cleaning: allow certain symbols."""
    if not isinstance(text, str):
        return ""
    text = text.lower()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^a-z0-9@._/%$ -]", "", text)
    return text.strip()


# =================================================================
# 3. PREPROCESS EXECUTIVE INBOX DATA
# =================================================================

def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    """Clean datetime, filter year, remove duplicates and spam."""
    df["DateTimeSent"] = clean_datetime(df["DateTimeSent"])
    df["DateTimeReceived"] = clean_datetime(df["DateTimeReceived"])

    # Remove rows with no valid dates
    df = df.dropna(subset=["DateTimeSent", "DateTimeReceived"], how="all")

    # Filter to 2025 only
    df = df[df["DateTimeReceived"].dt.year == 2025].copy()
    df.drop_duplicates(inplace=True)
    df.reset_index(drop=True, inplace=True)

    # Remove auto-replies & spam
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

CATEGORY_MAP = {
    "Anti-Bribery and Anti-Corruption (ABAC)": {
        "ISO 37001": {
            "strong": [r"\biso\s*37001\b", r"\baudit\b", r"\banti[- ]bribery\b"],
            "weak": [r"\bcertification\b", r"\bcompliance\b"]
        },
        "Gifts & Entertainment": {
            "strong": [r"\bgift\b", r"\bgifts\b", r"\bdeclaration\b", r"\boffered\b", r"\breceived\b"],
            "weak": [r"\bmeal\b", r"\btoken\b", r"\bhospitality\b"]
        },
        "ABAC eLearning / Training": {
            "strong": [r"\babac\b.*(training|elearning)", r"\bmandatory training\b"],
            "weak": [r"\bprocedures\b"]
        },
        "Third-Party Due Diligence / Screening": {
            "strong": [r"\bdow jones\b", r"\bscreening\b", r"\bthird[- ]party\b", r"\bdue diligence\b", r"\bvendor screening\b", r"\bdue diligence tool\b"],
            "weak": [r"\bcheck\b", r"\bmonitoring\b"]
        },
        "Charitable Donations, Sponsorship & Political Contributions": {
            "strong": [r"\bsponsorship\b", r"\bdonation\b", r"\bcharitable\b"],
            "weak": [r"\bcsr\b"]
        }
    },
    "Conflict of Interests (COI)": {
        "COI Declaration": {
            "strong": [r"\bconflict of interest\b", r"\bcoi\b", r"\binterest declaration\b"],
            "weak": [r"\bfamily relationship\b"]
        },
        "External Appointments": {
            "strong": [r"\bappointment\b", r"\bboard\b", r"\bexternal role\b"],
            "weak": [r"\badditional role\b"]
        }
    },
    "Data Protection": {
        "Data Incident / Breach": {
            "strong": [r"\bdata breach\b", r"\bphishing\b", r"\bransomware\b", r"\bcyber incident\b"],
            "weak": [r"\bsecurity incident\b", r"\bpersonal data\b"]
        },
        "Data Governance & Classification": {
            "strong": [r"\bdata classification\b", r"\bGDPR\b", r"\bgovernance\b"],
            "weak": [r"\bsensitive information\b"]
        }
    },
    "Interested Person Transactions (IPT)": {
        "IPT Policies & Procedures": {
            "strong": [r"\bipt\b.*policy", r"\bipt\b.*procedure"],
            "weak": []
        },
        "IPT Portal / System Issues": {
            "strong": [r"\bipt portal\b", r"\bipt system\b", r"\baccess\b"],
            "weak": [r"\blogin issue\b"]
        },
        "IPT Refreshers / Training": {
            "strong": [r"\bipt training\b", r"\bipt refresher\b"],
            "weak": []
        }
    },
    "Sanctions": {
        "Sanction Risk Framework": {
            "strong": [r"\bsanction\b", r"\brisk assessment\b", r"\bcompliance\b"],
            "weak": []
        },
        "Sanctions Policies & Procedures": {
            "strong": [r"\bsanctions operating\b", r"\bsanctions policy\b", r"\breview sanctions\b"],
            "weak": []
        }
    }
}

# Precompile regex patterns for efficiency
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


def map_category(text: str) -> Tuple[str, str, str, float]:
    """Map text to category and compute confidence score."""
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


def apply_category_mapping(df: pd.DataFrame) -> pd.DataFrame:
    """Apply category mapping to dataframe."""
    mapped = df["Body.TextBody"].apply(map_category)
    df[["Category", "Sub-Category", "Sub-Sub-Category", "Confidence"]] = pd.DataFrame(mapped.tolist(), index=df.index)
    return df


# =================================================================
# 5. CHATBOT ADDRESSABILITY ENGINE
# =================================================================

PATTERNS = {
    "high_confidence": {
        "weight": 3,
        "patterns": [
            # Direct "how/what/where" process queries
            r"\bhow to\b", r"\bhow do i\b", r"\bwhere do i\b",
            r"\bprocess\b", r"\bprocedure\b",

            # Access / portal / workflow issues
            r"\breset password\b", r"\blogin issue\b", r"\baccess denied\b",
            r"\bactivate account\b", r"\bcheck status\b",

            # Form or declaration queries
            r"\bsubmit form\b", r"\bform attached\b", r"\bsubmit declaration\b",

            # E&C-specific obligations
            r"\bg&e\b", r"\bgift(s?) & entertainment\b",
            r"\bconflict of interest\b", r"\bcoi\b",
            r"\btraining deadline\b", r"\bmandatory training\b",
            r"\bipt submission\b", r"\bannual declaration\b",
        ]
    },

    "medium_confidence": {
        "weight": 1,
        "patterns": [
            # Generic help-seeking
            r"\bquestion\b", r"\bclarification\b", r"\bhelp\b",
            r"\bsupport\b", r"\binquiry\b", r"\bissue\b",

            # Policy & procedural references
            r"\bpolicy\b", r"\bupdate procedure\b", r"\bworkflow\b",
            r"\bcompliance\b",

            # E&C review processes
            r"\brequest for check\b", r"\bdue diligence\b",
            r"\bscreening\b", r"\bsanctions\b",
        ]
    },

    # Signals that the user needs real judgement — borderline auto
    "borderline_human_indicators": {
        "weight": -1,
        "patterns": [
            r"\bplease advise\b", r"\bneed approval\b",
            r"\brequest approval\b", r"\bconfirm\b",

            # E&C escalations
            r"\bexception\b", r"\bescalation\b", r"\bwaiver\b",
            r"\bcase assessment\b"
        ]
    },

    # Very high risk or sensitive items → must be human-handled
    "human_required": {
        "weight": -3,
        "patterns": [
            r"\blegal\b", r"\bcomplaint\b", r"\bconfidential\b",
            r"\bresignation\b",

            # E&C critical topics
            r"\bbribery\b", r"\bcorruption\b",
            r"\bdisciplinary\b", r"\bmisconduct\b",
            r"\bretaliation\b", r"\bwhistleblow(ing)?\b"
        ]
    }
}

def compute_score(subject: str, body: str) -> int:
    """Compute chatbot addressability score."""
    text = clean_text_chatbot(subject + " " + body)
    score = 0
    for group, data in PATTERNS.items():
        weight = data["weight"]
        for pat in data["patterns"]:
            if re.search(pat, text):
                score += weight
    return score


def chatbot_addressability(row: pd.Series) -> Tuple[str, float, int]:
    """Determine if chatbot can address the query."""
    score = compute_score(row["Subject"], row["Body.TextBody"])
    if score >= 2:
        return "Yes", min(1, score / 4), score
    elif score <= -1:
        return "No", 0.1, score
    else:
        return "No", 0.3, score


def apply_chatbot(df: pd.DataFrame) -> pd.DataFrame:
    """Apply chatbot scoring to dataframe."""
    results = df.apply(chatbot_addressability, axis=1)
    df[["Chatbot_Addressable", "Chatbot_Confidence", "Chatbot_Score"]] = pd.DataFrame(results.tolist(), index=df.index)
    return df


# =================================================================
# 6. VISUALISATIONS (RETURN FIGURES FOR STREAMLIT)
# =================================================================

def plot_monthly(df: pd.DataFrame):
    """Plot monthly email volume."""
    monthly = df.groupby(df["DateTimeReceived"].dt.to_period("M")).size().reset_index(name="Count")
    monthly["Month"] = monthly["DateTimeReceived"].astype(str)

    fig, ax = plt.subplots(figsize=(10, 4))
    sns.lineplot(data=monthly, x="Month", y="Count", marker='o', color=PRIMARY_RED, ax=ax)
    ax.set_title("📈 Monthly Email Volume (2025)", fontsize=14)
    plt.xticks(rotation=45)
    plt.tight_layout()
    return fig


def plot_chatbot(df: pd.DataFrame):
    """Plot chatbot addressability breakdown."""
    counts = df["Chatbot_Addressable"].value_counts()
    fig, ax = plt.subplots(figsize=(6, 6))
    ax.pie(counts, labels=counts.index, autopct="%1.1f%%", colors=["#EE2536", "#CCCCCC"])
    ax.set_title("🤖 Chatbot Addressability Breakdown")
    return fig


# =================================================================
# 7. EXPORT
# =================================================================

def export(df: pd.DataFrame) -> str:
    """Export dataframe to Excel."""
    filename = f"ECInbox_Analysis_{datetime.today().strftime('%Y%m%d')}.xlsx"
    df.to_excel(filename, index=False)
    return filename


# =================================================================
# 8. MAIN PIPELINE FUNCTION (USED IN STREAMLIT)
# =================================================================

def run_full_pipeline(filepath: str) -> pd.DataFrame:
    """Run full analysis pipeline."""
    df = load_data(filepath)
    df = preprocess(df)
    df = apply_category_mapping(df)
    df = apply_chatbot(df)
    return df

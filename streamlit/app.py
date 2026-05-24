import io
import math
import os
import re
from datetime import datetime
from html import escape
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    from streamlit_autorefresh import st_autorefresh
except Exception:
    st_autorefresh = None

APP_TITLE = "UC2 - Underwriting Early-Claim Risk Dashboard"
DATA_DIR = Path(__file__).parent / "data"
DEFAULT_DATA_PATH = DATA_DIR / "master_data.xlsx"
DEFAULT_PARQUET_PATH = DATA_DIR / "master_data.parquet"
DEFAULT_MIN_SEGMENT_SIZE = 100

COLORS = {
    "bg": "#F0F4F8",
    "surface": "#FFFFFF",
    "border": "#DEE5EF",
    "text": "#1A2535",
    "text2": "#3D5068",
    "muted": "#6B7E96",
    "accent": "#4F46E5",
    "f2f": "#1D6FA4",
    "nf2f": "#C0392B",
    "green": "#1A7A44",
    "amber": "#B45309",
    "red": "#991B1B",
    "light": "#F7F9FC",
}

FRIENDLY_LABELS = {
    "application_channel": "Submission mode",
    "raw_channel": "Raw channel",
    "product_clean": "Product group",
    "uw_outcome_clean": "Underwriting outcome",
    "age_band_clean": "Age band",
    "sa_band_clean": "Sum assured band",
    "bmi_band_clean": "BMI band",
    "medical_exam_flag": "Medical exam",
    "material_finding_flag": "Material finding",
    "early_claim_flag": "Early claim",
    "early_claim_rate": "Early-claim rate",
    "material_finding_rate": "Material-finding rate",
    "medical_exam_rate": "Medical-exam rate",
    "lift_vs_portfolio": "Lift vs portfolio",
    "portfolio_share": "Portfolio share",
    "early_claim_share": "Early-claim share",
    "excess_share": "Excess contribution",
    "policies": "Policies",
    "early_claims": "Early claims",
    "medical_exams": "Medical exams",
    "material_findings": "Material findings",
    "policy_year": "Application year",
    "disclosure_rule_clean": "Disclosure rule",
    "disclosure_flag": "Disclosure present",
    "dvm_flag": "DVM present",
    "hd_disclosure_flag": "Health disclosure present",
    "icd_text": "ICD description",
    "claim_type_clean": "Claim type",
    "claim_decision_clean": "Claim decision",
    "annual_income": "Annual income",
    "ai_band_clean": "Annual income band",
    "aps_flag": "APS requested",
    "mpci_flag": "MPCI rider",
    "reopen_flag": "Reopen case",
    "source_business_clean": "Source of business",
    "fund_category_clean": "Fund category",
    "loading_flag": "Loading applied",
    "exclusion_flag": "Exclusion applied",
    "myinfo_flag": "MyInfo used",
    "efna_flag": "eFNA used",
    "gio_flag": "Guaranteed-issue case",
    "age": "Age",
    "base_sa": "Base sum assured",
    "bmi": "BMI",
    "threshold": "BMI threshold",
    "Extra Medical Exams": "Extra medical exams",
    "Estimated Cost": "Estimated cost",
    "Metric": "Metric",
    "Early-claim rate": "Early-claim rate",
}

st.set_page_config(page_title="UC2 UW Dashboard", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

BMI_BAND_ORDER = [
    "<=20",
    "21-25",
    "26-30",
    "31-35",
    "36-40",
    "> 40",
    "<18.5",
    "18.5-23",
    "23-25",
    "25-27.5",
    "27.5-30",
    "30-35",
    ">35",
    "Unknown",
]


def inject_css():
    css = """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500&display=swap');
    html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif; }
    .stApp { background: __BG__; color: __TEXT__; }
    section[data-testid="stSidebar"] { background: #FFFFFF; border-right: 1px solid __BORDER__; }
    div[data-testid="stMetric"] { background: #FFFFFF; border: 1px solid __BORDER__; border-radius: 14px; padding: 14px 16px; box-shadow: 0 1px 2px rgba(15,23,42,.04); }
    div[data-testid="stMetricLabel"] p { font-size: .72rem; font-weight: 800; letter-spacing: .08em; text-transform: uppercase; color: __MUTED__; }
    div[data-testid="stMetricValue"] { font-weight: 800; color: __TEXT__; }
    .topbar { background: __TEXT__; color: white; padding: 16px 22px; border-radius: 18px; display:flex; align-items:center; justify-content:space-between; margin-bottom: 14px; box-shadow: 0 8px 24px rgba(15,23,42,.18); }
    .topbar-title { font-weight: 800; letter-spacing: -0.02em; font-size: 1.05rem; }
    .topbar-sub { color: #94A3B8; font-size: .84rem; margin-left: 10px; }
    .badge { background: __ACCENT__; color: white; padding: 4px 10px; border-radius: 999px; font-size: .72rem; font-weight: 800; letter-spacing: .05em; }
    .story { background: #FFFFFF; border: 1px solid __BORDER__; border-left: 5px solid __ACCENT__; border-radius: 0 14px 14px 0; padding: 18px 22px; margin: 10px 0 18px 0; }
    .eyebrow { font-size: .72rem; color: __ACCENT__; font-weight: 800; letter-spacing: .12em; text-transform: uppercase; margin-bottom: 6px; }
    .headline { font-size: 1.25rem; color: __TEXT__; font-weight: 800; margin-bottom: 6px; }
    .bodycopy { color: #3D5068; font-size: .92rem; line-height: 1.7; max-width: 1120px; }
    .card { background: #FFFFFF; border: 1px solid __BORDER__; border-radius: 14px; padding: 18px 20px; box-shadow: 0 1px 2px rgba(15,23,42,.04); margin-bottom: 16px; }
    .card-title { font-size: .94rem; font-weight: 800; color: __TEXT__; margin-bottom: 4px; }
    .card-subtitle { font-size: .80rem; color: __MUTED__; margin-bottom: 12px; line-height: 1.55; }
    .insight { background: #EEF2FF; border-left: 4px solid __ACCENT__; border-radius: 0 10px 10px 0; padding: 12px 14px; margin: 10px 0 12px 0; color: #3D5068; font-size: .86rem; line-height: 1.65; }
    .warn { background: #FEF3C7; border-left-color: __AMBER__; }
    .danger { background: #FEE2E2; border-left-color: __NF2F__; }
    .good { background: #DCFCE7; border-left-color: __GREEN__; }
    .small-mono { font-family: 'IBM Plex Mono', monospace; font-size: .74rem; color: __MUTED__; }
    .section-label { display:flex; align-items:center; gap:10px; margin: 24px 0 14px 0; }
    .section-num { background:__ACCENT__; color:white; font-size:.72rem; font-weight:800; width:26px; height:26px; border-radius:8px; display:flex; align-items:center; justify-content:center; }
    .section-text { font-size:.92rem; font-weight:800; color:__TEXT__; }
    .pill { display:inline-flex; align-items:center; padding:4px 10px; border-radius:999px; background:#F7F9FC; border:1px solid __BORDER__; color:#3D5068; font-size:.76rem; font-weight:700; margin: 2px 4px 2px 0; }
    .risk-score { font-size: 2.5rem; font-weight: 900; letter-spacing: -0.04em; color: __TEXT__; line-height: 1; }
    .footer { color:__MUTED__; font-size:.76rem; border-top:1px solid __BORDER__; padding-top:16px; margin-top:28px; display:flex; justify-content:space-between; gap:1rem; }
    div[data-testid="stDataFrame"] { border: 1px solid __BORDER__; border-radius: 12px; overflow: hidden; }
    .kpi-card { background:#FFFFFF; border:1px solid __BORDER__; border-radius:14px; padding:18px 20px; min-height:138px; box-shadow:0 1px 2px rgba(15,23,42,.04); }
    .kpi-title { color:__TEXT__; font-size:.88rem; font-weight:800; margin-bottom:12px; }
    .kpi-value { color:__TEXT__; font-size:clamp(1.55rem, 2.4vw, 2.35rem); line-height:1.05; font-weight:900; letter-spacing:-.04em; margin-bottom:12px; overflow-wrap:anywhere; white-space:normal; }
    .kpi-sub { color:#3D5068; background:#F1F5F9; border:1px solid #E2E8F0; border-radius:12px; padding:7px 10px; font-size:.82rem; line-height:1.35; white-space:normal; overflow-wrap:anywhere; }
    .score-panel { background:#FFFFFF; border:1px solid __BORDER__; border-radius:16px; padding:18px 20px; box-shadow:0 1px 2px rgba(15,23,42,.04); margin-bottom:14px; }
    .score-label { color:__MUTED__; font-size:.76rem; text-transform:uppercase; letter-spacing:.08em; font-weight:800; margin-bottom:8px; }
    .score-main { color:__TEXT__; font-size:3rem; line-height:1; font-weight:900; letter-spacing:-.05em; margin-bottom:8px; }
    .score-explain { color:#3D5068; font-size:.9rem; line-height:1.55; margin-bottom:14px; }
    .score-bars { display:grid; gap:8px; }
    .score-row { display:grid; grid-template-columns: 145px 1fr 70px; gap:10px; align-items:center; color:#3D5068; font-size:.82rem; font-weight:700; }
    .score-track { height:12px; background:#E2E8F0; border-radius:999px; overflow:hidden; }
    .score-fill { height:12px; border-radius:999px; }
    .risk-banner { background:#FFFFFF; border:1px solid __BORDER__; border-radius:14px; padding:16px 18px; margin:14px 0; box-shadow:0 1px 2px rgba(15,23,42,.04); }
    .risk-banner-label { color:__MUTED__; font-size:.76rem; text-transform:uppercase; letter-spacing:.08em; font-weight:800; margin-bottom:6px; }
    .risk-banner-value { color:__TEXT__; font-size:2rem; line-height:1.15; font-weight:900; letter-spacing:-.03em; white-space:normal; overflow-wrap:break-word; }
    .risk-banner-sub { color:#3D5068; font-size:.86rem; line-height:1.45; margin-top:6px; }
    .decision-grid { display:grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap:14px; margin: 12px 0 18px 0; }
    .decision-panel { background:#FFFFFF; border:1px solid __BORDER__; border-radius:14px; padding:16px 18px; min-height:132px; box-shadow:0 1px 2px rgba(15,23,42,.04); }
    .decision-kicker { color:__ACCENT__; font-size:.68rem; text-transform:uppercase; letter-spacing:.12em; font-weight:900; margin-bottom:7px; }
    .decision-title { color:__TEXT__; font-size:1rem; line-height:1.35; font-weight:900; margin-bottom:6px; }
    .decision-copy { color:#3D5068; font-size:.84rem; line-height:1.55; }
    .method-table { border:1px solid __BORDER__; border-radius:12px; overflow-y:auto; max-height:440px; background:#F8FAFC; margin-top:8px; }
    .method-row { display:grid; grid-template-columns:minmax(150px, 230px) minmax(0, 1fr); border-bottom:1px solid __BORDER__; }
    .method-row:last-child { border-bottom:0; }
    .method-step { padding:12px 14px; color:__TEXT__; font-weight:800; font-size:.86rem; line-height:1.45; background:#F1F5F9; overflow-wrap:anywhere; }
    .method-detail { padding:12px 14px; color:#3D5068; font-size:.86rem; line-height:1.55; white-space:normal; overflow-wrap:anywhere; }
    .wrapped-table { border:1px solid __BORDER__; border-radius:12px; overflow:hidden; background:#FFFFFF; margin:8px 0 12px; }
    .wrapped-row { display:grid; border-bottom:1px solid __BORDER__; }
    .wrapped-row:last-child { border-bottom:0; }
    .wrapped-row.header { background:#F1F5F9; }
    .wrapped-cell { padding:11px 13px; color:#3D5068; font-size:.84rem; line-height:1.48; white-space:normal; overflow-wrap:anywhere; min-width:0; }
    .wrapped-row.header .wrapped-cell { color:__MUTED__; font-size:.76rem; text-transform:uppercase; letter-spacing:.06em; font-weight:900; }
    .wrapped-cell.strong { color:__TEXT__; font-weight:850; }
    .assumption-strip { display:grid; grid-template-columns:repeat(4, minmax(0, 1fr)); gap:10px; margin:14px 0 4px; }
    .assumption-item { background:#F8FAFC; border:1px solid __BORDER__; border-radius:12px; padding:12px 14px; min-height:82px; }
    .assumption-label { color:__MUTED__; font-size:.68rem; letter-spacing:.08em; text-transform:uppercase; font-weight:900; margin-bottom:6px; }
    .assumption-value { color:__TEXT__; font-size:1.05rem; line-height:1.3; font-weight:900; overflow-wrap:anywhere; }
    .assumption-note { color:#3D5068; font-size:.78rem; line-height:1.45; margin-top:5px; }
    .mini-finding { background:#FFFFFF; border:1px solid __BORDER__; border-radius:14px; padding:16px 18px; min-height:154px; box-shadow:0 1px 2px rgba(15,23,42,.04); margin-bottom:12px; }
    .mini-finding-tag { color:__ACCENT__; font-size:.72rem; text-transform:uppercase; letter-spacing:.08em; font-weight:900; margin-bottom:8px; }
    .mini-finding-title { color:__TEXT__; font-size:1rem; line-height:1.35; font-weight:900; margin-bottom:8px; }
    .mini-finding-copy { color:#3D5068; font-size:.84rem; line-height:1.55; }
    .funnel-stack { display:grid; gap:10px; margin-top:12px; }
    .funnel-step { background:#F8FAFC; border:1px solid __BORDER__; border-left:5px solid __ACCENT__; border-radius:12px; padding:12px 14px; }
    .funnel-step-title { color:__TEXT__; font-size:.95rem; font-weight:900; margin-bottom:4px; }
    .funnel-step-copy { color:#3D5068; font-size:.82rem; line-height:1.45; }
    .gap-grid { display:grid; grid-template-columns:repeat(2, minmax(0, 1fr)); gap:12px; margin:12px 0; }
    .gap-box { background:#F8FAFC; border:1px solid __BORDER__; border-radius:12px; padding:14px 16px; }
    .gap-title { color:__TEXT__; font-size:.86rem; font-weight:900; margin-bottom:8px; }
    .gap-item { color:#3D5068; font-size:.82rem; line-height:1.55; margin-bottom:4px; }
    @media (max-width: 760px) {
        .method-row { grid-template-columns:1fr; }
        .method-step { border-bottom:1px solid __BORDER__; }
        .assumption-strip { grid-template-columns:1fr; }
        .gap-grid { grid-template-columns:1fr; }
        .wrapped-row { grid-template-columns:1fr !important; }
        .wrapped-row.header { display:none; }
        .wrapped-cell::before { content: attr(data-label); display:block; color:__MUTED__; font-size:.68rem; letter-spacing:.08em; text-transform:uppercase; font-weight:900; margin-bottom:3px; }
    }
    </style>
    """
    for token, value in {
        "__BG__": COLORS["bg"],
        "__TEXT__": COLORS["text"],
        "__BORDER__": COLORS["border"],
        "__MUTED__": COLORS["muted"],
        "__ACCENT__": COLORS["accent"],
        "__AMBER__": COLORS["amber"],
        "__NF2F__": COLORS["nf2f"],
        "__GREEN__": COLORS["green"],
    }.items():
        css = css.replace(token, value)
    st.markdown(css, unsafe_allow_html=True)


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    cols = []
    seen = {}
    for c in out.columns:
        base = re.sub(r"\s+", " ", str(c)).strip()
        base = re.sub(r"\.\d+$", "", base)
        if base in seen:
            seen[base] += 1
            cols.append(f"{base}__{seen[base] + 1}")
        else:
            seen[base] = 0
            cols.append(base)
    out.columns = cols
    return out


def available(df: pd.DataFrame, names: Iterable[str]) -> str | None:
    normalized = {str(c).strip().lower(): c for c in df.columns}
    for name in names:
        if name in df.columns:
            return name
        key = str(name).strip().lower()
        if key in normalized:
            return normalized[key]
    return None


def text_col(df: pd.DataFrame, names: Iterable[str], default: str = "Unknown") -> pd.Series:
    col = available(df, names)
    if col is None:
        return pd.Series(default, index=df.index, dtype="object")
    s = df[col].astype("string").str.strip()
    s = s.replace({"": pd.NA, "#N/A": pd.NA, "N/A": pd.NA, "nan": pd.NA, "None": pd.NA, "-": pd.NA})
    return s.fillna(default).astype(str)


def numeric_col(df: pd.DataFrame, names: Iterable[str]) -> pd.Series:
    col = available(df, names)
    if col is None:
        return pd.Series(np.nan, index=df.index, dtype="float64")
    return pd.to_numeric(df[col].replace({"#N/A": np.nan, "N/A": np.nan, "": np.nan, "-": np.nan}), errors="coerce")


def coalesce_numeric(*series: pd.Series) -> pd.Series:
    if not series:
        return pd.Series(dtype="float64")
    out = series[0].copy()
    for s in series[1:]:
        out = out.where(out.notna(), s)
    return out


def yes_no_flag(s: pd.Series) -> pd.Series:
    raw = s.astype("string").str.strip().str.lower()
    return raw.isin(["yes", "y", "true", "1", "1.0", "medical examination", "medical", "me", "with medical examination", "f2f", "reopen case", "gio"])


def has_value_flag(s: pd.Series) -> pd.Series:
    raw = s.astype("string").str.strip().str.lower()
    return raw.notna() & ~raw.isin(["", "#n/a", "n/a", "nan", "none", "-", "unknown", "no", "0", "0.0"])


def normalize_f2f_indicator(primary: pd.Series, fallback_channel: pd.Series | None = None) -> pd.Series:
    raw = primary.astype("string").str.strip().str.upper()
    out = pd.Series("Unknown", index=primary.index, dtype="object")
    out[raw.isin(["F2F", "FACE TO FACE", "FACE-TO-FACE", "F", "FA", "AGENT", "AGENCY"])] = "F2F"
    out[raw.isin(["NF2F", "NON-F2F", "NON FACE TO FACE", "NON-FACE-TO-FACE", "P", "PHONE", "ONLINE", "DIGITAL", "TELE", "REMOTE"])] = "NF2F"
    if fallback_channel is not None:
        fb = fallback_channel.astype("string").str.strip().str.upper()
        unresolved = out.eq("Unknown")
        out[unresolved & fb.isin(["F2F", "F", "FACE TO FACE", "FACE-TO-FACE", "FA", "AGENT", "AGENCY"])] = "F2F"
        out[unresolved & fb.isin(["NF2F", "P", "NON-F2F", "PHONE", "ONLINE", "DIGITAL", "TELE", "REMOTE"])] = "NF2F"
    return out


def classify_uw(s: pd.Series) -> pd.Series:
    raw = s.astype("string").fillna("Unknown").str.replace(r"^\([a-z]\)\s*", "", regex=True).str.strip()
    low = raw.str.lower()
    out = pd.Series("Other / Unknown", index=s.index, dtype="object")
    out[low.str.contains("declin", na=False)] = "Declined"
    out[low.str.contains("postpon", na=False)] = "Postponed"
    out[low.str.contains("revised|substandard|loading|exclusion|rated|non standard|non-standard", na=False)] = "Revised Terms / Substandard"
    out[low.str.contains("standard", na=False) & ~low.str.contains("revised|substandard|non standard|non-standard", na=False)] = "Standard"
    return out


def make_age_band(age: pd.Series) -> pd.Series:
    bins = [-np.inf, 17, 25, 30, 35, 40, 45, 50, 55, np.inf]
    labels = ["a. 0-17", "b. 18-25", "c. 26-30", "d. 31-35", "e. 36-40", "f. 41-45", "g. 46-50", "h. 51-55", "i. above 56"]
    return pd.cut(age, bins=bins, labels=labels).astype("object").fillna("Unknown").astype(str)


def make_sa_band(sa: pd.Series) -> pd.Series:
    bins = [-np.inf, 50000, 100000, 150000, 200000, 250000, np.inf]
    labels = ["<= 50k", "50k-100k", "100k-150k", "150k-200k", "200k-250k", "> 250k"]
    return pd.cut(sa, bins=bins, labels=labels).astype("object").fillna("Unknown").astype(str)


def make_bmi_band(bmi: pd.Series) -> pd.Series:
    bins = [-np.inf, 20, 25, 30, 35, 40, np.inf]
    labels = ["<=20", "21-25", "26-30", "31-35", "36-40", "> 40"]
    return pd.cut(bmi, bins=bins, labels=labels).astype("object").fillna("Unknown").astype(str)


def get_early_claim_flag(df: pd.DataFrame) -> pd.Series:
    col = available(df, ["Early Claim", "EARLY_CLAIM", "early_claim"])
    if col is not None:
        raw = df[col].astype("string").str.strip().str.lower()
        flag = raw.isin(["yes", "y", "true", "1", "1.0", "early", "early claim", "ec"])
        flag = flag | (raw.str.contains("early", na=False) & ~raw.str.contains("non", na=False))
        return flag.fillna(False)
    claim_type = available(df, ["Type of Claim", "Claim Type"])
    if claim_type is not None:
        s = df[claim_type].astype("string").str.strip().str.lower()
        return (s.notna() & ~s.isin(["", "#n/a", "n/a", "nan", "none", "-"])).fillna(False)
    return pd.Series(False, index=df.index)


def add_derived_fields(df: pd.DataFrame) -> pd.DataFrame:
    df = clean_columns(df)
    out = df.copy()
    out["early_claim_flag"] = get_early_claim_flag(out).astype(int)
    out["policy_year"] = numeric_col(out, ["Year of App", "Inception Year", "Policy Year"]).astype("Int64")
    out["raw_channel"] = text_col(out, ["Channel", "SOURCE_OF_BUSINESS", "SOURCE_OF_BUSINESS_CD"], "Unknown")
    f2f_primary = text_col(out, ["F2F Indicator", "F2F_INDICATOR", "Face to Face Indicator"], "Unknown")
    out["application_channel"] = normalize_f2f_indicator(f2f_primary, out["raw_channel"])
    out["product_clean"] = text_col(out, ["Product Group", "POLICY_TYPE", "POLICY_TYPE_CD"], "Unknown")
    out["uw_raw"] = text_col(out, ["UW_OUTCOME1", "UW_OUTCOME2", "Decision BreakDown", "STATUS"], "Unknown")
    out["uw_outcome_clean"] = classify_uw(out["uw_raw"])
    out["uw_medical_clean"] = classify_uw(text_col(out, ["UW_OUTCOME2", "Decision BreakDown"], "Unknown"))
    out["age"] = numeric_col(out, ["AGE_NEXTBDAY", "Age", "age"])
    raw_age_band = text_col(out, ["AGE_NEXTBDAY_BAND", "Age Band"], "Unknown")
    out["age_band_clean"] = raw_age_band.where(raw_age_band.ne("Unknown"), make_age_band(out["age"]))
    out["base_sa"] = numeric_col(out, ["Base Plan SA", "SUM_ASSURED", "Sum Assured", "Base SA"])
    raw_sa_band = text_col(out, ["Base Plan SA Band", "SA Band", "Sum Assured Band"], "Unknown")
    out["sa_band_clean"] = raw_sa_band.where(raw_sa_band.ne("Unknown"), make_sa_band(out["base_sa"]))
    bmi_pf = numeric_col(out, ["BMI in PF", "BMI_PF"])
    bmi_me = numeric_col(out, ["BMI in ME", "BMI_ME"])
    out["bmi"] = coalesce_numeric(bmi_pf, bmi_me)
    raw_bmi_band = text_col(out, ["BMI Band in PF", "BMI Band"], "Unknown")
    out["bmi_band_clean"] = raw_bmi_band.where(raw_bmi_band.ne("Unknown"), make_bmi_band(out["bmi"]))
    out["medical_exam_flag"] = yes_no_flag(text_col(out, ["Medical Examination", "eMedex", "FU_MED"], "No")).astype(int)
    out["gender_clean"] = text_col(out, ["Gender", "Gender from Claim"], "Unknown")
    out["smoking_clean"] = text_col(out, ["Smoking Status", "SMOKING"], "Unknown")
    out["dvm_flag"] = has_value_flag(text_col(out, ["DVM Indicator"], "No")).astype(int)
    out["hd_disclosure_flag"] = has_value_flag(text_col(out, ["HD Disclosure"], "No")).astype(int)
    out["disclosure_rule_clean"] = text_col(out, ["Disclosure Rule", "Disclosure Alias"], "Unknown")
    out["disclosure_flag"] = (
        out["dvm_flag"].eq(1)
        | out["hd_disclosure_flag"].eq(1)
        | has_value_flag(out["disclosure_rule_clean"])
        | (numeric_col(out, ["ICD_CNT"]).fillna(0) > 0)
    ).astype(int)
    out["icd_text"] = text_col(out, ["ICD Description", "ICDS", "Condition Grouping"], "Unknown")
    out["claim_type_clean"] = text_col(out, ["Type of Claim"], "Unknown")
    out["claim_decision_clean"] = text_col(out, ["Claim Decision"], "Unknown")
    out["annual_income"] = numeric_col(out, ["Annual Income", "Income"])
    out["ai_band_clean"] = text_col(out, ["AI Band", "AI Band__2"], "Unknown")
    out["aps_flag"] = yes_no_flag(text_col(out, ["APS Indicator", "APS Indicator__2", "FU_APS"], "No")).astype(int)
    out["mpci_flag"] = yes_no_flag(text_col(out, ["MPCI Rider", "MPCI Component"], "No")).astype(int)
    out["reopen_flag"] = yes_no_flag(text_col(out, ["Reopen Cases Flag"], "No")).astype(int)
    out["source_business_clean"] = text_col(out, ["SOURCE_OF_BUSINESS", "SOURCE_OF_BUSINESS_CD", "Channel"], "Unknown")
    out["fund_category_clean"] = text_col(out, ["FUND_CATEGORY"], "Unknown")
    out["loading_flag"] = (
        has_value_flag(text_col(out, ["Loading", "LOADING"], "No"))
        | (numeric_col(out, ["LOADING1_MAX", "LOADING2_MAX"]).fillna(0) > 0)
    ).astype(int)
    out["exclusion_flag"] = (
        has_value_flag(text_col(out, ["Exclusion", "EXCS"], "No"))
        | (numeric_col(out, ["EXC_CNT"]).fillna(0) > 0)
    ).astype(int)
    out["myinfo_flag"] = yes_no_flag(text_col(out, ["MY_INFO_INDCTR"], "No")).astype(int)
    out["efna_flag"] = yes_no_flag(text_col(out, ["EFNA_INDCTR"], "No")).astype(int)
    out["gio_flag"] = yes_no_flag(text_col(out, ["GIO_IND"], "No")).astype(int)
    out["material_finding_flag"] = out["uw_outcome_clean"].isin(["Revised Terms / Substandard", "Declined", "Postponed"]).astype(int)
    return out


@st.cache_data(show_spinner="Loading Excel / Parquet data...")
def load_from_path(path: str, mtime: float) -> pd.DataFrame:
    path_obj = Path(path)
    if path_obj.suffix.lower() == ".parquet":
        df = pd.read_parquet(path_obj)
    else:
        df = pd.read_excel(path_obj, sheet_name=0, engine="openpyxl")
    return add_derived_fields(df)


@st.cache_data(show_spinner="Loading uploaded file...")
def load_from_bytes(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    suffix = Path(file_name).suffix.lower()
    if suffix == ".csv":
        df = pd.read_csv(io.BytesIO(file_bytes))
    elif suffix == ".parquet":
        df = pd.read_parquet(io.BytesIO(file_bytes))
    else:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, engine="openpyxl")
    return add_derived_fields(df)


def group_rate(df: pd.DataFrame, group_cols, min_count: int = 1) -> pd.DataFrame:
    if isinstance(group_cols, str):
        group_cols = [group_cols]
    if not group_cols or df.empty:
        return pd.DataFrame()
    g = (
        df.groupby(group_cols, dropna=False)
        .agg(policies=("early_claim_flag", "size"), early_claims=("early_claim_flag", "sum"), material_findings=("material_finding_flag", "sum"), medical_exams=("medical_exam_flag", "sum"))
        .reset_index()
    )
    g["early_claim_rate"] = np.where(g["policies"] > 0, g["early_claims"] / g["policies"], np.nan)
    g["material_finding_rate"] = np.where(g["policies"] > 0, g["material_findings"] / g["policies"], np.nan)
    g["medical_exam_rate"] = np.where(g["policies"] > 0, g["medical_exams"] / g["policies"], np.nan)
    baseline = df["early_claim_flag"].mean() if len(df) else np.nan
    g["lift_vs_portfolio"] = np.where(baseline > 0, g["early_claim_rate"] / baseline, np.nan)
    return g[g["policies"] >= min_count].sort_values(["early_claim_rate", "early_claims"], ascending=False)


def contribution_table(df: pd.DataFrame, col: str) -> pd.DataFrame:
    g = group_rate(df, col)
    if g.empty:
        return g
    total_policies = g["policies"].sum()
    total_claims = g["early_claims"].sum()
    g["portfolio_share"] = np.where(total_policies > 0, g["policies"] / total_policies, np.nan)
    g["early_claim_share"] = np.where(total_claims > 0, g["early_claims"] / total_claims, np.nan)
    g["excess_share"] = g["early_claim_share"] - g["portfolio_share"]
    return g


def fmt_pct(x, digits: int = 2) -> str:
    if pd.isna(x):
        return "—"
    return f"{x * 100:.{digits}f}%"


def fmt_n(x) -> str:
    try:
        if pd.isna(x):
            return "—"
        return f"{int(round(float(x))):,}"
    except Exception:
        return "—"


def fmt_money(x) -> str:
    try:
        return f"${float(x):,.0f}"
    except Exception:
        return "—"


def fmt_per_10k(rate) -> str:
    if pd.isna(rate):
        return "—"
    return f"{float(rate) * 10000:,.0f} per 10,000 policies"


def sort_bmi_band_frame(df: pd.DataFrame, col: str = "bmi_band_clean") -> pd.DataFrame:
    if df.empty or col not in df.columns:
        return df
    out = df.copy()
    observed = [str(x) for x in out[col].dropna().astype(str).unique()]
    ordered = [x for x in BMI_BAND_ORDER if x in observed]
    ordered += sorted([x for x in observed if x not in ordered and x != "Unknown"])
    if "Unknown" in observed and "Unknown" not in ordered:
        ordered.append("Unknown")
    out[col] = pd.Categorical(out[col].astype(str), categories=ordered, ordered=True)
    return out.sort_values(col)


def card(title: str, subtitle: str = "", kind: str = ""):
    extra = f" {kind}" if kind else ""
    st.markdown(f"<div class='card{extra}'><div class='card-title'>{title}</div><div class='card-subtitle'>{subtitle}</div>", unsafe_allow_html=True)


def end_card():
    st.markdown("</div>", unsafe_allow_html=True)


def story(eyebrow: str, headline: str, body: str):
    st.markdown(
        f"""
        <div class='story'>
            <div class='eyebrow'>{eyebrow}</div>
            <div class='headline'>{headline}</div>
            <div class='bodycopy'>{body}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def kpi_card(title: str, value: str, subtitle: str = ""):
    sub = f"<div class='kpi-sub'>{subtitle}</div>" if subtitle else ""
    st.markdown(
        f"""
        <div class='kpi-card'>
            <div class='kpi-title'>{title}</div>
            <div class='kpi-value'>{value}</div>
            {sub}
        </div>
        """,
        unsafe_allow_html=True,
    )


def finding_card(tag: str, title: str, copy: str):
    st.markdown(
        f"""
        <div class='mini-finding'>
            <div class='mini-finding-tag'>{escape(tag)}</div>
            <div class='mini-finding-title'>{escape(title)}</div>
            <div class='mini-finding-copy'>{escape(copy)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def scenario_score_panel(score: float, baseline: float, tier: str):
    max_value = max(score, baseline, 0.0001)
    score_width = min(100, score / max_value * 100)
    baseline_width = min(100, baseline / max_value * 100) if not pd.isna(baseline) else 0
    lift = score / baseline if baseline and not pd.isna(baseline) else np.nan
    st.markdown(
        f"""
        <div class='score-panel'>
            <div class='score-label'>Scenario early-claim risk score</div>
            <div class='score-main'>{fmt_pct(score, 3)}</div>
            <div class='score-explain'>
                Easier reading: about <strong>{fmt_per_10k(score)}</strong> for similar profiles.
                Portfolio baseline is about <strong>{fmt_per_10k(baseline)}</strong>.
                This is a review signal, not an automatic underwriting decision.
            </div>
            <div class='score-bars'>
                <div class='score-row'>
                    <div>Scenario</div>
                    <div class='score-track'><div class='score-fill' style='width:{score_width:.1f}%; background:{COLORS["accent"]};'></div></div>
                    <div>{fmt_pct(score, 3)}</div>
                </div>
                <div class='score-row'>
                    <div>Portfolio baseline</div>
                    <div class='score-track'><div class='score-fill' style='width:{baseline_width:.1f}%; background:{COLORS["green"]};'></div></div>
                    <div>{fmt_pct(baseline, 3)}</div>
                </div>
            </div>
            <div class='insight' style='margin-top:14px;'>
                <strong>How to read it:</strong> {tier}. {"" if pd.isna(lift) else f"This scenario is {lift:.2f}x the portfolio baseline."}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def risk_tier_banner(tier: str, subtitle: str):
    st.markdown(
        f"""
        <div class='risk-banner'>
            <div class='risk-banner-label'>Risk tier</div>
            <div class='risk-banner-value'>{tier}</div>
            <div class='risk-banner-sub'>{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def decision_summary(question: str, answer: str, action: str):
    st.markdown(
        f"""
        <div class='decision-grid'>
            <div class='decision-panel'>
                <div class='decision-kicker'>Business question</div>
                <div class='decision-title'>{question}</div>
            </div>
            <div class='decision-panel'>
                <div class='decision-kicker'>What the data says</div>
                <div class='decision-copy'>{answer}</div>
            </div>
            <div class='decision-panel'>
                <div class='decision-kicker'>Underwriting action</div>
                <div class='decision-copy'>{action}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def insight(text: str, kind: str = ""):
    st.markdown(f"<div class='insight {kind}'>{text}</div>", unsafe_allow_html=True)


def section(num: int, label: str):
    st.markdown(f"<div class='section-label'><div class='section-num'>{num}</div><div class='section-text'>{label}</div></div>", unsafe_allow_html=True)


def friendly_label(name: str | None) -> str | None:
    if name is None:
        return None
    return FRIENDLY_LABELS.get(str(name), str(name).replace("_clean", "").replace("_flag", "").replace("_", " ").title())


def friendly_labels_for(*cols: str | None) -> dict[str, str]:
    return {str(col): friendly_label(col) for col in cols if col is not None}


def prettify_plotly_axes(fig: go.Figure, x: str | None = None, y: str | None = None, color: str | None = None):
    if x:
        fig.update_xaxes(title_text=friendly_label(x))
    if y:
        fig.update_yaxes(title_text=friendly_label(y))
    fig.update_layout(legend_title_text=friendly_label(color) if color else "")
    fig.update_traces(
        hoverlabel=dict(bgcolor="white", font_size=13),
        hovertemplate=None,
    )
    return fig


def bar_chart(df: pd.DataFrame, x: str, y: str, color: str | None = None, title: str | None = None, y_is_pct: bool = False, text: str | None = None, height: int = 360):
    if df.empty:
        st.info("No data available for this selection.")
        return
    fig = px.bar(
        df,
        x=x,
        y=y,
        color=color,
        text=text,
        title=title,
        labels=friendly_labels_for(x, y, color, text),
        color_discrete_sequence=[COLORS["accent"], COLORS["f2f"], COLORS["nf2f"], COLORS["green"], COLORS["amber"]],
    )
    fig.update_layout(height=height, margin=dict(l=10, r=10, t=40 if title else 10, b=10), plot_bgcolor="white", paper_bgcolor="white", legend_title_text="")
    prettify_plotly_axes(fig, x, y, color)
    if y_is_pct:
        fig.update_yaxes(tickformat=".2%")
    st.plotly_chart(fig, use_container_width=True)


def line_chart(df: pd.DataFrame, x: str, y: str, color: str | None = None, y_is_pct: bool = False, height: int = 340):
    if df.empty:
        st.info("No data available for this selection.")
        return
    fig = px.line(
        df,
        x=x,
        y=y,
        color=color,
        markers=True,
        labels=friendly_labels_for(x, y, color),
        color_discrete_sequence=[COLORS["accent"], COLORS["f2f"], COLORS["nf2f"], COLORS["green"], COLORS["amber"]],
    )
    fig.update_layout(height=height, margin=dict(l=10, r=10, t=10, b=10), plot_bgcolor="white", paper_bgcolor="white", legend_title_text="")
    prettify_plotly_axes(fig, x, y, color)
    if y_is_pct:
        fig.update_yaxes(tickformat=".2%")
    st.plotly_chart(fig, use_container_width=True)


def heatmap_rate(df: pd.DataFrame, row: str, col: str, min_count: int = 25, title: str = ""):
    g = group_rate(df, [row, col], min_count=min_count)
    if g.empty:
        st.info("Not enough records for this heatmap.")
        return
    pivot = g.pivot(index=row, columns=col, values="early_claim_rate")
    fig = px.imshow(
        pivot,
        text_auto=".2%",
        aspect="auto",
        color_continuous_scale="Reds",
        title=title,
        labels=dict(x=friendly_label(col), y=friendly_label(row), color="Early-claim rate"),
    )
    fig.update_layout(height=380, margin=dict(l=10, r=10, t=40 if title else 10, b=10), plot_bgcolor="white", paper_bgcolor="white")
    fig.update_xaxes(title_text=friendly_label(col), tickangle=0)
    fig.update_yaxes(title_text=friendly_label(row))
    fig.update_traces(
        hovertemplate=f"{friendly_label(col)}: %{{x}}<br>{friendly_label(row)}: %{{y}}<br>Early-claim rate: %{{z:.2%}}<extra></extra>",
        hoverlabel=dict(bgcolor="white", font_size=13),
    )
    st.plotly_chart(fig, use_container_width=True)


def wrapped_table(rows, columns: list[str] | None = None, strong_cols: set[str] | None = None, widths: list[str] | None = None):
    frame = rows.copy() if isinstance(rows, pd.DataFrame) else pd.DataFrame(rows)
    if columns:
        frame = frame[columns]
    if frame.empty:
        st.info("No table rows available.")
        return

    cols = [str(col) for col in frame.columns]
    strong_cols = strong_cols or set()
    widths = widths or ["minmax(0, 1fr)"] * len(cols)
    grid = " ".join(widths)

    header = "".join(f"<div class='wrapped-cell'>{escape(col)}</div>" for col in cols)
    body_rows = []
    for _, row in frame.iterrows():
        cells = []
        for original_col, display_col in zip(frame.columns, cols):
            value = row[original_col]
            text = "" if pd.isna(value) else str(value)
            cls = "wrapped-cell strong" if display_col in strong_cols or original_col in strong_cols else "wrapped-cell"
            cells.append(f"<div class='{cls}' data-label='{escape(display_col, quote=True)}'>{escape(text)}</div>")
        body_rows.append(f"<div class='wrapped-row' style='grid-template-columns:{grid};'>{''.join(cells)}</div>")

    st.markdown(
        f"<div class='wrapped-table'><div class='wrapped-row header' style='grid-template-columns:{grid};'>{header}</div>{''.join(body_rows)}</div>",
        unsafe_allow_html=True,
    )


def two_prop_z_test(e1, n1, e2, n2):
    if min(n1, n2) <= 0:
        return np.nan, np.nan
    p1 = e1 / n1
    p2 = e2 / n2
    p = (e1 + e2) / (n1 + n2)
    se = math.sqrt(p * (1 - p) * (1 / n1 + 1 / n2)) if 0 < p < 1 else 0
    if se == 0:
        return np.nan, np.nan
    z = (p1 - p2) / se
    pval = math.erfc(abs(z) / math.sqrt(2))
    return z, pval


def display_rate_table(df: pd.DataFrame, key_cols, min_count: int = 1, top_n: int = 25):
    g = group_rate(df, key_cols, min_count=min_count).head(top_n).copy()
    if g.empty:
        st.info("No segment table available.")
        return
    for c in ["early_claim_rate", "material_finding_rate", "medical_exam_rate"]:
        if c in g.columns:
            g[c] = g[c].map(lambda x: fmt_pct(x, 3))
    if "lift_vs_portfolio" in g.columns:
        g["lift_vs_portfolio"] = g["lift_vs_portfolio"].map(lambda x: "—" if pd.isna(x) else f"{x:.2f}x")
    g = g.rename(
        columns={
            **friendly_labels_for(*key_cols),
            "policies": "Policies",
            "early_claims": "Early claims",
            "material_findings": "Material findings",
            "medical_exams": "Medical exams",
            "early_claim_rate": "Early-claim rate",
            "material_finding_rate": "Material-finding rate",
            "medical_exam_rate": "Medical-exam rate",
            "lift_vs_portfolio": "Lift vs portfolio",
        }
    )
    widths = ["minmax(130px, 1fr)"] * len(g.columns)
    wrapped_table(g, strong_cols={str(g.columns[0])}, widths=widths)


def filter_options(df: pd.DataFrame, col: str, label: str):
    values = sorted([x for x in df[col].dropna().astype(str).unique() if x and x != "Unknown"])
    return st.sidebar.multiselect(label, values, default=[])


def apply_sidebar_filters(df: pd.DataFrame) -> pd.DataFrame:
    with st.sidebar:
        st.markdown("### Portfolio filters")
        st.caption("Default view includes all policies. Select filters only when you want to narrow the analysis.")
        years = sorted([int(x) for x in df["policy_year"].dropna().unique()]) if "policy_year" in df else []
        selected_years = st.multiselect("Policy application year", years, default=[])
        selected_channel = filter_options(df, "application_channel", "F2F Indicator")
        selected_product = filter_options(df, "product_clean", "Product group")
        selected_uw = filter_options(df, "uw_outcome_clean", "UW outcome")
        st.session_state["min_count"] = DEFAULT_MIN_SEGMENT_SIZE
    out = df.copy()
    if selected_years:
        out = out[out["policy_year"].isin(selected_years)]
    if selected_channel:
        out = out[out["application_channel"].isin(selected_channel)]
    if selected_product:
        out = out[out["product_clean"].isin(selected_product)]
    if selected_uw:
        out = out[out["uw_outcome_clean"].isin(selected_uw)]
    with st.sidebar:
        st.markdown(f"<div class='small-mono'>Current view: {fmt_n(len(out))} of {fmt_n(len(df))} policies</div>", unsafe_allow_html=True)
    return out


def load_data_ui() -> tuple[pd.DataFrame, str, datetime | None]:
    with st.sidebar:
        st.markdown("### Data source")
        uploaded = st.file_uploader("Upload updated master data", type=["xlsx", "xls", "csv", "parquet"])
        auto_refresh = st.toggle("Auto-refresh every 60 seconds", value=False)
        if auto_refresh and st_autorefresh is not None:
            st_autorefresh(interval=60_000, key="data_autorefresh")
        elif auto_refresh:
            st.caption("Install streamlit-autorefresh to enable timed refresh. Manual refresh still works.")
        if st.button("Clear cache and reload now"):
            st.cache_data.clear()
            st.rerun()
    if uploaded is not None:
        st.sidebar.success(f"Using uploaded file: {uploaded.name}. All tabs and insights recalculate from this file.")
        df = load_from_bytes(uploaded.getvalue(), uploaded.name)
        return df, uploaded.name, None
    if DEFAULT_PARQUET_PATH.exists():
        path = DEFAULT_PARQUET_PATH
    else:
        path = DEFAULT_DATA_PATH
    if not path.exists():
        st.error(f"No data file found. Put your file here: {DEFAULT_DATA_PATH}")
        st.stop()
    mtime = path.stat().st_mtime
    df = load_from_path(str(path), mtime)
    return df, str(path.name), datetime.fromtimestamp(mtime)


def topbar(source_name: str, last_modified: datetime | None, df: pd.DataFrame):
    total = len(df)
    ec = int(df["early_claim_flag"].sum())
    sub = f"{fmt_n(total)} policies · {fmt_n(ec)} early claims · source: {source_name}"
    if last_modified:
        sub += f" · updated {last_modified:%Y-%m-%d %H:%M}"
    st.markdown(
        f"""
        <div class='topbar'>
            <div><span class='topbar-title'>Claims AI Analytics</span><span class='topbar-sub'>UC2 - Optimising Underwriting Guidelines</span></div>
            <div><span class='badge'>LIVE DASHBOARD</span></div>
        </div>
        <div class='small-mono'>{sub}</div>
        """,
        unsafe_allow_html=True,
    )


def rate_lift(num_rate: float, den_rate: float) -> float:
    if pd.isna(num_rate) or pd.isna(den_rate) or den_rate <= 0:
        return np.nan
    return float(num_rate / den_rate)


def render_overview(df: pd.DataFrame):
    story("UC2 · Executive View", "What should the underwriting team do with this dashboard?", "This dashboard turns policy, underwriting, disclosure, BMI, medical evidence, and claims data into a decision-support view for early-claim risk. It is not designed to auto-decline customers. It is designed to show where the current rules are working, where disclosure quality may be weak, and which segments deserve extra evidence before acceptance.")
    total = len(df)
    ec = int(df["early_claim_flag"].sum())
    rate = df["early_claim_flag"].mean() if total else np.nan
    me_yield = df.loc[df["medical_exam_flag"].eq(1), "material_finding_flag"].mean()
    non_me_yield = df.loc[df["medical_exam_flag"].eq(0), "material_finding_flag"].mean()
    nf2f_rate = df.loc[df["application_channel"].eq("NF2F"), "early_claim_flag"].mean()
    f2f_rate = df.loc[df["application_channel"].eq("F2F"), "early_claim_flag"].mean()
    nf2f_lift = rate_lift(nf2f_rate, f2f_rate)
    leak = df[df["early_claim_flag"].eq(1) & df["disclosure_flag"].eq(0) & df["medical_exam_flag"].eq(0)]
    if pd.notna(nf2f_lift) and nf2f_lift >= 1.25:
        overview_action = "Prioritise segment-level NF2F evidence checks, then test BMI and disclosure-rule refinements."
    elif len(leak) >= max(5, ec * 0.25):
        overview_action = "Prioritise disclosure-quality and medical-evidence leakage checks before broad channel changes."
    else:
        overview_action = "Use the rule tabs to monitor targeted segments; the current view does not support broad blanket action."
    decision_summary(
        "Where should underwriting focus first?",
        f"The current view has {fmt_n(total)} policies and {fmt_n(ec)} early claims. The portfolio early-claim rate is {fmt_pct(rate, 3)} ({fmt_per_10k(rate)}).",
        overview_action,
    )
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Policies analysed", fmt_n(total), "Filter context: after current filters")
    with c2:
        kpi_card("Early claims", fmt_n(ec), f"Early-claim rate: {fmt_pct(rate, 3)}")
    with c3:
        kpi_card(
            "NF2F vs F2F lift",
            "—" if pd.isna(nf2f_rate) or pd.isna(f2f_rate) or f2f_rate == 0 else f"{nf2f_rate / f2f_rate:.2f}x",
            f"Rates compared: NF2F {fmt_pct(nf2f_rate,3)} vs F2F {fmt_pct(f2f_rate,3)}",
        )
    with c4:
        kpi_card("Medical exam yield", fmt_pct(me_yield, 1), f"No-exam yield: {fmt_pct(non_me_yield,1)}")

    section(1, "What the data says")
    channel_table = contribution_table(df, "application_channel")
    uw_table = group_rate(df, "uw_outcome_clean")
    std_row = uw_table[uw_table["uw_outcome_clean"].eq("Standard")]
    rev_row = uw_table[uw_table["uw_outcome_clean"].eq("Revised Terms / Substandard")]
    std_rate = float(std_row["early_claim_rate"].iloc[0]) if not std_row.empty else np.nan
    rev_rate = float(rev_row["early_claim_rate"].iloc[0]) if not rev_row.empty else np.nan
    rev_lift = rate_lift(rev_rate, std_rate)
    bmi_table = group_rate(df[df["bmi"].notna()], "bmi_band_clean", min_count=25)
    bmi_table = sort_bmi_band_frame(bmi_table)
    top_bmi = bmi_table.sort_values("early_claim_rate", ascending=False).head(1)
    top_bmi_label = str(top_bmi["bmi_band_clean"].iloc[0]) if not top_bmi.empty else "not enough BMI data"
    top_bmi_rate = float(top_bmi["early_claim_rate"].iloc[0]) if not top_bmi.empty else np.nan
    me_lift = rate_lift(me_yield, non_me_yield)
    f1, f2 = st.columns(2)
    with f1:
        finding_card(
            "B2 · F2F Indicator",
            "NF2F risk is checked, not assumed",
            "NF2F vs F2F lift is "
            + ("not available" if pd.isna(nf2f_lift) else f"{nf2f_lift:.2f}x")
            + ". Use B2 to see whether the gap is concentrated by age, sum assured, product, or UW outcome.",
        )
    with f2:
        finding_card(
            "B4 · UW outcome",
            "UW decisions should separate risk",
            "Revised/substandard vs Standard lift is "
            + ("not available" if pd.isna(rev_lift) else f"{rev_lift:.2f}x")
            + ". B4 checks whether higher-risk decisions actually carry higher early-claim rates.",
        )
    f3, f4 = st.columns(2)
    with f3:
        finding_card(
            "B1 · BMI threshold",
            f"Highest observed BMI band: {top_bmi_label}",
            f"Observed early-claim rate is {fmt_pct(top_bmi_rate, 3)} in that band. Use B1 to test whether a lower ME trigger catches more risk without too much friction.",
        )
    with f4:
        finding_card(
            "B5 · Disclosure yield",
            "Medical exams are a risk signal, not a causal test",
            "ME material-finding yield is "
            + fmt_pct(me_yield, 1)
            + " vs "
            + fmt_pct(non_me_yield, 1)
            + " without ME"
            + ("" if pd.isna(me_lift) else f" ({me_lift:.1f}x)."),
        )
    section(2, "Where underwriting should look first")
    if not channel_table.empty and set(["F2F", "NF2F"]).intersection(set(channel_table["application_channel"])):
        nf = channel_table[channel_table["application_channel"].eq("NF2F")]
        if not nf.empty:
            nf = nf.iloc[0]
            insight(f"<strong>Channel view:</strong> Use <strong>F2F Indicator</strong> as the official F2F/NF2F split. NF2F portfolio share is {fmt_pct(nf['portfolio_share'],1)} and early-claim share is {fmt_pct(nf['early_claim_share'],1)}. The right action is targeted evidence for high-risk NF2F segments, not a blanket channel surcharge.")
    else:
        insight("<strong>Channel view:</strong> F2F Indicator could not be clearly split after filters. Check the uploaded file mapping for F2F Indicator / Channel.", "warn")

    c1, c2 = st.columns([1.05, 0.95])
    with c1:
        card("Early-claim contribution by F2F Indicator", "This checks whether NF2F is over-represented in early claims compared with its portfolio share.")
        plot_df = channel_table.copy()
        if not plot_df.empty:
            share_long = plot_df.melt(id_vars="application_channel", value_vars=["portfolio_share", "early_claim_share"], var_name="share_type", value_name="share")
            share_long["share_type"] = share_long["share_type"].replace({"portfolio_share": "Portfolio share", "early_claim_share": "Early-claim share"})
            bar_chart(share_long, "application_channel", "share", "share_type", y_is_pct=True, height=320)
        end_card()
    with c2:
        card("Segments to review first", "This is an advanced ranking. Open it only when you want to see which channel, age, and sum-assured groups are driving the headline risk.")
        st.caption("How to read it: start with Policies, then Early claims, then Early-claim rate. Lift vs portfolio above 1.00x means the segment is higher than the current portfolio average.")
        with st.expander("Show technical segment ranking"):
            display_rate_table(df, ["application_channel", "age_band_clean", "sa_band_clean"], st.session_state.get("min_count", 100), top_n=12)
        end_card()

    section(3, "Evidence views")
    c1, c2 = st.columns(2)
    with c1:
        card("Age band x F2F Indicator", "Darker cells indicate higher observed early-claim incidence.")
        heatmap_rate(df, "age_band_clean", "application_channel", min_count=st.session_state.get("min_count", 100))
        end_card()
    with c2:
        card("Sum assured band x F2F Indicator", "Use this to identify where NF2F needs stronger evidence requirements.")
        heatmap_rate(df, "sa_band_clean", "application_channel", min_count=st.session_state.get("min_count", 100))
        end_card()

    section(4, "Recommended actions")
    action_rows = []
    if pd.notna(nf2f_lift) and nf2f_lift >= 1.25:
        action_rows.append(["High", "Target NF2F segments with excess early-claim lift", f"NF2F is {nf2f_lift:.2f}x F2F in the current data. Use B2 segment views before changing any blanket rule."])
    else:
        action_rows.append(["Monitor", "Keep NF2F changes targeted", "The current data does not justify treating all NF2F cases as higher risk. Review only concentrated age, SA, product, or UW-outcome pockets."])
    if df["bmi"].notna().mean() >= 0.5:
        action_rows.append(["High", "Test BMI evidence thresholds", f"BMI coverage is {fmt_pct(df['bmi'].notna().mean(), 1)}. Use B1 to estimate extra ME volume, cost, and risk capture."])
    else:
        action_rows.append(["Medium", "Fix BMI data coverage before rule changes", f"BMI coverage is only {fmt_pct(df['bmi'].notna().mean(), 1)}, so BMI threshold conclusions may be incomplete."])
    if pd.notna(std_rate) and std_rate >= max(rate * 1.25, 0.003):
        action_rows.append(["High", "Review Standard-case leakage", f"Standard early-claim rate is {fmt_pct(std_rate, 3)}, which is high relative to the portfolio baseline of {fmt_pct(rate, 3)}."])
    else:
        action_rows.append(["Medium", "Monitor Standard approvals", "Standard-case leakage is not the dominant signal under the current filters, but B4 should be reviewed after each data refresh."])
    if len(leak):
        action_rows.append(["High", "Use post-claim ICDs to refine disclosure questions", f"{fmt_n(len(leak))} early-claim cases have no disclosure flag and no medical exam. Use B5 to identify condition themes."])
    else:
        action_rows.append(["Monitor", "Maintain disclosure-quality monitoring", "No no-disclosure/no-exam early-claim leakage is visible under the current filters."])
    wrapped_table(
        pd.DataFrame(action_rows, columns=["Priority", "UW action", "Why it matters"]),
        strong_cols={"Priority", "UW action"},
        widths=["120px", "minmax(240px, .85fr)", "minmax(360px, 1.4fr)"],
    )


def render_b1_bmi(df: pd.DataFrame):
    story("B1.1 · BMI threshold simulator", "Would lowering the medical-exam BMI trigger catch more risk without overwhelming the portfolio?", "Use this page to test the rule described in the use-case: lower the BMI threshold from the current rule to a proposed rule, restrict by age and disclosure profile, then quantify additional medical exams, estimated evidence cost, incremental material findings, early-claim catch proxy, withdrawal impact, and whether the change is risk-neutral or risk-improving.")
    usable = df[df["bmi"].notna()].copy()
    decision_summary(
        "Should the BMI medical-exam trigger be lowered?",
        f"This tab uses {fmt_n(len(usable))} policies with usable BMI. It compares extra evidence volume against material-finding proxy, observed early-claim capture, estimated cost, and possible withdrawals.",
        "Use this as a retrospective what-if, not causal proof. Pilot any lower threshold only where the added evidence justifies operational load and customer friction.",
    )
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Policies with BMI", fmt_n(len(usable)), f"BMI coverage: {fmt_pct(len(usable)/len(df) if len(df) else np.nan,1)} of filtered data")
    with c2:
        kpi_card("Current ME coverage", fmt_pct(usable["medical_exam_flag"].mean(), 1), "Policies already sent for medical exam")
    with c3:
        kpi_card("ME material-finding yield", fmt_pct(usable.loc[usable["medical_exam_flag"].eq(1), "material_finding_flag"].mean(), 1), "Material findings among ME cases")
    with c4:
        kpi_card("No-ME material-finding yield", fmt_pct(usable.loc[usable["medical_exam_flag"].eq(0), "material_finding_flag"].mean(), 1), "Material findings among non-ME cases")

    section(1, "Data readiness and BMI risk")
    c1, c2 = st.columns(2)
    with c1:
        card("Early-claim rate by BMI band", "This shows whether a BMI range has observed claim risk above the portfolio average.")
        g = group_rate(usable, "bmi_band_clean", min_count=25)
        g = sort_bmi_band_frame(g)
        if g.empty:
            st.info("No usable BMI band data after the current filters. Clear filters or check the BMI columns in the uploaded file.")
        else:
            bar_chart(g, "bmi_band_clean", "early_claim_rate", y_is_pct=True, text="early_claims", height=340)
        end_card()
    with c2:
        card("Medical-exam coverage and material findings by BMI band", "This checks whether high-BMI bands are actually receiving evidence and whether the evidence changes UW terms.")
        g2 = group_rate(usable, "bmi_band_clean", min_count=25)
        g2 = sort_bmi_band_frame(g2)
        if g2.empty:
            st.info("No usable BMI band data after the current filters. Clear filters or check the BMI columns in the uploaded file.")
        else:
            long = g2.melt(id_vars="bmi_band_clean", value_vars=["medical_exam_rate", "material_finding_rate"], var_name="metric", value_name="rate")
            long["metric"] = long["metric"].replace({"medical_exam_rate": "Medical-exam coverage", "material_finding_rate": "Material-finding rate"})
            bar_chart(long, "bmi_band_clean", "rate", "metric", y_is_pct=True, height=340)
        end_card()

    section(2, "Rule simulator")
    with st.container(border=True):
        st.markdown("<div class='card-title' style='font-size:1.05rem;'>Rule assumptions</div>", unsafe_allow_html=True)
        st.caption("Set the current rule, proposed rule, and operating assumptions for the medical-exam threshold test.")
        threshold_col, operating_col = st.columns([1, 1.25], gap="large")
        with threshold_col:
            st.markdown("**BMI trigger**")
            t1, t2 = st.columns(2)
            with t1:
                current_threshold = st.number_input("Current trigger", min_value=15.0, max_value=50.0, value=30.0, step=0.5, format="%.1f")
            with t2:
                proposed_threshold = st.number_input("Proposed trigger", min_value=15.0, max_value=50.0, value=27.5, step=0.5, format="%.1f")
        with operating_col:
            st.markdown("**Eligibility and friction**")
            age_min, age_max = st.slider("Age range", 0, 90, (31, 55), help="Only applications within this age range are included in the what-if test.")
            o1, o2 = st.columns(2)
            with o1:
                exam_cost = st.number_input("ME cost per case", min_value=0.0, value=150.0, step=25.0, format="%.0f")
            with o2:
                withdrawal_pct = st.slider("Withdrawal rate (%)", 0, 50, 5, 1, help="Estimated share of customers who may withdraw after an extra medical request.")
                assume_withdrawal = withdrawal_pct / 100
        no_disclosed_only = st.toggle("Limit to customers with no disclosed health impairment", value=True)

        base = usable[(usable["age"].between(age_min, age_max, inclusive="both"))]
        if no_disclosed_only:
            base = base[base["disclosure_flag"].eq(0)]
        current_trigger = base["bmi"].ge(current_threshold)
        proposed_trigger = base["bmi"].ge(proposed_threshold)
        additional = base[proposed_trigger & ~current_trigger & base["medical_exam_flag"].eq(0)].copy()
        extra_me = len(additional)
        extra_cost = extra_me * exam_cost
        hist_yield = usable.loc[usable["medical_exam_flag"].eq(1), "material_finding_flag"].mean()
        no_me_yield = usable.loc[usable["medical_exam_flag"].eq(0), "material_finding_flag"].mean()
        yield_uplift = max((hist_yield if pd.notna(hist_yield) else 0) - (no_me_yield if pd.notna(no_me_yield) else 0), 0)
        expected_new_findings = extra_me * yield_uplift
        captured_claims = int(additional["early_claim_flag"].sum())
        expected_withdrawals = extra_me * assume_withdrawal
        substandard = int(additional["uw_outcome_clean"].eq("Revised Terms / Substandard").sum())
        declined = int(additional["uw_outcome_clean"].eq("Declined").sum())
        postponed = int(additional["uw_outcome_clean"].eq("Postponed").sum())
        disclosure_scope = "No disclosed health impairment only" if no_disclosed_only else "All disclosure profiles"
        st.markdown(
            f"""
            <div class='assumption-strip'>
                <div class='assumption-item'>
                    <div class='assumption-label'>BMI rule change</div>
                    <div class='assumption-value'>{current_threshold:.1f} to {proposed_threshold:.1f}</div>
                    <div class='assumption-note'>Newly triggered cases sit between the two thresholds.</div>
                </div>
                <div class='assumption-item'>
                    <div class='assumption-label'>Population</div>
                    <div class='assumption-value'>Ages {age_min}-{age_max}</div>
                    <div class='assumption-note'>{disclosure_scope}</div>
                </div>
                <div class='assumption-item'>
                    <div class='assumption-label'>Evidence cost</div>
                    <div class='assumption-value'>{fmt_money(exam_cost)} per ME</div>
                    <div class='assumption-note'>Applied to each extra medical exam.</div>
                </div>
                <div class='assumption-item'>
                    <div class='assumption-label'>Friction assumption</div>
                    <div class='assumption-value'>{fmt_pct(assume_withdrawal, 0)} withdrawal</div>
                    <div class='assumption-note'>Used to estimate potential drop-off.</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.divider()
        st.markdown("<div class='card-title' style='font-size:1.05rem;'>Calculated impact</div>", unsafe_allow_html=True)
        if extra_me:
            decision_summary(
                "What happens if this threshold is changed?",
                f"The selected rule would trigger {fmt_n(extra_me)} extra medical exams, costing about {fmt_money(extra_cost)}. Historically, this newly triggered group contains {fmt_n(captured_claims)} early claims.",
                "This is a historical proxy, not a randomized estimate. If expected findings materially exceed withdrawals, pilot the rule on targeted segments first. If not, narrow it by channel, age, sum assured, or disclosure status.",
            )

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            kpi_card("Additional MEs", fmt_n(extra_me), "Extra medical exams triggered")
        with c2:
            kpi_card("Estimated ME cost", fmt_money(extra_cost), "Assumed cost per exam")
        with c3:
            kpi_card("Estimated findings proxy", fmt_n(expected_new_findings), f"Yield uplift assumption: {fmt_pct(yield_uplift,1)}")
        with c4:
            kpi_card("Historical ECs captured", fmt_n(captured_claims), f"EC rate in added ME group: {fmt_pct(additional['early_claim_flag'].mean() if extra_me else np.nan, 3)}")
        with c5:
            kpi_card("Potential withdrawals", fmt_n(expected_withdrawals), f"Withdrawal assumption: {fmt_pct(assume_withdrawal,1)}")

        if extra_me == 0:
            insight("No additional cases are triggered by this rule after the selected filters. Try widening the age band or lowering the proposed threshold.", "warn")
        elif expected_new_findings > expected_withdrawals:
            insight(f"<strong>Conclusion:</strong> This rule is likely <strong>risk-improving</strong> under the selected assumptions because expected new material findings ({fmt_n(expected_new_findings)}) exceed expected withdrawals ({fmt_n(expected_withdrawals)}). Review operational capacity and customer friction before implementation.", "good")
        else:
            insight(f"<strong>Conclusion:</strong> This rule may be <strong>risk-neutral or operationally heavy</strong> under the selected assumptions because expected findings ({fmt_n(expected_new_findings)}) do not clearly exceed potential withdrawals ({fmt_n(expected_withdrawals)}). Consider a more targeted rule, e.g. NF2F, high SA, or older age only.", "warn")

    c1, c2 = st.columns(2)
    with c1:
        card("Additional triggered cases by F2F Indicator", "This tells UW where the lower BMI trigger would create extra evidence volume.")
        if extra_me:
            g = group_rate(additional, ["application_channel"])
            bar_chart(g, "application_channel", "policies", text="policies", height=300)
        end_card()
    with c2:
        card("Additional triggered cases: observed outcomes", "Historical outcomes among the cases that would have been newly triggered.")
        outcome = pd.DataFrame({"Outcome": ["Substandard / Revised", "Declined", "Postponed", "Early claims"], "Cases": [substandard, declined, postponed, captured_claims]})
        bar_chart(outcome, "Outcome", "Cases", text="Cases", height=300)
        end_card()

    section(3, "Optional segment detail")
    if extra_me:
        st.caption("This table is only for analysts who want to inspect the newly triggered population behind the KPI cards above.")
        with st.expander("Show newly triggered case profile"):
            display_rate_table(additional, ["application_channel", "age_band_clean", "sa_band_clean", "product_clean"], min_count=1, top_n=30)
    else:
        st.info("No additional cases to profile.")


def render_b2_f2f(df: pd.DataFrame):
    story("B2.1 · F2F Indicator / NF2F analysis", "Does NF2F show higher early-claim incidence, and where is the excess concentrated?", "This page uses the actual <strong>F2F Indicator</strong> column as the official split. The underwriting question is not simply 'is NF2F bad?' The senior-underwriter question is: after controlling for age, sum assured, product, and UW outcome, which NF2F segments need more evidence or stronger disclosure capture?")
    g = contribution_table(df, "application_channel")
    f2f = g[g["application_channel"].eq("F2F")].iloc[0] if not g[g["application_channel"].eq("F2F")].empty else None
    nf2f = g[g["application_channel"].eq("NF2F")].iloc[0] if not g[g["application_channel"].eq("NF2F")].empty else None
    if f2f is not None and nf2f is not None:
        z, p = two_prop_z_test(nf2f["early_claims"], nf2f["policies"], f2f["early_claims"], f2f["policies"])
        lift = nf2f["early_claim_rate"] / f2f["early_claim_rate"] if f2f["early_claim_rate"] else np.nan
    else:
        z, p, lift = np.nan, np.nan, np.nan
    ratio = np.nan
    if nf2f is not None:
        ratio = nf2f["early_claim_share"] / nf2f["portfolio_share"] if nf2f["portfolio_share"] else np.nan
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("F2F early-claim rate", fmt_pct(f2f["early_claim_rate"], 3) if f2f is not None else "—", f"Early claims / policies: {fmt_n(f2f['early_claims'])} / {fmt_n(f2f['policies'])}" if f2f is not None else "")
    with c2:
        kpi_card("NF2F early-claim rate", fmt_pct(nf2f["early_claim_rate"], 3) if nf2f is not None else "—", f"Early claims / policies: {fmt_n(nf2f['early_claims'])} / {fmt_n(nf2f['policies'])}" if nf2f is not None else "")
    with c3:
        kpi_card("NF2F relative risk", "—" if pd.isna(lift) else f"{lift:.2f}x", "Formula: NF2F rate / F2F rate")
    with c4:
        kpi_card("Significance check", "—" if pd.isna(p) else f"p={p:.4f}", "Test used: two-proportion z-test")
    if nf2f is not None and f2f is not None:
        if pd.notna(lift) and lift >= 1.25 and pd.notna(p) and p < 0.05:
            answer = f"NF2F is higher than F2F in this view: {fmt_pct(nf2f['early_claim_rate'], 3)} vs {fmt_pct(f2f['early_claim_rate'], 3)}. The difference is statistically visible, with {ratio:.2f}x contribution versus portfolio share."
            action = "Do not treat all NF2F as bad. Tighten evidence only in the age, sum assured, product, and UW-outcome pockets shown below."
        else:
            answer = "NF2F does not show a large enough standalone signal under the current filters to justify a blanket policy change."
            action = "Use targeted monitoring and segment-level checks instead of changing the whole NF2F rule."
        decision_summary("Does NF2F create materially higher early-claim risk?", answer, action)
        if min(f2f["early_claims"], nf2f["early_claims"]) < 10:
            insight("One channel has fewer than 10 early claims under the current filters. Treat the p-value and relative-risk estimate as directional, not definitive.", "warn")

    if nf2f is not None:
        insight(f"<strong>Contribution check:</strong> NF2F contributes {fmt_pct(nf2f['early_claim_share'],1)} of early claims vs {fmt_pct(nf2f['portfolio_share'],1)} of policies. That is {ratio:.2f}x its portfolio share. If this is materially below 2.0x, the result supports targeted rule changes rather than a blanket NF2F loading.")

    section(1, "Main comparison")
    c1, c2 = st.columns(2)
    with c1:
        card("Portfolio share vs early-claim share", "If early-claim share is higher than portfolio share, the channel is over-represented in early claims.")
        if not g.empty:
            long = g.melt(id_vars="application_channel", value_vars=["portfolio_share", "early_claim_share"], var_name="share_type", value_name="share")
            long["share_type"] = long["share_type"].replace({"portfolio_share": "Portfolio share", "early_claim_share": "Early-claim share"})
            bar_chart(long, "application_channel", "share", "share_type", y_is_pct=True, height=330)
        end_card()
    with c2:
        card("Early-claim rate by F2F Indicator", "This is the direct rate comparison: claims per policy by application channel.")
        bar_chart(g, "application_channel", "early_claim_rate", y_is_pct=True, text="early_claims", height=330)
        end_card()

    section(2, "Where the gap is concentrated")
    st.caption("These are stratified comparisons, not causal adjustments. Use them to locate concentrated pockets before recommending rule changes.")
    c1, c2 = st.columns(2)
    with c1:
        card("By age band and F2F Indicator", "Check whether NF2F risk is concentrated in older ages.")
        age = group_rate(df, ["age_band_clean", "application_channel"], st.session_state.get("min_count", 100))
        line_chart(age.sort_values("age_band_clean"), "age_band_clean", "early_claim_rate", "application_channel", y_is_pct=True, height=350)
        end_card()
    with c2:
        card("By sum assured band and F2F Indicator", "Check whether higher coverage is the real driver behind channel risk.")
        sa = group_rate(df, ["sa_band_clean", "application_channel"], st.session_state.get("min_count", 100))
        line_chart(sa.sort_values("sa_band_clean"), "sa_band_clean", "early_claim_rate", "application_channel", y_is_pct=True, height=350)
        end_card()

    c1, c2 = st.columns(2)
    with c1:
        card("By UW outcome and F2F Indicator", "If NF2F Standard is high, the likely issue is disclosure quality rather than the revised-terms process.")
        uw = group_rate(df, ["uw_outcome_clean", "application_channel"], st.session_state.get("min_count", 100))
        bar_chart(uw, "uw_outcome_clean", "early_claim_rate", "application_channel", y_is_pct=True, height=350)
        end_card()
    with c2:
        card("By product and F2F Indicator", "Controls for product mix, which often explains apparent channel differences.")
        prod = group_rate(df, ["product_clean", "application_channel"], st.session_state.get("min_count", 100)).head(30)
        bar_chart(prod, "product_clean", "early_claim_rate", "application_channel", y_is_pct=True, height=350)
        end_card()

    section(3, "Trend check")
    if "policy_year" in df.columns:
        trend = group_rate(df, ["policy_year", "application_channel"], min_count=1)
        if trend.empty:
            st.info("No yearly trend is available for the current filters.")
        else:
            card("Early-claim rate by application year", "This mirrors the trend view in the original dashboard, but recalculates from the uploaded file and current filters.")
            line_chart(trend.sort_values(["policy_year", "application_channel"]), "policy_year", "early_claim_rate", "application_channel", y_is_pct=True, height=340)
            end_card()
    else:
        st.info("No application-year field is available in this file.")

    section(4, "Optional technical ranking")
    st.caption("The charts above are the main view. This table is for analysts who need the exact segment ranking behind those charts.")
    with st.expander("Show technical rule-tightening table"):
        display_rate_table(df, ["application_channel", "age_band_clean", "sa_band_clean", "uw_outcome_clean", "product_clean"], st.session_state.get("min_count", 100), top_n=35)


def render_b4_uw(df: pd.DataFrame):
    story("B4.1 · UW outcome risk", "Is underwriting identifying the right applicants as higher risk?", "A healthy underwriting process should show higher observed early-claim incidence among applicants receiving revised terms or other adverse decisions than applicants accepted at Standard. If Standard cases still have meaningful early-claim leakage, the team should look for non-disclosure signals, product/channel mix, and evidence thresholds.")
    g = group_rate(df, "uw_outcome_clean")
    std = g[g["uw_outcome_clean"].eq("Standard")].iloc[0] if not g[g["uw_outcome_clean"].eq("Standard")].empty else None
    rev = g[g["uw_outcome_clean"].eq("Revised Terms / Substandard")].iloc[0] if not g[g["uw_outcome_clean"].eq("Revised Terms / Substandard")].empty else None
    rev_lift = rev["early_claim_rate"] / std["early_claim_rate"] if std is not None and rev is not None and std["early_claim_rate"] else np.nan
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Standard EC rate", fmt_pct(std["early_claim_rate"], 3) if std is not None else "—", "Early-claim rate for standard decisions")
    with c2:
        kpi_card("Revised/Substandard EC rate", fmt_pct(rev["early_claim_rate"], 3) if rev is not None else "—", "Early-claim rate for revised/substandard decisions")
    with c3:
        kpi_card("Revised vs Standard lift", "—" if pd.isna(rev_lift) else f"{rev_lift:.2f}x", "Formula: revised rate / standard rate")
    with c4:
        kpi_card("Non-standard decision rate", fmt_pct(df["material_finding_flag"].mean(), 1), "Share with revised, declined, or postponed outcome")
    if std is not None and rev is not None:
        if pd.notna(rev_lift) and rev_lift >= 1.25:
            answer = f"Revised/substandard cases show higher early-claim incidence than Standard cases ({fmt_pct(rev['early_claim_rate'], 3)} vs {fmt_pct(std['early_claim_rate'], 3)}). This means underwriting is directionally separating risk."
            action = "The next focus should be Standard-case leakage: Standard approvals that still have high early-claim rates by channel, age, SA, and product."
        else:
            answer = f"The revised/substandard rate is not clearly above Standard in this filtered view ({fmt_pct(rev['early_claim_rate'], 3)} vs {fmt_pct(std['early_claim_rate'], 3)})."
            action = "Before changing UW rules, check whether product mix, small samples, or missing outcome mapping is weakening the signal."
        decision_summary("Are UW outcomes separating higher-risk applicants?", answer, action)

    if std is not None and rev is not None and rev["early_claim_rate"] > std["early_claim_rate"]:
        insight("<strong>Interpretation:</strong> Revised/substandard lives have higher observed early-claim incidence than Standard lives, which means the UW process is directionally identifying risk. The next question is whether some Standard segments still need better evidence capture.", "good")
    else:
        insight("<strong>Interpretation:</strong> The revised-vs-standard gap is weak under the current filters. A senior underwriter should inspect Standard early claims by channel, SA, and product before changing rules.", "warn")

    section(1, "Main outcome check")
    c1, c2 = st.columns([1.15, .85])
    with c1:
        card("Early-claim rate by UW outcome", "This is the core test of whether higher-risk decisions actually correspond to higher claims.")
        bar_chart(g, "uw_outcome_clean", "early_claim_rate", y_is_pct=True, text="early_claims", height=360)
        end_card()
    with c2:
        card("10,000 application funnel", "A practical underwriting view: expected decisions and early claims per 10,000 cases.")
        if not g.empty:
            funnel = g.copy()
            funnel["per_10k"] = funnel["policies"] / funnel["policies"].sum() * 10000
            funnel["expected_ec_per_10k"] = funnel["per_10k"] * funnel["early_claim_rate"]
            order = {
                "Standard": 0,
                "Revised Terms / Substandard": 1,
                "Declined": 2,
                "Postponed": 3,
                "Withdrawn": 4,
                "Other / Unknown": 5,
            }
            funnel["sort_order"] = funnel["uw_outcome_clean"].map(order).fillna(99)
            funnel = funnel.sort_values(["sort_order", "policies"], ascending=[True, False])
            rows = []
            for _, row in funnel.iterrows():
                outcome = str(row["uw_outcome_clean"])
                rows.append(
                    f"""
                    <div class='funnel-step'>
                        <div class='funnel-step-title'>{fmt_n(row['per_10k'])} per 10,000 -> {escape(outcome)}</div>
                        <div class='funnel-step-copy'>Expected early claims: {row['expected_ec_per_10k']:.2f} · Observed rate: {fmt_pct(row['early_claim_rate'], 3)} · Source policies: {fmt_n(row['policies'])}</div>
                    </div>
                    """
                )
            st.markdown("<div class='funnel-stack'>" + "".join(rows) + "</div>", unsafe_allow_html=True)
            with st.expander("Show funnel table"):
                funnel_display = funnel[["uw_outcome_clean", "per_10k", "expected_ec_per_10k", "early_claim_rate"]].assign(
                    per_10k=lambda x: x["per_10k"].round(0).astype(int),
                    expected_ec_per_10k=lambda x: x["expected_ec_per_10k"].round(2),
                    early_claim_rate=lambda x: x["early_claim_rate"].map(lambda v: fmt_pct(v, 3)),
                ).rename(columns={
                    "uw_outcome_clean": "Underwriting outcome",
                    "per_10k": "Cases per 10,000",
                    "expected_ec_per_10k": "Expected early claims per 10,000",
                    "early_claim_rate": "Early-claim rate",
                })
                wrapped_table(
                    funnel_display,
                    strong_cols={"Underwriting outcome"},
                    widths=["minmax(180px, 1fr)", "minmax(130px, .8fr)", "minmax(190px, 1fr)", "minmax(130px, .8fr)"],
                )
        end_card()

    section(2, "Standard-case leakage checks")
    c1, c2 = st.columns(2)
    with c1:
        card("UW outcome by F2F Indicator", "Looks for Standard NF2F pockets with elevated early-claim rates.")
        heatmap_rate(df, "uw_outcome_clean", "application_channel", min_count=st.session_state.get("min_count", 100))
        end_card()
    with c2:
        card("UW outcome by product", "Checks whether product mix is driving revised-terms risk.")
        heatmap_rate(df, "uw_outcome_clean", "product_clean", min_count=st.session_state.get("min_count", 100))
        end_card()

    section(3, "Optional leakage table")
    standard = df[df["uw_outcome_clean"].eq("Standard")]
    st.caption("This table is for investigating Standard approvals with elevated early-claim rates. It is hidden by default because the chart and decision summary above are the main readout.")
    with st.expander("Show Standard-case leakage table"):
        display_rate_table(standard, ["application_channel", "age_band_clean", "sa_band_clean", "product_clean"], st.session_state.get("min_count", 100), top_n=35)


def split_icd_values(s: pd.Series) -> pd.Series:
    vals = []
    for item in s.dropna().astype(str):
        if item.strip().lower() in ["", "unknown", "#n/a", "nan", "none"]:
            continue
        pieces = re.split(r"[,;/|]+", item)
        for p in pieces:
            p = p.strip()
            if p and p.lower() not in ["unknown", "#n/a", "nan", "none"]:
                vals.append(p)
    return pd.Series(vals, dtype="object")


def render_b5_disclosure(df: pd.DataFrame):
    story("B5.1 · Disclosure yield and post-claim signals", "Which conditions are missed at application, and how useful are medical exams?", "This page compares what customers disclose, what medical exams find, and what appears in early claims. The objective is to turn post-claim underwriting findings into better application questions, medical evidence triggers, and exclusion/revised-term rules.")
    me = df[df["medical_exam_flag"].eq(1)]
    no = df[df["medical_exam_flag"].eq(0)]
    leakage = df[df["early_claim_flag"].eq(1) & df["disclosure_flag"].eq(0) & df["medical_exam_flag"].eq(0)]
    decision_summary(
        "Where is disclosure or evidence capture weakest?",
        f"There are {fmt_n(len(leakage))} early-claim cases with no disclosure flag and no medical exam in the current view. ME material-finding yield is {fmt_pct(me['material_finding_flag'].mean(), 1)}.",
        "Use the ICD and disclosure charts to refine application questions, targeted evidence triggers, and post-claim review priorities.",
    )
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("Medical-exam cases", fmt_n(len(me)), f"Share of filtered policies: {fmt_pct(len(me)/len(df) if len(df) else np.nan, 1)}")
    with c2:
        kpi_card("ME material-finding yield", fmt_pct(me["material_finding_flag"].mean(), 1), "Material findings among ME cases")
    with c3:
        kpi_card("No-ME material-finding yield", fmt_pct(no["material_finding_flag"].mean(), 1), "Material findings among non-ME cases")
    with c4:
        kpi_card("Disclosure flag present", fmt_pct(df["disclosure_flag"].mean(), 1), "Policies with disclosure signal")

    section(1, "Medical-exam signal")
    st.caption("Medical-exam and no-medical-exam groups are selected by underwriting rules, so this comparison is observational and should not be read as a randomized effectiveness estimate.")
    c1, c2 = st.columns(2)
    with c1:
        card("With vs without medical exam", "Yield means a case that resulted in revised/substandard, postponed, or declined outcome. This is selection-biased because exams are triggered for higher-concern cases.")
        y = pd.DataFrame({
            "Evidence group": ["Medical exam", "No medical exam"],
            "Material-finding yield": [me["material_finding_flag"].mean(), no["material_finding_flag"].mean()],
            "Early-claim rate": [me["early_claim_flag"].mean(), no["early_claim_flag"].mean()],
        })
        long = y.melt(id_vars="Evidence group", var_name="Metric", value_name="Rate")
        bar_chart(long, "Evidence group", "Rate", "Metric", y_is_pct=True, height=340)
        end_card()
    with c2:
        card("UW outcomes by medical exam", "If exams are working, the exam group should have a materially higher non-standard rate.")
        g = group_rate(df, ["medical_exam_flag", "uw_outcome_clean"])
        g["medical_exam_flag"] = g["medical_exam_flag"].map({1: "Medical exam", 0: "No medical exam"})
        bar_chart(g, "uw_outcome_clean", "policies", "medical_exam_flag", text="policies", height=340)
        end_card()

    section(2, "Claimed vs disclosed conditions")
    early = df[df["early_claim_flag"].eq(1)]
    icds = split_icd_values(early["icd_text"]).value_counts().head(15).reset_index()
    icds.columns = ["ICD / condition", "Count"]
    disc = df.loc[df["disclosure_flag"].eq(1), "disclosure_rule_clean"].replace("Unknown", np.nan).dropna().value_counts().head(15).reset_index()
    disc.columns = ["Disclosure rule", "Count"]
    c1, c2 = st.columns(2)
    with c1:
        card("Top ICD / condition text in early claims", "These are conditions that should be reviewed for application question and evidence-rule relevance.")
        bar_chart(icds, "ICD / condition", "Count", text="Count", height=380)
        end_card()
    with c2:
        card("Top disclosed rules / aliases at application", "Compare this with early-claim ICDs to spot disclosure gaps.")
        bar_chart(disc, "Disclosure rule", "Count", text="Count", height=380)
        end_card()

    card("Disclosure gap summary", "This is the screenshot-style claimed-vs-told view, recalculated from the uploaded file.")
    early_items = "".join(
        f"<div class='gap-item'>{idx + 1}. {escape(str(row['ICD / condition']))} ({fmt_n(row['Count'])})</div>"
        for idx, row in icds.head(5).iterrows()
    ) or "<div class='gap-item'>No early-claim condition text available.</div>"
    disclosed_items = "".join(
        f"<div class='gap-item'>{idx + 1}. {escape(str(row['Disclosure rule']))} ({fmt_n(row['Count'])})</div>"
        for idx, row in disc.head(5).iterrows()
    ) or "<div class='gap-item'>No disclosed-rule text available.</div>"
    top_claim = str(icds["ICD / condition"].iloc[0]) if not icds.empty else "No condition"
    top_disclosure = str(disc["Disclosure rule"].iloc[0]) if not disc.empty else "No disclosure rule"
    st.markdown(
        f"""
        <div class='gap-grid'>
            <div class='gap-box'>
                <div class='gap-title'>Top in early claims</div>
                {early_items}
            </div>
            <div class='gap-box'>
                <div class='gap-title'>Top disclosed at application</div>
                {disclosed_items}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    insight(f"<strong>How to read this:</strong> compare the most common early-claim condition ({escape(top_claim)}) with the most common disclosed condition/rule ({escape(top_disclosure)}). Large mismatches are where disclosure questions or evidence triggers may need refinement.")
    end_card()

    c1, c2 = st.columns(2)
    with c1:
        card("DVM / HD disclosure flags", "Checks whether digital verification or health disclosure flags are associated with lower leakage or higher material findings.")
        st.caption("This is an audit table. Read it as directional evidence only, not proof that the flag caused the outcome.")
        tmp = df.copy()
        tmp["DVM flag"] = tmp["dvm_flag"].map({1: "DVM present", 0: "No DVM"})
        tmp["HD disclosure"] = tmp["hd_disclosure_flag"].map({1: "HD disclosure", 0: "No HD disclosure"})
        g1 = group_rate(tmp, "DVM flag")[["DVM flag", "early_claim_rate", "material_finding_rate", "policies"]]
        g2 = group_rate(tmp, "HD disclosure")[["HD disclosure", "early_claim_rate", "material_finding_rate", "policies"]]
        signal_display = pd.concat([g1.rename(columns={"DVM flag": "Signal"}), g2.rename(columns={"HD disclosure": "Signal"})], ignore_index=True).assign(
            early_claim_rate=lambda x: x["early_claim_rate"].map(lambda v: fmt_pct(v, 3)),
            material_finding_rate=lambda x: x["material_finding_rate"].map(lambda v: fmt_pct(v, 1)),
        ).rename(columns={
            "early_claim_rate": "Early-claim rate",
            "material_finding_rate": "Material-finding rate",
            "policies": "Policies",
        })
        wrapped_table(
            signal_display,
            strong_cols={"Signal"},
            widths=["minmax(170px, 1fr)", "minmax(140px, .8fr)", "minmax(160px, .9fr)", "minmax(110px, .7fr)"],
        )
        end_card()
    with c2:
        card("Early-claim cases requiring post-claim review", "Prioritise cases with early claim + no disclosure + no medical exam.")
        leak = df[df["early_claim_flag"].eq(1) & df["disclosure_flag"].eq(0) & df["medical_exam_flag"].eq(0)]
        kpi_card("Potential non-disclosure leakage cases", fmt_n(len(leak)), f"Share of early claims: {fmt_pct(len(leak) / max(1, int(df['early_claim_flag'].sum())), 1)}")
        with st.expander("Show leakage-case segment table"):
            display_rate_table(leak, ["application_channel", "product_clean", "age_band_clean", "sa_band_clean"], min_count=1, top_n=15)
        end_card()
    if not icds.empty:
        section(3, "Translate disclosure gaps into action")
        action_df = icds.head(8).copy()
        action_df["Underwriting action"] = "Review application question, evidence trigger, and exclusion/revised-term rule for this condition."
        action_df["Priority"] = np.where(action_df["Count"].rank(method="first", ascending=False) <= 3, "High", "Monitor")
        wrapped_table(
            action_df[["Priority", "ICD / condition", "Count", "Underwriting action"]],
            strong_cols={"Priority", "ICD / condition"},
            widths=["110px", "minmax(220px, .85fr)", "90px", "minmax(360px, 1.4fr)"],
        )


def credibility_adjusted_rate(early_claims: int, policies: int, baseline: float, prior_weight: int = 750) -> float:
    """Shrink small comparison groups toward the portfolio baseline to avoid overreacting to tiny samples."""
    if policies <= 0 or pd.isna(baseline):
        return baseline
    return float((early_claims + baseline * prior_weight) / (policies + prior_weight))


def historical_rate_for_scenario(df: pd.DataFrame, filters: dict) -> tuple[float, int, str]:
    summary = scenario_match_summary(df, filters)
    used = summary[summary["Used for score"]].head(1)
    if used.empty:
        return float(df["early_claim_flag"].mean()), len(df), "portfolio baseline"
    row = used.iloc[0]
    return float(row["Credibility-adjusted rate"]), int(row["Policies"]), str(row["Comparison level"])


def scenario_match_summary(df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    levels = [
        ("Detailed underwriting match", ["application_channel", "age_band_clean", "sa_band_clean", "bmi_band_clean", "product_clean", "smoking_clean"]),
        ("Core underwriting match", ["application_channel", "age_band_clean", "sa_band_clean", "bmi_band_clean", "product_clean"]),
        ("Channel + age + SA + BMI", ["application_channel", "age_band_clean", "sa_band_clean", "bmi_band_clean"]),
        ("Channel + age + SA", ["application_channel", "age_band_clean", "sa_band_clean"]),
        ("Channel + age", ["application_channel", "age_band_clean"]),
        ("Channel + SA", ["application_channel", "sa_band_clean"]),
        ("Channel only", ["application_channel"]),
        ("Portfolio", []),
    ]
    rows = []
    baseline = float(df["early_claim_flag"].mean()) if len(df) else np.nan
    for label, cols in levels:
        mask = pd.Series(True, index=df.index)
        matched_labels = []
        for col in cols:
            if col in filters and str(filters[col]) != "Unknown":
                mask &= df[col].astype(str).eq(str(filters[col]))
                matched_labels.append(friendly_label(col))
        sample = df.loc[mask]
        early_claims = int(sample["early_claim_flag"].sum()) if len(sample) else 0
        adjusted = credibility_adjusted_rate(early_claims, len(sample), baseline)
        rows.append(
            {
                "Comparison level": label,
                "Matched fields": " + ".join(matched_labels) if matched_labels else "None",
                "Policies": len(sample),
                "Early claims": early_claims,
                "Observed early-claim rate": sample["early_claim_flag"].mean() if len(sample) else np.nan,
                "Credibility-adjusted rate": adjusted,
                "Used for score": False,
            }
        )
    for idx, row in enumerate(rows):
        if row["Policies"] >= 30:
            rows[idx]["Used for score"] = True
            break
    if rows and not any(row["Used for score"] for row in rows):
        rows[-1]["Used for score"] = True
    return pd.DataFrame(rows)


def scenario_confidence(n: int) -> tuple[str, str]:
    if n >= 1000:
        return "High confidence", COLORS["green"]
    if n >= 200:
        return "Usable confidence", COLORS["accent"]
    if n >= 50:
        return "Limited confidence", COLORS["amber"]
    return "Low confidence", COLORS["red"]


def scenario_risk_tier(score: float, baseline: float) -> tuple[str, str]:
    if pd.isna(score):
        return "Not enough data", COLORS["muted"]
    if score >= max(0.01, baseline * 4):
        return "High", COLORS["red"]
    if score >= max(0.005, baseline * 2):
        return "Moderate", COLORS["amber"]
    return "Normal", COLORS["green"]


def scenario_action(tier: str) -> str:
    if tier == "High":
        return "Escalate for targeted evidence before standard acceptance."
    if tier == "Moderate":
        return "Review disclosure quality and consider additional evidence."
    if tier == "Not enough data":
        return "Use manual underwriting judgement because the comparable population is too thin."
    return "Proceed with normal underwriting checks and monitor if profile changes."


def render_score_gauge(score: float, baseline: float):
    max_axis = max(0.02, score * 1.35, baseline * 5)
    fig = go.Figure(
        go.Indicator(
            mode="gauge+number+delta",
            value=score * 100,
            number={"suffix": "%", "valueformat": ".3f"},
            delta={"reference": baseline * 100, "suffix": "% vs portfolio"},
            gauge={
                "axis": {"range": [0, max_axis * 100], "tickformat": ".2f"},
                "bar": {"color": COLORS["accent"]},
                "steps": [
                    {"range": [0, max(0.005, baseline * 2) * 100], "color": "#DCFCE7"},
                    {"range": [max(0.005, baseline * 2) * 100, max(0.01, baseline * 4) * 100], "color": "#FEF3C7"},
                    {"range": [max(0.01, baseline * 4) * 100, max_axis * 100], "color": "#FEE2E2"},
                ],
            },
        )
    )
    fig.update_layout(height=280, margin=dict(l=20, r=20, t=20, b=20), paper_bgcolor="white")
    st.plotly_chart(fig, use_container_width=True)


def option_values(df: pd.DataFrame, col: str, fallback: list[str] | None = None) -> list[str]:
    values = sorted([x for x in df[col].dropna().astype(str).unique() if x and x != "Unknown"]) if col in df else []
    return values or (fallback or ["Unknown"])


def options_with_unknown(df: pd.DataFrame, col: str, fallback: list[str] | None = None) -> list[str]:
    values = option_values(df, col, fallback)
    return ["Unknown"] + [value for value in values if value != "Unknown"]


def simulator_output_card(title: str, value: str, subtitle: str = "", tone: str = ""):
    tone_class = {"good": "good", "warn": "warn", "danger": "danger"}.get(tone, "")
    st.markdown(
        f"""
        <div class='card {tone_class}' style='min-height:150px;'>
            <div class='score-label'>{title}</div>
            <div class='card-title' style='font-size:1.35rem; margin-top:8px;'>{value}</div>
            <div class='card-subtitle' style='font-size:.88rem;'>{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def run_application_scenario_score(df: pd.DataFrame, age: int, bmi: float, sa: float, channel: str, product: str, smoker: str, disclosure: str, medical_exam: str, loading: str, exclusion: str, myinfo: str, efna: str, mpci: str, gio: str, reopen: str) -> dict:
    age_band = make_age_band(pd.Series([age])).iloc[0]
    sa_band = make_sa_band(pd.Series([sa])).iloc[0]
    bmi_band = make_bmi_band(pd.Series([bmi])).iloc[0]
    filters = {
        "application_channel": channel,
        "age_band_clean": age_band,
        "sa_band_clean": sa_band,
        "bmi_band_clean": bmi_band,
        "product_clean": product,
        "smoking_clean": smoker,
    }
    match_summary = scenario_match_summary(df, filters)
    used_match = match_summary[match_summary["Used for score"]].head(1)
    baseline = float(df["early_claim_flag"].mean()) if len(df) else np.nan
    used_level = used_match["Comparison level"].iloc[0] if not used_match.empty else "Portfolio"
    matched_cases = int(used_match["Policies"].iloc[0]) if not used_match.empty else len(df)
    observed_rate = float(used_match["Observed early-claim rate"].iloc[0]) if not used_match.empty else baseline
    score = float(used_match["Credibility-adjusted rate"].iloc[0]) if not used_match.empty else baseline
    adjustments = []
    if bmi >= 35:
        score *= 1.55
        adjustments.append("BMI >= 35")
    elif bmi >= 30:
        score *= 1.35
        adjustments.append("BMI 30-35")
    elif bmi >= 27.5:
        score *= 1.15
        adjustments.append("BMI 27.5-30")
    if channel == "NF2F" and disclosure == "No" and medical_exam == "No":
        score *= 1.25
        adjustments.append("NF2F + no disclosure + no exam")
    if sa >= 250000:
        score *= 1.20
        adjustments.append("High sum assured")
    if "smok" in str(smoker).lower() or smoker == "Yes":
        score *= 1.12
        adjustments.append("Smoker")
    if loading == "Yes":
        score *= 1.20
        adjustments.append("Loading applied")
    if exclusion == "Yes":
        score *= 1.10
        adjustments.append("Exclusion applied")
    if gio == "Yes":
        score *= 1.10
        adjustments.append("Guaranteed-issue case")
    if reopen == "Yes":
        score *= 1.10
        adjustments.append("Reopen case")
    if myinfo == "Yes":
        score *= 0.95
        adjustments.append("MyInfo used")
    if efna == "Yes":
        score *= 0.97
        adjustments.append("eFNA used")
    if mpci == "Yes":
        score *= 1.05
        adjustments.append("MPCI rider")
    score = min(score, 0.50)
    tier, color = scenario_risk_tier(score, baseline)
    high_threshold = max(0.01, baseline * 4) if baseline and not pd.isna(baseline) else 0.01
    risk_index = int(min(100, max(1, round(score / high_threshold * 100))))
    confidence, confidence_color = scenario_confidence(matched_cases)
    if tier == "High":
        outlook = "Likely early-claim review case"
        tone = "danger"
    elif tier == "Moderate":
        outlook = "Possible early-claim review case"
        tone = "warn"
    else:
        outlook = "Lower early-claim outlook"
        tone = "good"
    review_action = scenario_action(tier)
    review_reason = "Based on the final score tier, confidence label, and selected underwriting flags."
    return {
        "age_band": age_band,
        "sa_band": sa_band,
        "bmi_band": bmi_band,
        "baseline": baseline,
        "score": score,
        "risk_index": risk_index,
        "tier": tier,
        "tone": tone,
        "outlook": outlook,
        "review_action": review_action,
        "review_reason": review_reason,
        "matched_cases": matched_cases,
        "observed_rate": observed_rate,
        "used_level": used_level,
        "confidence": confidence,
        "confidence_color": confidence_color,
        "adjustments": adjustments,
        "match_summary": match_summary,
    }


def render_application_scenario_simulator(df: pd.DataFrame):
    story(
        "New Application Triage",
        "If a new applicant profile is entered, what is the early-claim outlook?",
        "This page mirrors an application-time triage tool. It captures practical fields known at submission, including BMI, then maps the raw inputs back to the same historical bands used by the dashboard.",
    )
    section(1, "Application input and triage output")
    left, right = st.columns([1.15, .85])
    with left:
        with st.container(border=True):
            st.markdown("#### Application-time inputs")
            st.caption("Age, BMI, and sum assured are entered as raw values. The dashboard maps them into age, BMI, and sum assured bands before scoring. Gender is intentionally excluded from scoring.")
            core_tab, evidence_tab, context_tab = st.tabs(["Core profile", "Evidence", "Other flags"])
            with core_tab:
                r1c1, r1c2 = st.columns(2)
                age = r1c1.number_input("Age next birthday", 0, 90, 35, key="appsim_age")
                bmi = r1c2.number_input("BMI", 10.0, 60.0, 27.5, 0.1, key="appsim_bmi")
                r2c1, r2c2 = st.columns(2)
                sa = r2c1.number_input("Desired sum assured (SGD)", 0.0, 10_000_000.0, 100_000.0, 10_000.0, key="appsim_sa")
                smoker = r2c2.selectbox("Smoker", options_with_unknown(df, "smoking_clean", ["No", "Yes"]), key="appsim_smoker")
                r3c1, r3c2 = st.columns(2)
                product = r3c1.selectbox("Type of insurance / product group", option_values(df, "product_clean"), key="appsim_product")
                channel = r3c2.selectbox("Submission mode", ["F2F", "NF2F"], key="appsim_channel")
            with evidence_tab:
                r4c1, r4c2 = st.columns(2)
                disclosure = r4c1.selectbox("Health disclosure present", ["No", "Yes"], key="appsim_disclosure")
                medical_exam = r4c2.selectbox("Medical exam required", ["No", "Yes"], key="appsim_medical")
                r5c1, r5c2 = st.columns(2)
                loading = r5c1.selectbox("Loading applied", ["No", "Yes"], key="appsim_loading")
                exclusion = r5c2.selectbox("Exclusion applied", ["No", "Yes"], key="appsim_exclusion")
                r6c1, r6c2 = st.columns(2)
                myinfo = r6c1.selectbox("MyInfo used", ["No", "Yes"], key="appsim_myinfo")
                efna = r6c2.selectbox("eFNA used", ["No", "Yes"], key="appsim_efna")
            with context_tab:
                r7c1, r7c2 = st.columns(2)
                source = r7c1.selectbox("Source of business", options_with_unknown(df, "source_business_clean"), key="appsim_source")
                fund = r7c2.selectbox("Fund category", options_with_unknown(df, "fund_category_clean"), key="appsim_fund")
                r8c1, r8c2 = st.columns(2)
                mpci = r8c1.selectbox("MPCI rider included", ["No", "Yes"], key="appsim_mpci")
                gio = r8c2.selectbox("Guaranteed-issue / GIO case", ["No", "Yes"], key="appsim_gio")
                reopen = st.selectbox("Reopen case", ["No", "Yes"], key="appsim_reopen")
            st.caption("The score uses channel, age, BMI, sum assured, product, smoking, evidence, and adverse flags. Source of business and fund category are captured for context. Gender is excluded and the output does not auto-approve or auto-decline.")

    result = run_application_scenario_score(df, age, bmi, sa, channel, product, smoker, disclosure, medical_exam, loading, exclusion, myinfo, efna, mpci, gio, reopen)
    with right:
        with st.container(border=True):
            st.markdown("#### Scenario output")
            st.caption("The simulator returns a risk index, early-claim outlook, and underwriting review action. It is a triage aid, not an automatic decision.")
            o1, o2 = st.columns(2)
            with o1:
                simulator_output_card("Scenario risk index", f"{result['risk_index']} / 100", f"Estimated early-claim rate: {fmt_pct(result['score'], 3)}", result["tone"])
            with o2:
                simulator_output_card("Early-claim outlook", result["outlook"], f"Portfolio baseline: {fmt_pct(result['baseline'], 3)} · Tier: {result['tier']}", result["tone"])
            o3, o4 = st.columns(2)
            with o3:
                simulator_output_card("Review action", result["review_action"], result["review_reason"], result["tone"])
            with o4:
                simulator_output_card("Derived risk bands", f"{result['age_band']} · {result['sa_band']} · {result['bmi_band']}", "Raw inputs are mapped back to dashboard bands.")
            st.progress(result["risk_index"] / 100)
            st.markdown(
                f"""
                <div class='insight'>
                    <strong>Scenario interpretation:</strong> This case maps to <strong>{result['age_band']}</strong>,
                    <strong>{result['sa_band']}</strong>, and <strong>{result['bmi_band']}</strong>.
                    The simulator classifies it as <strong>{result['outlook'].lower()}</strong>.
                    Use this as a triage signal, not an automatic underwriting decision.
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.markdown(
                f"<span class='pill'>Matched cases: {fmt_n(result['matched_cases'])}</span>"
                f"<span class='pill'>Matched using: {result['used_level']}</span>"
                f"<span class='pill' style='border-color:{result['confidence_color']}; color:{result['confidence_color']}'>{result['confidence']}</span>"
                f"<span class='pill'>Observed comparable rate: {fmt_pct(result['observed_rate'], 3)}</span>",
                unsafe_allow_html=True,
            )

    section(2, "Why the simulator gave this result")
    c1, c2 = st.columns([.9, 1.1])
    with c1:
        card("Fields captured", "These are the practical inputs used in the scenario.")
        captured = pd.DataFrame(
            [
                ["Age band", result["age_band"]],
                ["BMI band", result["bmi_band"]],
                ["Sum assured band", result["sa_band"]],
                ["Product group", product],
                ["Submission mode", channel],
                ["Source of business", source],
                ["Fund category", fund],
                ["Smoker", smoker],
                ["Health disclosure present", disclosure],
                ["Medical exam required", medical_exam],
                ["Loading applied", loading],
                ["Exclusion applied", exclusion],
                ["MyInfo used", myinfo],
                ["eFNA used", efna],
                ["MPCI rider", mpci],
                ["GIO case", gio],
                ["Reopen case", reopen],
            ],
            columns=["Input field", "Selected value"],
        )
        wrapped_table(captured, strong_cols={"Input field"}, widths=["minmax(180px, .75fr)", "minmax(220px, 1fr)"])
        end_card()
    with c2:
        card("Score basis and adjustments", "The score starts from similar historical cases, then applies simple underwriting review adjustments.")
        adjustments = result["adjustments"] or ["No additional review adjustment applied"]
        st.markdown("**Adjustments applied:** " + ", ".join(adjustments))
        compare = pd.DataFrame(
            {
                "Metric": ["Portfolio baseline", "Observed comparable rate", "Final scenario score"],
                "Early-claim rate": [result["baseline"], result["observed_rate"], result["score"]],
            }
        )
        bar_chart(compare, "Metric", "Early-claim rate", y_is_pct=True, text="Early-claim rate", height=300)
        end_card()
    with st.expander("Show scenario matching fallback table"):
        match_display = result["match_summary"].copy()
        match_display["Observed early-claim rate"] = match_display["Observed early-claim rate"].map(lambda x: fmt_pct(x, 3))
        match_display["Credibility-adjusted rate"] = match_display["Credibility-adjusted rate"].map(lambda x: fmt_pct(x, 3))
        wrapped_table(
            match_display,
            strong_cols={"Comparison level", "Used for score"},
            widths=["minmax(150px, .9fr)", "minmax(230px, 1.4fr)", "90px", "110px", "minmax(140px, .8fr)", "minmax(160px, .9fr)", "110px"],
        )

    section(3, "How this triage is built")
    card(
        "Scenario scoring method",
        "The triage score is built from historical comparison groups and underwriting review adjustments, not from a machine-learning model.",
    )
    build_notes = [
        ("1. Capture inputs", "The form captures application-time fields such as age, BMI, sum assured, product, F2F/NF2F mode, smoking, disclosure, medical exam, and adverse flags. Gender is not used in scoring."),
        ("2. Convert to bands", "Raw age, BMI, and sum assured are mapped into the same cleaned bands used throughout the dashboard."),
        ("3. Find comparable cases", "The app searches historical policies with the closest matching channel, age band, sum assured band, BMI band, product, and smoking profile."),
        ("4. Use fallback matching", "If the closest group is too small, the matching table falls back to broader groups so the score is not driven by a tiny sample."),
        ("5. Credibility adjust", "Observed early-claim rates are blended toward the portfolio baseline so small segments do not overstate risk."),
        ("6. Apply review factors", "Simple underwriting adjustments are applied for BMI threshold zones, NF2F with no disclosure or exam, high sum assured, smoker status, loading, exclusions, GIO, reopen, MyInfo, eFNA, and MPCI."),
        ("7. Return triage output", "The final score is translated into a risk index, early-claim outlook, confidence label, and underwriting review action."),
    ]
    rows = "".join(
        f"<div class='method-row'><div class='method-step'>{escape(step)}</div><div class='method-detail'>{escape(detail)}</div></div>"
        for step, detail in build_notes
    )
    st.markdown(f"<div class='method-table'>{rows}</div>", unsafe_allow_html=True)
    st.info("This is a decision-support triage method. It should help underwriters prioritise review, not automatically approve, decline, load, or exclude an application.")
    end_card()


def main():
    inject_css()
    df, source_name, last_modified = load_data_ui()
    if df.empty:
        st.error("Loaded data is empty.")
        st.stop()
    filtered = apply_sidebar_filters(df)
    topbar(source_name, last_modified, filtered)
    if len(filtered) == 0:
        st.warning("No records left after filters. Clear the sidebar filters.")
        st.stop()
    tabs = st.tabs([
        "Overview",
        "BMI Rule",
        "F2F / NF2F",
        "UW Outcomes",
        "Disclosure & Exams",
        "New Application Triage",
    ])
    with tabs[0]:
        render_overview(filtered)
    with tabs[1]:
        render_b1_bmi(filtered)
    with tabs[2]:
        render_b2_f2f(filtered)
    with tabs[3]:
        render_b4_uw(filtered)
    with tabs[4]:
        render_b5_disclosure(filtered)
    with tabs[5]:
        render_application_scenario_simulator(filtered)
    st.markdown(f"<div class='footer'><div>UC2 Underwriting Dashboard · Internal decision support only</div><div>{fmt_n(len(filtered))} policies · generated {datetime.now():%Y-%m-%d %H:%M}</div></div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()

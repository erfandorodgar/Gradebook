
"""
Streamlit app: Grade Bot — Student-ID-only version
Reads a multi-sheet Excel workbook (upload or cloud link) and shows grades by Student ID.
- No teacher Access Code required.
- Optional per-student Secret/PIN supported if present in data.
- If Student IDs are unique across the workbook, we skip course selection and show summaries.
"""

from __future__ import annotations
import io
from typing import Dict, List, Optional, Tuple
import pandas as pd
import streamlit as st
import requests
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

st.set_page_config(page_title="NEC Grade Bot (ID only)", page_icon="✅", layout="centered")
st.title("NEC Grade Bot")
st.caption("Enter your Student ID to view your grades.")

# -----------------------------
# Utilities
# -----------------------------

DEFAULT_SYNONYMS: Dict[str, List[str]] = {
    "student id": ["student id", "id", "sid", "student_number", "student num", "studentno", "student#"],
    "first name": ["first name", "first", "given"],
    "last name": ["last name", "last", "family", "surname"],
    "course": ["course", "class", "section"],
    "term": ["term", "semester", "session"],
    "assessment": ["assessment", "assignment", "quiz", "test", "task", "name"],
    "score": ["score", "mark", "points", "grade"],
    "out of": ["out of", "max", "total", "points possible", "denominator"],
    "weight": ["weight", "weight %", "%", "percent", "percentage"],
    "secret": ["secret", "pin", "dob_last4"],
}

STANDARD_COLUMNS = [
    "student id",
    "first name",
    "last name",
    "course",
    "term",
    "assessment",
    "score",
    "out of",
    "weight",
    "secret",
]

@st.cache_data(show_spinner=False)
def read_workbook(file_bytes: bytes) -> Tuple[pd.DataFrame, List[str]]:
    """
    Read all sheets, normalize columns via synonyms and optional _aliases sheet, and stack into a long DataFrame.
    Returns (df, sheet_names_read)
    """
    xl = pd.ExcelFile(io.BytesIO(file_bytes))

    # Optional alias sheet: two columns Key, Value (case-insensitive)
    custom_map: Dict[str, str] = {}
    if any(name.strip().lower() == "_aliases" for name in xl.sheet_names):
        ali = pd.read_excel(io.BytesIO(file_bytes), sheet_name="_aliases")
        if {c.strip().lower() for c in ali.columns} >= {"key", "value"}:
            for _, row in ali.iterrows():
                k = str(row["key"]).strip().lower()
                v = str(row["value"]).strip().lower()
                custom_map[v] = k  # map sheet header v to canonical key k

    def canonical(col: str) -> Optional[str]:
        c = str(col).strip().lower()
        # apply custom aliasing first
        if c in custom_map:
            return custom_map[c]
        for standard, syns in DEFAULT_SYNONYMS.items():
            if c == standard or c in syns:
                return standard
        return None

    frames: List[pd.DataFrame] = []
    read_names: List[str] = []

    for name in xl.sheet_names:
        if name.strip().lower() in {"_aliases"}:
            continue
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=name)
        # Map columns
        colmap = {}
        for c in df.columns:
            can = canonical(c)
            if can:
                colmap[c] = can
        df = df.rename(columns=colmap)
        # Keep only recognized columns, add missing ones as NaN
        for col in STANDARD_COLUMNS:
            if col not in df.columns:
                df[col] = pd.NA
        # Drop rows without Student ID
        df = df.dropna(subset=["student id"]).copy()
        # Standardize types
        df["student id"] = df["student id"].astype(str).str.strip()
        for c in ["first name", "last name", "course", "term", "assessment", "secret"]:
            df[c] = df[c].astype(str).str.strip()
        # Numeric fields
        for c in ["score", "out of", "weight"]:
            df[c] = pd.to_numeric(df[c], errors="coerce")
        # Defaults
        df.loc[df["out of"].isna(), "out of"] = 100
        # Attach sheet info
        df["_sheet"] = name
        frames.append(df[STANDARD_COLUMNS + ["_sheet"]])
        read_names.append(name)

    if not frames:
        return pd.DataFrame(columns=STANDARD_COLUMNS + ["_sheet"]), []

    all_df = pd.concat(frames, ignore_index=True)
    return all_df, read_names


def compute_course_total(q: pd.DataFrame) -> float:
    """Compute one course overall percent for a subset q (single course)."""
    if q.empty:
        return float("nan")
    if q["weight"].notna().any():
        w = q["weight"].fillna(0)
        if w.sum() == 0:
            denom = q["out of"].fillna(100).sum()
            return (q["score"].fillna(0).sum() / denom) * 100 if denom else float("nan")
        contrib = (q["score"].fillna(0) / q["out of"].replace(0, pd.NA).fillna(100)) * w
        return (contrib.sum() / w.sum()) * 100
    denom = q["out of"].fillna(100).sum()
    return (q["score"].fillna(0).sum() / denom) * 100 if denom else float("nan")


def summarize_by_course(rows: pd.DataFrame) -> pd.DataFrame:
    """Return per-course summary with overall percent."""
    out = []
    for course, sub in rows.groupby(rows["course"].fillna("")).groups.items():
        part = rows.loc[sub]
        pct = compute_course_total(part)
        out.append({"Course": course or "(Unspecified)", "Overall %": pct, "Assessments": len(part)})
    if not out:
        return pd.DataFrame(columns=["Course", "Overall %", "Assessments"])
    df = pd.DataFrame(out).sort_values("Course").reset_index(drop=True)
    return df


def _coerce_download_url(u: str) -> str:
    try:
        pr = urlparse(u)
        q = parse_qs(pr.query)
        # SharePoint/OneDrive: add download=1 if missing
        if pr.netloc.endswith("sharepoint.com") or pr.netloc.endswith("1drv.ms"):
            if "download" not in q:
                q["download"] = ["1"]
        # Dropbox: dl=1 for direct bytes
        if pr.netloc.endswith("dropbox.com"):
            q["dl"] = ["1"]
        new_query = urlencode({k: v[0] for k, v in q.items()})
        return urlunparse((pr.scheme, pr.netloc, pr.path, pr.params, new_query, pr.fragment))
    except Exception:
        return u

# -----------------------------
# Sidebar: Teacher setup
# -----------------------------
st.sidebar.header("Teacher setup")
st.sidebar.write("Upload your Excel gradebook OR paste a cloud link. No access code needed in this version.")
workbook = st.sidebar.file_uploader("Upload .xlsx gradebook", type=["xlsx"])  # kept in session

cloud_url = st.sidebar.text_input("Or paste a cloud link to .xlsx (OneDrive/SharePoint/Dropbox/GDrive)")
fetch_btn = st.sidebar.button("Fetch from cloud link")

if "grade_df" not in st.session_state:
    st.session_state["grade_df"] = None

if workbook is not None:
    try:
        df_all, used_sheets = read_workbook(workbook.getvalue())
        st.sidebar.success(f"Loaded {len(df_all):,} rows from {len(used_sheets)} sheet(s).")
        if used_sheets:
            st.sidebar.caption("Sheets: " + ", ".join(used_sheets))
        st.session_state["grade_df"] = df_all
    except Exception as e:
        st.sidebar.error(f"Problem reading workbook: {e}")

if fetch_btn and cloud_url:
    try:
        direct = _coerce_download_url(cloud_url.strip())
        with st.sidebar.status("Fetching workbook...", expanded=False):
            r = requests.get(direct, timeout=30)
        if r.status_code != 200:
            st.sidebar.error(f"Could not download file (HTTP {r.status_code}). If this is SharePoint/OneDrive, set link permissions to 'Anyone with the link' and try again. Also try adding &download=1 to the link.")
        else:
            df_all, used_sheets = read_workbook(r.content)
            st.sidebar.success(f"Loaded {len(df_all):,} rows from {len(used_sheets)} sheet(s) via cloud link.")
            if used_sheets:
                st.sidebar.caption("Sheets: " + ", ".join(used_sheets))
            st.session_state["grade_df"] = df_all
    except Exception as e:
        st.sidebar.error(f"Problem fetching workbook: {e}")

# -----------------------------
# Student view
# -----------------------------
st.subheader("Student portal")

col1, col2 = st.columns([2,1])
with col1:
    sid = st.text_input("Student ID", help="Enter your Student ID exactly as provided by your instructor.")
with col2:
    secret = st.text_input("Optional PIN / Secret (if your teacher gave you one)", type="password")

show = st.button("Show my grades")

if show:
    df = st.session_state.get("grade_df")
    if df is None:
        st.error("The gradebook is not loaded yet. Ask your instructor to upload or connect the workbook.")
    else:
        sid_rows = df[df["student id"].astype(str).str.lower() == str(sid).strip().lower()]
        if secret:
            sid_rows = sid_rows[(sid_rows["secret"].astype(str).str.strip() == secret.strip()) | (sid_rows["secret"].isna())]

        if sid_rows.empty:
            st.warning("No records found. Check your Student ID and optional secret.")
        else:
            # Per-course summary
            summary = summarize_by_course(sid_rows)
            if not summary.empty:
                st.markdown("### Course summary")
                st.dataframe(summary, use_container_width=True)
            # Detailed assessments
            st.markdown("### Assessment details")
            details = sid_rows.copy()
            details["Percent"] = (details["score"].fillna(0) / details["out of"].replace(0, pd.NA).fillna(100)) * 100
            details = details[[
                "course", "term", "assessment", "score", "out of", "Percent", "weight", "_sheet"
            ]].rename(columns={
                "course": "Course",
                "term": "Term",
                "assessment": "Assessment",
                "score": "Score",
                "out of": "Out of",
                "weight": "Weight %",
                "_sheet": "Sheet"
            })
            st.dataframe(details.sort_values(["Course", "Assessment"]), use_container_width=True)
            st.caption("If any assessment is missing or looks off, please contact your instructor.")

st.divider()
st.caption("Made with ♥ for instructors and students. © 2025")

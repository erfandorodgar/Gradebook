
"""
Streamlit app: Grade Bot — secure self-serve grade lookup from a multi‑sheet Excel workbook
With direct cloud link support (SharePoint/OneDrive/Dropbox/Google Drive).
"""

from __future__ import annotations
import io
from typing import Dict, List, Optional, Tuple
import pandas as pd
import streamlit as st
import requests
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

st.set_page_config(page_title="NEC Grade Bot", page_icon="✅", layout="centered")
st.title("NEC Grade Bot")
st.caption("Private, self-serve grade lookup from your teacher's Excel workbook.")

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


def gradebook_summary(df: pd.DataFrame, sid: str, secret: Optional[str], course: Optional[str]):
    q = df[df["student id"].str.lower() == sid.lower()]
    if secret:
        q = q[(q["secret"].astype(str).str.strip() == str(secret).strip()) | (q["secret"].isna())]
    if course:
        q = q[q["course"].str.lower() == course.lower()]
    return q


def compute_totals(q: pd.DataFrame) -> Tuple[pd.DataFrame, float]:
    """Return (detail table, overall percent).
    If weights are present, use normalized weights. Otherwise compute by points (sum(score)/sum(out of)).
    """
    if q.empty:
        return q, float("nan")
    if q["weight"].notna().any():
        w = q["weight"].fillna(0)
        if w.sum() == 0:
            pct = (q["score"].fillna(0).sum() / q["out of"].fillna(100).sum()) * 100 if q["out of"].fillna(100).sum() else float("nan")
        else:
            contrib = (q["score"].fillna(0) / q["out of"].replace(0, pd.NA).fillna(100)) * w
            pct = contrib.sum() / w.sum() * 100
    else:
        denom = q["out of"].fillna(100).sum()
        pct = (q["score"].fillna(0).sum() / denom) * 100 if denom else float("nan")

    details = q.copy()
    details["Percent"] = (details["score"].fillna(0) / details["out of"].replace(0, pd.NA).fillna(100)) * 100
    details = details[[
        "assessment", "score", "out of", "Percent", "weight", "course", "term", "_sheet"
    ]].rename(columns={
        "assessment": "Assessment",
        "score": "Score",
        "out of": "Out of",
        "weight": "Weight %",
        "course": "Course",
        "term": "Term",
        "_sheet": "Sheet",
    })
    return details, float(pct)

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
st.sidebar.write("Upload your Excel gradebook or paste a cloud link, then set an access code. Share the URL and access code with your students.")
access_code = st.sidebar.text_input("Access code (you choose)", type="password")
workbook = st.sidebar.file_uploader("Upload .xlsx gradebook", type=["xlsx"])  # kept in session

cloud_url = st.sidebar.text_input("Or paste a cloud link to .xlsx (OneDrive/SharePoint/Dropbox/GDrive)")
fetch_btn = st.sidebar.button("Fetch from cloud link")

if workbook is not None:
    try:
        df_all, used_sheets = read_workbook(workbook.getvalue())
        st.sidebar.success(f"Loaded {len(df_all):,} rows from {len(used_sheets)} sheet(s).")
        if used_sheets:
            st.sidebar.caption("Sheets: " + ", ".join(used_sheets))
        st.session_state["grade_df"] = df_all
        st.session_state["access_code_set"] = access_code.strip() if access_code else None
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
            st.session_state["access_code_set"] = access_code.strip() if access_code else None
    except Exception as e:
        st.sidebar.error(f"Problem fetching workbook: {e}")

# -----------------------------
# Tabs: Student view | Teacher tips
# -----------------------------
student_tab, teacher_tab = st.tabs(["Student", "Teacher tips"])

with student_tab:
    st.subheader("Student portal")
    st.write("Enter the access code your teacher gave you, then your Student ID.")

    user_code = st.text_input("Access code", type="password")
    col1, col2 = st.columns(2)
    with col1:
        sid = st.text_input("Student ID")
    with col2:
        secret = st.text_input("Optional PIN / Secret (if your teacher gave you one)")

    proceed = st.button("Show my grades")

    if proceed:
        df = st.session_state.get("grade_df")
        code_set = st.session_state.get("access_code_set")
        if df is None:
            st.error("The teacher hasn't uploaded a gradebook yet. Please try again later.")
        elif not code_set:
            st.error("Access code is not configured by the teacher.")
        elif user_code.strip() != code_set:
            st.error("Access code is incorrect.")
        else:
            # Determine course options for this SID
            sid_rows = df[df["student id"].astype(str).str.lower() == sid.strip().lower()]
            if secret:
                sid_rows = sid_rows[(sid_rows["secret"].astype(str).str.strip() == secret.strip()) | (sid_rows["secret"].isna())]
            if sid_rows.empty:
                st.warning("No records found. Check your Student ID and optional secret.")
            else:
                courses = sorted([c for c in sid_rows["course"].dropna().unique() if str(c).strip()])
                course_choice = st.selectbox("Pick your course", options=courses)
                q = gradebook_summary(df, sid, secret if secret else None, course_choice)
                details, overall = compute_totals(q)
                st.markdown(f"### Overall grade: **{overall:.2f}%**" if pd.notna(overall) else "### Overall grade: N/A")
                st.dataframe(details, use_container_width=True)
                st.caption("If any assessment is missing or looks off, please contact your instructor.")

with teacher_tab:
    st.subheader("Workbook format and tips")
    st.markdown(
        """
        **Minimum columns per sheet**: Student ID, Course, Assessment, Score. Optional: Out Of (default 100), Weight %, Term, First Name, Last Name, Secret.

        **Multiple sheets**: Put different assessment groups or courses on separate sheets. The app stacks them automatically.

        **Weights vs points**:
        - If Weight % is provided on any row, the app uses a normalized weighted model for that course.
        - Otherwise it uses points (sum of scores divided by sum of Out Of).

        **Security**:
        - Use an Access code so only your class can see the data.
        - You may add a per-student **Secret/PIN** column for two-factor lookup.
        - Never include sensitive data like full DOBs. Use last 4 only if you must.
        - Using a cloud link? Ensure the sharing setting is **Anyone with the link can view**. For SharePoint/OneDrive, append `download=1` to force a direct file download.

        **Aliases sheet**: You can include a sheet named `_aliases` with two columns **Key** and **Value** to map unusual headers to canonical ones.

        **Common gotchas**:
        - Ensure Student IDs match exactly across sheets.
        - Avoid merged cells and multi-row headers.
        - Replace blanks with real zeros where intended.
        - When using weights, ensure they sum to something meaningful across a course.
        """
    )

st.divider()
st.caption("Made with ♥ for instructors and students. © 2025")

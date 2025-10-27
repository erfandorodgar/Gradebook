
"""
Streamlit app: Grade Bot — Login via Credentials Sheet (Student ID + Access Code)
- Expects a credentials sheet with columns: Student ID, Access Code (synonyms supported).
- Credentials sheet can be named 'credentials' or 'login', OR the first sheet containing both fields will be used.
- All *other* sheets are treated as grade sheets and stacked.
- Supports upload OR cloud link (SharePoint/OneDrive/Dropbox/GDrive).
"""
from __future__ import annotations
import io
from typing import Dict, List, Optional, Tuple
import pandas as pd
import streamlit as st
import requests
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

st.set_page_config(page_title="NEC Grade Bot (Login via Access Code)", page_icon="✅", layout="centered")
st.title("NEC Grade Bot")
st.caption("Log in with your Student ID and Access Code to view your grades.")

DEFAULT_SYNONYMS: Dict[str, List[str]] = {
    "student id": ["student id", "id", "sid", "student_number", "student num", "studentno", "student#"],
    "access code": ["access code", "code", "login code", "passcode", "access_code"],
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

GRADE_STANDARD_COLUMNS = [
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

CRED_REQUIRED = ["student id", "access code"]

def canonical_name(col: str, custom_map: Dict[str, str]) -> Optional[str]:
    c = str(col).strip().lower()
    if c in custom_map:
        return custom_map[c]
    for standard, syns in DEFAULT_SYNONYMS.items():
        if c == standard or c in syns:
            return standard
    return None

def looks_like_credentials(df: pd.DataFrame, custom_map: Dict[str, str]) -> bool:
    mapped = set()
    for c in df.columns:
        cn = canonical_name(c, custom_map)
        if cn:
            mapped.add(cn)
    return all(req in mapped for req in CRED_REQUIRED)

@st.cache_data(show_spinner=False)
def parse_workbook(file_bytes: bytes) -> Tuple[pd.DataFrame, pd.DataFrame, List[str], str]:
    xl = pd.ExcelFile(io.BytesIO(file_bytes))

    custom_map: Dict[str, str] = {}
    if any(name.strip().lower() == "_aliases" for name in xl.sheet_names):
        ali = pd.read_excel(io.BytesIO(file_bytes), sheet_name="_aliases")
        if {c.strip().lower() for c in ali.columns} >= {"key", "value"}:
            for _, row in ali.iterrows():
                k = str(row["key"]).strip().lower()
                v = str(row["value"]).strip().lower()
                custom_map[v] = k

    creds_name = None
    for name in xl.sheet_names:
        low = name.strip().lower()
        if low in {"credentials", "login"}:
            creds_name = name
            break
    if creds_name is None:
        for name in xl.sheet_names:
            if name.strip().lower() == "_aliases":
                continue
            df_try = pd.read_excel(io.BytesIO(file_bytes), sheet_name=name)
            if looks_like_credentials(df_try, custom_map):
                creds_name = name
                break

    creds_df = pd.DataFrame(columns=["student id", "access code"])
    grade_frames: List[pd.DataFrame] = []
    used_grade_sheets: List[str] = []

    for name in xl.sheet_names:
        low = name.strip().lower()
        if low == "_aliases":
            continue
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=name)

        colmap = {}
        for c in df.columns:
            can = canonical_name(c, custom_map)
            if can:
                colmap[c] = can
        df = df.rename(columns=colmap)

        if creds_name and name == creds_name:
            tmp = df.copy()
            for req in CRED_REQUIRED:
                if req not in tmp.columns:
                    tmp[req] = pd.NA
            tmp = tmp[CRED_REQUIRED].dropna(subset=["student id"]).copy()
            tmp["student id"] = tmp["student id"].astype(str).str.strip()
            tmp["access code"] = tmp["access code"].astype(str).str.strip()
            creds_df = tmp
        else:
            for col in GRADE_STANDARD_COLUMNS:
                if col not in df.columns:
                    df[col] = pd.NA
            df = df.dropna(subset=["student id"]).copy()
            df["student id"] = df["student id"].astype(str).str.strip()
            for c in ["first name", "last name", "course", "term", "assessment", "secret"]:
                df[c] = df[c].astype(str).str.strip()
            for c in ["score", "out of", "weight"]:
                df[c] = pd.to_numeric(df[c], errors="coerce")
            df.loc[df["out of"].isna(), "out of"] = 100
            df["_sheet"] = name
            used_grade_sheets.append(name)
            grade_frames.append(df[GRADE_STANDARD_COLUMNS + ["_sheet"]])

    grades_df = pd.concat(grade_frames, ignore_index=True) if grade_frames else pd.DataFrame(columns=GRADE_STANDARD_COLUMNS + ["_sheet"])
    return grades_df, creds_df, used_grade_sheets, creds_name or "(auto-detected/none)"

def _coerce_download_url(u: str) -> str:
    try:
        pr = urlparse(u)
        q = parse_qs(pr.query)
        if pr.netloc.endswith("sharepoint.com") or pr.netloc.endswith("1drv.ms"):
            if "download" not in q:
                q["download"] = ["1"]
        if pr.netloc.endswith("dropbox.com"):
            q["dl"] = ["1"]
        new_query = urlencode({k: v[0] for k, v in q.items()})
        return urlunparse((pr.scheme, pr.netloc, pr.path, pr.params, new_query, pr.fragment))
    except Exception:
        return u

def compute_course_total(q: pd.DataFrame) -> float:
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
    out = []
    for course, sub in rows.groupby(rows["course"].fillna("")).groups.items():
        part = rows.loc[sub]
        pct = compute_course_total(part)
        out.append({"Course": course or "(Unspecified)", "Overall %": pct, "Assessments": len(part)})
    if not out:
        return pd.DataFrame(columns=["Course", "Overall %", "Assessments"])
    df = pd.DataFrame(out).sort_values("Course").reset_index(drop=True)
    return df

st.sidebar.header("Teacher setup")
st.sidebar.write("Upload your Excel OR paste a cloud link. The app finds a credentials sheet (Student ID + Access Code) and uses other sheets for grades.")
workbook = st.sidebar.file_uploader("Upload .xlsx gradebook", type=["xlsx"])
cloud_url = st.sidebar.text_input("Or paste a cloud link to .xlsx (OneDrive/SharePoint/Dropbox/GDrive)")
fetch_btn = st.sidebar.button("Fetch from cloud link")

if "grade_df" not in st.session_state:
    st.session_state["grade_df"] = None
    st.session_state["creds_df"] = None
    st.session_state["creds_sheet_name"] = None

def _load_bytes(file_bytes: bytes):
    grades, creds, grade_sheets, creds_name = parse_workbook(file_bytes)
    st.session_state["grade_df"] = grades
    st.session_state["creds_df"] = creds
    st.session_state["creds_sheet_name"] = creds_name
    st.sidebar.success(f"Loaded {len(grades):,} grade rows from {len(grade_sheets)} sheet(s). Credentials sheet: {creds_name}")
    if grade_sheets:
        st.sidebar.caption("Grade sheets: " + ", ".join(grade_sheets))

if workbook is not None:
    try:
        _load_bytes(workbook.getvalue())
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
            _load_bytes(r.content)
    except Exception as e:
        st.sidebar.error(f"Problem fetching workbook: {e}")

st.subheader("Student login")
col1, col2 = st.columns([1,1])
with col1:
    sid = st.text_input("Student ID", help="Enter your Student ID exactly as provided by your instructor.")
with col2:
    code = st.text_input("Access Code", type="password", help="Use the Access Code from your teacher.")

show = st.button("Log in and show my grades")

if show:
    grades = st.session_state.get("grade_df")
    creds = st.session_state.get("creds_df")

    if grades is None or creds is None or creds.empty:
        st.error("Workbook not loaded or credentials sheet missing. Ensure a sheet named 'credentials'/'login' exists, or the first sheet contains 'Student ID' and 'Access Code'.")
    else:
        sid_norm = str(sid).strip().lower()
        code_norm = str(code).strip()
        creds_cmp = creds.copy()
        creds_cmp["student id"] = creds_cmp["student id"].astype(str).str.strip().str.lower()
        creds_cmp["access code"] = creds_cmp["access code"].astype(str).str.strip()

        match = creds_cmp[(creds_cmp["student id"] == sid_norm) & (creds_cmp["access code"] == code_norm)]
        if match.empty:
            st.warning("Invalid Student ID or Access Code.")
        else:
            rows = grades[grades["student id"].astype(str).str.strip().str.lower() == sid_norm]
            if rows.empty:
                st.info("Login OK, but no grade rows were found for this Student ID.")
            else:
                summary = summarize_by_course(rows)
                if not summary.empty:
                    st.markdown("### Course summary")
                    st.dataframe(summary, use_container_width=True)

                st.markdown("### Assessment details")
                details = rows.copy()
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

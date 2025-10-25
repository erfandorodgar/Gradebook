
# Grade Bot (Streamlit)

A lightweight, privacy-conscious Streamlit app that lets students view their grades from your multi‑sheet Excel workbook.

## Quick Deploy (Streamlit Cloud)

1. Create a new **public** GitHub repo and upload these files:
   - `app.py`
   - `requirements.txt`
2. Go to Streamlit Community Cloud and click **Deploy app**, selecting your repo.
3. In the app sidebar:
   - Set an **Access code**.
   - Either **upload** your `.xlsx` or **paste a cloud link** (SharePoint/OneDrive/Dropbox/GDrive), then click **Fetch from cloud link**.

> For SharePoint/OneDrive, ensure the file is shared as **Anyone with the link can view**, and add `&download=1` to the URL so the app can download the bytes.

## Excel Format

Minimum columns per sheet:
- `Student ID`, `Course`, `Assessment`, `Score`

Optional columns:
- `Out Of` (default 100), `Weight %`, `Term`, `First Name`, `Last Name`, `Secret`

You can add a sheet `_aliases` with two columns `Key`, `Value` to map your custom header names to the canonical names above.

## Student Privacy
- Protect the app with an **Access code**.
- Optionally require a per-student `Secret`/PIN column before showing grades.

## Notes
- The app supports multiple sheets. It stacks them and keeps a `Sheet` column for traceability.
- If `Weight %` exists for any rows in a course, the app computes a weighted average; otherwise it uses points.

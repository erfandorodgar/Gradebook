[README.md](https://github.com/user-attachments/files/23144799/README.md)

# Grade Bot (Student-ID-only)

A streamlined Streamlit app that lets students view grades by **Student ID only** (no access code). Supports multi-sheet Excel workbooks and direct SharePoint/OneDrive links.

## Quick Deploy

1. Create a new **public** GitHub repo and upload:
   - `app.py`
   - `requirements.txt`
2. Deploy on Streamlit Community Cloud.
3. In the app sidebar:
   - Upload `.xlsx` or paste your SharePoint/OneDrive link (set to **Anyone with the link can view** and add `&download=1`).

## Excel Format

Minimum columns per sheet:
- `Student ID`, `Course`, `Assessment`, `Score`

Optional columns:
- `Out Of` (default 100), `Weight %`, `Term`, `First Name`, `Last Name`, `Secret` (for optional PIN)

The app aggregates across sheets. If an `_aliases` sheet exists with `Key`, `Value`, it maps custom headers to the canonical names.

## Behavior
- Students enter **Student ID** (and optional **Secret/PIN** if present).
- App shows **per-course summary** and a **detailed assessment table**.
- If Student IDs are unique across the workbook, no course selection is needed.

## Privacy
- No access code in this variant. If you want an extra layer, add a per-student `Secret`/PIN column to the workbook.

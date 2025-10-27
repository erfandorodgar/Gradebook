[README.md](https://github.com/user-attachments/files/23173927/README.md)

# Grade Bot — Login via Credentials Sheet

This Streamlit app authenticates students with **Student ID + Access Code** from a credentials sheet,
then displays grades from other sheets.

## Workbook Structure
- A sheet named **`credentials`** or **`login`**, OR the **first** sheet that contains both:
  - `Student ID`
  - `Access Code`
- All other sheets contain grade rows with at least:
  - `Student ID`, `Course`, `Assessment`, `Score`
- Optional columns: `Out Of` (default 100), `Weight %`, `Term`, `First Name`, `Last Name`, `Secret`
- Optional: `_aliases` sheet mapping `Key` -> `Value` to support custom headers.

## Deploy (Streamlit Cloud)
1. Create a new **public** GitHub repo and upload:
   - `app.py`
   - `requirements.txt`
2. Deploy on Streamlit Community Cloud.
3. In the app sidebar:
   - Upload `.xlsx` or paste your SharePoint/OneDrive link (set to **Anyone with the link can view** and add `&download=1`).

## Student Flow
1. Enter **Student ID** and **Access Code**.
2. See **Course summary** and **Assessment details**.

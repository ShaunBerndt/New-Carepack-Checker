# HP Care Pack Checker (Streamlit)

This app:
- Uploads an exported HP warranty CSV
- Displays the computed **Warranty End Date Table** (includes Coverage status)
- Lets you download an **updated Excel workbook** based on the supplied template, preserving all formatting and formulas (CSV lines are pasted into `imported data csv` column A)
- Generates **consolidated calendar reminders** (one 10-minute event per day) at **30 and 15 days before expiry** for **non-expired** items.

## Files in this pack
- `app.py` — Streamlit app
- `requirements.txt` — Python dependencies
- `Carepacks Tool_TEMPLATE_UPDATED.xlsx` — updated template (formulas/formatting preserved; added a header for Coverage status on the output sheet)

## Setup

```bash
python -m venv .venv
# Windows: .venv\Scriptsctivate
# Mac/Linux: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Place the Excel template in the same folder as `app.py` (already included in this pack).

## Calendar reminders (.ics)
- Consolidated: **one event per reminder date**
- Duration: **10 minutes**
- Availability: **Available/Free** (does not block time)
- Color: best-effort red using Outlook category name `HP Care Packs` (or choose `Urgent` / `Red`) (create that category in Outlook for consistent red)

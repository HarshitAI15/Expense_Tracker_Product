# StatementSense — Bank Statement Analyzer

Upload your SBI PDF bank statements and get a clean Excel spending report in seconds. No data is stored anywhere — everything runs locally on your machine.

## What you get

- **Transaction Summary** — merchant-wise totals grouped by month
- **Category Details** — every transaction tagged (Food, Transport, Shopping, etc.)
- **Analytics** — credit vs debit, week-wise spend, category breakdown
- **Merged PDF** — all uploaded statements combined into one file

## Project structure

```
├── app.py                   # Streamlit web app (this is what users see)
├── parse_bank_statement.py  # Core parsing and Excel logic
├── requirements.txt         # Python dependencies
└── README.md
```

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

App opens at http://localhost:8501

## Deploy free on Streamlit Cloud

1. Push this folder to a GitHub repo
2. Go to https://share.streamlit.io
3. Click **New app** → select your repo → set `app.py` as entry point
4. Click **Deploy** — live link in ~2 minutes

## Supported statements

- SBI account statements (text-based PDF, not scanned)
- Single month or full year — both work
- Multiple files can be uploaded together

## Known limitations

- Only SBI format supported right now
- Password-protected PDFs not supported
- Scanned/image PDFs won't work (no OCR)

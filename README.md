# Singlife UC2 Underwriting Dashboard

Streamlit dashboard for UC2: optimising underwriting guidelines using early-claim, disclosure, BMI, medical-exam, and underwriting outcome data.

## App

The current app lives in:

```bash
streamlit/app.py
```

Tabs included:

- Overview
- BMI Rule
- F2F / NF2F
- UW Outcomes
- Disclosure & Exams
- New Application Triage

## Run Locally

```bash
cd streamlit
pip install -r requirements.txt
streamlit run app.py
```

The app supports Excel, CSV, and Parquet upload from the sidebar. Raw workbook files are intentionally not committed to this public repository.

## Optional Local Preload

To preload data locally, place a workbook at:

```text
streamlit/data/master_data.xlsx
```

That folder is ignored by git so private data is not pushed publicly by accident.

## Deploy

For Streamlit Community Cloud:

1. Connect this GitHub repository.
2. Set the main file path to `streamlit/app.py`.
3. Use `streamlit/requirements.txt` for dependencies.
4. Upload data in the sidebar after launch, or add a safe public sample file if the app needs to open with preloaded data.

## Governance

This is decision-support analytics. It should help underwriters prioritise review and rule testing, not automatically approve, decline, load, or exclude an application.

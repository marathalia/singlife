# UC2 Streamlit App

This folder contains the Streamlit dashboard for underwriting early-claim risk analysis.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Data

Use the sidebar uploader to load an Excel, CSV, or Parquet file. All metrics, charts, insights, and scenario outputs recalculate from the uploaded file.

For local preload only, place the default workbook here:

```text
data/master_data.xlsx
```

The `data/` folder is ignored by git because the GitHub repository is public.

## Main Features

- Executive overview and recommended underwriting actions
- BMI medical-exam threshold simulator
- F2F / NF2F early-claim comparison
- UW outcome risk and leakage checks
- Disclosure and medical-exam yield analysis
- New application triage using historical comparison groups

The app excludes gender from new-application scoring and is intended for decision support only.

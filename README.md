# AO School Performance Analyzer

A Streamlit web app for analysing SPE 2025 Mathematics Paper 1 and Paper 2 item-analysis Excel files.

## What it calculates

The app calculates AO performance using a marks-based method:

`AO Performance (%) = Total marks scored in AO / Total marks available in AO × 100`

This is suitable for AO2 and AO3 because many items are multi-mark and may show partial understanding.

## Files

- `app.py` - Streamlit application
- `requirements.txt` - Python dependencies
- `data/ao_mapping.csv` - AO reference mapping

## How to run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## How to deploy on Streamlit Cloud

1. Upload these files to a GitHub repository.
2. Go to Streamlit Cloud.
3. Create a new app from the GitHub repository.
4. Set the main file path to `app.py`.
5. Deploy.

## Expected upload format

Upload the official SPE 2025 item-analysis Excel template for each school. The app detects:

- School name
- Question-number row
- Paper 1 columns
- Paper 2 columns
- Student score rows

Blank, `X`, and `A` values are treated as non-attempts and excluded from the denominator.

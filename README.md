# AO School Performance Analyzer

A Streamlit web application for analyzing school performance by Assessment Objective classification: AO1, AO2 and AO3.

## What the app does

The app lets you upload Excel files from individual schools and calculates:

- AO1 performance percentage
- AO2 performance percentage
- AO3 performance percentage
- Separate results for Paper 1 and Paper 2
- Item-level detail report
- Downloadable Excel summary report

The main calculation is marks-based:

```text
AO Performance % = Total marks scored in AO / Total marks available in AO × 100
```

This method is recommended because AO2 and AO3 questions often have multiple marks and students can receive partial credit.

## Files included

```text
app.py                 Main Streamlit application
requirements.txt       Python dependencies
README.md              Project instructions
data/ao_mapping.csv    Editable AO and max-mark reference table
.gitignore             Git ignore file
```

## How to run locally

1. Install Python 3.10 or newer.
2. Open this folder in your terminal.
3. Install requirements:

```bash
pip install -r requirements.txt
```

4. Run the app:

```bash
streamlit run app.py
```

## Expected Excel format

The app works best when each school file has columns similar to:

```text
Q1, Q2, Q3, Q4A, Q5AI, Q5AII, Q15B, etc.
```

The app also normalizes common formats such as:

```text
1, Q1, Question 1, Q11(a), 11a
```

## Important setup note

The max-mark values included are starting values. Please review the mapping table inside the app and adjust the maximum marks if your Excel item columns use different subpart totals.

## Upload modes

The app supports two modes:

1. **One uploaded file per school**  
   The school name is taken from the Excel filename.

2. **Uploaded file already contains a School column**  
   Use this when one Excel file contains multiple schools.

## Deploying on Streamlit Community Cloud

1. Upload this folder to a GitHub repository.
2. Go to Streamlit Community Cloud.
3. Select your GitHub repository.
4. Set the main file path as:

```text
app.py
```

5. Deploy.

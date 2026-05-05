
# (FULL FILE WITH STRUGGLING COLUMN INCLUDED)

import io
import re
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st

st.set_page_config(page_title="AO School Performance Analyzer", layout="wide")

st.title("AO1 / AO2 / AO3 School Performance Analyzer")
st.caption("Upload individual school Excel files, classify items by AO level, and calculate performance by paper.")

PAPER_1_AO_MAP = {"Q1":"AO1","Q2":"AO1","Q3":"AO1","Q4":"AO1","Q5":"AO1","Q6":"AO1","Q7":"AO1","Q8":"AO1","Q10":"AO1",
                 "Q9":"AO2","Q11A":"AO2","Q11B":"AO2","Q12":"AO2","Q13":"AO2","Q14":"AO2","Q15A":"AO2",
                 "Q15B":"AO3","Q16":"AO3","Q17":"AO3"}

PAPER_2_AO_MAP = {"Q1":"AO1","Q4A":"AO1","Q5AI":"AO1","Q6A":"AO1","Q8A":"AO1","Q9A":"AO1","Q12A":"AO1",
                 "Q3":"AO2","Q4B":"AO2","Q5AII":"AO2","Q5B":"AO2","Q6B":"AO2","Q7":"AO2","Q8B":"AO2","Q8C":"AO2",
                 "Q9B":"AO2","Q10":"AO2","Q11":"AO2","Q12B":"AO2","Q13":"AO2",
                 "Q14":"AO3","Q15":"AO3","Q16":"AO3"}

DEFAULT_PAPER_1_MAX_MARKS = {"Q1":3,"Q2":3,"Q3":2,"Q4":2,"Q5":4,"Q6":2,"Q7":2,"Q8":3,
                            "Q9":3,"Q10":5,"Q11A":2,"Q11B":2,"Q12":4,"Q13":3,"Q14":4,
                            "Q15A":3,"Q15B":2,"Q16":5,"Q17":6}

DEFAULT_PAPER_2_MAX_MARKS = {"Q1":5,"Q3":4,"Q4A":1,"Q4B":3,"Q5AI":2,"Q5AII":1,"Q5B":3,
                            "Q6A":2,"Q6B":2,"Q7":3,"Q8A":1,"Q8B":3,"Q8C":2,"Q9A":2,
                            "Q9B":2,"Q10":5,"Q11":5,"Q12A":2,"Q12B":3,"Q13":7,
                            "Q14":4,"Q15":4,"Q16":6}

def normalize_question_name(value):
    text = str(value).upper().strip()
    text = re.sub(r"\s+", "", text)
    if text.startswith("Q") is False:
        text = "Q" + text
    return text

def read_excel_file(uploaded_file):
    return pd.read_excel(uploaded_file, sheet_name=None)

def clean(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df.dropna(how="all")

def find_columns(df, ao_map):
    norm = {normalize_question_name(c): c for c in df.columns}
    return {q: norm[q] for q in ao_map if q in norm}

def calc(df, ao_map, max_marks, school, paper):
    rows = []
    for q, col in find_columns(df, ao_map).items():
        s = pd.to_numeric(df[col], errors="coerce")
        attempts = s.notna().sum()
        scored = s.fillna(0).sum()
        possible = attempts * max_marks.get(q,1)
        pct = (scored/possible*100) if possible else 0
        rows.append({"School":school,"Paper":paper,"AO":ao_map[q],
                     "Total Marks Scored":scored,"Total Marks Available":possible,
                     "Performance %":pct})
    detail = pd.DataFrame(rows)
    if detail.empty:
        return pd.DataFrame(), pd.DataFrame()
    summary = detail.groupby(["School","Paper","AO"], as_index=False).sum()
    summary["Performance %"] = summary["Total Marks Scored"]/summary["Total Marks Available"]*100
    return summary, detail

uploaded_files = st.file_uploader("Upload files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    summaries, details = [], []
    for f in uploaded_files:
        sheets = read_excel_file(f)
        school = f.name.split(".")[0]
        p1 = clean(sheets[list(sheets.keys())[0]])
        p2 = clean(sheets[list(sheets.keys())[-1]])

        s1,d1 = calc(p1,PAPER_1_AO_MAP,DEFAULT_PAPER_1_MAX_MARKS,school,"Paper 1")
        s2,d2 = calc(p2,PAPER_2_AO_MAP,DEFAULT_PAPER_2_MAX_MARKS,school,"Paper 2")

        summaries += [s1,s2]
        details += [d1,d2]

    valid_s = [x for x in summaries if not x.empty]
    valid_d = [x for x in details if not x.empty]

    if not valid_s:
        st.error("No data found")
        st.stop()

    summary_df = pd.concat(valid_s)
    detail_df = pd.concat(valid_d)

    pivot = summary_df.pivot_table(index=["School","Paper"], columns="AO", values="Performance %").reset_index()

    for c in ["AO1","AO2","AO3"]:
        if c in pivot:
            pivot[c] = pivot[c].round(2)

    # STRUGGLING COLUMN
    if "AO2" in pivot and "AO3" in pivot:
        pivot["Struggling"] = ((pivot["AO2"] < 50) | (pivot["AO3"] < 50)).map({True:"Yes",False:"No"})
    else:
        pivot["Struggling"] = "Unknown"

    st.dataframe(pivot)

    pivot_chart = summary_df.copy()
    pivot_chart["School - Paper"] = pivot_chart["School"]+" - "+pivot_chart["Paper"]
    pivot_chart = pivot_chart.pivot_table(index="School - Paper", columns="AO", values="Performance %")

    st.bar_chart(pivot_chart)

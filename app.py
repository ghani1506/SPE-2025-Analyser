import io
import re
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st

st.set_page_config(page_title="AO School Performance Analyzer", layout="wide")

st.title("AO1 / AO2 / AO3 School Performance Analyzer")
st.caption("Upload individual school Excel files, classify items by AO level, and calculate performance by paper.")

PAPER_1_AO_MAP: Dict[str, str] = {
    "Q1": "AO1","Q2": "AO1","Q3": "AO1","Q4": "AO1","Q5": "AO1",
    "Q6": "AO1","Q7": "AO1","Q8": "AO1","Q10": "AO1",
    "Q9": "AO2","Q11A": "AO2","Q11B": "AO2","Q12": "AO2","Q13": "AO2","Q14": "AO2","Q15A": "AO2",
    "Q15B": "AO3","Q16": "AO3","Q17": "AO3",
}

PAPER_2_AO_MAP: Dict[str, str] = {
    "Q1": "AO1","Q4A": "AO1","Q5AI": "AO1","Q6A": "AO1","Q8A": "AO1","Q9A": "AO1","Q12A": "AO1",
    "Q3": "AO2","Q4B": "AO2","Q5AII": "AO2","Q5B": "AO2","Q6B": "AO2","Q7": "AO2","Q8B": "AO2","Q8C": "AO2","Q9B": "AO2","Q10": "AO2","Q11": "AO2","Q12B": "AO2","Q13": "AO2",
    "Q14": "AO3","Q15": "AO3","Q16": "AO3",
}

DEFAULT_PAPER_1_MAX_MARKS = {"Q1":3,"Q2":3,"Q3":2,"Q4":2,"Q5":4,"Q6":2,"Q7":2,"Q8":3,"Q9":3,"Q10":5,"Q11A":2,"Q11B":2,"Q12":4,"Q13":3,"Q14":4,"Q15A":3,"Q15B":2,"Q16":5,"Q17":6}
DEFAULT_PAPER_2_MAX_MARKS = {"Q1":5,"Q3":4,"Q4A":1,"Q4B":3,"Q5AI":2,"Q5AII":1,"Q5B":3,"Q6A":2,"Q6B":2,"Q7":3,"Q8A":1,"Q8B":3,"Q8C":2,"Q9A":2,"Q9B":2,"Q10":5,"Q11":5,"Q12A":2,"Q12B":3,"Q13":7,"Q14":4,"Q15":4,"Q16":6}

def normalize_question_name(value: object) -> str:
    text = str(value).upper().strip()
    text = text.replace("QUESTION","Q")
    text = re.sub(r"\s+","",text)
    text = text.replace("(","").replace(")","").replace(".","").replace("-","")
    if re.match(r"^\d", text):
        text = "Q"+text
    return text

def read_excel_file(uploaded_file):
    return pd.read_excel(uploaded_file, sheet_name=None)

def choose_likely_sheet(sheets, keywords):
    for s in sheets:
        if any(k in s.lower() for k in keywords):
            return s
    return list(sheets.keys())[0]

def clean_school_dataframe(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df.dropna(how="all")

def find_question_columns(df, ao_map):
    norm = {normalize_question_name(c): c for c in df.columns}
    matched = {}
    for q in ao_map:
        qn = normalize_question_name(q)
        if qn in norm:
            matched[q] = norm[qn]
    return matched

def calculate_ao_performance(df, ao_map, max_marks, school, paper):
    matched = find_question_columns(df, ao_map)
    rows=[]
    for q, col in matched.items():
        ao = ao_map[q]
        maxm = float(max_marks.get(q,1))
        s = pd.to_numeric(df[col], errors="coerce")
        attempts = s.notna().sum()
        scored = s.fillna(0).sum()
        possible = attempts * maxm
        pct = (scored/possible*100) if possible else 0
        rows.append({"School":school,"Paper":paper,"AO":ao,"Total Marks Scored":scored,"Total Marks Available":possible,"Performance %":pct})
    detail = pd.DataFrame(rows)
    if detail.empty:
        return pd.DataFrame(), pd.DataFrame()
    summary = detail.groupby(["School","Paper","AO"], as_index=False).sum()
    summary["Performance %"] = summary["Total Marks Scored"]/summary["Total Marks Available"]*100
    return summary, detail

def create_download_excel(summary, detail):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as w:
        summary.to_excel(w, sheet_name="AO Summary", index=False)
        detail.to_excel(w, sheet_name="Item Details", index=False)
    return output.getvalue()

uploaded_files = st.file_uploader("Upload school Excel files", type=["xlsx","xls"], accept_multiple_files=True)
if not uploaded_files:
    st.stop()

all_summaries=[]
all_details=[]

for f in uploaded_files:
    sheets = read_excel_file(f)
    p1 = clean_school_dataframe(sheets[choose_likely_sheet(sheets, ["paper 1","p1"])])
    p2 = clean_school_dataframe(sheets[choose_likely_sheet(sheets, ["paper 2","p2"])])
    school = f.name.split(".")[0]

    s1,d1 = calculate_ao_performance(p1, PAPER_1_AO_MAP, DEFAULT_PAPER_1_MAX_MARKS, school, "Paper 1")
    s2,d2 = calculate_ao_performance(p2, PAPER_2_AO_MAP, DEFAULT_PAPER_2_MAX_MARKS, school, "Paper 2")

    all_summaries.extend([s1,s2])
    all_details.extend([d1,d2])

summary_df = pd.concat([x for x in all_summaries if not x.empty], ignore_index=True)
detail_df = pd.concat([x for x in all_details if not x.empty], ignore_index=True)

pivot = summary_df.pivot_table(index=["School","Paper"], columns="AO", values="Performance %", aggfunc="mean").reset_index()

for col in ["AO1","AO2","AO3"]:
    if col in pivot.columns:
        pivot[col] = pivot[col].round(2)

# 🔴 Struggling logic
pivot["Struggling"] = ((pivot.get("AO2",0) < 50) | (pivot.get("AO3",0) < 50))
pivot["Struggling"] = pivot["Struggling"].map({True:"Yes", False:"No"})

st.dataframe(pivot)

pivot_chart = summary_df.copy()
pivot_chart["School - Paper"] = pivot_chart["School"]+" - "+pivot_chart["Paper"]
pivot_chart = pivot_chart.pivot_table(index="School - Paper", columns="AO", values="Performance %", aggfunc="mean")

st.bar_chart(pivot_chart)

excel_bytes = create_download_excel(pivot, detail_df)
st.download_button("Download report", data=excel_bytes, file_name="ao_report.xlsx")

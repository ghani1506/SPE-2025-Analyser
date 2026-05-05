import io
import re
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st

st.set_page_config(page_title="AO School Performance Analyzer", layout="wide")

st.title("AO1 / AO2 / AO3 School Performance Analyzer")
st.caption("Upload individual school Excel files, classify items by AO level, and calculate marks-based performance by paper.")

# -----------------------------------------------------------------------------
# AO mappings based on SPE 2025 Paper 1 and Paper 2 classification
# -----------------------------------------------------------------------------

PAPER_1_AO_MAP: Dict[str, str] = {
    "Q1": "AO1", "Q2": "AO1", "Q3": "AO1", "Q4": "AO1", "Q5": "AO1",
    "Q6": "AO1", "Q7": "AO1", "Q8": "AO1", "Q10": "AO1",
    "Q9": "AO2", "Q11A": "AO2", "Q11B": "AO2", "Q12": "AO2", "Q13": "AO2",
    "Q14": "AO2", "Q15A": "AO2",
    "Q15B": "AO3", "Q16": "AO3", "Q17": "AO3",
}

PAPER_2_AO_MAP: Dict[str, str] = {
    "Q1": "AO1", "Q4A": "AO1", "Q5AI": "AO1", "Q6A": "AO1",
    "Q8A": "AO1", "Q9A": "AO1", "Q12A": "AO1",
    "Q3": "AO2", "Q4B": "AO2", "Q5AII": "AO2", "Q5B": "AO2", "Q6B": "AO2",
    "Q7": "AO2", "Q8B": "AO2", "Q8C": "AO2", "Q9B": "AO2", "Q10": "AO2",
    "Q11": "AO2", "Q12B": "AO2", "Q13": "AO2",
    "Q14": "AO3", "Q15": "AO3", "Q16": "AO3",
}

# IMPORTANT: Review and edit these values in the app if the Excel item marks differ.
DEFAULT_PAPER_1_MAX_MARKS: Dict[str, float] = {
    "Q1": 3, "Q2": 3, "Q3": 2, "Q4": 2, "Q5": 4, "Q6": 2, "Q7": 2, "Q8": 3,
    "Q9": 3, "Q10": 5, "Q11A": 2, "Q11B": 2, "Q12": 4, "Q13": 3, "Q14": 4,
    "Q15A": 3, "Q15B": 2, "Q16": 5, "Q17": 6,
}

DEFAULT_PAPER_2_MAX_MARKS: Dict[str, float] = {
    "Q1": 5, "Q3": 4, "Q4A": 1, "Q4B": 3, "Q5AI": 2, "Q5AII": 1, "Q5B": 3,
    "Q6A": 2, "Q6B": 2, "Q7": 3, "Q8A": 1, "Q8B": 3, "Q8C": 2, "Q9A": 2,
    "Q9B": 2, "Q10": 5, "Q11": 5, "Q12A": 2, "Q12B": 3, "Q13": 7,
    "Q14": 4, "Q15": 4, "Q16": 6,
}


def normalize_question_name(value: object) -> str:
    """Normalize question column names so Q11(a), 11a, Q11A match."""
    text = str(value).upper().strip()
    text = text.replace("QUESTION", "Q")
    text = re.sub(r"\s+", "", text)
    text = text.replace("(", "").replace(")", "")
    text = text.replace(".", "").replace("-", "").replace("_", "")
    text = text.replace("/", "")
    if re.match(r"^\d", text):
        text = "Q" + text
    return text


def read_excel_file(uploaded_file) -> Dict[str, pd.DataFrame]:
    return pd.read_excel(uploaded_file, sheet_name=None)


def choose_likely_sheet(sheets: Dict[str, pd.DataFrame], keywords: List[str]) -> str:
    lower_keywords = [k.lower() for k in keywords]
    for sheet_name in sheets:
        if any(k in sheet_name.lower() for k in lower_keywords):
            return sheet_name
    return list(sheets.keys())[0]


def clean_school_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    return df


def find_question_columns(df: pd.DataFrame, ao_map: Dict[str, str]) -> Dict[str, str]:
    normalized_to_actual = {normalize_question_name(col): col for col in df.columns}
    matched = {}
    for question in ao_map:
        q_norm = normalize_question_name(question)
        if q_norm in normalized_to_actual:
            matched[question] = normalized_to_actual[q_norm]
    return matched


def calculate_ao_performance(
    df: pd.DataFrame,
    ao_map: Dict[str, str],
    max_marks: Dict[str, float],
    school_name: str,
    paper_name: str,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    matched_columns = find_question_columns(df, ao_map)
    detail_rows = []

    for question, actual_col in matched_columns.items():
        ao = ao_map[question]
        max_mark = float(max_marks.get(question, 1))
        scores = pd.to_numeric(df[actual_col], errors="coerce")
        attempts = scores.notna().sum()
        total_scored = scores.fillna(0).sum()
        total_possible = attempts * max_mark
        percentage = (total_scored / total_possible * 100) if total_possible else 0

        detail_rows.append({
            "School": school_name,
            "Paper": paper_name,
            "Question": question,
            "Excel Column": actual_col,
            "AO": ao,
            "Max Mark": max_mark,
            "Students Attempted": attempts,
            "Total Marks Scored": total_scored,
            "Total Marks Available": total_possible,
            "Performance %": percentage,
        })

    detail_df = pd.DataFrame(detail_rows)
    if detail_df.empty:
        return pd.DataFrame(), detail_df

    summary_df = detail_df.groupby(["School", "Paper", "AO"], as_index=False).agg({
        "Total Marks Scored": "sum",
        "Total Marks Available": "sum",
    })
    summary_df["Performance %"] = summary_df.apply(
        lambda r: (r["Total Marks Scored"] / r["Total Marks Available"] * 100)
        if r["Total Marks Available"] else 0,
        axis=1,
    )
    return summary_df, detail_df


def create_download_excel(summary_df: pd.DataFrame, detail_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="AO Summary", index=False)
        detail_df.to_excel(writer, sheet_name="Item Details", index=False)
    return output.getvalue()


st.sidebar.header("Settings")
analysis_mode = st.sidebar.radio(
    "Excel structure",
    ["One uploaded file per school", "Uploaded file already contains a School column"],
)

school_column_name = None
if analysis_mode == "Uploaded file already contains a School column":
    school_column_name = st.sidebar.text_input("School column name", value="School")

st.sidebar.markdown("---")
st.sidebar.subheader("Method")
st.sidebar.write("AO Performance = Total marks scored in AO / Total marks available in AO × 100")

with st.expander("Review / edit AO mapping and max marks", expanded=False):
    p1_ref = pd.DataFrame([
        {"Paper": "Paper 1", "Question": q, "AO": ao, "Max Mark": DEFAULT_PAPER_1_MAX_MARKS.get(q, 1)}
        for q, ao in PAPER_1_AO_MAP.items()
    ])
    p2_ref = pd.DataFrame([
        {"Paper": "Paper 2", "Question": q, "AO": ao, "Max Mark": DEFAULT_PAPER_2_MAX_MARKS.get(q, 1)}
        for q, ao in PAPER_2_AO_MAP.items()
    ])
    ref_table = pd.concat([p1_ref, p2_ref], ignore_index=True)
    edited_ref = st.data_editor(ref_table, use_container_width=True, num_rows="dynamic")

paper1_ao_map = dict(zip(edited_ref[edited_ref["Paper"] == "Paper 1"]["Question"], edited_ref[edited_ref["Paper"] == "Paper 1"]["AO"]))
paper2_ao_map = dict(zip(edited_ref[edited_ref["Paper"] == "Paper 2"]["Question"], edited_ref[edited_ref["Paper"] == "Paper 2"]["AO"]))
paper1_max_marks = dict(zip(edited_ref[edited_ref["Paper"] == "Paper 1"]["Question"], edited_ref[edited_ref["Paper"] == "Paper 1"]["Max Mark"]))
paper2_max_marks = dict(zip(edited_ref[edited_ref["Paper"] == "Paper 2"]["Question"], edited_ref[edited_ref["Paper"] == "Paper 2"]["Max Mark"]))

uploaded_files = st.file_uploader(
    "Upload school Excel files",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info("Upload one or more Excel files to begin.")
    st.stop()

all_summaries = []
all_details = []

for uploaded_file in uploaded_files:
    sheets = read_excel_file(uploaded_file)
    file_school_name = uploaded_file.name.rsplit(".", 1)[0]

    with st.expander(f"File: {uploaded_file.name}", expanded=False):
        st.write("Detected sheets:", list(sheets.keys()))

    p1_sheet = choose_likely_sheet(sheets, ["paper 1", "p1", "maths 1"])
    p2_sheet = choose_likely_sheet(sheets, ["paper 2", "p2", "maths 2"])

    p1_df = clean_school_dataframe(sheets[p1_sheet])
    p2_df = clean_school_dataframe(sheets[p2_sheet])

    if analysis_mode == "Uploaded file already contains a School column" and school_column_name in p1_df.columns:
        school_names = p1_df[school_column_name].dropna().unique().tolist()
    else:
        school_names = [file_school_name]

    for school in school_names:
        if analysis_mode == "Uploaded file already contains a School column" and school_column_name:
            p1_school_df = p1_df[p1_df[school_column_name] == school] if school_column_name in p1_df.columns else p1_df
            p2_school_df = p2_df[p2_df[school_column_name] == school] if school_column_name in p2_df.columns else p2_df
        else:
            p1_school_df = p1_df
            p2_school_df = p2_df

        p1_summary, p1_detail = calculate_ao_performance(
            p1_school_df, paper1_ao_map, paper1_max_marks, str(school), "Paper 1"
        )
        p2_summary, p2_detail = calculate_ao_performance(
            p2_school_df, paper2_ao_map, paper2_max_marks, str(school), "Paper 2"
        )

        all_summaries.extend([p1_summary, p2_summary])
        all_details.extend([p1_detail, p2_detail])

summary_df = pd.concat([df for df in all_summaries if not df.empty], ignore_index=True) if all_summaries else pd.DataFrame()
detail_df = pd.concat([df for df in all_details if not df.empty], ignore_index=True) if all_details else pd.DataFrame()

if summary_df.empty:
    st.error("No matching question columns were found. Check whether your Excel columns match names like Q1, Q4A, Q15B, etc.")
    st.stop()

st.subheader("AO Performance Summary")
st.caption("Calculation: total marks scored in AO ÷ total marks available in AO × 100")

pivot_summary = summary_df.pivot_table(
    index=["School", "Paper"],
    columns="AO",
    values="Performance %",
    aggfunc="mean",
).reset_index()

for col in ["AO1", "AO2", "AO3"]:
    if col in pivot_summary.columns:
        pivot_summary[col] = pivot_summary[col].round(2)

st.dataframe(pivot_summary, use_container_width=True)

chart_df = summary_df.copy()
chart_df["School - Paper"] = chart_df["School"] + " - " + chart_df["Paper"]
st.bar_chart(chart_df, x="School - Paper", y="Performance %", color="AO")

st.subheader("Item-Level Details")
detail_display = detail_df.copy()
detail_display["Performance %"] = detail_display["Performance %"].round(2)
st.dataframe(detail_display, use_container_width=True)

excel_bytes = create_download_excel(pivot_summary, detail_display)
st.download_button(
    label="Download AO analysis Excel report",
    data=excel_bytes,
    file_name="ao_school_performance_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.success("Analysis complete.")

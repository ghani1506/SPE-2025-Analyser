import io
import re
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st

st.set_page_config(page_title="AO School Performance Analyzer", layout="wide")

st.title("AO1 / AO2 / AO3 School Performance Analyzer")
st.caption("Upload individual school Excel files and calculate marks-based AO performance for SPE 2025 Mathematics Paper 1 and Paper 2.")

# -----------------------------------------------------------------------------
# AO mapping matched to the SPE 2025 Excel template columns.
# Method: AO Performance = total marks scored in AO / total marks available in AO × 100
# -----------------------------------------------------------------------------

PAPER_1_AO_MAP: Dict[str, str] = {
    "1a": "AO1", "1b": "AO1", "1c": "AO1",
    "2a": "AO1", "2b": "AO1", "2c": "AO1",
    "3": "AO1", "4": "AO1",
    "5a": "AO1", "5b": "AO1", "5c": "AO1", "5d": "AO1",
    "6": "AO1", "7": "AO1", "8": "AO1",
    "10a": "AO1", "10b": "AO1", "10c": "AO1",
    "9": "AO2",
    "11ai": "AO2", "11aii": "AO2", "11b": "AO2",
    "12a": "AO2", "12b": "AO2", "13": "AO2",
    "14a": "AO2", "14b": "AO2", "15a": "AO2",
    "15b": "AO3", "16a": "AO3", "16b": "AO3",
    "17a": "AO3", "17b": "AO3", "17c": "AO3",
}

PAPER_2_AO_MAP: Dict[str, str] = {
    "1a": "AO1", "1b": "AO1", "1c": "AO1",
    "2a": "AO1", "2b": "AO1", "2c": "AO1", "2d": "AO1",
    "4a": "AO1", "5ai": "AO1", "6a": "AO1", "8a": "AO1", "9a": "AO1", "12a": "AO1",
    "3": "AO2", "4b": "AO2", "5aii": "AO2", "5b": "AO2", "6b": "AO2",
    "7a": "AO2", "7b": "AO2", "7c": "AO2",
    "8b": "AO2", "8c": "AO2", "9b": "AO2",
    "10a": "AO2", "10b": "AO2", "10c": "AO2", "10d": "AO2",
    "11a": "AO2", "11b": "AO2", "11c": "AO2",
    "12b": "AO2", "13a": "AO2", "13b": "AO2", "13ci": "AO2", "13cii": "AO2",
    "14a": "AO3", "14b": "AO3", "15a": "AO3", "15b": "AO3", "16a": "AO3", "16b": "AO3",
}

PAPER_1_MAX_MARKS: Dict[str, float] = {
    "1a": 1, "1b": 1, "1c": 1, "2a": 1, "2b": 1, "2c": 1, "3": 2, "4": 2,
    "5a": 1, "5b": 1, "5c": 1, "5d": 1, "6": 2, "7": 2, "8": 3, "9": 3,
    "10a": 1, "10b": 2, "10c": 2, "11ai": 1, "11aii": 1, "11b": 2,
    "12a": 2, "12b": 2, "13": 3, "14a": 2, "14b": 2, "15a": 3, "15b": 2,
    "16a": 2, "16b": 3, "17a": 2, "17b": 2, "17c": 2,
}

PAPER_2_MAX_MARKS: Dict[str, float] = {
    "1a": 2, "1b": 1, "1c": 2, "2a": 1, "2b": 2, "2c": 3, "2d": 2, "3": 4,
    "4a": 1, "4b": 3, "5ai": 2, "5aii": 1, "5b": 3, "6a": 2, "6b": 2,
    "7a": 1, "7b": 1, "7c": 1, "8a": 1, "8b": 3, "8c": 2, "9a": 2, "9b": 2,
    "10a": 2, "10b": 1, "10c": 1, "10d": 1, "11a": 1, "11b": 3, "11c": 1,
    "12a": 2, "12b": 3, "13a": 2, "13b": 1, "13ci": 2, "13cii": 2,
    "14a": 2, "14b": 2, "15a": 2, "15b": 2, "16a": 2, "16b": 4,
}


def normalize_item(value: object) -> str:
    text = str(value).lower().strip()
    text = text.replace("question", "")
    text = re.sub(r"\s+", "", text)
    text = text.replace("(", "").replace(")", "")
    text = text.replace(".", "").replace("-", "").replace("_", "").replace("/", "")
    text = text.replace("q", "", 1) if text.startswith("q") else text
    return text


def read_template_sheet(uploaded_file) -> pd.DataFrame:
    """Read the first non-empty Excel sheet without assuming headers."""
    sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)
    for _, df in sheets.items():
        if not df.dropna(how="all").empty:
            return df
    return pd.DataFrame()


def extract_school_name(raw_df: pd.DataFrame, fallback: str) -> str:
    for r in range(min(len(raw_df), 30)):
        row_values = raw_df.iloc[r].tolist()
        for c, value in enumerate(row_values):
            if isinstance(value, str) and "NAMA MAKTAB" in value.upper():
                for next_c in range(c + 1, min(c + 5, len(row_values))):
                    candidate = row_values[next_c]
                    if pd.notna(candidate):
                        return str(candidate).strip()
    return fallback


def find_question_row(raw_df: pd.DataFrame) -> int | None:
    """Find the row containing item labels such as 1a, 1b, 2a, etc."""
    for r in range(min(len(raw_df), 60)):
        values = [normalize_item(v) for v in raw_df.iloc[r].tolist() if pd.notna(v)]
        hits = sum(v in PAPER_1_MAX_MARKS or v in PAPER_2_MAX_MARKS for v in values)
        if hits >= 10:
            return r
    return None


def find_paper_ranges(raw_df: pd.DataFrame, question_row: int) -> Tuple[Tuple[int, int], Tuple[int, int]]:
    """Detect Paper 1 and Paper 2 question-column ranges from the template."""
    # In the official template, row 18 has PAPER 1 at col 5 and PAPER 2 at col 40,
    # and row 21 has total columns at 39 and 82. This fallback also works if row numbers shift.
    paper_row = max(0, question_row - 3)
    p1_start, p2_start = None, None
    for r in range(max(0, question_row - 6), question_row + 1):
        for c, value in enumerate(raw_df.iloc[r].tolist()):
            text = str(value).upper() if pd.notna(value) else ""
            if "PAPER 1" in text and p1_start is None:
                p1_start = c
            if "PAPER 2" in text and p2_start is None:
                p2_start = c
    if p1_start is None:
        p1_start = 5
    if p2_start is None:
        # first repeated item 1a after Paper 1 usually indicates Paper 2
        labels = [normalize_item(v) for v in raw_df.iloc[question_row].tolist()]
        one_a_cols = [i for i, v in enumerate(labels) if v == "1a"]
        p2_start = one_a_cols[1] if len(one_a_cols) > 1 else 40

    labels = [normalize_item(v) for v in raw_df.iloc[question_row].tolist()]
    p1_end = p2_start
    p2_end = len(labels)
    for c in range(p1_start, p2_start):
        if str(raw_df.iloc[question_row, c]).strip() in ["60", "60.0"]:
            p1_end = c
            break
    for c in range(p2_start, len(labels)):
        if str(raw_df.iloc[question_row, c]).strip() in ["80", "80.0"]:
            p2_end = c
            break
    return (p1_start, p1_end), (p2_start, p2_end)


def extract_student_scores(raw_df: pd.DataFrame, question_row: int, col_range: Tuple[int, int]) -> pd.DataFrame:
    start, end = col_range
    headers = [normalize_item(v) for v in raw_df.iloc[question_row, start:end].tolist()]
    data = raw_df.iloc[question_row + 1 :, start:end].copy()
    data.columns = headers

    # Keep rows that look like student rows. Column 1 in template is student name.
    names = raw_df.iloc[question_row + 1 :, 1] if raw_df.shape[1] > 1 else pd.Series(dtype=object)
    data = data[names.notna().values]

    # Convert X/A/blank to missing; numeric marks remain.
    for col in data.columns:
        data[col] = pd.to_numeric(data[col], errors="coerce")
    return data


def calculate_ao_performance(scores_df: pd.DataFrame, ao_map: Dict[str, str], max_marks: Dict[str, float], school: str, paper: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    detail_rows = []
    available_cols = set(scores_df.columns)
    for item, ao in ao_map.items():
        item_norm = normalize_item(item)
        if item_norm not in available_cols:
            continue
        marks = pd.to_numeric(scores_df[item_norm], errors="coerce")
        attempts = int(marks.notna().sum())
        scored = float(marks.fillna(0).sum())
        max_mark = float(max_marks[item])
        possible = attempts * max_mark
        pct = scored / possible * 100 if possible else 0
        detail_rows.append({
            "School": school,
            "Paper": paper,
            "Item": item,
            "AO": ao,
            "Max Mark": max_mark,
            "Students Attempted": attempts,
            "Total Marks Scored": scored,
            "Total Marks Available": possible,
            "Performance %": pct,
        })

    detail_df = pd.DataFrame(detail_rows)
    if detail_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    summary_df = detail_df.groupby(["School", "Paper", "AO"], as_index=False).agg({
        "Total Marks Scored": "sum",
        "Total Marks Available": "sum",
    })
    summary_df["Performance %"] = summary_df["Total Marks Scored"] / summary_df["Total Marks Available"] * 100
    return summary_df, detail_df


def safe_concat(frames: List[pd.DataFrame]) -> pd.DataFrame:
    valid = [f for f in frames if isinstance(f, pd.DataFrame) and not f.empty]
    return pd.concat(valid, ignore_index=True) if valid else pd.DataFrame()


def create_download_excel(summary: pd.DataFrame, details: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="AO Summary", index=False)
        details.to_excel(writer, sheet_name="Item Details", index=False)
    return output.getvalue()


st.sidebar.header("Method")
st.sidebar.write("AO Performance = Total marks scored in AO ÷ Total marks available in AO × 100")
st.sidebar.write("Blank/X/A entries are excluded from attempts.")

with st.expander("Review AO mapping and max marks", expanded=False):
    ref = pd.concat([
        pd.DataFrame([{"Paper": "Paper 1", "Item": k, "AO": PAPER_1_AO_MAP[k], "Max Mark": PAPER_1_MAX_MARKS[k]} for k in PAPER_1_AO_MAP]),
        pd.DataFrame([{"Paper": "Paper 2", "Item": k, "AO": PAPER_2_AO_MAP[k], "Max Mark": PAPER_2_MAX_MARKS[k]} for k in PAPER_2_AO_MAP]),
    ], ignore_index=True)
    st.dataframe(ref, use_container_width=True)

uploaded_files = st.file_uploader("Upload individual school Excel files", type=["xlsx", "xls"], accept_multiple_files=True)

if not uploaded_files:
    st.info("Upload one or more school Excel files to begin.")
    st.stop()

summary_frames: List[pd.DataFrame] = []
detail_frames: List[pd.DataFrame] = []
problems: List[str] = []

for uploaded_file in uploaded_files:
    raw = read_template_sheet(uploaded_file)
    fallback_school = uploaded_file.name.rsplit(".", 1)[0]
    school = extract_school_name(raw, fallback_school)
    q_row = find_question_row(raw)

    if raw.empty or q_row is None:
        problems.append(f"{uploaded_file.name}: could not find the question-number row.")
        continue

    p1_range, p2_range = find_paper_ranges(raw, q_row)
    p1_scores = extract_student_scores(raw, q_row, p1_range)
    p2_scores = extract_student_scores(raw, q_row, p2_range)

    p1_summary, p1_detail = calculate_ao_performance(p1_scores, PAPER_1_AO_MAP, PAPER_1_MAX_MARKS, school, "Paper 1")
    p2_summary, p2_detail = calculate_ao_performance(p2_scores, PAPER_2_AO_MAP, PAPER_2_MAX_MARKS, school, "Paper 2")

    summary_frames.extend([p1_summary, p2_summary])
    detail_frames.extend([p1_detail, p2_detail])

    with st.expander(f"Detected structure: {uploaded_file.name}", expanded=False):
        st.write(f"School detected: **{school}**")
        st.write(f"Question row detected: Excel row {q_row + 1}")
        st.write(f"Paper 1 columns: {p1_range[0] + 1} to {p1_range[1]}")
        st.write(f"Paper 2 columns: {p2_range[0] + 1} to {p2_range[1]}")

summary_df = safe_concat(summary_frames)
detail_df = safe_concat(detail_frames)

if problems:
    st.warning("Some files could not be processed:\n\n" + "\n".join(f"- {p}" for p in problems))

if summary_df.empty or detail_df.empty:
    st.error("No AO results could be produced. Please confirm the uploaded file uses the SPE 2025 item-analysis Excel template.")
    st.stop()

st.subheader("AO Performance Summary")
pivot_chart = summary_df.pivot_table(
    index="School - Paper",
    columns="AO",
    values="Performance %",
    aggfunc="mean"
)

st.bar_chart(pivot_chart)


st.subheader("Item-Level Details")
show_detail = detail_df.copy()
show_detail["Performance %"] = show_detail["Performance %"].round(2)
st.dataframe(show_detail, use_container_width=True)

st.download_button(
    "Download AO analysis Excel report",
    data=create_download_excel(pivot, show_detail),
    file_name="ao_school_performance_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.success("Analysis complete.")

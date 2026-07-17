from __future__ import annotations

from collections import Counter
from io import BytesIO
from pathlib import Path

import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill


APP_TITLE = "KDHW Children Verification"
SAMPLE_WORKBOOK = Path(__file__).with_name("KDHW Sample dataset.xlsx")
SOURCE_SHEET = "children"
CODE_HEADER = "children_code"
VERIFICATION_SHEET = "children_verification"
RED_FILL = PatternFill(fill_type="solid", fgColor="FFC7CE")
RED_FONT = Font(color="9C0006")


def normalize_code(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def load_source_workbook(uploaded_file):
    if uploaded_file is not None:
        return load_workbook(uploaded_file), None
    if not SAMPLE_WORKBOOK.exists():
        raise FileNotFoundError(f"Sample workbook not found: {SAMPLE_WORKBOOK}")
    return load_workbook(SAMPLE_WORKBOOK), SAMPLE_WORKBOOK


def get_children_sheet(workbook: Workbook):
    if SOURCE_SHEET not in workbook.sheetnames:
        available = ", ".join(workbook.sheetnames)
        raise KeyError(f"Sheet '{SOURCE_SHEET}' was not found. Available sheets: {available}")
    return workbook[SOURCE_SHEET]


def build_verification_report(source_sheet) -> tuple[Workbook, dict[str, int]]:
    headers = [source_sheet.cell(1, column).value for column in range(1, source_sheet.max_column + 1)]
    if CODE_HEADER not in headers:
        raise KeyError(f"Column '{CODE_HEADER}' was not found in sheet '{SOURCE_SHEET}'.")

    code_column = headers.index(CODE_HEADER) + 1
    rows = list(source_sheet.iter_rows(min_row=2, values_only=True))
    normalized_codes = [normalize_code(row[code_column - 1]) for row in rows]
    non_empty_codes = [code for code in normalized_codes if code]
    code_counts = Counter(non_empty_codes)
    duplicate_codes = {code for code, count in code_counts.items() if count > 1}

    output = Workbook()
    report = output.active
    report.title = VERIFICATION_SHEET

    report_headers = headers + ["verification_status"]
    report.append(report_headers)

    missing_count = 0
    duplicate_row_count = 0

    for row_index, row in enumerate(rows, start=2):
        row_values = list(row)
        code_value = normalize_code(row_values[code_column - 1])
        status = "OK"

        if not code_value:
            status = "Missing children_code"
            missing_count += 1
        elif code_value in duplicate_codes:
            status = "Duplicate children_code"
            duplicate_row_count += 1

        row_values.append(status)
        report.append(row_values)

        if status != "OK":
            code_cell = report.cell(row=row_index, column=code_column)
            code_cell.fill = RED_FILL
            code_cell.font = RED_FONT

            status_cell = report.cell(row=row_index, column=len(report_headers))
            status_cell.fill = RED_FILL
            status_cell.font = RED_FONT

    duplicate_code_count = len(duplicate_codes)
    affected_rows = missing_count + duplicate_row_count
    summary = {
        "rows": len(rows),
        "missing_rows": missing_count,
        "duplicate_rows": duplicate_row_count,
        "duplicate_codes": duplicate_code_count,
        "affected_rows": affected_rows,
    }
    return output, summary


def workbook_to_bytes(workbook: Workbook) -> bytes:
    buffer = BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer.read()


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)
    st.write(
        "Upload a KDHW workbook or use the bundled sample file. The app checks the children sheet, "
        "flags missing or duplicate children_code values, and generates a verification workbook with red highlights."
    )

    uploaded_file = st.file_uploader("Upload an .xlsx workbook", type=["xlsx"])

    try:
        source_workbook, source_path = load_source_workbook(uploaded_file)
        source_sheet = get_children_sheet(source_workbook)
        report_workbook, summary = build_verification_report(source_sheet)
    except Exception as exc:
        st.error(str(exc))
        return

    if source_path is not None:
        st.caption(f"Using sample workbook: {source_path.name}")
    else:
        st.caption(f"Using uploaded workbook: {uploaded_file.name}")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Rows checked", summary["rows"])
    col2.metric("Missing children_code", summary["missing_rows"])
    col3.metric("Duplicate rows", summary["duplicate_rows"])
    col4.metric("Duplicate codes", summary["duplicate_codes"])

    st.info(f"Rows with issues: {summary['affected_rows']}")

    report_bytes = workbook_to_bytes(report_workbook)
    st.download_button(
        label="Download verification workbook",
        data=report_bytes,
        file_name="KDHW_children_verification.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.subheader("Rules applied")
    st.write("Missing children_code values are flagged red.")
    st.write("Duplicate children_code values are flagged red on every matching row.")


if __name__ == "__main__":
    main()
from __future__ import annotations

from collections import Counter
from io import BytesIO
from datetime import date, datetime

import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill


APP_TITLE = "KDHW Children Verification"
SOURCE_SHEET = "children"
CODE_HEADER = "children_code"
VERIFICATION_SHEET = "children_verification"
RED_FILL = PatternFill(fill_type="solid", fgColor="FFC7CE")
RED_FONT = Font(color="9C0006")
DATE_COLUMNS = [
    "registered_date",
    "BCG",
    "OPV_first_time_dose",
    "OPV_second_time_dose",
    "OPV_third_time_dose",
    "MMR_first_time_dose",
    "MMR_second_time_dose",
    "IPV",
    "PCV_first_time_dose",
    "PCV_second_time_dose",
    "PCV_third_time_dose",
    "Rota_first_time_dose",
    "Rota_second_time_dose",
]


def normalize_code(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_date(value: object):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def load_source_workbook(uploaded_file):
    if uploaded_file is None:
        raise ValueError("Please upload a workbook to verify.")
    return load_workbook(uploaded_file), uploaded_file.name


def get_children_sheet(workbook: Workbook):
    if SOURCE_SHEET not in workbook.sheetnames:
        available = ", ".join(workbook.sheetnames)
        raise KeyError(f"Sheet '{SOURCE_SHEET}' was not found. Available sheets: {available}")
    return workbook[SOURCE_SHEET]


def build_verification_report(source_sheet) -> tuple[Workbook, dict[str, int]]:
    headers = [source_sheet.cell(1, column).value for column in range(1, source_sheet.max_column + 1)]
    if CODE_HEADER not in headers:
        raise KeyError(f"Column '{CODE_HEADER}' was not found in sheet '{SOURCE_SHEET}'.")
    if "date_of_birth" not in headers:
        raise KeyError("Column 'date_of_birth' was not found in sheet 'children'.")

    code_column = headers.index(CODE_HEADER) + 1
    dob_column = headers.index("date_of_birth") + 1
    registered_column = headers.index("registered_date") + 1 if "registered_date" in headers else None
    comparison_columns = [headers.index(name) + 1 for name in DATE_COLUMNS if name in headers and name != "registered_date"]
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
    dob_order_row_count = 0

    for row_index, row in enumerate(rows, start=2):
        row_values = list(row)
        code_value = normalize_code(row_values[code_column - 1])
        dob_value = normalize_date(row_values[dob_column - 1])
        status = "OK"
        issues: list[str] = []

        if not code_value:
            status = "Missing children_code"
            missing_count += 1
        elif code_value in duplicate_codes:
            status = "Duplicate children_code"
            duplicate_row_count += 1

        if dob_value is not None and registered_column is not None:
            registered_value = normalize_date(row_values[registered_column - 1])
            if registered_value is not None and dob_value > registered_value:
                issues.append("date_of_birth later than registered_date")

        if dob_value is not None:
            for column in comparison_columns:
                compared_value = normalize_date(row_values[column - 1])
                if compared_value is not None and dob_value > compared_value:
                    issues.append(f"date_of_birth later than {headers[column - 1]}")

        if issues:
            dob_order_row_count += 1
            if status == "OK":
                status = "; ".join(issues)
            else:
                status = f"{status}; " + "; ".join(issues)

        row_values.append(status)
        report.append(row_values)

        if status != "OK":
            code_cell = report.cell(row=row_index, column=code_column)
            code_cell.fill = RED_FILL
            code_cell.font = RED_FONT

            dob_cell = report.cell(row=row_index, column=dob_column)
            dob_cell.fill = RED_FILL
            dob_cell.font = RED_FONT

            if registered_column is not None:
                registered_value = normalize_date(row_values[registered_column - 1])
                if dob_value is not None and registered_value is not None and dob_value > registered_value:
                    registered_cell = report.cell(row=row_index, column=registered_column)
                    registered_cell.fill = RED_FILL
                    registered_cell.font = RED_FONT

            for column in comparison_columns:
                compared_value = normalize_date(row_values[column - 1])
                if dob_value is not None and compared_value is not None and dob_value > compared_value:
                    compared_cell = report.cell(row=row_index, column=column)
                    compared_cell.fill = RED_FILL
                    compared_cell.font = RED_FONT

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
        "dob_order_rows": dob_order_row_count,
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

    if uploaded_file is None:
        st.info("Upload a workbook to generate the verification file.")
        return

    try:
        source_workbook, source_path = load_source_workbook(uploaded_file)
        source_sheet = get_children_sheet(source_workbook)
        report_workbook, summary = build_verification_report(source_sheet)
    except Exception as exc:
        st.error(str(exc))
        return

    st.caption(f"Using uploaded workbook: {source_path}")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Rows checked", summary["rows"])
    col2.metric("Missing children_code", summary["missing_rows"])
    col3.metric("Duplicate rows", summary["duplicate_rows"])
    col4.metric("Duplicate codes", summary["duplicate_codes"])

    st.metric("Rows with date order issues", summary["dob_order_rows"])

    st.info(f"Rows with issues: {summary['affected_rows']}")

    report_bytes = workbook_to_bytes(report_workbook)
    st.download_button(
        label="Download new verification file",
        data=report_bytes,
        file_name="KDHW_children_verification.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.subheader("Rules applied")
    st.write("Missing children_code values are flagged red.")
    st.write("Duplicate children_code values are flagged red on every matching row.")
    st.write("date_of_birth is checked against registered_date and the listed vaccine date columns; later values are flagged red.")


if __name__ == "__main__":
    main()

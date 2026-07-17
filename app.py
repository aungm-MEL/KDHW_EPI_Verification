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
SOURCE_FIELD_RULES = [
    ("BCG", "BCG_other", "BCG_source"),
    ("OPV_first_time_dose", "OPV_first_time_dose_other", "OPV1_source"),
    ("OPV_second_time_dose", "OPV_second_time_dose_other", "OPV2_source"),
    ("OPV_third_time_dose", "OPV_third_time_dose_other", "OPV3_source"),
    ("Penta_first_time_dose", "Penta_first_time_dose_other", "Penta1_source"),
    ("Penta_second_time_dose", "Penta_second_time_dose_other", "Penta2_source"),
    ("Penta_third_time_dose", "Penta_third_time_dose_other", "Penta3_source"),
    ("Penta_fourth_time_dose", "Penta_fourth_time_dose_other", "Penta4_source"),
    ("MMR_first_time_dose", "MMR_first_time_dose_other", "MMR1_source"),
    ("MMR_second_time_dose", "MMR_second_time_dose_other", "MMR2_source"),
    ("JE_first_time_dose", "JE_first_time_dose_other", "JE1_source"),
    ("JE_second_time_dose", "JE_second_time_dose_other", "JE2_source"),
    ("IPV", "IPV_other", "IPV_source"),
    ("Rota_first_time_dose", "Rota_first_time_dose_other", "Rota1_source"),
    ("Rota_second_time_dose", "Rota_second_time_dose_other", "Rota2_source"),
]
EXCLUDED_OUTPUT_HEADERS = {
    "HPV_first_time_dose",
    "HPV_first_time_dose_other",
    "HPV_first_time_dose_reporting_month",
    "HPV_second_time_dose",
    "HPV_second_time_dose_other",
    "HPV_second_time_dose_reporting_month",
    "HepB",
    "HepB_other",
    "HepB_reporting_month",
    "PCV_first_time_dose",
    "PCV_first_time_dose_other",
    "PCV_first_time_dose_reporting_month",
    "PCV_second_time_dose",
    "PCV_second_time_dose_other",
    "PCV_second_time_dose_reporting_month",
    "PCV_third_time_dose",
    "PCV_third_time_dose_other",
    "PCV_third_time_dose_reporting_month",
    "full_prevention",
    "full_prevention_reporting_month",
    "complete_dose",
    "complete_dose_reporting_month",
    "comments",
}
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


def build_source_value(date_value: object, other_value: object) -> str:
    if normalize_code(other_value).casefold() == "yes":
        return "Other"
    if normalize_date(date_value) is not None:
        return "KDHW"
    return "Not received yet"


def should_exclude_output_column(header: object) -> bool:
    return isinstance(header, str) and header in EXCLUDED_OUTPUT_HEADERS


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

    included_columns = [index for index, header in enumerate(headers, start=1) if not should_exclude_output_column(header)]
    present_source_rules = [rule for rule in SOURCE_FIELD_RULES if rule[0] in headers and rule[1] in headers]
    source_columns = [source_header for _, _, source_header in present_source_rules]
    report_headers = [headers[index - 1] for index in included_columns] + source_columns + ["verification_status"]
    report.append(report_headers)

    report_column_map = {source_index: report_index for report_index, source_index in enumerate(included_columns, start=1)}

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

        filtered_row_values = [row_values[index - 1] for index in included_columns]
        derived_source_values = []
        for date_header, other_header, _source_header in present_source_rules:
            date_value = row_values[headers.index(date_header)]
            other_value = row_values[headers.index(other_header)]
            derived_source_values.append(build_source_value(date_value, other_value))
        report.append(filtered_row_values + derived_source_values + [status])

        if status != "OK":
            code_cell = report.cell(row=row_index, column=report_column_map[code_column])
            code_cell.fill = RED_FILL
            code_cell.font = RED_FONT

            dob_cell = report.cell(row=row_index, column=report_column_map[dob_column])
            if issues:
                dob_cell.fill = RED_FILL
                dob_cell.font = RED_FONT

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

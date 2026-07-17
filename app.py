from __future__ import annotations

from collections import Counter
from io import BytesIO
from datetime import date, datetime

import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill


APP_TITLE = "KDHW Children Verification"
SOURCE_SHEET = "children"
PW_SHEET = "PW"
CODE_HEADER = "children_code"
PW_CODE_HEADER = "mother_code"
VERIFICATION_SHEET = "children_verification"
LONG_FORM_SHEET = "children_long_form"
PW_VERIFICATION_SHEET = "pw_verification"
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
DOSE_DATE_HEADERS = {
    "BCG": "BCG",
    "OPV1": "OPV_first_time_dose",
    "OPV2": "OPV_second_time_dose",
    "OPV3": "OPV_third_time_dose",
    "Penta1": "Penta_first_time_dose",
    "Penta2": "Penta_second_time_dose",
    "Penta3": "Penta_third_time_dose",
    "Penta4": "Penta_fourth_time_dose",
    "MMR1": "MMR_first_time_dose",
    "MMR2": "MMR_second_time_dose",
    "JE1": "JE_first_time_dose",
    "JE2": "JE_second_time_dose",
    "IPV": "IPV",
    "Rota1": "Rota_first_time_dose",
    "Rota2": "Rota_second_time_dose",
}
DOSE_SOURCE_HEADERS = {
    "BCG": "BCG_source",
    "OPV1": "OPV1_source",
    "OPV2": "OPV2_source",
    "OPV3": "OPV3_source",
    "Penta1": "Penta1_source",
    "Penta2": "Penta2_source",
    "Penta3": "Penta3_source",
    "Penta4": "Penta4_source",
    "MMR1": "MMR1_source",
    "MMR2": "MMR2_source",
    "JE1": "JE1_source",
    "JE2": "JE2_source",
    "IPV": "IPV_source",
    "Rota1": "Rota1_source",
    "Rota2": "Rota2_source",
}
DOSE_REPORTING_HEADERS = {
    "BCG": "BCG_reporting_month",
    "OPV1": "OPV_first_time_dose_reporting_month",
    "OPV2": "OPV_second_time_dose_reporting_month",
    "OPV3": "OPV_third_time_dose_reporting_month",
    "Penta1": "Penta_first_time_dose_reporting_month",
    "Penta2": "Penta_second_time_dose_reporting_month",
    "Penta3": "Penta_third_time_dose_reporting_month",
    "Penta4": "Penta_fourth_time_dose_reporting_month",
    "MMR1": "MMR_first_time_dose_reporting_month",
    "MMR2": "MMR_second_time_dose_reporting_month",
    "JE1": "JE_first_time_dose_reporting_month",
    "JE2": "JE_second_time_dose_reporting_month",
    "IPV": "IPV_reporting_month",
    "Rota1": "Rota_first_time_dose_reporting_month",
    "Rota2": "Rota_second_time_dose_reporting_month",
}
LONG_FORM_DOSES = [
    "BCG",
    "OPV1",
    "OPV2",
    "OPV3",
    "Penta1",
    "Penta2",
    "Penta3",
    "Penta4",
    "MMR1",
    "MMR2",
    "JE1",
    "JE2",
    "IPV",
    "Rota1",
    "Rota2",
]
U1_REQUIRED_DOSES = ["BCG", "OPV1", "OPV2", "OPV3", "Penta1", "Penta2", "Penta3", "MMR1"]
ONE_TO_FIVE_REQUIRED_DOSES = ["OPV1", "OPV2", "OPV3", "Penta1", "Penta2", "Penta3", "MMR1"]
COMPLETION_CHECK_DOSES = ["BCG", "OPV1", "OPV2", "OPV3", "Penta1", "Penta2", "Penta3", "MMR1", "MMR2"]
SEQUENCE_RULES = {
    "OPV": ["OPV1", "OPV2", "OPV3"],
    "Penta": ["Penta1", "Penta2", "Penta3"],
    "MMR": ["MMR1", "MMR2"],
}
PW_DROP_HEADERS = {
    "td_third_dose",
    "td_third_dose_other",
    "td_third_dose_reporting_month",
    "td_fourth_dose",
    "td_fourth_dose_other",
    "td_fourth_dose_reporting_month",
    "td_fifth_dose",
    "td_fifth_dose_other",
    "td_fifth_dose_reporting_month",
    "prevent_td_newborn",
    "comments",
}
PW_SOURCE_FIELD_RULES = [
    ("td_first_dose", "td_first_dose_other", "Td1_source"),
    ("td_second_dose", "td_second_dose_other", "Td2_source"),
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


def normalize_age_months(value: object):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return int(value)
    text = normalize_code(value)
    if not text:
        return None
    try:
        return int(float(text))
    except ValueError:
        return None


def quarter_label(value: date) -> str:
    quarter = ((value.month - 1) // 3) + 1
    return f"Q{quarter} {value.year}"


def calculate_completed_dose(age_months, source_values: dict[str, str], date_values: dict[str, date | None]) -> str:
    if age_months is None:
        return "not complete yet"

    if 0 <= age_months <= 11:
        group_label = "U1"
        required_doses = U1_REQUIRED_DOSES
    elif 12 <= age_months <= 59:
        group_label = "1-5"
        required_doses = ONE_TO_FIVE_REQUIRED_DOSES
    else:
        return "not complete yet"

    if any(source_values.get(dose, "Not received yet") == "Not received yet" for dose in required_doses):
        return "not complete yet"

    kdhw_dates = [
        date_values.get(dose)
        for dose in COMPLETION_CHECK_DOSES
        if source_values.get(dose) == "KDHW" and date_values.get(dose) is not None
    ]
    if not kdhw_dates:
        return "not complete yet"

    last_kdhw_date = max(kdhw_dates)
    return f"{group_label} completed in {quarter_label(last_kdhw_date)}"


def has_received(source_value: str | None) -> bool:
    return normalize_code(source_value) not in {"", "Not received yet"}


def find_unlogical_sequence_issues(source_values: dict[str, str], date_values: dict[str, date | None]) -> list[str]:
    issues: list[str] = []

    for group_name, doses in SEQUENCE_RULES.items():
        received_flags = [has_received(source_values.get(dose)) for dose in doses]

        for idx, dose in enumerate(doses[:-1]):
            if not received_flags[idx] and any(received_flags[idx + 1 :]):
                later_doses = ", ".join(doses[idx + 1 :])
                issues.append(f"{dose} not received but later dose(s) received ({later_doses})")

        for earlier_idx, earlier_dose in enumerate(doses[:-1]):
            earlier_date = date_values.get(earlier_dose)
            if earlier_date is None:
                continue

            for later_dose in doses[earlier_idx + 1 :]:
                later_date = date_values.get(later_dose)
                if later_date is not None and later_date < earlier_date:
                    issues.append(f"{later_dose} date earlier than {earlier_dose} date")

    return issues


def datediff_months(start_date: date | None, end_date: date | None):
    if start_date is None or end_date is None:
        return None
    if end_date < start_date:
        return None

    months = (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month)
    if end_date.day < start_date.day:
        months -= 1
    return months


def build_source_value(date_value: object, other_value: object) -> str:
    if normalize_code(other_value).casefold() == "yes":
        return "Other"
    if normalize_date(date_value) is not None:
        return "KDHW"
    return "Not received yet"


def should_exclude_output_column(header: object) -> bool:
    return isinstance(header, str) and header in EXCLUDED_OUTPUT_HEADERS


def build_output_columns(headers: list[object]) -> tuple[list[int], list[str], dict[int, int]]:
    included_columns: list[int] = []
    output_headers: list[str] = []
    source_column_map: dict[int, int] = {}

    other_to_source = {other_header: source_header for _, other_header, source_header in SOURCE_FIELD_RULES}

    for source_index, header in enumerate(headers, start=1):
        if should_exclude_output_column(header):
            continue

        if isinstance(header, str) and header in other_to_source:
            included_columns.append(source_index)
            source_column_map[source_index] = len(output_headers) + 1
            output_headers.append(other_to_source[header])
            continue

        included_columns.append(source_index)
        source_column_map[source_index] = len(output_headers) + 1
        output_headers.append(header if isinstance(header, str) else "")

    return included_columns, output_headers, source_column_map


def load_source_workbook(uploaded_file):
    if uploaded_file is None:
        raise ValueError("Please upload a workbook to verify.")
    return load_workbook(uploaded_file), uploaded_file.name


def get_children_sheet(workbook: Workbook):
    if SOURCE_SHEET not in workbook.sheetnames:
        available = ", ".join(workbook.sheetnames)
        raise KeyError(f"Sheet '{SOURCE_SHEET}' was not found. Available sheets: {available}")
    return workbook[SOURCE_SHEET]


def get_pw_sheet(workbook: Workbook):
    if PW_SHEET not in workbook.sheetnames:
        available = ", ".join(workbook.sheetnames)
        raise KeyError(f"Sheet '{PW_SHEET}' was not found. Available sheets: {available}")
    return workbook[PW_SHEET]


def build_pw_verification_sheet(output: Workbook, pw_sheet) -> dict[str, int]:
    headers = [pw_sheet.cell(1, column).value for column in range(1, pw_sheet.max_column + 1)]
    if PW_CODE_HEADER not in headers:
        raise KeyError(f"Column '{PW_CODE_HEADER}' was not found in sheet '{PW_SHEET}'.")

    mother_code_column = headers.index(PW_CODE_HEADER) + 1
    rows = list(pw_sheet.iter_rows(min_row=2, values_only=True))

    normalized_codes = [normalize_code(row[mother_code_column - 1]) for row in rows]
    non_empty_codes = [code for code in normalized_codes if code]
    code_counts = Counter(non_empty_codes)
    duplicate_codes = {code for code, count in code_counts.items() if count > 1}

    report = output.create_sheet(title=PW_VERIFICATION_SHEET)

    present_source_rules = [rule for rule in PW_SOURCE_FIELD_RULES if rule[0] in headers and rule[1] in headers]
    other_to_source = {other_header: source_header for _, other_header, source_header in present_source_rules}

    included_columns: list[int] = []
    output_headers: list[str] = []
    report_column_map: dict[int, int] = {}

    for source_index, header in enumerate(headers, start=1):
        if isinstance(header, str) and header in PW_DROP_HEADERS:
            continue

        included_columns.append(source_index)
        report_column_map[source_index] = len(output_headers) + 1
        if isinstance(header, str) and header in other_to_source:
            output_headers.append(other_to_source[header])
        else:
            output_headers.append(header if isinstance(header, str) else "")

    report.append(output_headers + ["verification_status"])

    missing_row_count = 0
    duplicate_row_count = 0
    unlogical_row_count = 0
    affected_row_count = 0

    rule_by_other_header = {rule[1]: rule for rule in present_source_rules}

    for row_index, row in enumerate(rows, start=2):
        row_values = list(row)
        mother_code = normalize_code(row_values[mother_code_column - 1])
        status = "OK"
        issues: list[str] = []

        if not mother_code:
            status = "Missing mother_code"
            missing_row_count += 1
        elif mother_code in duplicate_codes:
            status = "Duplicate mother_code"
            duplicate_row_count += 1

        output_row_values: list[object] = []
        td_source_values: dict[str, str] = {}

        for source_index in included_columns:
            header = headers[source_index - 1]
            if isinstance(header, str) and header in rule_by_other_header:
                date_header, other_header, source_header = rule_by_other_header[header]
                date_value = row_values[headers.index(date_header)]
                other_value = row_values[headers.index(other_header)]
                source_value = build_source_value(date_value, other_value)
                output_row_values.append(source_value)
                if source_header == "Td1_source":
                    td_source_values["TD1"] = source_value
                elif source_header == "Td2_source":
                    td_source_values["TD2"] = source_value
            else:
                output_row_values.append(row_values[source_index - 1])

        td1_date = normalize_date(row_values[headers.index("td_first_dose")]) if "td_first_dose" in headers else None
        td2_date = normalize_date(row_values[headers.index("td_second_dose")]) if "td_second_dose" in headers else None
        td1_received = has_received(td_source_values.get("TD1"))
        td2_received = has_received(td_source_values.get("TD2"))

        if not td1_received and td2_received:
            issues.append("Td1 not received but Td2 received")
        if td1_date is not None and td2_date is not None and td2_date < td1_date:
            issues.append("Td2 date earlier than Td1 date")

        if issues:
            unlogical_row_count += 1
            if status == "OK":
                status = "; ".join(issues)
            else:
                status = f"{status}; " + "; ".join(issues)

        report.append(output_row_values + [status])

        if status != "OK":
            affected_row_count += 1
            code_cell = report.cell(row=row_index, column=report_column_map[mother_code_column])
            code_cell.fill = RED_FILL
            code_cell.font = RED_FONT

    summary = {
        "rows": len(rows),
        "missing_rows": missing_row_count,
        "duplicate_rows": duplicate_row_count,
        "duplicate_codes": len(duplicate_codes),
        "unlogical_rows": unlogical_row_count,
        "affected_rows": affected_row_count,
    }
    return summary


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
    long_form = output.create_sheet(title=LONG_FORM_SHEET)

    included_columns, report_headers, report_column_map = build_output_columns(headers)
    present_source_rules = [rule for rule in SOURCE_FIELD_RULES if rule[0] in headers and rule[1] in headers]
    source_header_to_dose = {source_header: dose_key for dose_key, source_header in DOSE_SOURCE_HEADERS.items()}
    other_header_to_rule = {rule[1]: rule for rule in present_source_rules}

    report_headers = report_headers + ["completed_dose", "verification_status"]
    report.append(report_headers)

    if "IDP" not in headers:
        raise KeyError("Column 'IDP' was not found in sheet 'children'.")
    idp_column = headers.index("IDP") + 1
    static_columns = list(range(1, idp_column + 1))
    long_form_headers = [headers[index - 1] for index in static_columns] + [
        "completed_dose",
        "vaccine_dose",
        "reporting_month",
        "source",
        "age_at_dose",
    ]
    long_form.append(long_form_headers)

    missing_count = 0
    duplicate_row_count = 0
    dob_order_row_count = 0
    unlogical_row_count = 0
    affected_row_count = 0

    for row_index, row in enumerate(rows, start=2):
        row_values = list(row)
        code_value = normalize_code(row_values[code_column - 1])
        dob_value = normalize_date(row_values[dob_column - 1])
        status = "OK"
        dob_issues: list[str] = []

        if not code_value:
            status = "Missing children_code"
            missing_count += 1
        elif code_value in duplicate_codes:
            status = "Duplicate children_code"
            duplicate_row_count += 1

        if dob_value is not None and registered_column is not None:
            registered_value = normalize_date(row_values[registered_column - 1])
            if registered_value is not None and dob_value > registered_value:
                dob_issues.append("date_of_birth later than registered_date")

        if dob_value is not None:
            for column in comparison_columns:
                compared_value = normalize_date(row_values[column - 1])
                if compared_value is not None and dob_value > compared_value:
                    dob_issues.append(f"date_of_birth later than {headers[column - 1]}")

        output_row_values: list[object] = []
        source_values_by_dose: dict[str, str] = {}
        for index in included_columns:
            header = headers[index - 1]
            if isinstance(header, str) and header in other_header_to_rule:
                date_header, other_header, source_header = other_header_to_rule[header]
                date_value = row_values[headers.index(date_header)]
                other_value = row_values[headers.index(other_header)]
                source_value = build_source_value(date_value, other_value)
                output_row_values.append(source_value)
                dose_key = source_header_to_dose.get(source_header)
                if dose_key is not None:
                    source_values_by_dose[dose_key] = source_value
            else:
                output_row_values.append(row_values[index - 1])

        dose_date_values = {
            dose_key: normalize_date(row_values[headers.index(date_header)])
            for dose_key, date_header in DOSE_DATE_HEADERS.items()
            if date_header in headers
        }
        unlogical_issues = find_unlogical_sequence_issues(source_values_by_dose, dose_date_values)

        if dob_issues:
            dob_order_row_count += 1
        if unlogical_issues:
            unlogical_row_count += 1

        combined_issues = dob_issues + unlogical_issues
        if combined_issues:
            if status == "OK":
                status = "; ".join(combined_issues)
            else:
                status = f"{status}; " + "; ".join(combined_issues)

        age_months = normalize_age_months(row_values[headers.index("ageInMonths")]) if "ageInMonths" in headers else None
        completed_dose_value = calculate_completed_dose(age_months, source_values_by_dose, dose_date_values)

        report.append(output_row_values + [completed_dose_value, status])
        if status != "OK":
            affected_row_count += 1

        static_values = [row_values[index - 1] for index in static_columns]
        dob_for_age = normalize_date(row_values[dob_column - 1])
        for dose_key in LONG_FORM_DOSES:
            date_header = DOSE_DATE_HEADERS.get(dose_key)
            reporting_header = DOSE_REPORTING_HEADERS.get(dose_key)
            source_value = source_values_by_dose.get(dose_key, "Not received yet")

            dose_date_value = dose_date_values.get(dose_key)
            reporting_month_value = row_values[headers.index(reporting_header)] if reporting_header in headers else None
            age_at_dose = datediff_months(dob_for_age, dose_date_value)

            long_form.append(
                static_values
                + [
                    completed_dose_value,
                    dose_key,
                    reporting_month_value,
                    source_value,
                    age_at_dose,
                ]
            )

        if status != "OK":
            code_cell = report.cell(row=row_index, column=report_column_map[code_column])
            code_cell.fill = RED_FILL
            code_cell.font = RED_FONT

            dob_cell = report.cell(row=row_index, column=report_column_map[dob_column])
            if dob_issues:
                dob_cell.fill = RED_FILL
                dob_cell.font = RED_FONT

    duplicate_code_count = len(duplicate_codes)
    summary = {
        "rows": len(rows),
        "missing_rows": missing_count,
        "duplicate_rows": duplicate_row_count,
        "duplicate_codes": duplicate_code_count,
        "dob_order_rows": dob_order_row_count,
        "unlogical_rows": unlogical_row_count,
        "affected_rows": affected_row_count,
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
        pw_sheet = get_pw_sheet(source_workbook)
        report_workbook, summary = build_verification_report(source_sheet)
        pw_summary = build_pw_verification_sheet(report_workbook, pw_sheet)
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
    st.metric("Rows with unlogical dose records", summary["unlogical_rows"])

    st.info(f"Rows with issues: {summary['affected_rows']}")
    if summary["unlogical_rows"] > 0:
        st.warning("Unlogical dose records detected for OPV, Penta, or MMR sequences. Check highlighted children_code rows.")

    st.subheader("PW sheet checks")
    pw_col1, pw_col2, pw_col3, pw_col4 = st.columns(4)
    pw_col1.metric("PW rows checked", pw_summary["rows"])
    pw_col2.metric("Missing mother_code rows", pw_summary["missing_rows"])
    pw_col3.metric("Duplicate mother_code rows", pw_summary["duplicate_rows"])
    pw_col4.metric("Duplicate mother_code values", pw_summary["duplicate_codes"])
    st.metric("PW rows with unlogical Td records", pw_summary["unlogical_rows"])
    st.info(f"PW rows with issues: {pw_summary['affected_rows']}")

    report_bytes = workbook_to_bytes(report_workbook)
    st.download_button(
        label="Download new verification file",
        data=report_bytes,
        file_name="KDHW_verification.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.subheader("Rules applied")
    st.write("Missing children_code values are flagged red.")
    st.write("Duplicate children_code values are flagged red on every matching row.")
    st.write("date_of_birth is checked against registered_date and the listed vaccine date columns; later values are flagged red.")
    st.write("Unlogical sequences are checked for OPV, Penta, and MMR: earlier dose missing while later dose received, or later dose date earlier than primary dose date.")
    st.write("PW sheet drops Td third/fourth/fifth groups, prevent_td_newborn, and comments; checks duplicate mother_code; replaces td_first_dose_other and td_second_dose_other with Td1_source and Td2_source.")
    st.write("PW unlogical checks: Td1 not received while Td2 received, and Td2 date earlier than Td1 date. Rows are flagged and mother_code is highlighted red.")


if __name__ == "__main__":
    main()

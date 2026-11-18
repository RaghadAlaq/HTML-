"""
Simple weekly employee / resignation reconciliation.
Install: pip install openpyxl
Edit only the four paths below and run this file.
No fuzzy ID matching: unclear headers or duplicate IDs stop the comparison.
"""
from pathlib import Path
from collections import Counter
from datetime import datetime, date
import re
import unicodedata
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ==================== YOUR FOUR FILE PATHS ====================
EMPLOYEES_LAST_WEEK = r"C:\YourPath\employees_last_week.xlsx"
EMPLOYEES_THIS_WEEK = r"C:\YourPath\employees_this_week.xlsx"
RESIGNATIONS_LAST_WEEK = r"C:\YourPath\resignations_last_week.xlsx"
RESIGNATIONS_THIS_WEEK = r"C:\YourPath\resignations_this_week.xlsx"
# =============================================================

RESIGNATION_SHEET = "HC_HCS_Resignations_Report_V2"
OUTPUT_FILE = "Simple_Employee_Reconciliation.xlsx"

ARABIC_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
ID_HEADERS = {
    "employee id", "employee number", "employee no", "employee code",
    "emp id", "emp no", "emp number", "staff id", "staff number",
    "personnel number", "personnel id", "worker id",
    "رقم الموظف", "الرقم الوظيفي", "رقم وظيفي",
    "معرف الموظف", "كود الموظف", "رقم العامل",
}
# If the ID header is not one of these, add its exact text here.
# Never use National ID, Manager ID, or Transaction ID as Employee ID.


def clean_header(value):
    text = unicodedata.normalize("NFKC", str(value or "")).translate(ARABIC_DIGITS)
    text = re.sub(r"[\u064b-\u065f\u0670\u0640]", "", text)
    text = re.sub(r"\s+", " ", text.strip().lower().replace("_", " ").replace("-", " "))
    return text


def clean_id(value):
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, (datetime, date)):
        return ""
    text = unicodedata.normalize("NFKC", str(value)).translate(ARABIC_DIGITS).strip()
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".")[0]
    text = re.sub(r"\s+", "", text)
    return text.upper()


def read_report(path, kind):
    file = Path(path)
    if not file.is_file():
        raise ValueError(f"File not found: {file}")
    book = load_workbook(file, read_only=True, data_only=True)
    try:
        if kind == "employees":
            sheet = book.worksheets[0]
        else:
            matches = [s for s in book.worksheets
                       if clean_header(s.title) == clean_header(RESIGNATION_SHEET)]
            if len(matches) != 1:
                raise ValueError(
                    f"{file.name}: resignation sheet '{RESIGNATION_SHEET}' not found exactly. "
                    f"Available: {[s.title for s in book.worksheets]}"
                )
            sheet = matches[0]

        iterator = sheet.iter_rows(values_only=True)
        header_row = None
        id_index = None
        for row_number, row in enumerate(iterator, 1):
            if row_number > 40:
                break
            hits = [i for i, value in enumerate(row)
                    if clean_header(value) in ID_HEADERS]
            if len(hits) > 1:
                raise ValueError(f"{file.name}: multiple possible Employee ID columns at row {row_number}.")
            if hits:
                header_row = row_number
                id_index = hits[0]
                headers = [str(v).strip() if v is not None else f"Column_{i+1}"
                           for i, v in enumerate(row)]
                break
        if header_row is None:
            raise ValueError(
                f"{file.name} / {sheet.title}: Employee ID header not found in first 40 rows. "
                "Add its exact header text to ID_HEADERS; no guessing will be performed."
            )
        records = {}
        duplicates = set()
        blank_ids = 0
        for row in iterator:
            if not any(v is not None and str(v).strip() for v in row):
                continue
            values = list(row[:len(headers)])
            values += [None] * (len(headers) - len(values))
            employee_id = clean_id(values[id_index])
            if not employee_id:
                blank_ids += 1
                continue
            if employee_id in records:
                duplicates.add(employee_id)
            else:
                records[employee_id] = values
        if duplicates:
            raise ValueError(
                f"{file.name} / {sheet.title}: duplicate Employee IDs found "
                f"({len(duplicates)}): {', '.join(sorted(duplicates)[:15])}. "
                "Please resolve duplicates before trusting the comparison."
            )
        if blank_ids:
            raise ValueError(
                f"{file.name} / {sheet.title}: {blank_ids} nonempty rows have no Employee ID. "
                "Please review these rows before running the audit."
            )
        print(f"  {file.name} | sheet: {sheet.title} | unique employees: {len(records)}", flush=True)
        return {"headers": headers, "records": records, "id_index": id_index}
    finally:
        book.close()


def output_rows(source, ids):
    return [["Employee ID"] +
            [header for i, header in enumerate(source["headers"]) if i != source["id_index"]]] + [
            [employee_id] + [
                value for i, value in enumerate(source["records"][employee_id])
                if i != source["id_index"]
            ]
            for employee_id in sorted(ids, key=lambda x: (not x.isdigit(), int(x) if x.isdigit() else x))
        ]


def write_sheet(book, title, rows):
    ws = book.create_sheet(title)
    for row in rows:
        ws.append([v if not isinstance(v, (list, dict, set)) else str(v) for v in row])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="17324D")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(wrap_text=True)
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        longest = max((len(str(c.value or "")) for c in list(col)[:150]), default=10)
        ws.column_dimensions[letter].width = min(45, max(16, longest + 2))
    ws.sheet_view.showGridLines = False
    return ws


def main():
    print("Reading four reports...", flush=True)
    old_e = read_report(EMPLOYEES_LAST_WEEK, "employees")
    new_e = read_report(EMPLOYEES_THIS_WEEK, "employees")
    old_r = read_report(RESIGNATIONS_LAST_WEEK, "resignations")
    new_r = read_report(RESIGNATIONS_THIS_WEEK, "resignations")

    old_e_ids, new_e_ids = set(old_e["records"]), set(new_e["records"])
    old_r_ids, new_r_ids = set(old_r["records"]), set(new_r["records"])
    added_e, missing_e = new_e_ids - old_e_ids, old_e_ids - new_e_ids
    added_r, missing_r = new_r_ids - old_r_ids, old_r_ids - new_r_ids

    out = Workbook()
    out.remove(out.active)
    summary = [
        ["Metric", "Count"],
        ["Employees last week", len(old_e_ids)],
        ["Employees this week", len(new_e_ids)],
        ["NEW employee IDs", len(added_e)],
        ["MISSING employee IDs", len(missing_e)],
        ["Resignation IDs last week", len(old_r_ids)],
        ["Resignation IDs this week", len(new_r_ids)],
        ["NEW resignation IDs", len(added_r)],
        ["MISSING resignation IDs", len(missing_r)],
        ["Missing employees present in THIS week's resignation report",
         len(missing_e & new_r_ids)],
        ["Missing employees NOT in THIS week's resignation report",
         len(missing_e - new_r_ids)],
        ["Note", "New/missing means present in one snapshot only; not proof of hire/resignation."],
    ]
    write_sheet(out, "SUMMARY", summary)
    write_sheet(out, "NEW EMPLOYEES", output_rows(new_e, added_e))
    write_sheet(out, "MISSING EMPLOYEES", output_rows(old_e, missing_e))
    write_sheet(out, "NEW RESIGNATIONS", output_rows(new_r, added_r))
    write_sheet(out, "MISSING RESIGNATIONS", output_rows(old_r, missing_r))

    destination = Path(__file__).resolve().parent / OUTPUT_FILE
    print("Writing:", destination, flush=True)
    out.save(destination)
    print("\nCOMPLETE", flush=True)
    for label, ids in [
        ("New employee IDs", added_e), ("Missing employee IDs", missing_e),
        ("New resignation IDs", added_r), ("Missing resignation IDs", missing_r)
    ]:
        print(f"{label}: {len(ids)}", flush=True)
    print(f"Output: {destination}", flush=True)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nSTOPPED: {exc}", flush=True)
        raise SystemExit(1)

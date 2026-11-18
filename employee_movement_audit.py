"""Employee movement and resignation audit across four weekly Excel workbooks.

Install once: pip install pandas openpyxl
Run: python employee_movement_audit.py

Only edit the four paths in CONFIGURATION. Never treats a missing roster entry as
proof of resignation. All output is an evidence-based comparison of snapshots.
"""

from __future__ import annotations

import re
import sys
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    raise SystemExit("Missing packages. Install with: pip install pandas openpyxl")

# ===================== CONFIGURATION: EDIT ONLY THESE =====================
EMPLOYEES_LAST_WEEK = r"C:\YourPath\employees_last_week.xlsx"
EMPLOYEES_THIS_WEEK = r"C:\YourPath\employees_this_week.xlsx"
RESIGNATIONS_LAST_WEEK = r"C:\YourPath\resignations_last_week.xlsx"
RESIGNATIONS_THIS_WEEK = r"C:\YourPath\resignations_this_week.xlsx"

OUTPUT_FILE = "Employee_Movement_Audit.xlsx"  # Saved beside this script
# ===========================================================================

ARABIC_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
ARABIC_MARKS = re.compile(r"[\u064b-\u065f\u0670\u0640]")

ID_TERMS = (
    "employee id", "employee number", "employee no", "staff id", "staff no",
    "personnel number", "personnel no", "personnel id", "emp id", "emp no",
    "worker id", "worker number", "national id", "iqama", "badge number",
    "رقم الموظف", "الرقم الوظيفي", "رقم وظيفي", "رقم العامل", "رقم الهوية",
    "رقم الاقامة", "الرقم الشخصي", "رقم السجل", "كود الموظف", "معرف الموظف",
)
NAME_TERMS = (
    "employee name", "full name", "staff name", "worker name", "person name",
    "اسم الموظف", "اسم العامل", "اسم الموظفة", "الاسم الكامل", "اسم كامل",
)
DATE_TERMS = (
    "resignation date", "termination date", "last working day", "last day",
    "effective date", "separation date", "exit date", "leaving date",
    "تاريخ الاستقالة", "تاريخ انتهاء الخدمة", "اخر يوم عمل", "تاريخ ترك العمل",
    "تاريخ الانفصال", "تاريخ السريان", "تاريخ انهاء الخدمة",
)
STATUS_TERMS = (
    "employment status", "employee status", "staff status", "work status",
    "الحالة الوظيفية", "حالة الموظف", "حالة العامل", "حالة الخدمة",
)
RESIGN_TERMS = (
    "resign", "termination", "separation", "leaver", "exit", "turnover",
    "استقال", "المستقيل", "انهاء خدم", "انتهاء خدم", "ترك العمل",
)
ID_EXCLUSIONS = (
    "manager", "supervisor", "مباشر", "مدير", "رئيس", "request", "طلب",
    "transaction", "معاملة", "قرار", "serial", "sequence", "تسلسل",
    "تاريخ", "date", "phone", "mobile", "جوال", "هاتف",
)


def norm(value):
    if pd.isna(value):
        return ""
    s = unicodedata.normalize("NFKC", str(value)).translate(ARABIC_DIGITS)
    s = ARABIC_MARKS.sub("", s)
    s = s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    s = s.replace("ى", "ي").replace("ة", "ه")
    s = re.sub(r"[^\w]+", " ", s.lower(), flags=re.UNICODE)
    return " ".join(s.split())


def cell(value):
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    s = str(value).strip().translate(ARABIC_DIGITS)
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


def norm_id(value):
    return re.sub(r"\s+", "", cell(value)).upper()


def score_header(header, terms):
    h = norm(header)
    if not h:
        return 0
    scores = []
    for term in terms:
        t = norm(term)
        if h == t:
            scores.append(100)
        elif t in h and len(t) >= 5:
            scores.append(75)
        elif h in t and len(h) >= 6:
            scores.append(55)
    return max(scores, default=0)


def best_column(df, terms, kind="general"):
    candidates = []
    for col in df.columns:
        h = norm(col)
        if kind == "id" and any(x in h for x in (norm(z) for z in ID_EXCLUSIONS)):
            continue
        score = score_header(col, terms)
        if score:
            nonblank = df[col].map(cell).ne("").sum()
            if nonblank:
                if kind == "id":
                    vals = df[col].map(norm_id)
                    distinct_ratio = vals[vals.ne("")].nunique() / max(nonblank, 1)
                    score += min(10, 10 * distinct_ratio)
                candidates.append((score, col))
    return max(candidates, default=(0, None))[1]


def infer_id_column(df):
    """Conservative value-based fallback if no recognizable ID heading exists."""
    candidates = []
    for col in df.columns:
        h = norm(col)
        if any(x in h for x in (norm(z) for z in ID_EXCLUSIONS)):
            continue
        vals = df[col].map(norm_id)
        vals = vals[vals.ne("")]
        if len(vals) < 3:
            continue
        numeric_like = vals.map(lambda x: bool(re.fullmatch(r"[A-Z]*\d{4,14}", x))).mean()
        unique_ratio = vals.nunique() / len(vals)
        if numeric_like >= .90 and unique_ratio >= .90:
            candidates.append((numeric_like + unique_ratio, col))
    return max(candidates, default=(0, None))[1]


def make_unique_headers(raw):
    seen = Counter()
    names = []
    for i, v in enumerate(raw):
        label = str(v).strip() if not pd.isna(v) else ""
        if not label or label.lower().startswith("unnamed:"):
            label = f"Unlabeled_{i + 1}"
        seen[label] += 1
        names.append(label if seen[label] == 1 else f"{label}__{seen[label]}")
    return names


def read_workbook(path, role):
    book = pd.ExcelFile(path, engine="openpyxl")
    tables, metadata = [], []
    for sheet in book.sheet_names:
        # Header detection for reports with titles or introductory rows.
        preview = pd.read_excel(book, sheet_name=sheet, header=None, nrows=20, dtype=object)
        if preview.empty:
            metadata.append({"File": role, "Sheet": sheet, "Rows": 0, "Status": "Empty"})
            continue
        choices = []
        for i in range(min(len(preview), 15)):
            entries = [cell(x) for x in preview.iloc[i].tolist()]
            nonblank = sum(bool(v) for v in entries)
            keywords = sum(bool(score_header(v, ID_TERMS + NAME_TERMS + DATE_TERMS + STATUS_TERMS)) for v in entries)
            choices.append((keywords * 8 + min(nonblank, 15), -i, i))
        header_pos = max(choices)[2]
        raw = pd.read_excel(book, sheet_name=sheet, header=header_pos, dtype=object)
        raw.columns = make_unique_headers(raw.columns)
        raw = raw.dropna(how="all")
        raw = raw.loc[:, ~raw.columns.astype(str).str.match(r"^Unlabeled_\d+$") | raw.notna().any(axis=0)]
        raw = raw.reset_index(drop=True)
        id_col = best_column(raw, ID_TERMS, "id")
        id_source = "recognized header" if id_col else ""
        if not id_col:
            id_col = infer_id_column(raw)
            id_source = "inferred from values (review)" if id_col else "not found"
        name_col = best_column(raw, NAME_TERMS)
        date_col = best_column(raw, DATE_TERMS)
        status_col = best_column(raw, STATUS_TERMS)
        item = {
            "role": role, "sheet": sheet, "df": raw, "id_col": id_col,
            "name_col": name_col, "date_col": date_col, "status_col": status_col,
            "id_source": id_source, "header_row": header_pos + 1,
        }
        tables.append(item)
        metadata.append({
            "File": role, "Sheet": sheet, "Rows": len(raw), "Columns": len(raw.columns),
            "Header Row": header_pos + 1, "Employee ID Column": id_col or "NOT FOUND",
            "ID Detection": id_source, "Name Column": name_col or "NOT FOUND",
            "Event Date Column": date_col or "NOT FOUND",
            "Status Column": status_col or "NOT FOUND", "Status": "Read",
        })
    return tables, metadata


def sheet_similarity(a, b):
    ah, bh = {norm(c) for c in a["df"].columns}, {norm(c) for c in b["df"].columns}
    overlap = len(ah & bh) / max(len(ah | bh), 1)
    name_sim = SequenceMatcher(None, norm(a["sheet"]), norm(b["sheet"])).ratio()
    identity = int(bool(a["id_col"]) and bool(b["id_col"]))
    return .55 * overlap + .30 * name_sim + .15 * identity


def match_sheets(previous, current, kind):
    pairs, notes = [], []
    remaining = set(range(len(current)))
    for a in previous:
        possibilities = sorted(((sheet_similarity(a, current[j]), j) for j in remaining), reverse=True)
        if possibilities and possibilities[0][0] >= .32:
            confidence, j = possibilities[0]
            remaining.remove(j)
            pairs.append((a, current[j], confidence))
            notes.append({"Report": kind, "Last Week Sheet": a["sheet"], "This Week Sheet": current[j]["sheet"],
                          "Match Score": round(confidence, 3), "Decision": "Compared" if confidence >= .55 else "Compared — review sheet match"})
        else:
            notes.append({"Report": kind, "Last Week Sheet": a["sheet"], "This Week Sheet": "",
                          "Match Score": 0, "Decision": "Unmatched last-week sheet"})
    for j in sorted(remaining):
        notes.append({"Report": kind, "Last Week Sheet": "", "This Week Sheet": current[j]["sheet"],
                      "Match Score": 0, "Decision": "Unmatched current-week sheet"})
    return pairs, notes


def align_columns(a, b):
    """Match headers: exact normalized name first, then conservative fuzzy matches."""
    used_b, pairs = set(), []
    for ca in a.columns:
        matches = [cb for cb in b.columns if cb not in used_b and norm(ca) == norm(cb)]
        if matches:
            cb = matches[0]
            used_b.add(cb)
            pairs.append((ca, cb, 1.0))
    remaining_a = [c for c in a.columns if c not in {p[0] for p in pairs}]
    for ca in remaining_a:
        ranked = sorted(((SequenceMatcher(None, norm(ca), norm(cb)).ratio(), cb)
                         for cb in b.columns if cb not in used_b), reverse=True)
        if ranked and ranked[0][0] >= .84 and min(len(norm(ca)), len(norm(ranked[0][1]))) >= 5:
            score, cb = ranked[0]
            used_b.add(cb)
            pairs.append((ca, cb, round(score, 3)))
    return pairs


def row_dict(table, idx):
    row = table["df"].iloc[idx]
    result = {str(k): cell(v) for k, v in row.items()}
    result["_source_file"] = table["role"]
    result["_source_sheet"] = table["sheet"]
    result["_excel_row_estimate"] = idx + table["header_row"] + 1
    return result


def gather(tables):
    """All sheets; records without confident IDs are retained for audit, not counted."""
    result, no_id, duplicates = defaultdict(list), [], []
    for table in tables:
        df = table["df"]
        if not table["id_col"]:
            for idx in range(len(df)):
                no_id.append(row_dict(table, idx))
            continue
        for idx, value in enumerate(df[table["id_col"]]):
            eid = norm_id(value)
            record = row_dict(table, idx)
            if not eid:
                no_id.append(record)
            else:
                record["_employee_id"] = eid
                record["_employee_name"] = cell(df.iloc[idx][table["name_col"]]) if table["name_col"] else ""
                record["_event_date"] = cell(df.iloc[idx][table["date_col"]]) if table["date_col"] else ""
                record["_employment_status"] = cell(df.iloc[idx][table["status_col"]]) if table["status_col"] else ""
                result[eid].append((table, idx, record))
    for eid, occurrences in result.items():
        if len(occurrences) > 1:
            duplicates.append({"Employee ID": eid, "Occurrences": len(occurrences),
                               "Sheets": "; ".join(sorted({r["_source_sheet"] for _, _, r in occurrences})),
                               "Note": "May be legitimately repeated across sheets; review before interpreting"})
    return result, no_id, duplicates


def first_record(items):
    return items[0][2] if items else {}


def chosen_values(rec):
    return {k: v for k, v in rec.items() if not k.startswith("_")}


def make_record(eid, rec, tag, explanation=""):
    out = {"Employee ID": eid, "Employee Name": rec.get("_employee_name", ""),
           "Classification": tag, "Explanation": explanation,
           "Source File": rec.get("_source_file", ""), "Source Sheet": rec.get("_source_sheet", ""),
           "Excel Row (approx.)": rec.get("_excel_row_estimate", ""),
           "Event Date": rec.get("_event_date", ""), "Employee Status": rec.get("_employment_status", "")}
    out.update({f"DATA | {k}": v for k, v in chosen_values(rec).items()})
    return out


def fingerprint(rec):
    """Detect a new report row, not just an employee present in a cumulative report."""
    date = rec.get("_event_date", "")
    # If no date, include the full row so that modified records are highlighted
    # as changed/ambiguous, not claimed to be a new resignation.
    return norm(date) if date else ""


def compare_resignations(old, new):
    new_event, modified_event, disappeared_event = [], [], []
    for eid in sorted(set(old) | set(new)):
        older, newer = old.get(eid, []), new.get(eid, [])
        if not older:
            for _, _, rec in newer:
                new_event.append(make_record(eid, rec, "New resignation report entry",
                                            "ID did not appear in last week's resignation report; verify effective date"))
            continue
        if not newer:
            for _, _, rec in older:
                disappeared_event.append(make_record(eid, rec, "Resignation entry removed",
                                                     "Entry was present last week but absent this week"))
            continue
        old_dates = {fingerprint(r) for _, _, r in older if fingerprint(r)}
        for _, _, rec in newer:
            d = fingerprint(rec)
            if d and old_dates and d not in old_dates:
                modified_event.append(make_record(eid, rec, "Changed/new event date — REVIEW",
                                                  f"Last week event date(s): {', '.join(sorted(old_dates))}"))
        old_rec = first_record(older)
        new_rec = first_record(newer)
        old_values, new_values = chosen_values(old_rec), chosen_values(new_rec)
        if old_values != new_values and not (old_dates and any(fingerprint(r) not in old_dates for _, _, r in newer)):
            modified_event.append(make_record(eid, new_rec, "Existing resignation record changed",
                                              "Same employee appeared in both weekly resignation reports"))
    return new_event, modified_event, disappeared_event


def compare_fields(pairs, kind):
    differences, mappings = [], []
    for a, b, sheet_score in pairs:
        columns = align_columns(a["df"], b["df"])
        for ca, cb, confidence in columns:
            mappings.append({"Report": kind, "Last Week Sheet": a["sheet"], "This Week Sheet": b["sheet"],
                             "Last Week Column": ca, "This Week Column": cb, "Similarity": confidence,
                             "Sheet Match Score": round(sheet_score, 3)})
        if not (a["id_col"] and b["id_col"]):
            continue
        ai, bi = {}, {}
        for idx, v in enumerate(a["df"][a["id_col"]]):
            eid = norm_id(v)
            if eid and eid not in ai:
                ai[eid] = idx
        for idx, v in enumerate(b["df"][b["id_col"]]):
            eid = norm_id(v)
            if eid and eid not in bi:
                bi[eid] = idx
        for eid in sorted(set(ai) & set(bi)):
            for ca, cb, conf in columns:
                if ca == a["id_col"] or cb == b["id_col"]:
                    continue
                before = cell(a["df"].iloc[ai[eid]][ca])
                after = cell(b["df"].iloc[bi[eid]][cb])
                if norm(before) != norm(after):
                    differences.append({"Report": kind, "Employee ID": eid, "Last Week Sheet": a["sheet"],
                                        "This Week Sheet": b["sheet"], "Column Last Week": ca,
                                        "Column This Week": cb, "Previous Value": before, "Current Value": after,
                                        "Column Match Score": conf, "Note": "Value changed (or representation differs)"})
    return differences, mappings


def excel_safe(value):
    if isinstance(value, str) and value[:1] in ("=", "+", "-", "@"):
        return "'" + value  # Prevent formula injection from source cell text.
    return value


def write_tab(writer, sheet, data, empty_message="No records found"):
    df = pd.DataFrame(data)
    if df.empty:
        df = pd.DataFrame([{"Result": empty_message}])
    df = df.fillna("").map(excel_safe)
    df.to_excel(writer, sheet_name=sheet[:31], index=False)
    ws = writer.sheets[sheet[:31]]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    for c in ws[1]:
        c.fill = PatternFill("solid", fgColor="243447")
        c.font = Font(color="FFFFFF", bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 30
    for j, col in enumerate(df.columns, 1):
        sample = [len(str(v)) for v in df.iloc[:100, j - 1].tolist()]
        width = max(14, min(48, max([len(str(col))] + sample) + 2))
        ws.column_dimensions[get_column_letter(j)].width = width


def main():
    files = {
        "Employees last week": EMPLOYEES_LAST_WEEK,
        "Employees this week": EMPLOYEES_THIS_WEEK,
        "Resignations last week": RESIGNATIONS_LAST_WEEK,
        "Resignations this week": RESIGNATIONS_THIS_WEEK,
    }
    for role, path in files.items():
        if "YourPath" in path or not Path(path).is_file():
            raise SystemExit(f"Please set a valid file path for '{role}': {path}")

    all_tables, all_meta = {}, []
    for role, path in files.items():
        print("Reading:", role)
        tables, metadata = read_workbook(path, role)
        all_tables[role] = tables
        all_meta.extend(metadata)

    old_emp, old_emp_missing, old_emp_dups = gather(all_tables["Employees last week"])
    new_emp, new_emp_missing, new_emp_dups = gather(all_tables["Employees this week"])
    old_res, old_res_missing, old_res_dups = gather(all_tables["Resignations last week"])
    new_res, new_res_missing, new_res_dups = gather(all_tables["Resignations this week"])

    old_keys, new_keys = set(old_emp), set(new_emp)
    added_ids, missing_ids = sorted(new_keys - old_keys), sorted(old_keys - new_keys)
    retained_ids = sorted(old_keys & new_keys)
    newly_res, changed_res, removed_res = compare_resignations(old_res, new_res)
    new_res_ids = {x["Employee ID"] for x in newly_res}

    added = [make_record(eid, first_record(new_emp[eid]), "New in comprehensive roster",
                         "Could be new hire, transfer-in, reactivation, or report correction") for eid in added_ids]
    missing = [make_record(eid, first_record(old_emp[eid]), "Missing from current roster",
                           "Not proof of resignation; could be transfer, termination, filter or data issue") for eid in missing_ids]
    both = [make_record(eid, first_record(new_emp[eid]), "Appears in both weekly rosters") for eid in retained_ids]

    crosscheck = []
    for eid in sorted(old_keys | new_keys | set(old_res) | set(new_res)):
        last_emp, this_emp = eid in old_keys, eid in new_keys
        last_res, this_res = eid in old_res, eid in new_res
        if eid in new_res_ids and this_emp:
            finding = "New resignation entry but still in current roster"
            note = "May be normal for a future resignation date; verify status/effective date"
        elif eid in new_res_ids and not this_emp and last_emp:
            finding = "New resignation entry + removed from roster"
            note = "Strong cross-report consistency, but check effective date and employment status"
        elif last_emp and not this_emp and not this_res:
            finding = "Disappeared from roster; no current resignation entry"
            note = "Investigate movement, termination, transfer, filtering, or missing resignation record"
        elif not last_emp and this_emp and this_res:
            finding = "New roster ID also present in resignation report"
            note = "Check historical entries, rehire, future departure or incorrect ID"
        elif not last_emp and this_emp:
            finding = "New roster appearance without current resignation entry"
            note = "Potential new hire / transfer-in / reactivation; verify hire date"
        elif this_res and not this_emp and not last_emp:
            finding = "Resignation entry but employee absent from both weekly rosters"
            note = "Could be historical or employee excluded from comprehensive report"
        elif last_res and not this_res:
            finding = "Previously resigned entry absent from current resignation report"
            note = "Resignation report may be periodic rather than cumulative"
        else:
            continue
        reference = first_record(new_emp.get(eid) or old_emp.get(eid) or new_res.get(eid) or old_res.get(eid))
        crosscheck.append({"Employee ID": eid, "Employee Name": reference.get("_employee_name", ""),
                           "Last Week Roster": "Yes" if last_emp else "No", "This Week Roster": "Yes" if this_emp else "No",
                           "Last Week Resignation": "Yes" if last_res else "No", "This Week Resignation": "Yes" if this_res else "No",
                           "Finding": finding, "Interpretation": note,
                           "Current Resignation Date": first_record(new_res.get(eid)).get("_event_date", ""),
                           "Current Employment Status": first_record(new_emp.get(eid)).get("_employment_status", "")})

    employee_pairs, employee_pair_notes = match_sheets(all_tables["Employees last week"], all_tables["Employees this week"], "Employees")
    resign_pairs, resign_pair_notes = match_sheets(all_tables["Resignations last week"], all_tables["Resignations this week"], "Resignations")
    emp_diffs, emp_cols = compare_fields(employee_pairs, "Employees")
    res_diffs, res_cols = compare_fields(resign_pairs, "Resignations")

    unresolved = (old_emp_missing + new_emp_missing + old_res_missing + new_res_missing)
    duplicates = (old_emp_dups + new_emp_dups + old_res_dups + new_res_dups)
    summary = [
        {"Metric": "Employees in last week's comprehensive roster (unique IDs)", "Count": len(old_keys)},
        {"Metric": "Employees in this week's comprehensive roster (unique IDs)", "Count": len(new_keys)},
        {"Metric": "Net change in comprehensive roster (not confirmed net hiring)", "Count": len(new_keys) - len(old_keys)},
        {"Metric": "New appearances in comprehensive roster (NOT necessarily hires)", "Count": len(added_ids)},
        {"Metric": "Disappeared from comprehensive roster (NOT necessarily resignations)", "Count": len(missing_ids)},
        {"Metric": "Present in both comprehensive rosters", "Count": len(retained_ids)},
        {"Metric": "New employee IDs in resignation report (NOT necessarily effective this week)", "Count": len(new_res_ids)},
        {"Metric": "Changed or potentially new resignation events requiring review", "Count": len(changed_res)},
        {"Metric": "Employee IDs absent from current resignation report", "Count": len(set(old_res) - set(new_res))},
        {"Metric": "Roster column-level value changes", "Count": len(emp_diffs)},
        {"Metric": "Resignation report column-level value changes", "Count": len(res_diffs)},
        {"Metric": "Source rows without identifiable employee ID (not included in counts)", "Count": len(unresolved)},
        {"Metric": "Duplicate-ID groups within a weekly report (review)", "Count": len(duplicates)},
        {"Metric": "Confirmed hires this week", "Count": "NOT DETERMINABLE without reliable hire dates/status"},
        {"Metric": "Confirmed resignations effective this week", "Count": "NOT DETERMINABLE without reporting period and reliable effective dates"},
    ]
    warnings = [
        {"Issue": "Snapshot ≠ event", "Action": "A person newly present in the roster may be a transfer, rehire, correction or new hire."},
        {"Issue": "Missing ≠ resignation", "Action": "A person absent from the new roster may have transferred, been filtered out or otherwise left."},
        {"Issue": "Resignation entry ≠ effective resignation", "Action": "A resignation may be historical, future-dated, pending or cancelled. Confirm the effective date/status."},
        {"Issue": "Unmatched or fuzzy sheets", "Action": "Review Sheet Mapping; results from unmatched sheets are still used in ID-level changes."},
        {"Issue": "Weak/absent ID mapping", "Action": "Check Sheet Inventory and Missing IDs; those rows are excluded from employee counts."},
        {"Issue": "Repeated employee IDs", "Action": "Review Duplicate IDs, especially if multiple assignment records exist in one report."},
        {"Issue": "Multiple resignation events for same ID", "Action": "A person may have multiple records. Review event dates and raw records rather than equating IDs with events."},
        {"Issue": "Sheet and column mismatches", "Action": "Review Column Mapping; nonmatching columns cannot be reliably compared field by field."},
    ]
    for note in employee_pair_notes + resign_pair_notes:
        if note["Decision"] != "Compared":
            warnings.append({"Issue": "Sheet match: " + note["Decision"],
                             "Action": f"{note['Report']}: {note['Last Week Sheet']} -> {note['This Week Sheet']}"})
    out = Path(__file__).resolve().parent / OUTPUT_FILE
    print("Writing audit:", out)
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        tabs = [
            ("SUMMARY", summary), ("CROSS REPORT CHECK", crosscheck),
            ("NEW ROSTER ENTRIES", added), ("MISSING ROSTER ENTRIES", missing),
            ("NEW RESIGNATION ENTRIES", newly_res), ("CHANGED RESIGNATIONS", changed_res),
            ("REMOVED RESIGNATIONS", removed_res), ("EMPLOYEE FIELD CHANGES", emp_diffs),
            ("RESIGN FIELD CHANGES", res_diffs), ("SHEET MAPPING", employee_pair_notes + resign_pair_notes),
            ("COLUMN MAPPING", emp_cols + res_cols), ("SHEET INVENTORY", all_meta),
            ("DUPLICATE IDS", duplicates), ("MISSING IDS", unresolved),
            ("RETAINED EMPLOYEES", both), ("AUDIT NOTES", warnings),
        ]
        for name, records in tabs:
            write_tab(writer, name, records)
    print("\nCOMPLETE")
    print("New roster appearances:", len(added_ids))
    print("Roster disappearances:", len(missing_ids))
    print("New IDs in resignation report:", len(new_res_ids))
    print("Field-level roster changes:", len(emp_diffs))
    print("Open:", out)


if __name__ == "__main__":
    main()

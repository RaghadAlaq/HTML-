#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Resignation prediction preparation, v1.1.
Windows: py -m pip install openpyxl
         py prepare_predictions.py
Only reads the two prediction workbooks. Never opens or changes actual HR data.
Produces Predictions_Ready.xlsx and ready-to-paste Power BI instructions.
Paste the two workbook paths in HISTORY_FILE and NOVEMBER_FILE below.
Output is saved beside HISTORY_FILE; rerunning replaces the prepared output.
"""
from __future__ import annotations
import csv
import math
import re
import sys
import traceback
import unicodedata
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

# ===== الصقي المسارين بين علامتي التنصيص فقط، مع اسم الملف وامتداده =====
HISTORY_FILE = r""   # ملف التنبؤات التاريخية: مارس 2025 إلى أكتوبر 2026
NOVEMBER_FILE = r""  # ملف نوفمبر 2026: ملخص الموظفين على رأس العمل
# الناتج يحفظ بجانب الملف التاريخي. إعادة التشغيل تحدث الناتج نفسه.
START_MONTH = "2026-01"
END_MONTH = "2026-11"

SECTOR_TEXT = '''BOD General Secreteriat|الأمانة العامة لمجلس الإدارة
Corporate Communications and awareness|التواصل المؤسسي والتوعية
Data Management Office|مكتب إدارة البيانات
Engineering Affairs|الشؤون الهندسية
Finance and Administration|الشؤون المالية والإدارية
General Secretariat of Zakat, Tax and Customs Committees|الأمانة العامة للجان الزكوية والضريبية والجمركية
Governor Office|مكتب المحافظ
Human Capital|رأس المال البشري
Information Technology and Digitization|الأنظمة التقنية والرقمنة
Internal Audit|المراجعة الداخلية
Legal and Enforcement|الشؤون القانونية والالتزام
Operations|العمليات
Research and Zakat Advisory|البحوث والاستشارات الزكوية
Risk and Analytics|التحليل والمخاطر
Security Affairs|الشؤون الأمنية
Strategy and Trade Facilitation|الاستراتيجية و تيسير التجارة'''

# Translations supplied in the user's original script. No invented replacements
# for new organizational names. Unmapped names remain visible in English.
DEPT_TEXT = '''Business Support|دعم الأعمال
BoD Affairs & Governance|شؤون وحوكمة المجلس
Committees Affairs & Governance|شؤون وحوكمة اللجان
Awareness|التوعية
Brand & Content Management|المحتوى وهوية الهيئة
Events and Executive Protocols|الفعاليات والمراسم التنفيذية
External Communications|التواصل الخارجي
Internal Communications|التواصل الداخلي
Availability and Sharing Data|اتاحة ومشاركة البيانات
Data Governance and Compliance|حوكمة والتزام البيانات
Data Management|إدارة البيانات
Data Privacy and Protection|حماية وخصوصية البيانات
Engineering Affairs Planning|تخطيط الشؤون الهندسية
Engineering Projects|المشاريع الهندسية
Facilities|المرافق
Health, Safety and Environment|الصحة والسلامة والبيئة
Quality and Performance Management|الجودة والأداء
Work Environment|بيئة العمل
Accounting|المحاسبة
Documentation and Archives Center|مركز الوثائق والمحفوظات
Finance and Investment Management|المالية والاستثمار
Procurement|المشتريات
Support Services|الخدمات المساندة
1st Instance Committees Legal Studies|الدراسات القانونية للجان الفصل
Appeals Committees Legal Studies|الدراسات القانونية للجان الاستئناف
Appeals Technical and Accounting Studies Committees|الدراسات الفنية والمحاسبية للجان الاستئناف
First Instance Technical and Accounting Studies Committees|الدراسات الفنية والمحاسبية للجان الفصل
Committees' Support|دعم اللجان
Corporate Services|الخدمات المؤسسية
GSTSC Governance and Compliance|الحوكمة والالتزام بالأمانة العامة للجان
Strategy and Performance Management|الاستراتيجية وإدارة الأداء
Administrative Support|الدعم الإداري
Governor Office Affairs and Follow-up|إدارة شؤون مكتب المحافظ والمتابعة
Employee Experience|تجربة الموظف
HC Business Partner|شركاء أعمال الموارد البشرية
HC Center of Excellence|مركز تميز رأس المال البشري
Human Capital Services|خدمات رأس المال البشري
Talent Acquisition|استقطاب الكفاءات
Talent and Performance|المواهب والأداء الوظيفي
ZATCA Academy|أكاديمية الزكاة والضريبة والجمارك
Information Technology Planning and Management|تخطيط وإدارة تقنية المعلومات
IT Infrastructure and Information Security|البنية التحتية و أمن المعلومات
IT Operations|عمليات تقنية المعلومات
Smart Facilities|المرافق الذكية
Business Solutions Analysis and Design|تحليل وتصميم حلول الأعمال
Business Solutions Development|تطوير حلول الأعمال
Digital and Emerging Technology Audit|مراجعة تقنية المعلومات
Financial and Administrative Audit|المراجعة المالية والإدارية
Internal Audit Excellence|تميز المراجعة الداخلية
Operational Audit|مراجعة العمليات التشغيلية
Security and Customs Operations Audit|مراجعة المنافذ والفروع
Customs Investigations, Anti-Money Laundering and Counter Terrorism Financing|التحقيقات الجمركية و مكافحة غسل الأموال وتمويل الإرهاب
Customs Policy and Legislation|التشريعات والسياسات الجمركية
Governance and Compliance|الحوكمة والالتزام
Legal Affairs and Execution|الشؤون القانونية والتنفيذ
Legal Research and Development|البحوث القانونية والتطوير
Quality Assurance and Enforcement|ضمان الجودة والإمتثال
Review and Litigation|المراجعة والتقاضي
Zakat and Tax Policy and Legislation|التشريعات والسياسات الضريبية والزكوية
Anti-Evasion|مكافحة التهرب
Audit|التدقيق
Central Operations|العمليات المركزية
Customs Operations|العمليات الجمركية
Electronic Invoicing|الفاتورة الإلكترونية
Monitoring and QA|المراقبة وضمان الجودة
Operations Improvement|تحسين وتطوير العمليات
Business Empowerment|تمكين الأعمال
Settlement|التسويات
Zakat Advisory|الاستشارات الزكوية
Business Continuity And Crisis Management|إستمرارية الأعمال وإدارة الأزمات
Compliance Risk|مخاطر الالتزام
Enterprise Risk Management|إدارة المخاطر المؤسسية
Information and Cybersecurity|الامن المعلوماتي والسيبراني
Risk Analytics|تحليل المخاطر
Control and Security Intelligence|التحكم و المعلومات الأمنية
Ports Management|إدارة المنافذ
Security Support|المساندة الأمنية
Customer Experience and Services Improvement|تحسين الخدمات وتجربة العميل
Data and Analytics|البيانات والتحليل
Organizational Excellence|التميز المؤسسي
Organizational Performance|الأداء المؤسسي
Strategy and Enterprise Project Management Office|الاستراتيجية ومكتب إدارة المشاريع المؤسسية
Treaties and International and Strategic Partnerships|الإتفاقيات والشراكات الدولية والإستراتيجية'''


def clean(v):
    s = unicodedata.normalize('NFKC', '' if v is None else str(v))
    s = s.replace('\u200e', '').replace('\u200f', '').replace('\ufeff', '')
    return re.sub(r'\s+', ' ', s).strip()


def key(v):
    return clean(v).replace('’', "'").replace('‘', "'").casefold()


SECTORS = dict(line.split('|', 1) for line in SECTOR_TEXT.splitlines())
DEPTS = {key(a): b for a, b in (line.split('|', 1) for line in DEPT_TEXT.splitlines())}
SECTOR_LOOKUP = {key(s): s for s in SECTORS}
SECTOR_LOOKUP[key('BOD General Secretariat')] = 'BOD General Secreteriat'
SECTOR_LOOKUP[key('Corporate Communications and Awareness')] = 'Corporate Communications and awareness'
EXPECTED_RUNS = {date(2025, m, 1) for m in (3, 6, 9, 11)} | {date(2026, m, 1) for m in (4, 10)}


def total_label(v):
    return key(v) in {'grand total', 'total', 'الإجمالي', 'الاجمالي', 'الإجمالي العام', 'الاجمالي العام'}


def month(v):
    if isinstance(v, (date, datetime)):
        return date(v.year, v.month, 1)
    s = clean(v)
    for fmt in ('%Y-%m', '%Y-%m-%d', '%b-%y', '%b-%Y', '%B-%Y', '%m/%d/%Y', '%d/%m/%Y'):
        try:
            d = datetime.strptime(s, fmt)
            return date(d.year, d.month, 1)
        except ValueError:
            pass
    return None


def months(start, end):
    d = start
    while d <= end:
        yield d
        d = date(d.year + (d.month == 12), d.month % 12 + 1, 1)


def count(v, where, optional=False):
    if v is None or clean(v) == '':
        if optional:
            return None
        raise ValueError(f'{where}: blank numeric cell. Fill it with 0 ONLY if it really means zero. If it is a formula, open and save the source in Excel first.')
    if isinstance(v, bool):
        raise ValueError(f'{where}: boolean instead of a count')
    s = clean(v).replace(',', '').replace('٬', '')
    try:
        n = float(s)
    except (TypeError, ValueError):
        raise ValueError(f'{where}: invalid count {v!r}') from None
    if not math.isfinite(n) or n < 0 or not n.is_integer():
        raise ValueError(f'{where}: expected a non-negative whole count, got {v!r}')
    return int(n)


def canonical_sector(v):
    return SECTOR_LOOKUP.get(key(v), clean(v))


def canonical_dept(sector, dept):
    d = clean(dept)
    # Only strip sector prefix for the two known repeating office/support names.
    for prefix in (sector, sector.replace('Secreteriat', 'Secretariat')):
        if key(d).startswith(key(prefix) + ' '):
            suffix = d[len(prefix):].strip()
            if key(suffix).replace(' ', '') in {'level1office', 'businesssupport'}:
                d = suffix
                break
    if key(d).replace(' ', '') == 'level1office':
        return 'Level 1 office'
    return d


def load_overrides(path):
    """Optional CSV adjacent to this script: Sector_EN,Department_EN,Level 1,Level 2.
    Copy a row from Mapping_Review, fill Arabic names, save as UTF-8 CSV.
    Only nonempty translations are applied. Never changes source names.
    """
    result = {}
    if path.exists():
        with path.open(encoding='utf-8-sig', newline='') as f:
            reader = csv.DictReader(f)
            needed = {'Sector_EN', 'Department_EN', 'Level 1', 'Level 2'}
            if not needed.issubset(reader.fieldnames or []):
                raise ValueError('mapping_overrides.csv must have: Sector_EN,Department_EN,Level 1,Level 2')
            for row in reader:
                pair = (key(canonical_sector(row['Sector_EN'])), key(canonical_dept(canonical_sector(row['Sector_EN']), row['Department_EN'])))
                if row['Level 1'].strip() and row['Level 2'].strip():
                    if pair in result:
                        raise ValueError(f'Duplicate mapping override: {pair}')
                    result[pair] = (clean(row['Level 1']), clean(row['Level 2']))
    return result


def translate(sector, dept, overrides):
    pair = (key(sector), key(dept))
    if pair in overrides:
        return *overrides[pair], 'Override'
    ar1 = SECTORS.get(sector, sector)
    if dept == 'Level 1 office':
        # Same convention as the user's supplied script.
        ar2 = ar1
    else:
        ar2 = DEPTS.get(key(dept), dept)
    status = 'Mapped' if sector in SECTORS and (dept == 'Level 1 office' or key(dept) in DEPTS) else 'Review'
    return ar1, ar2, status


def record(run, sector, dept, emp, pred, file, sheet, row):
    sector = canonical_sector(sector)
    dept = canonical_dept(sector, dept)
    if pred > emp:
        raise ValueError(f'{file.name} / {sheet} row {row}: prediction {pred} exceeds employees {emp}')
    return dict(run=run, sector=sector, dept=dept, employees=emp, predicted=pred,
                file=file.name, sheet=sheet, row=row)


def header_row(ws, historical):
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 15), max_col=min(ws.max_column, 16)):
        vals = [clean(c.value) for c in row]
        joined = ' '.join(vals)
        if 'عدد الموظفين' in joined and 'عدد الاستقالات' in joined:
            # History has distinct sector and department columns. November has one.
            if historical and len(vals) >= 4 and 'الإدارة' in vals[1] and 'عدد الموظفين' in vals[2]:
                return row[0].row
            if not historical and len(vals) >= 3 and 'عدد الموظفين' in vals[1] and 'عدد الاستقالات' in vals[2]:
                return row[0].row
    return None


def open_source(path, historical):
    from openpyxl import load_workbook
    if path.suffix.lower() not in ('.xlsx', '.xlsm'):
        raise ValueError(f'{path.name}: save as .xlsx or .xlsm first')
    wb = load_workbook(path, data_only=True)
    candidates = [(ws, header_row(ws, historical)) for ws in wb.worksheets]
    candidates = [(ws, r) for ws, r in candidates if r is not None]
    if len(candidates) != 1:
        wb.close()
        raise ValueError(f'{path.name}: expected one matching {"historical" if historical else "November summary"} sheet; found {len(candidates)}. Check that the correct file was selected.')
    ws, r = candidates[0]
    return wb, ws, r


def check_total(checks, label, got_emp, got_pred, source_emp, source_pred):
    for metric, got, expected in [('Employees', got_emp, source_emp), ('Predicted', got_pred, source_pred)]:
        status = 'No source total' if expected is None else ('OK' if got == expected else 'MISMATCH')
        checks.append([label, metric, got, expected, status])
        if status == 'MISMATCH':
            raise ValueError(f'{label}: {metric} sum={got}, source total={expected}. No output created; check sector/department rows in source.')


def read_history(path, checks):
    wb, ws, hr = open_source(path, True)
    try:
        pairs = []
        for c in range(3, ws.max_column + 1):
            if 'عدد الموظفين' not in clean(ws.cell(hr, c).value):
                continue
            if 'عدد الاستقالات' not in clean(ws.cell(hr, c+1).value):
                raise ValueError(f'{path.name}: expected prediction column after employee column {c}')
            run = None
            for r in range(hr - 1, 0, -1):
                run = month(ws.cell(r, c).value) or month(ws.cell(r, c+1).value)
                if run:
                    break
            if not run:
                raise ValueError(f'{path.name}: no readable run month above column {c}')
            pairs.append((run, c))
        if len({d for d, _ in pairs}) != len(pairs):
            raise ValueError('Duplicate run month in historical headers')
        if {d for d, _ in pairs} != EXPECTED_RUNS:
            raise ValueError(f'Historical headers must contain Mar/Jun/Sep/Nov 2025 and Apr/Oct 2026; found {[str(d) for d,c in pairs]}')
        records, totals = [], {}
        for r in range(hr + 1, ws.max_row + 1):
            sector, dept = clean(ws.cell(r, 1).value), clean(ws.cell(r, 2).value)
            if total_label(sector) or total_label(dept):
                for run, c in pairs:
                    if run in totals:
                        raise ValueError('More than one Grand Total row in history')
                    totals[run] = (count(ws.cell(r, c).value, f'{ws.title}!{r}/{c}', True), count(ws.cell(r, c+1).value, f'{ws.title}!{r}/{c+1}', True))
                continue
            has_numbers = any(ws.cell(r, c).value is not None or ws.cell(r, c+1).value is not None for _, c in pairs)
            if not sector and not dept and not has_numbers:
                continue
            if not sector or not dept:
                raise ValueError(f'{path.name} row {r}: sector and department must both be present')
            for run, c in pairs:
                emp = count(ws.cell(r, c).value, f'{path.name} row {r}, employees {run:%Y-%m}')
                pred = count(ws.cell(r, c+1).value, f'{path.name} row {r}, prediction {run:%Y-%m}')
                records.append(record(run, sector, dept, emp, pred, path, ws.title, r))
        for run, _ in pairs:
            rs = [x for x in records if x['run'] == run]
            check_total(checks, f'History {run:%Y-%m}', sum(x['employees'] for x in rs), sum(x['predicted'] for x in rs), *totals.get(run, (None, None)))
        return records
    finally:
        wb.close()


def read_november(path, checks):
    wb, ws, hr = open_source(path, False)
    try:
        run = date(2026, 11, 1)
        records, sector_totals = [], {}
        sector, grand = None, None
        for r in range(hr + 1, ws.max_row + 1):
            label = clean(ws.cell(r, 1).value)
            b, c = ws.cell(r, 2).value, ws.cell(r, 3).value
            if not label and b is None and c is None:
                continue
            if not label:
                raise ValueError(f'{path.name} row {r}: values without a sector/department name')
            if total_label(label):
                if grand is not None:
                    raise ValueError('More than one Grand Total in November')
                grand = (count(b, f'{path.name} total B', True), count(c, f'{path.name} total C', True))
                continue
            emp, pred = count(b, f'{path.name} B{r}'), count(c, f'{path.name} C{r}')
            normalized = canonical_sector(label)
            if normalized in SECTORS:
                sector = normalized
                if sector in sector_totals:
                    raise ValueError(f'Duplicate sector header in November: {sector}')
                sector_totals[sector] = (emp, pred)
            else:
                if sector is None:
                    raise ValueError(f'{path.name} row {r}: department before a recognized sector: {label}')
                # Do not use bold alone to classify rows: some department names are bold too.
                records.append(record(run, sector, label, emp, pred, path, ws.title, r))
        if not records:
            raise ValueError('No November department records found')
        for sector, totals in sector_totals.items():
            rs = [x for x in records if x['sector'] == sector]
            check_total(checks, f'November / {sector}', sum(x['employees'] for x in rs), sum(x['predicted'] for x in rs), *totals)
        check_total(checks, 'November grand total', sum(x['employees'] for x in records), sum(x['predicted'] for x in records), *(grand or (None, None)))
        return records
    finally:
        wb.close()


def validate_unique(records):
    seen = set()
    for x in records:
        k = (x['run'], key(x['sector']), key(x['dept']))
        if k in seen:
            raise ValueError(f'Duplicate run/sector/department: {k}. Not summed automatically.')
        seen.add(k)


def build_monthly(records, start, end, overrides):
    by_run = defaultdict(list)
    for x in records:
        by_run[x['run']].append(x)
    result, review = [], {}
    for m in months(start, end):
        available = [d for d in by_run if d <= m]
        if not available:
            raise ValueError(f'No previous run available for {m:%Y-%m}; choose a later start month.')
        run = max(available)
        for x in by_run[run]:
            ar1, ar2, status = translate(x['sector'], x['dept'], overrides)
            result.append([m, ar1, ar2, x['employees'], x['predicted'], run, m != run, status,
                           x['sector'], x['dept']])
            if status != 'Mapped' and status != 'Override':
                review[(x['sector'], x['dept'])] = [x['sector'], x['dept'],
                    ar1 if x['sector'] in SECTORS else '', '', 'Missing translation: English retained in Predictions']
    return result, list(review.values())


PRED_HEADERS = ['Month_Start', 'Level 1', 'Level 2', 'Employees', 'Predicted', 'Source_Run',
                'Is_Carried_Forward', 'Mapping_Status', 'Sector_EN', 'Department_EN']


def write_workbook(path, rows, records, checks, review):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    wb.remove(wb.active)
    def sheet(name, headers, values, table_name=None):
        ws = wb.create_sheet(name)
        ws.append(headers)
        for row in values:
            ws.append(row)
        for row in ws:
            for cell in row:
                if isinstance(cell.value, str):
                    cell.data_type = 's'  # never interpret source text as Excel formulas
                cell.font = Font(name='Calibri', size=11)
                if isinstance(cell.value, (date, datetime)):
                    cell.number_format = 'yyyy-mm-dd'
        for cell in ws[1]:
            cell.font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor='152446')
            cell.alignment = Alignment(wrap_text=True)
        ws.freeze_panes = 'A2'
        ws.sheet_view.showGridLines = False
        for i, h in enumerate(headers, 1):
            ws.column_dimensions[get_column_letter(i)].width = 48 if h in ('Level 1','Level 2','Sector_EN','Department_EN','Check','Note') else 23
        if table_name and values:
            t = Table(displayName=table_name, ref=f'A1:{get_column_letter(len(headers))}{ws.max_row}')
            t.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
            ws.add_table(t)
        return ws
    sheet('Predictions', PRED_HEADERS, rows, 'tblPredictions')
    sheet('Checks', ['Check', 'Metric', 'Calculated', 'Source_Total', 'Status'], checks, 'tblChecks')
    sheet('Mapping_Review', ['Sector_EN','Department_EN','Level 1','Level 2','Note'], review)
    # Original snapshots are audit data, never append them to monthly Predictions.
    sheet('Run_Audit', ['Source_Run','Sector_EN','Department_EN','Employees','Predicted','File','Sheet','Source_Row'],
          [[x[k] for k in ('run','sector','dept','employees','predicted','file','sheet','row')] for x in records], 'tblRunAudit')
    sheet('Read_Me', ['Item','Note'], [
        ['Import', 'Import tblPredictions only. Do not import audit/checks into the chart.'],
        ['Meaning', 'Latest available model snapshot carried forward; NOT independent monthly predictions.'],
        ['Dates', 'All dates use day 1 as a month key, not the exact execution day.'],
        ['Scope', f'{rows[0][0]:%Y-%m} to {rows[-1][0]:%Y-%m}. November file assigned to November 2026.'],
        ['Totals', 'Never SUM repeated snapshots across months. Use the supplied Predicted measure.'],
        ['Unknown names', 'English retained. Review Mapping_Review. No departments dropped.'],
        ['Office mapping', 'Level 1 office uses sector Arabic name, as in the supplied original script.'],
        ['Actual data', 'This Python script does not read or modify the actual resignations workbook.'],
        ['Validation', 'Source totals checked where supplied. Blanks/errors/duplicate rows fail rather than becoming zero.'],
        ['Power BI', 'Use the generated setup text. Actual follows the visible existing filters, including ALL Event Reasons.'],
    ])
    tmp = path.with_name(path.stem + '.tmp.xlsx')
    try:
        wb.save(tmp)
        tmp.replace(path)
    finally:
        wb.close()
        if tmp.exists():
            tmp.unlink()


M_CODE = r'''let
    // Read existing query only. Do not paste into the existing query!
    A0 = #"الاستقالات",
    P0 = Predictions,
    P = Table.TransformColumnTypes(P0, {{"Month_Start", type date}, {"Source_Run", type date},
        {"Employees", Int64.Type}, {"Predicted", Int64.Type}, {"Level 1", type text}, {"Level 2", type text}}),
    StartMonth = List.Min(P[Month_Start]),
    EndMonth = List.Max(P[Month_Start]),
    // Same visible filters as the existing dashboard. No new Event Reason filter.
    A1 = Table.SelectRows(A0, each [Employee Class] = "موظفين"
        and [Nationality] = "المملكة العربية السعودية"
        and [Employment Details Termination Date] <> null
        and Date.From([Employment Details Termination Date]) >= StartMonth
        and Date.From([Employment Details Termination Date]) < Date.AddMonths(EndMonth, 1)),
    A2 = Table.AddColumn(A1, "Month_Start", each Date.StartOfMonth(Date.From([Employment Details Termination Date])), type date),
    A3 = Table.TransformColumns(A2, {{"Level 1", each if _ = null then "غير محدد" else Text.Trim(Text.Clean(_)), type text},
        {"Level 2", each if _ = null then "غير محدد" else Text.Trim(Text.Clean(_)), type text}}),
    A4 = Table.Group(A3, {"Month_Start", "Level 1", "Level 2"},
        {{"Value", each List.NonNullCount([Person ID]), Int64.Type}}),
    A5 = Table.AddColumn(A4, "Series", each "Actual", type text),
    P1 = Table.RenameColumns(P, {{"Predicted", "Value"}}),
    P2 = Table.AddColumn(P1, "Series", each "Predicted", type text),
    Combined = Table.Combine({A5, P2}),
    WithYear = Table.AddColumn(Combined, "Year", each Date.Year([Month_Start]), Int64.Type),
    WithMonth = Table.AddColumn(WithYear, "Month_Num", each Date.Month([Month_Start]), Int64.Type),
    WithLabel = Table.AddColumn(WithMonth, "Month_Label", each Date.ToText([Month_Start], "MMM yyyy", "en-US"), type text),
    WithMonthName = Table.AddColumn(WithLabel, "Month", each Date.ToText([Month_Start], "MMMM", "en-US"), type text)
in
    WithMonthName
'''


DAX = '''Actual =
CALCULATE(
    SUM(Comparison[Value]),
    Comparison[Series] = "Actual"
)

Predicted =
IF(
    HASONEVALUE(Comparison[Month_Start]),
    CALCULATE(SUM(Comparison[Value]), Comparison[Series] = "Predicted"),
    BLANK()
)

Employees Studied =
IF(
    HASONEVALUE(Comparison[Month_Start]),
    CALCULATE(SUM(Comparison[Employees]), Comparison[Series] = "Predicted"),
    BLANK()
)

Prediction Source =
IF(
    HASONEVALUE(Comparison[Month_Start]),
    VAR RunMonth = CALCULATE(MAX(Comparison[Source_Run]), Comparison[Series] = "Predicted")
    RETURN IF(NOT ISBLANK(RunMonth), FORMAT(RunMonth, "MMM yyyy"))
)
'''


def setup_text(output):
    return '''دليل سريع — بيانات التنبؤ ومقارنتها بالفعلي

الملف الناتج: ''' + str(output) + '''

السكربت لا يغير أي ملف مصدر. ملف نوفمبر يمثل تشغيل نوفمبر 2026.
الافتراضي: يناير إلى نوفمبر 2026. يناير–مارس من نوفمبر 2025، أبريل–سبتمبر من أبريل 2026، أكتوبر من أكتوبر، نوفمبر من ملف نوفمبر.
التكرار يمثل آخر تحديث متاح، وليس تنبؤًا مستقلًا لكل شهر. لا تقسمي الأعداد على 12 ولا تجمعيها عبر الأشهر.

1) افتحي Checks في Excel: الإجماليات المتاحة يجب أن تكون OK.
   Mapping_Review يظهر الأسماء التي لم توجد في قاموسك السابق؛ لم تُحذف أرقامها وتظهر بالإنجليزية.
   إذا عندك ترجمتها المعتمدة، احفظي CSV باسم mapping_overrides.csv بجانب السكربت، بالأعمدة:
   Sector_EN,Department_EN,Level 1,Level 2
   ثم شغلي السكربت ثانية. لا حاجة لتعديل الكود.

2) في نسخة عمل من تقرير Power BI، قبل الاستيراد:
   File > Options and settings > Options > Current File > Data Load
   ألغِي Autodetect new relationships after data is loaded.
   لا تغيّري Auto date/time ولا تحذفي علاقات قائمة.

3) Get data > Excel > اختاري الملف الناتج > حددي tblPredictions فقط > Transform Data.
   غيّري اسم الاستعلام المستورد إلى Predictions بالضبط.
   لا تستوردي Run_Audit أو Checks أو Mapping_Review للشارتات.

4) في Power Query: Home > New Source > Blank Query.
   سمّيه Comparison. افتحي Advanced Editor والصقي كود M الموجود أدناه.
   الصقيه في الاستعلام الجديد فقط، وليس الاستقالات.
   الكود يقرأ الاستعلام الحالي «الاستقالات» ويجمعه في جدول عرض جديد مع التنبؤ.
   الفعلي يستخدم Count of Person ID غير الفارغ، فئة موظفين، الجنسية المملكة العربية السعودية، وكل Event Reasons، مثل الصور المرسلة.
   لا يفرض «استقالة فقط» ولا يغير تعريف الفعلي. لذلك المعروض هو نفس نطاق أسباب المغادرة الحالي، وليس بالضرورة الاستقالة وحدها.
   فلاتر/slicers أخرى غير ظاهرة في الصور لن تُنسخ تلقائيًا.

5) Close & Apply. تأكدي في Model view أن Comparison وPredictions ليس لهما أي علاقات جديدة.
   لا نحتاج Calendar أو علاقات جديدة. لا تُضاف سجلات التنبؤ إلى الاستقالات الأصلية.
   إذا ظهر تنبيه خصوصية لمصادر البيانات، استخدمي سياسة المؤسسة؛ لا تعطلي الحماية.

6) أنشئي كل Measure أدناه منفصلًا داخل Comparison (New measure).
   لا تستخدمي SUM(Value) مباشرة للتنبؤ، ولا تضعي Series في slicer.

7) صفحة جديدة للمقارنة:
   - كل المحاور والسلايسرات لهذه الصفحة من Comparison.
   - Month_Label: Sort by column > Month_Start.
   - Month: Sort by column > Month_Num.
   - شارت Line and clustered column: X = Month_Label؛ الأعمدة Actual؛ الخط Predicted.
   - عنوان الخط: «التنبؤ حسب آخر تحديث متاح». Tooltip = Prediction Source.
   - شارت القطاعات: Clustered bar؛ المحور Level 1 ثم Level 2؛ القيم Actual وPredicted.
   - لشارت القطاعات اختاري شهرًا واحدًا بسلايسر Month_Label (Single select).
   - بدون اختيار شهر واحد، Predicted يرجع فراغًا عمدًا، كي لا يجمع نسخ التنبؤ الشهرية.
   - لمراجعة أشهر متعددة معًا استخدمي Matrix: الصفوف Level 1 ثم Level 2، الأعمدة Month_Label، القيم Actual وPredicted.
   - اختاري قطاعًا وشهرًا وقارني Actual مع الشارت القديم بنفس الفترة والفلاتر قبل اعتماد المقارنة.
   - فلتر الجنسية/الفئة الأصلي لا يؤثر في الجدول الجديد؛ تطبيقهما موجود في كود المقارنة نفسه. لا توجد تفاصيل جنس/موقع/أفراد في التنبؤ لتقسيمه بها.
   - لا تُحسب دقة شهرية من هذه اللقطات ولا إجمالي سنوي للتنبؤ.
   - الشهر غير المكتمل يعرض الفعلي المتوفر فقط. غياب سجلات فعلية يبقى فراغًا؛ لم نفترض صفرًا.

المصادر الفنية:
https://learn.microsoft.com/en-us/power-query/append-queries
https://learn.microsoft.com/en-us/power-bi/transform-model/desktop-create-and-manage-relationships

================ كود M: استعلام جديد باسم Comparison ================
''' + M_CODE + '''
================ DAX: كل مقياس على حدة ================
''' + DAX


def build(history, november, output, start=date(2026,1,1), end=date(2026,11,1), override_path=None):
    if start > end:
        raise ValueError('Start month must not be after end month')
    if history.resolve() == november.resolve():
        raise ValueError('Select two different source workbooks')
    if output.resolve() in {history.resolve(), november.resolve()}:
        raise ValueError('Output must not replace either source workbook')
    checks = []
    records = read_history(history, checks) + read_november(november, checks)
    validate_unique(records)
    overrides = load_overrides(override_path or Path(__file__).with_name('mapping_overrides.csv'))
    rows, review = build_monthly(records, start, end, overrides)
    output.parent.mkdir(parents=True, exist_ok=True)
    write_workbook(output, rows, records, checks, review)
    setup = output.with_name(output.stem + '_PowerBI.txt')
    setup.write_text(setup_text(output), encoding='utf-8-sig')
    return dict(output=output, setup=setup, rows=len(rows), review=len(review),
                november=sum(x['predicted'] for x in records if x['run'] == date(2026,11,1)))


def main():
    try:
        import openpyxl  # noqa: F401
        def source_path(value, label):
            value = value.strip().strip('"').strip("'")
            if not value:
                raise ValueError(f'Paste the full workbook path into {label} at the top of this script.')
            path = Path(value).expanduser()
            if not path.is_file():
                raise FileNotFoundError(f'{label}: file not found: {path}')
            return path
        history = source_path(HISTORY_FILE, 'HISTORY_FILE')
        november = source_path(NOVEMBER_FILE, 'NOVEMBER_FILE')
        output = history.parent / 'Predictions_Ready.xlsx'
        start, end = month(START_MONTH), month(END_MONTH)
        if not start or not end:
            raise ValueError('Use YYYY-MM for START_MONTH and END_MONTH')
        result = build(history, november, output, start, end)
        msg = (f"Saved: {result['output']}\nPower BI instructions: {result['setup']}\n"
               f"Monthly rows: {result['rows']}\nNovember predicted total: {result['november']}\n"
               f"Names to review: {result['review']}\n\nImport tblPredictions only. See the PowerBI text file for comparison setup.")
        print(msg)
    except Exception as exc:
        detail = traceback.format_exc()
        log = Path(__file__).with_name('prediction_error.txt')
        try:
            log.write_text(detail, encoding='utf-8')
        except OSError:
            log = None
        if isinstance(exc, ModuleNotFoundError) and exc.name == 'openpyxl':
            msg = 'Install once: py -m pip install openpyxl\nThen run this script again.'
        elif isinstance(exc, PermissionError):
            msg = 'Close the output Excel/text file and check folder permissions, then run again.\n' + str(exc)
        else:
            msg = str(exc)
        if log:
            msg += f'\nDetails: {log}'
        print(msg, file=sys.stderr)
        raise SystemExit(1)


if __name__ == '__main__':
    main()

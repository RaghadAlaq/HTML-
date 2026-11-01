import re
import sys
from datetime import timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ============================================================
# إعدادات التشغيل
# ============================================================
RAW_FILE = r"Ports Leaves - Dummy.xlsx"
OUTPUT_FILE = r"تقرير تدقيق تداخل الإجازات - Dummy.xlsx"
PREPARED_BY = "إدارة أداء وكالة رأس المال البشري"
OVERLAP_THRESHOLD = 2

# أسماء الحقول المطلوبة من ملف المصدر
REQUIRED = {
    "first": "(Employee) First Name",
    "last": "(Employee) Last Name",
    "id": "(Employee) Person ID",
    "l2": "(Employee) Level 2 (Label)",
    "l3": "(Employee) Level 3 (Label)",
    "loc": "(Employee) Location (Name)",
    "mgr": "Manager",
    "start": "Actual Start Date",
    "end": "Actual End Date",
    "approval": "Approval Status",
    "type": "Time Type (Label)",
    "qty": "Quantity (in Days)",
    "class": "(Employee) Employee Class (Label)",
}

LEVEL2_AR = {"Ports Management": "إدارة المنافذ"}
LEVEL3_AR = {
    "Land ports Management": "إدارة المنافذ البرية",
    "Airports Management": "إدارة المنافذ الجوية",
    "Seaports Management": "إدارة المنافذ البحرية",
}
LOCATION_AR = {
    "Al-Durra Land Port": "منفذ الدرة",
    "Najran Airport": "مطار نجران",
    "Jeddah Islamic Seaport Seaport": "ميناء جدة الإسلامي",
}
LEAVE_TYPE_AR = {
    "ZATCA Annual Leave": "إجازة سنوية",
    "ZATCA Annual Leave for Trainee": "إجازة سنوية (متدرب)",
}
CLASS_AR = {"Employees": "موظف", "Employee": "موظف", "Trainee": "متدرب", "Trainees": "متدرب"}

# الهوية البصرية
NAVY, TIFFANY = "152446", "4BB8C4"
FONT = "Somar"
F_TITLE = Font(name=FONT, bold=True, size=16, color=NAVY)
F_SECTION = Font(name=FONT, bold=True, size=13, color=NAVY)
F_META = Font(name=FONT, size=10, color="555555")
F_HEAD = Font(name=FONT, bold=True, size=10, color="FFFFFF")
F_BODY = Font(name=FONT, size=10, color="222222")
F_BOLD = Font(name=FONT, bold=True, size=10, color="222222")
F_KPI = Font(name=FONT, bold=True, size=17, color=NAVY)
F_KPI_LABEL = Font(name=FONT, size=9, color="555555")
F_BAD = Font(name=FONT, bold=True, size=10, color="9C0006")
F_WARN = Font(name=FONT, bold=True, size=10, color="7A5A00")
F_OK = Font(name=FONT, bold=True, size=10, color="3F5A22")
FILL_HEAD = PatternFill("solid", fgColor=NAVY)
FILL_BAD = PatternFill("solid", fgColor="F8D7DA")
FILL_WARN = PatternFill("solid", fgColor="FFF4D6")
FILL_OK = PatternFill("solid", fgColor="E3EED8")
FILL_KPI = PatternFill("solid", fgColor="F2F6F8")
FILL_BAND = PatternFill("solid", fgColor="F7F8FA")
LINE = Side(style="thin", color="C9CED6")
BORDER = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)

SUMMARY, CONFLICTS, LOG, RAW = "ملخص التدقيق", "التعارضات", "سجل الإجازات", "البيانات الخام"
T_EMP, T_TRAINEE = "بين موظفين", "يشمل متدربين"


def normalize_text(v):
    if pd.isna(v):
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def fix_manager(v):
    s = normalize_text(v)
    if not s:
        return "غير محدد"
    if "," in s or "،" in s:
        parts = [p.strip() for p in re.split(r"[,،]", s, maxsplit=1)]
        if len(parts) == 2:
            return f"{parts[1]} {parts[0]}".strip()
    return s


def tr(mapping, v):
    s = normalize_text(v)
    return mapping.get(s, s or "غير محدد")


def parse_dates(col):
    # تحليل كل قيمة بشكل مستقل: ISO أولاً، ثم الصيغ الشائعة، ثم محاولة احتياطية.
    def parse_one(value):
        if pd.isna(value) or normalize_text(value) == "":
            return pd.NaT
        if isinstance(value, (pd.Timestamp, )):
            return pd.Timestamp(value).normalize()
        if hasattr(value, "year") and hasattr(value, "month") and hasattr(value, "day"):
            return pd.Timestamp(value).normalize()
        if isinstance(value, (int, float)):
            return (pd.Timestamp("1899-12-30") + pd.to_timedelta(float(value), unit="D")).normalize()
        s = normalize_text(value)
        for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d-%b-%Y", "%d-%b-%y",
                    "%d/%m/%Y", "%d/%m/%y", "%d %b %Y", "%b %d, %Y"):
            try:
                return pd.Timestamp(pd.to_datetime(s, format=fmt)).normalize()
            except (ValueError, TypeError):
                pass
        return pd.to_datetime(s, errors="coerce", dayfirst=True).normalize()
    return pd.Series([parse_one(v) for v in col], index=col.index, dtype="datetime64[ns]")


def load():
    p = Path(RAW_FILE)
    if not p.exists():
        sys.exit(f"Raw file not found: {p.resolve()}")
    raw = pd.read_csv(p, encoding="utf-8-sig") if p.suffix.lower() == ".csv" else pd.read_excel(p)
    raw.columns = [normalize_text(c) for c in raw.columns]

    # مطابقة أسماء الأعمدة بعد إزالة المسافات الزائدة
    cols = {}
    for key, wanted in REQUIRED.items():
        wanted_n = normalize_text(wanted).lower()
        cols[key] = next((c for c in raw.columns if normalize_text(c).lower() == wanted_n), None)

    missing = [REQUIRED[k] for k, v in cols.items() if v is None
               and k in ("first", "last", "id", "loc", "mgr", "start", "end", "approval", "class")]
    if missing:
        sys.exit(f"Missing required columns: {missing}\nFound: {list(raw.columns)}")
    return raw, cols


def build_log(raw, c):
    get = lambda k: raw[c[k]] if c.get(k) else pd.Series([None] * len(raw), index=raw.index)
    start, end = parse_dates(get("start")), parse_dates(get("end"))

    log = pd.DataFrame({
        "اسم الموظف": (get("first").fillna("").astype(str).str.strip() + " " +
                       get("last").fillna("").astype(str).str.strip()).str.strip(),
        "الرقم الوظيفي": get("id").map(normalize_text),
        "الفئة": get("class").map(lambda v: CLASS_AR.get(normalize_text(v), normalize_text(v) or "غير محدد")),
        "القطاع": get("l2").map(lambda v: tr(LEVEL2_AR, v)),
        "الإدارة": get("l3").map(lambda v: tr(LEVEL3_AR, v)),
        "المنفذ": get("loc").map(lambda v: tr(LOCATION_AR, v)),
        "المدير": get("mgr").map(fix_manager),
        "نوع الإجازة": get("type").map(lambda v: tr(LEAVE_TYPE_AR, v)),
        "تاريخ البداية": start,
        "تاريخ النهاية": end,
        "عدد الأيام": pd.to_numeric(get("qty"), errors="coerce"),
        "حالة الاعتماد": get("approval").map(normalize_text).str.upper().map(
            {"APPROVED": "معتمدة"}).fillna(get("approval").map(normalize_text)),
    })

    # لا نعتمد سجلًا بلا رقم وظيفي أو تاريخ صالح أو فترة منطقية.
    log["سبب الاستبعاد"] = ""
    log.loc[log["الرقم الوظيفي"].eq(""), "سبب الاستبعاد"] = "رقم وظيفي مفقود"
    log.loc[log["تاريخ البداية"].isna() | log["تاريخ النهاية"].isna(), "سبب الاستبعاد"] = "تاريخ غير صالح"
    log.loc[log["تاريخ النهاية"] < log["تاريخ البداية"], "سبب الاستبعاد"] = "تاريخ النهاية يسبق البداية"
    log["حالة التدقيق"] = log["حالة الاعتماد"].eq("معتمدة").map({True: "داخل التدقيق", False: "مستبعدة: غير معتمدة"})
    log.loc[log["سبب الاستبعاد"].ne(""), "حالة التدقيق"] = "مستبعدة: " + log.loc[log["سبب الاستبعاد"].ne(""), "سبب الاستبعاد"]
    log["ضمن تعارض"] = ""
    log["أرقام حالات التعارض"] = ""
    return log


def find_conflicts(log):
    """يحسب الأشخاص الفريدين في كل يوم، بعد استبعاد غير المعتمد والسجلات غير الصالحة."""
    rows = []
    eligible = log[
        log["حالة التدقيق"].eq("داخل التدقيق")
        & log["الرقم الوظيفي"].ne("")
        & log["تاريخ البداية"].notna()
        & log["تاريخ النهاية"].notna()
        & (log["تاريخ النهاية"] >= log["تاريخ البداية"])
    ]

    case_no = 0
    for (mgr, loc), g in eligible.groupby(["المدير", "المنفذ"], sort=True):
        if len(g) < OVERLAP_THRESHOLD:
            continue

        # لكل يوم: مجموعة أرقام وظيفية فريدة، لا عدد السجلات.
        people_by_day = {}
        for _, rec in g.iterrows():
            for day in pd.date_range(rec["تاريخ البداية"], rec["تاريخ النهاية"], freq="D"):
                people_by_day.setdefault(day, set()).add(rec["الرقم الوظيفي"])

        hot_days = sorted(day for day, people in people_by_day.items()
                          if len(people) >= OVERLAP_THRESHOLD)
        if not hot_days:
            continue

        # تقسيم الأيام المتتالية إلى حالات مستقلة.
        runs, current = [], []
        for day in hot_days:
            if current and day - current[-1] > timedelta(days=1):
                runs.append(current)
                current = []
            current.append(day)
        if current:
            runs.append(current)

        for run in runs:
            case_no += 1
            start_case, end_case = run[0], run[-1]
            hot_set = set(run)

            # السجلات التي تشارك فعليًا في يوم تعارض واحد على الأقل.
            participating = []
            for idx, rec in g.iterrows():
                overlap_days = pd.date_range(
                    max(rec["تاريخ البداية"], start_case),
                    min(rec["تاريخ النهاية"], end_case), freq="D"
                )
                if any(day in hot_set for day in overlap_days):
                    participating.append(idx)

            members = g.loc[participating].copy()
            unique_people = members.drop_duplicates("الرقم الوظيفي")
            peak = max(len(people_by_day[d]) for d in run)

            # تصنيف الحالة بناءً على الأشخاص الفريدين في اليوم نفسه.
            employee_peak = 0
            for d in run:
                active = g[(g["تاريخ البداية"] <= d) & (g["تاريخ النهاية"] >= d)]
                employee_peak = max(
                    employee_peak,
                    active[active["الفئة"].ne("متدرب")]["الرقم الوظيفي"].nunique()
                )
            conflict_type = T_EMP if employee_peak >= OVERLAP_THRESHOLD else T_TRAINEE

            # لا نكرر الشخص في قائمة الأسماء، ونعرض كل فترات إجازته المشاركة.
            person_lines = []
            for person_id, person_rows in members.groupby("الرقم الوظيفي", sort=False):
                first = person_rows.iloc[0]
                periods = "، ".join(
                    f"{r['تاريخ البداية']:%d/%m/%Y} إلى {r['تاريخ النهاية']:%d/%m/%Y}"
                    for _, r in person_rows.sort_values("تاريخ البداية").iterrows()
                )
                suffix = " (متدرب)" if first["الفئة"] == "متدرب" else ""
                person_lines.append(f"{first['اسم الموظف']}{suffix} — {periods}")

            rows.append({
                "رقم الحالة": case_no,
                "المدير": mgr,
                "المنفذ": loc,
                "الإدارة": g["الإدارة"].mode().iat[0],
                "من": start_case,
                "إلى": end_case,
                "عدد أيام التعارض": len(run),
                "أكبر عدد أشخاص في إجازة باليوم نفسه": peak,
                "عدد الأشخاص": unique_people["الرقم الوظيفي"].nunique(),
                "نوع التعارض": conflict_type,
                "الأشخاص وفترات إجازاتهم": "\n".join(person_lines),
            })

            # اربط فقط السجلات المشاركة فعليًا بالحالة.
            for idx in participating:
                log.at[idx, "ضمن تعارض"] = "نعم"
                old = log.at[idx, "أرقام حالات التعارض"]
                log.at[idx, "أرقام حالات التعارض"] = f"{old}، {case_no}" if old else str(case_no)

    return pd.DataFrame(rows)


def sheet(wb, title, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    return ws


def head_row(ws, row, headers, height=30):
    for j, h in enumerate(headers, 1):
        c = ws.cell(row, j, h)
        c.fill, c.font, c.alignment, c.border = FILL_HEAD, F_HEAD, CENTER, BORDER
    ws.row_dimensions[row].height = height


def cell(ws, row, col, value, font=F_BODY, fill=None, fmt=None, align=CENTER):
    c = ws.cell(row, col, value)
    c.font, c.alignment, c.border = font, align, BORDER
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    return c


def section(ws, row, title, span):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row, 1, title)
    c.font, c.alignment = F_SECTION, RIGHT
    c.border = Border(bottom=Side(style="medium", color=TIFFANY))


def title_block(ws, title, meta, span):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    ws["A1"] = title
    ws["A1"].font, ws["A1"].alignment = F_TITLE, RIGHT
    ws.row_dimensions[1].height = 34
    for row, (label, value) in enumerate(meta, 2):
        ws.cell(row, 1, label).font = F_BOLD
        ws.cell(row, 1).alignment = RIGHT
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=span)
        ws.cell(row, 2, value).font = F_META
        ws.cell(row, 2).alignment = RIGHT


def print_setup(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0


def write_df(ws, df, row, date_cols=(), widths=None):
    head_row(ws, row, list(df.columns))
    for i, rec in enumerate(df.itertuples(index=False, name=None), row + 1):
        for j, (header, value) in enumerate(zip(df.columns, rec), 1):
            if isinstance(value, pd.Timestamp):
                value = value.to_pydatetime()
            elif not isinstance(value, str) and pd.isna(value):
                value = None
            cell(ws, i, j, value, fill=FILL_BAND if (i-row) % 2 == 0 else None,
                 fmt="yyyy-mm-dd" if header in date_cols else None,
                 align=RIGHT if header in ("اسم الموظف", "المدير", "المنفذ", "الأشخاص وفترات إجازاتهم") else CENTER)
    if widths:
        for j, header in enumerate(df.columns, 1):
            ws.column_dimensions[get_column_letter(j)].width = widths.get(header, 15)
    return row + len(df)


def build_report(raw, log, conf):
    wb = Workbook()
    ws = sheet(wb, SUMMARY, first=True)
    widths = [25, 30, 26, 16, 16, 16, 16, 16, 16]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    valid_approved = log[log["حالة التدقيق"].eq("داخل التدقيق")]
    period_from = valid_approved["تاريخ البداية"].min()
    period_to = valid_approved["تاريخ النهاية"].max()
    period = f"من {period_from:%d/%m/%Y} إلى {period_to:%d/%m/%Y}" if pd.notna(period_from) and pd.notna(period_to) else "—"
    title_block(ws, "تقرير تدقيق تداخل الإجازات المجدولة لموظفي المنافذ", [
        ("نطاق البيانات", f"الإجازات المعتمدة {period}"),
        ("معيار التدقيق", f"وجود {OVERLAP_THRESHOLD} أشخاص مختلفين أو أكثر في إجازة في اليوم نفسه، تحت المدير نفسه وفي المنفذ نفسه"),
        ("الإعداد", f"{PREPARED_BY} — {pd.Timestamp.today():%d/%m/%Y}"),
    ], 9)

    r = 6
    section(ws, r, "أولاً: المؤشرات الرئيسية", 9)
    approved_count = len(valid_approved)
    manager_count = valid_approved["المدير"].nunique()
    conflict_count = len(conf)
    manager_conflict_count = conf["المدير"].nunique() if conflict_count else 0
    emp_count = int((conf["نوع التعارض"] == T_EMP).sum()) if conflict_count else 0
    trainee_count = int((conf["نوع التعارض"] == T_TRAINEE).sum()) if conflict_count else 0
    kpis = [
        ("إجمالي الإجازات المعتمدة", approved_count),
        ("عدد المديرين", manager_count),
        ("حالات التعارض", conflict_count),
        ("مديرون لديهم تعارض", manager_conflict_count),
        ("تعارض بين موظفين", emp_count),
        ("تعارض يشمل متدربين", trainee_count),
    ]
    for k, (label, value) in enumerate(kpis):
        col = 1 + (k % 3) * 3
        rr = 7 + (k // 3) * 2
        ws.merge_cells(start_row=rr, start_column=col, end_row=rr, end_column=col+2)
        ws.merge_cells(start_row=rr+1, start_column=col, end_row=rr+1, end_column=col+2)
        lc, vc = ws.cell(rr, col, label), ws.cell(rr+1, col, value)
        lc.font, lc.alignment, lc.fill = F_KPI_LABEL, CENTER, FILL_KPI
        vc.font, vc.alignment, vc.fill = F_KPI, CENTER, FILL_KPI
        ws.row_dimensions[rr+1].height = 30

    r = 12
    section(ws, r, "ثانياً: الملخص حسب المدير والمنفذ", 9)
    r += 1
    headers = ["المدير", "المنفذ", "الإدارة", "عدد الإجازات المعتمدة", "حالات التعارض",
               "أيام التعارض", "أكبر عدد أشخاص في إجازة باليوم نفسه", "يشمل متدربين", "النتيجة"]
    head_row(ws, r, headers, 34)
    groups = (valid_approved.groupby(["المدير", "المنفذ"])
              .agg(الإدارة=("الإدارة", lambda s: s.mode().iat[0]),
                   الإجازات=("الرقم الوظيفي", "size")).reset_index())
    for _, group in groups.sort_values(["المدير", "المنفذ"]).iterrows():
        r += 1
        m, loc = group["المدير"], group["المنفذ"]
        matched = conf[(conf["المدير"] == m) & (conf["المنفذ"] == loc)] if conflict_count else conf
        has = not matched.empty
        cell(ws, r, 1, m, F_BOLD, align=RIGHT)
        cell(ws, r, 2, loc, align=RIGHT)
        cell(ws, r, 3, group["الإدارة"], align=RIGHT)
        cell(ws, r, 4, int(group["الإجازات"]))
        cell(ws, r, 5, len(matched))
        cell(ws, r, 6, int(matched["عدد أيام التعارض"].sum()) if has else 0)
        cell(ws, r, 7, int(matched["أكبر عدد أشخاص في إجازة باليوم نفسه"].max()) if has else "—")
        cell(ws, r, 8, "نعم" if has and (matched["نوع التعارض"] == T_TRAINEE).any() else "—")
        result = cell(ws, r, 9, "يوجد تعارض" if has else "لا يوجد تعارض")
        result.fill, result.font = (FILL_WARN, F_WARN) if has else (FILL_OK, F_OK)

    ws.freeze_panes = "A6"
    ws.auto_filter.ref = f"A13:I{r}"
    print_setup(ws)

    # ورقة التعارضات
    wc = sheet(wb, CONFLICTS)
    conflict_cols = ["رقم الحالة", "المدير", "المنفذ", "الإدارة", "من", "إلى", "عدد أيام التعارض",
                     "أكبر عدد أشخاص في إجازة باليوم نفسه", "عدد الأشخاص", "نوع التعارض",
                     "الأشخاص وفترات إجازاتهم"]
    if not conf.empty:
        write_df(wc, conf[conflict_cols], 1, date_cols=("من", "إلى"), widths={
            "رقم الحالة": 10, "المدير": 22, "المنفذ": 28, "الإدارة": 24, "من": 13, "إلى": 13,
            "عدد أيام التعارض": 13, "أكبر عدد أشخاص في إجازة باليوم نفسه": 20, "عدد الأشخاص": 11,
            "نوع التعارض": 17, "الأشخاص وفترات إجازاتهم": 58})
        for i in range(2, len(conf)+2):
            t = wc.cell(i, 10)
            t.fill, t.font = (FILL_BAD, F_BAD) if t.value == T_EMP else (FILL_WARN, F_WARN)
            wc.cell(i, 11).alignment = RIGHT
            wc.row_dimensions[i].height = max(28, 18 * (str(wc.cell(i, 11).value).count("\n")+1))
        wc.auto_filter.ref = f"A1:K{len(conf)+1}"
    else:
        head_row(wc, 1, conflict_cols)
        wc["A2"] = "لم تُرصد حالات تعارض"
    wc.freeze_panes = "B2"
    print_setup(wc)

    # سجل الإجازات
    wl = sheet(wb, LOG)
    end_row = write_df(wl, log, 1, date_cols=("تاريخ البداية", "تاريخ النهاية"), widths={
        "اسم الموظف": 24, "الرقم الوظيفي": 13, "الفئة": 10, "القطاع": 22, "الإدارة": 24,
        "المنفذ": 28, "المدير": 22, "نوع الإجازة": 20, "تاريخ البداية": 13, "تاريخ النهاية": 13,
        "عدد الأيام": 10, "حالة الاعتماد": 13, "سبب الاستبعاد": 22, "حالة التدقيق": 22,
        "ضمن تعارض": 12, "أرقام حالات التعارض": 18})
    flag_col = list(log.columns).index("ضمن تعارض") + 1
    for i in range(2, end_row+1):
        if wl.cell(i, flag_col).value == "نعم":
            wl.cell(i, flag_col).fill, wl.cell(i, flag_col).font = FILL_BAD, F_BAD
    wl.auto_filter.ref = f"A1:{get_column_letter(len(log.columns))}{end_row}"
    wl.freeze_panes = "B2"
    print_setup(wl)

    # البيانات الخام كما وردت
    wr = sheet(wb, RAW)
    write_df(wr, raw, 1, widths={h: 22 for h in raw.columns})
    wr.auto_filter.ref = f"A1:{get_column_letter(len(raw.columns))}{len(raw)+1}"
    wr.freeze_panes = "A2"

    wb.save(OUTPUT_FILE)


def main():
    raw, cols = load()
    log = build_log(raw, cols)
    conf = find_conflicts(log)
    build_report(raw, log, conf)
    print(f"Done: {OUTPUT_FILE}")
    print(f"إجمالي السجلات: {len(log)} | المعتمدة: {(log['حالة التدقيق']=='داخل التدقيق').sum()} | حالات التعارض: {len(conf)}")
    if len(conf):
        print(conf[["رقم الحالة", "المدير", "المنفذ", "من", "إلى", "أكبر عدد أشخاص في إجازة باليوم نفسه", "عدد الأشخاص", "نوع التعارض"]].to_string(index=False))


if __name__ == "__main__":
    main()

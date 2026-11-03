import re
import unicodedata
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
RAW_FILE = r"Ports Leaves.xlsx"
EMPLOYEE_FILE = r"ضعي هنا مسار ملف الموظفين.xlsx"
OUTPUT_FILE = r"تقرير تدقيق تداخل الإجازات.xlsx"
PREPARED_BY = "إدارة أداء وكالة رأس المال البشري"
OVERLAP_THRESHOLD = 2

# أسماء الحقول المطلوبة من ملف المصدر
REQUIRED = {
    "first": "(Employee) First Name",
    "last": "(Employee) Last Name",
    "id": "(Employee) Person ID",
    "l1": "(Employee) Level 1 (Label)",
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

LEVEL1_AR = {"Security Affairs": "الشؤون الأمنية"}
LEVEL2_AR = {
    "Business Support": "دعم الأعمال",
    "Ports Management": "إدارة المنافذ",
}
# تُترك المسميات غير المؤكدة بالإنجليزية لتجنب اعتماد ترجمة تنظيمية غير دقيقة.
LEVEL3_AR = {
    "Land Ports Management": "إدارة المنافذ البرية",
    "Land ports Management": "إدارة المنافذ البرية",
    "Airports Management": "إدارة المطارات",
    "Seaports Management": "إدارة الموانئ البحرية",
}
LOCATION_AR = {
    "Al-Durra Land Port": "منفذ الدرة البري",
    "King Abdulaziz Intl. AirPort": "مطار الملك عبدالعزيز الدولي",
    "Abha Airport": "مطار أبها",
    "Salwa Land Port": "منفذ سلوى البري",
    "Najran Airport": "مطار نجران",
    "Alhadithah Land Port": "منفذ الحديثة البري",
    "Headquarter": "المقر الرئيسي",
    "Jeddah Islamic Seaport Seaport": "ميناء جدة الإسلامي",
    "Jeddah Islamic Seaport": "ميناء جدة الإسلامي",
    "King Khalid Intl., Airport": "مطار الملك خالد الدولي",
    "King Khalid Intl. Airport": "مطار الملك خالد الدولي",
    "King Abdulaziz Seaport": "ميناء الملك عبدالعزيز",
    "Yanbu Commercial Seaport": "ميناء ينبع التجاري",
    "Al-Wadia Land Port": "منفذ الوديعة البري",
    "Riggae Land Port": "منفذ الرقعي البري",
    "King Fahd Causeway": "جسر الملك فهد",
    "King Fahd Intl. Airport": "مطار الملك فهد الدولي",
    "Albatha Land Port": "منفذ البطحاء البري",
    "P.Mohammad Bin Abdulaziz Airport": "مطار الأمير محمد بن عبدالعزيز",
    "P.Naif Bin Abdulaziz Airport": "مطار الأمير نايف بن عبدالعزيز",
    "P. Naif Bin Abdulaziz Airport": "مطار الأمير نايف بن عبدالعزيز",
    "P.Abdulmohsen Bin Abdulaziz Air": "مطار الأمير عبدالمحسن بن عبدالعزيز",
    "P. Abdulmohsen Bin Abdulaziz Air": "مطار الأمير عبدالمحسن بن عبدالعزيز",
    "Alap Land Port": "منفذ علب البري",
    "Al Khafji Land Port": "منفذ الخفجي البري",
    "Al-Ahsa Airport": "مطار الأحساء",
    "Jubail Seaport": "ميناء الجبيل",
    "Jazan Seaport": "ميناء جازان",
    "Neom Seaport": "ميناء نيوم",
    "King Abdullah Airport": "مطار الملك عبدالله",
    "Al Khadhra Land Port": "منفذ الخضراء البري",
    "Ar Rub Al Khali Land Port": "منفذ الربع الخالي البري",
    "Halat Ammar Land Port": "منفذ حالة عمار البري",
    "Dry Port": "الميناء الجاف",
    "Ras Tanura Seaport": "ميناء رأس تنورة",
    "Jadeedath Arar Land Port": "منفذ جديدة عرعر البري",
    "Taif Airport": "مطار الطائف",
    "King Abdullah Seaport Rabigh": "ميناء الملك عبدالله - رابغ",
    "P.Sultan Bin Abdulaziz Airport": "مطار الأمير سلطان بن عبدالعزيز",
    "Hail Airport": "مطار حائل",
    "Jeddah Branch": "فرع جدة",
    "Red Sea Int`l Airport": "مطار البحر الأحمر الدولي",
    "AlUla International Airport": "مطار العلا الدولي",
    "Medina Branch": "فرع المدينة المنورة",
    "Al Jouf Airport": "مطار الجوف",
    "Tabuk Branch": "فرع تبوك",
    "NEOM Airport": "مطار نيوم",
    "Al-Ahsa Branch": "فرع الأحساء",
    "Al Wajh Airport": "مطار الوجه",
    "Hafr Al, Batin Branch": "فرع حفر الباطن",
    "Yanbu Branch": "فرع ينبع",
    "King Khalid Air Force Base": "قاعدة الملك خالد الجوية",
    "Sharurah Airport": "مطار شرورة",
    "Dammam Branch": "فرع الدمام",
    "Ras Al Khair Seaport": "ميناء رأس الخير",
    "Baish Seaport": "ميناء بيش",
    "Farasan Seaport": "ميناء فرسان",
    "Jazan Branch": "فرع جازان",
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
    if not s:
        return ""
    for key, value in mapping.items():
        if key.casefold() == s.casefold():
            return value
    return s


def clean_id(value):
    s = normalize_text(value)
    return re.sub(r"\.0$", "", s) if s else ""


def location_key(value):
    s = normalize_text(value)
    s = unicodedata.normalize("NFKC", s).casefold()
    return re.sub(r"[^\w\u0600-\u06ff]+", "", s)


LOCATION_LOOKUP = {location_key(k): v for k, v in LOCATION_AR.items()}
LOCATION_LOOKUP.update({location_key(v): v for v in LOCATION_AR.values()})


def canonical_location(value):
    s = normalize_text(value)
    return LOCATION_LOOKUP.get(location_key(s), s)


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



def load_employee_master():
    p = Path(EMPLOYEE_FILE)
    if not p.exists():
        sys.exit(f"Employee master file not found: {p.resolve()}")

    df = pd.read_csv(p, encoding="utf-8-sig") if p.suffix.lower() == ".csv" else pd.read_excel(p)
    df.columns = [normalize_text(c) for c in df.columns]

    required = {
        "id": ["الرقم الوظيفي"],
        "manager_id": ["الرقم الوظيفي للمدير المباشر"],
        "manager_first": ["المدير المباشر - الاسم الاول", "المدير المباشر - الاسم الأول"],
        "manager_last": ["المدير المباشر - الاسم الاخير", "المدير المباشر - الاسم الأخير"],
        "location": ["Location (Name)"],
        "level3": ["المستوى الثالث"],
        "level1": ["المستوى الأول", "المستوى الاول", "Level 1", "Level 1 (Label)", "(Employee) Level 1 (Label)"],
    }
    cols = {}
    for key, candidates in required.items():
        wanted = {name.casefold() for name in candidates}
        cols[key] = next((c for c in df.columns if normalize_text(c).casefold() in wanted), None)

    missing = [" / ".join(required[k]) for k, v in cols.items() if v is None]
    if missing:
        sys.exit(f"Missing employee master columns: {missing}\nFound: {list(df.columns)}")

    # يقتصر حصر أعداد الموظفين على قطاع الشؤون الأمنية فقط.
    # لا يطبق هذا الفلتر على ملف الإجازات أو سجل التدقيق.
    df = df[df[cols["level1"]].map(normalize_text).str.casefold().eq("security affairs")].copy()

    employees = pd.DataFrame({
        "الرقم الوظيفي": df[cols["id"]].map(clean_id),
        "معرف المدير": df[cols["manager_id"]].map(clean_id),
        "المدير": (df[cols["manager_first"]].fillna("").astype(str).str.strip() + " " +
                    df[cols["manager_last"]].fillna("").astype(str).str.strip()).str.strip(),
        "المنفذ": df[cols["location"]].map(canonical_location),
        "المستوى الثالث": df[cols["level3"]].map(lambda v: tr(LEVEL3_AR, v)),
    })
    employees = employees[
        employees["الرقم الوظيفي"].ne("") & employees["المنفذ"].ne("")
    ].copy()
    employees = employees.drop_duplicates(subset=["الرقم الوظيفي", "معرف المدير", "المنفذ"])
    return employees


def build_log(raw, c):
    get = lambda k: raw[c[k]] if c.get(k) else pd.Series([None] * len(raw), index=raw.index)
    start, end = parse_dates(get("start")), parse_dates(get("end"))

    log = pd.DataFrame({
        "اسم الموظف": (get("first").fillna("").astype(str).str.strip() + " " +
                       get("last").fillna("").astype(str).str.strip()).str.strip(),
        "الرقم الوظيفي": get("id").map(clean_id),
        "الفئة": get("class").map(lambda v: CLASS_AR.get(normalize_text(v), normalize_text(v) or "غير محدد")),
        "المستوى الأول": get("l1").map(lambda v: tr(LEVEL1_AR, v)),
        "المستوى الثاني": get("l2").map(lambda v: tr(LEVEL2_AR, v)),
        "المستوى الثالث": get("l3").map(lambda v: tr(LEVEL3_AR, v)),
        "المنفذ": get("loc").map(canonical_location),
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
                "المستوى الثالث": g["المستوى الثالث"].mode().iat[0],
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



def reconcile_master(employees, log):
    """Assign leave to master group only when one unambiguous master assignment exists."""
    keys = ["الرقم الوظيفي", "معرف المدير", "المدير", "المنفذ"]
    assignments = employees[keys].drop_duplicates()
    unique_counts = assignments.groupby("الرقم الوظيفي").size()
    valid_ids = set(unique_counts[unique_counts == 1].index)
    unique_master = assignments[assignments["الرقم الوظيفي"].isin(valid_ids)].drop_duplicates("الرقم الوظيفي")
    merged = log.merge(unique_master, on="الرقم الوظيفي", how="left", suffixes=("_الإجازة", "_الموظفين"), indicator=True)
    ambiguous_ids = set(unique_counts[unique_counts > 1].index)
    merged["حالة المطابقة"] = "مطابق"
    merged.loc[merged["_merge"].eq("left_only"), "حالة المطابقة"] = "غير موجود في ملف الموظفين"
    merged.loc[merged["الرقم الوظيفي"].isin(ambiguous_ids), "حالة المطابقة"] = "توزيع متعدد في ملف الموظفين"
    loc_diff = merged["حالة المطابقة"].eq("مطابق") & merged["المنفذ_الإجازة"].ne(merged["المنفذ_الموظفين"])
    merged.loc[loc_diff, "حالة المطابقة"] = "اختلاف موقع الإجازة عن ملف الموظفين"
    manager_diff = merged["حالة المطابقة"].eq("مطابق") & merged["المدير_الإجازة"].ne(merged["المدير_الموظفين"])
    merged.loc[manager_diff, "حالة المطابقة"] = "اختلاف اسم المدير (راجع)"
    return merged


def build_employee_summary(employees, log):
    """Raw leave file determines manager/location groups; master supplies employee counts only."""
    q4_start, q4_end = pd.Timestamp("2026-10-01"), pd.Timestamp("2026-12-31")
    eligible = log[
        log["حالة التدقيق"].eq("داخل التدقيق")
        & log["تاريخ البداية"].le(q4_end)
        & log["تاريخ النهاية"].ge(q4_start)
        & log["المدير"].ne("") & log["المنفذ"].ne("")
    ].copy()

    # The manager and port labels come ONLY from the leave data.
    groups = (eligible.groupby(["المدير", "المنفذ"], dropna=False)
              .agg(المستوى_الثالث=("المستوى الثالث", lambda x: x.mode().iat[0] if not x.mode().empty else ""),
                   **{"موظفون لديهم إجازة": ("الرقم الوظيفي", "nunique"),
                      "عدد الإجازات المعتمدة": ("الرقم الوظيفي", "size")})
              .reset_index())

    # Infer a master manager-ID -> raw manager crosswalk from matching employee IDs
    # at the SAME port. Never overwrite the raw manager or port.
    links = (eligible[["الرقم الوظيفي", "المدير", "المنفذ"]].drop_duplicates()
             .merge(employees[["الرقم الوظيفي", "معرف المدير", "المنفذ"]].drop_duplicates(),
                    on=["الرقم الوظيفي", "المنفذ"], how="inner"))
    links = links[links["معرف المدير"].ne("")]
    vote = (links.groupby(["معرف المدير", "المنفذ", "المدير"])["الرقم الوظيفي"]
            .nunique().reset_index(name="دليل"))
    # A manager ID with competing raw-manager names in one location is ambiguous.
    unique = vote.groupby(["معرف المدير", "المنفذ"])["المدير"].nunique()
    valid_keys = set(unique[unique == 1].index)
    vote = vote[vote.apply(lambda x: (x["معرف المدير"], x["المنفذ"]) in valid_keys, axis=1)]
    # Also allow exact normalized names where no ID-based evidence is available.
    emp = employees.copy()
    emp["المدير_من_الخام"] = ""
    lookup = {(x["معرف المدير"], x["المنفذ"]): x["المدير"] for _, x in vote.iterrows()}
    emp["المدير_من_الخام"] = emp.apply(
        lambda x: lookup.get((x["معرف المدير"], x["المنفذ"]), ""), axis=1)
    raw_names = set(zip(groups["المدير"], groups["المنفذ"]))
    fallback = emp["المدير_من_الخام"].eq("") & emp.apply(
        lambda x: (fix_manager(x["المدير"]), x["المنفذ"]) in raw_names, axis=1)
    emp.loc[fallback, "المدير_من_الخام"] = emp.loc[fallback, "المدير"].map(fix_manager)

    assigned = emp[emp["المدير_من_الخام"].ne("")].copy()
    # One employee can be assigned multiple times in the master; avoid double-counting
    # within each manager+port, and flag uncertain population separately.
    totals = (assigned.groupby(["المدير_من_الخام", "المنفذ"])["الرقم الوظيفي"]
              .nunique().reset_index(name="إجمالي_الموظفين")
              .rename(columns={"المدير_من_الخام": "المدير"}))
    groups = groups.merge(totals, on=["المدير", "المنفذ"], how="left")
    groups["حالة حصر الموظفين"] = groups["إجمالي_الموظفين"].apply(
        lambda x: "لا توجد مطابقة موثوقة للحصر" if pd.isna(x) else "تمت المطابقة")
    groups["إجمالي_الموظفين"] = groups["إجمالي_الموظفين"].astype("Int64")
    groups["نسبة الموظفين ذوي الإجازات"] = groups.apply(
        lambda x: x["موظفون لديهم إجازة"] / x["إجمالي_الموظفين"]
        if pd.notna(x["إجمالي_الموظفين"]) and x["إجمالي_الموظفين"] > 0
        and x["موظفون لديهم إجازة"] <= x["إجمالي_الموظفين"] else float("nan"), axis=1)
    return groups


def audit_conflicts_on_master(log, reconciled):
    """Preserve raw leave assignments as the authority for conflict analysis."""
    return log.copy()


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


def presentation_metrics(log, conf, summary):
    """Daily unique-person concurrency; Q4 2026 only; raw manager and location authoritative."""
    start, end = pd.Timestamp("2026-10-01"), pd.Timestamp("2026-12-31")
    eligible = log[(log["حالة التدقيق"] == "داخل التدقيق")
                   & (log["تاريخ البداية"] <= end)
                   & (log["تاريخ النهاية"] >= start)].copy()
    day_people = {}
    people_meta = {}
    for _, row in eligible.iterrows():
        person = row["الرقم الوظيفي"]
        if not person or not row["المدير"] or not row["المنفذ"]:
            continue
        key = (row["المدير"], row["المنفذ"])
        people_meta[(key, person)] = (row["اسم الموظف"], row["الفئة"])
        a = max(row["تاريخ البداية"], start)
        b = min(row["تاريخ النهاية"], end)
        for day in pd.date_range(a, b):
            day_people.setdefault((key, day), set()).add(person)

    person_days = {}
    daily_rows = []
    for (key, day), persons in sorted(day_people.items(), key=lambda x: (x[0][0], x[0][1])):
        n = len(persons)
        if n < OVERLAP_THRESHOLD:
            continue
        daily_rows.append({"المدير": key[0], "المنفذ": key[1], "التاريخ": day,
                           "عدد الموظفين المتزامنين": n})
        for person in persons:
            person_days.setdefault((key, person), set()).add(day)

    # Case count per person: one case is a consecutive run of overlap days
    # within one manager/location; use case boundaries from the original audit.
    person_case_ids = {}
    for _, case in conf.iterrows():
        key = (case["المدير"], case["المنفذ"])
        for (k, person), days in person_days.items():
            if k == key and any(case["من"] <= d <= case["إلى"] for d in days):
                person_case_ids.setdefault((k, person), set()).add(case["رقم الحالة"])

    people_rows = []
    for (key, person), days in person_days.items():
        name, category = people_meta.get((key, person), ("", ""))
        people_rows.append({"المدير": key[0], "المنفذ": key[1],
                            "الرقم الوظيفي": person, "اسم الموظف": name,
                            "الفئة": category,
                            "عدد حالات التعارض": len(person_case_ids.get((key, person), set())),
                            "أيام التعارض الفريدة": len(days)})
    people = pd.DataFrame(people_rows, columns=[
        "المدير", "المنفذ", "الرقم الوظيفي", "اسم الموظف", "الفئة",
        "عدد حالات التعارض", "أيام التعارض الفريدة"])
    daily = pd.DataFrame(daily_rows, columns=[
        "المدير", "المنفذ", "التاريخ", "عدد الموظفين المتزامنين"])

    # Non-overlapping bands: 2–4, 5–9, 10–15, 16+.
    bands = [(2, 4, "2–4"), (5, 9, "5–9"), (10, 15, "10–15"), (16, None, "16 فأكثر")]
    distribution = []
    for lo, hi, label in bands:
        subset = daily[daily["عدد الموظفين المتزامنين"] >= lo]
        if hi is not None:
            subset = subset[subset["عدد الموظفين المتزامنين"] <= hi]
        distribution.append({"شريحة التزامن": label, "عدد أيام المدير والمنفذ": len(subset),
                             "النسبة من أيام التزامن": len(subset) / len(daily) if len(daily) else 0})
    dist = pd.DataFrame(distribution)

    # Group-level unique employees and peak concurrency, including groups without overlaps.
    group_rows = []
    for _, row in summary.iterrows():
        key = (row["المدير"], row["المنفذ"])
        group_people = people[(people["المدير"] == key[0]) & (people["المنفذ"] == key[1])]
        group_daily = [len(v) for (k, _), v in day_people.items() if k == key]
        peak = max(group_daily, default=0)
        total = row["إجمالي_الموظفين"]
        group_rows.append({"المدير": key[0], "المنفذ": key[1],
                           "عدد الموظفين المتأثرين بالتعارض": group_people["الرقم الوظيفي"].nunique(),
                           "أعلى عدد متزامن": peak,
                           "أعلى نسبة تزامن": peak / total if pd.notna(total) and total > 0 and peak <= total else None})
    return people, daily, dist, pd.DataFrame(group_rows)


def build_report(raw, log, conf, employee_summary, reconciled, employees):
    people, daily, dist, group_stats = presentation_metrics(log, conf, employee_summary)
    wb = Workbook()
    ws = sheet(wb, SUMMARY, first=True)
    widths = [25, 30, 26, 16, 18, 19, 16, 14, 14, 22, 15, 16]
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
    ], 12)

    r = 6
    section(ws, r, "أولاً: المؤشرات الرئيسية", 12)
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
    section(ws, r, "ثانياً: الملخص حسب المدير والمنفذ", 12)
    r += 1
    headers = [
        "المدير", "المنفذ", "المستوى الثالث", "إجمالي الموظفين",
        "موظفون لديهم إجازة", "نسبة الموظفين ذوي الإجازات", "عدد الإجازات المعتمدة",
        "حالات التعارض", "أيام التعارض", "أكبر عدد أشخاص في إجازة باليوم نفسه",
        "يشمل متدربين", "النتيجة"
    ]
    head_row(ws, r, headers, 34)

    groups = employee_summary.copy()

    for _, group in groups.sort_values(["المدير", "المنفذ"]).iterrows():
        r += 1
        m, loc = group["المدير"], group["المنفذ"]
        population_row = employee_summary[(employee_summary["المدير"] == m) & (employee_summary["المنفذ"] == loc)]
        total_employees = int(group["إجمالي_الموظفين"]) if pd.notna(group["إجمالي_الموظفين"]) else None
        employees_with_leave = int(group["موظفون لديهم إجازة"])
        leave_rate = group["نسبة الموظفين ذوي الإجازات"]

        matched = conf[(conf["المدير"] == m) & (conf["المنفذ"] == loc)] if conflict_count else conf
        has = not matched.empty

        cell(ws, r, 1, m, F_BOLD, align=RIGHT)
        cell(ws, r, 2, loc, align=RIGHT)
        cell(ws, r, 3, group["المستوى_الثالث"], align=RIGHT)
        cell(ws, r, 4, total_employees if total_employees is not None else "غير مطابق")
        cell(ws, r, 5, employees_with_leave)
        cell(ws, r, 6, float(leave_rate) if pd.notna(leave_rate) else "تتطلب مراجعة", fmt="0.0%")
        cell(ws, r, 7, int(group["عدد الإجازات المعتمدة"]))
        cell(ws, r, 8, len(matched))
        cell(ws, r, 9, int(matched["عدد أيام التعارض"].sum()) if has else 0)
        cell(ws, r, 10, int(matched["أكبر عدد أشخاص في إجازة باليوم نفسه"].max()) if has else "—")
        cell(ws, r, 11, "نعم" if has and (matched["نوع التعارض"] == T_TRAINEE).any() else "—")
        result = cell(ws, r, 12, "يوجد تعارض" if has else "لا يوجد تعارض")
        result.fill, result.font = (FILL_WARN, F_WARN) if has else (FILL_OK, F_OK)

    ws.freeze_panes = "A6"
    ws.auto_filter.ref = f"A13:L{r}"
    print_setup(ws)

    # ورقة التعارضات
    wc = sheet(wb, CONFLICTS)
    conflict_cols = ["رقم الحالة", "المدير", "المنفذ", "المستوى الثالث", "من", "إلى", "عدد أيام التعارض",
                     "أكبر عدد أشخاص في إجازة باليوم نفسه", "عدد الأشخاص", "نوع التعارض",
                     "الأشخاص وفترات إجازاتهم"]
    if not conf.empty:
        write_df(wc, conf[conflict_cols], 1, date_cols=("من", "إلى"), widths={
            "رقم الحالة": 10, "المدير": 22, "المنفذ": 28, "المستوى الثالث": 28, "من": 13, "إلى": 13,
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
        "اسم الموظف": 24, "الرقم الوظيفي": 13, "الفئة": 10, "المستوى الأول": 20,
        "المستوى الثاني": 24, "المستوى الثالث": 28,
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

    # كشف الاستثناءات: لا نخفي الاختلافات أو نعتبرها صفرًا.
    wa = sheet(wb, "التحقق من المطابقة")
    check_cols = ["الرقم الوظيفي", "اسم الموظف", "المدير_الإجازة", "المنفذ_الإجازة",
                  "المدير_الموظفين", "المنفذ_الموظفين", "معرف المدير", "حالة المطابقة",
                  "تاريخ البداية", "تاريخ النهاية", "حالة الاعتماد"]
    write_df(wa, reconciled[check_cols], 1, date_cols=("تاريخ البداية", "تاريخ النهاية"),
             widths={c: 27 for c in check_cols})
    wa.auto_filter.ref = f"A1:K{len(reconciled)+1}"
    wa.freeze_panes = "C2"
    print_setup(wa)

    wm = sheet(wb, "حصر الموظفين")
    write_df(wm, employees, 1, widths={c: 25 for c in employees.columns})
    wm.auto_filter.ref = f"A1:E{len(employees)+1}"
    wm.freeze_panes = "C2"
    print_setup(wm)

    notes = sheet(wb, "ملاحظات الحصر")
    issues = employee_summary[
        employee_summary["حالة حصر الموظفين"].ne("تمت المطابقة")
        | employee_summary["نسبة الموظفين ذوي الإجازات"].isna()
    ][["المدير", "المنفذ", "إجمالي_الموظفين", "موظفون لديهم إجازة", "حالة حصر الموظفين"]]
    write_df(notes, issues, 1, widths={c: 29 for c in issues.columns})
    # One-slide, presentation-ready figures. Each row in daily = one port-manager-day.
    wp = sheet(wb, "Presentation Ready")
    overall_people = people["الرقم الوظيفي"].nunique() if not people.empty else 0
    affected_groups = len(group_stats[group_stats["عدد الموظفين المتأثرين بالتعارض"] > 0])
    peak = int(group_stats["أعلى عدد متزامن"].max()) if not group_stats.empty else 0
    approved_people = log[(log["حالة التدقيق"] == "داخل التدقيق")
        & (log["تاريخ البداية"] <= pd.Timestamp("2026-12-31"))
        & (log["تاريخ النهاية"] >= pd.Timestamp("2026-10-01"))]["الرقم الوظيفي"].nunique()
    wp.append(["مؤشرات العرض | الربع الرابع 2026", "القيمة", "تعريف المؤشر"])
    for item in [
        ("حالات التعارض", len(conf), "فترات متصلة فيها موظفان أو أكثر بإجازة تحت نفس المدير والمنفذ؛ قد تمتد الحالة عدة أيام"),
        ("الموظفون المتأثرون بالتعارض", overall_people, "أرقام وظيفية فريدة؛ الموظف يحسب مرة واحدة على مستوى التقرير"),
        ("نسبة المتأثرين من أصحاب الإجازات", overall_people/approved_people if approved_people else None, "عدد المتأثرين ÷ أصحاب الإجازات المعتمدة؛ ليست نسبة من إجمالي القوى العاملة"),
        ("المديرون والمنافذ المتأثرة", affected_groups, "عدد مجموعات المدير والمنفذ التي ظهر فيها تعارض"),
        ("أعلى تزامن في يوم واحد", peak, "أكبر عدد أشخاص بإجازة في اليوم نفسه لدى مدير ومنفذ واحد"),
        ("أيام التزامن على مستوى المدير والمنفذ", len(daily), "اليوم في منفذين مختلفين يحسب مرتين؛ لا يمثل الأيام التقويمية الفريدة"),
    ]:
        wp.append(list(item))
    wp.append([])
    wp.append(["توزيع شدة التزامن", "عدد أيام المدير والمنفذ", "الحصة من أيام التزامن"])
    for _, d in dist.iterrows():
        wp.append([d["شريحة التزامن"], int(d["عدد أيام المدير والمنفذ"]), float(d["النسبة من أيام التزامن"])])
    wp.append([])
    wp.append(["أعلى المديرين والمنافذ من حيث عدد المتأثرين", "عدد الموظفين المتأثرين", "أعلى عدد متزامن", "أعلى نسبة تزامن"])
    top = group_stats.sort_values("عدد الموظفين المتأثرين بالتعارض", ascending=False).head(8)
    for _, g in top.iterrows():
        wp.append([f'{g["المدير"]} | {g["المنفذ"]}', int(g["عدد الموظفين المتأثرين بالتعارض"]),
                   int(g["أعلى عدد متزامن"]), float(g["أعلى نسبة تزامن"]) if pd.notna(g["أعلى نسبة تزامن"]) else "غير متاح"])
    wp.column_dimensions["A"].width = 56
    wp.column_dimensions["B"].width = 25
    wp.column_dimensions["C"].width = 78
    wp.column_dimensions["D"].width = 22
    for rr in (4, 12, 13, 14, 15, 16, 17, 18, 19, 20):
        if rr <= wp.max_row and isinstance(wp.cell(rr, 2).value, float):
            wp.cell(rr, 2).number_format = "0.0%"
    for row in wp.iter_rows():
        for c in row:
            c.alignment = Alignment(vertical="center", wrap_text=True)
    wp.freeze_panes = "B2"

    wpeople = sheet(wb, "تعارضات الموظفين")
    write_df(wpeople, people.sort_values(["أيام التعارض الفريدة", "عدد حالات التعارض"], ascending=False),
             1, widths={c: 25 for c in people.columns})
    wpeople.freeze_panes = "C2"
    if len(people):
        wpeople.auto_filter.ref = f"A1:G{len(people)+1}"

    wdaily = sheet(wb, "التزامن اليومي")
    write_df(wdaily, daily, 1, date_cols=("التاريخ",),
             widths={c: 27 for c in daily.columns})
    wdaily.freeze_panes = "C2"
    if len(daily):
        wdaily.auto_filter.ref = f"A1:D{len(daily)+1}"

    wgroups = sheet(wb, "مؤشرات التعارض حسب المنفذ")
    write_df(wgroups, group_stats, 1, widths={c: 29 for c in group_stats.columns})
    for row in range(2, len(group_stats)+2):
        wgroups.cell(row, 5).number_format = "0.0%"
    wgroups.freeze_panes = "C2"

    wb.save(OUTPUT_FILE)


def main():
    raw, cols = load()
    employees = load_employee_master()
    log = build_log(raw, cols)
    reconciled = reconcile_master(employees, log)
    audit_log = audit_conflicts_on_master(log, reconciled)
    conf = find_conflicts(audit_log)
    employee_summary = build_employee_summary(employees, log)
    build_report(raw, audit_log, conf, employee_summary, reconciled, employees)
    print("حالات المطابقة:")
    print(reconciled["حالة المطابقة"].value_counts().to_string())
    print(f"Done: {OUTPUT_FILE}")
    print(f"إجمالي سجلات الإجازات: {len(log)} | المعتمدة: {(log['حالة التدقيق']=='داخل التدقيق').sum()} | حالات التعارض: {len(conf)}")
    print(f"إجمالي الموظفين في ملف الموظفين: {employees['الرقم الوظيفي'].nunique()}")
    if len(conf):
        print(conf[["رقم الحالة", "المدير", "المنفذ", "من", "إلى", "أكبر عدد أشخاص في إجازة باليوم نفسه", "عدد الأشخاص", "نوع التعارض"]].to_string(index=False))


if __name__ == "__main__":
    main()

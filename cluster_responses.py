import re
import math
from collections import Counter, defaultdict
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment

INPUT_FILE = "responses.xlsx"
SHEET_NAME = 0
OUTPUT_FILE = "top_comments.xlsx"

POSITIVE_COLUMN = "key strengths"
NEGATIVE_COLUMN = "practices and procedures may not add tangible value"

MATCH_THRESHOLD = 0.42
DISCOVERY_THRESHOLD = 0.45
MIN_COUNT = 2
TOP_N = 30

POSITIVE_THEMES = {
    "المزايا المالية": [
        "مزايا مالية", "ميزة مالية", "ميزه ماليه", "المزايا الماليه", "ميزتي المالية", "مميزات مالية", "الراتب", "الرواتب",
        "راتب ممتاز", "الرواتب عالية", "رواتب مجزية", "البدلات", "بدل السكن", "بدل النقل", "المكافآت", "مكافأة نهاية السنة",
        "الحوافز", "الحوافز المالية", "العلاوة السنوية", "العلاوات", "الدخل الجيد", "التعويضات", "salary", "good salary",
        "financial benefits", "compensation", "bonus", "allowances",
    ],
    "بيئة عمل داعمة": [
        "بيئة داعمة", "بيئه داعمه", "بيئة عمل داعمة", "بيئة العمل", "بيئة عمل محفزة", "بيئة إيجابية", "بيئة صحية",
        "اجواء العمل", "أجواء مريحة", "جو العمل حلو", "بيئة مريحة", "بيئة مشجعة", "الدعم المستمر", "supportive environment",
        "work environment", "positive environment",
    ],
    "الزملاء وروح الفريق": [
        "الزملاء", "زملاء العمل", "الزملاء متعاونين", "تعاون الزملاء", "التعاون", "روح الفريق", "روح الفريق الواحد",
        "العمل الجماعي", "فريق العمل", "الفريق متعاون", "العلاقات بين الموظفين", "الألفة بين الموظفين", "teamwork",
        "team spirit", "colleagues", "collaboration",
    ],
    "القيادة والإدارة": [
        "القيادة", "القيادات", "الإدارة العليا", "المدير المباشر", "مديري متعاون", "دعم المدير", "قيادة ملهمة",
        "الإدارة متفهمة", "تفهم المسؤولين", "المسؤول يدعمنا", "سياسة الباب المفتوح", "قرب القيادة من الموظفين",
        "leadership", "management support", "good manager",
    ],
    "التدريب والتطوير": [
        "الدورات التدريبية", "الدورة التدريبية", "الدوره التدريبيه", "الدورات", "التدريب", "برامج التدريب",
        "التطوير المهني", "تطوير الموظفين", "فرص التعلم", "الابتعاث", "الشهادات المهنية", "ورش العمل", "التعليم المستمر",
        "training", "courses", "development programs", "learning opportunities",
    ],
    "البيئة المكانية والمكاتب": [
        "البيئة المكانية", "البيئه المكانيه", "المكاتب", "المكاتب حديثة", "المبنى", "المقر", "المرافق", "تجهيزات المكاتب",
        "المكان مريح", "التصميم الداخلي", "الكافتيريا", "النادي الرياضي", "مصليات", "office", "offices", "facilities", "workplace",
    ],
    "الاستقرار والأمان الوظيفي": [
        "الاستقرار الوظيفي", "الأمان الوظيفي", "الامان الوظيفي", "الاستقرار", "ثبات الوظيفة", "الوظيفة مضمونة",
        "استقرار وظيفي", "أمان وظيفي", "job security", "stability",
    ],
    "المرونة وساعات العمل": [
        "المرونة", "مرونة الدوام", "مرونه الدوام", "ساعات عمل مرنة", "الدوام المرن", "العمل عن بعد", "العمل عن بعد أحيانا",
        "التوازن بين الحياة والعمل", "ساعات العمل مناسبة", "flexibility", "flexible hours", "remote work", "work life balance",
    ],
    "الإتقان والجودة": [
        "الاتقان", "إتقان", "إتقان العمل", "اتقان العمل", "الجودة", "جودة العمل", "الاحترافية", "الدقة", "الدقة في العمل",
        "الحرص على الجودة", "التميز", "quality", "professionalism",
    ],
    "سمعة ومكانة الجهة": [
        "سمعة الهيئة", "مكانة الجهة", "اسم الهيئة", "الفخر بالانتماء", "الانتماء", "الولاء", "رسالة الهيئة",
        "أهمية العمل الوطني", "خدمة الوطن", "العمل له أثر", "reputation", "prestige", "national impact",
    ],
    "التقنية والأنظمة الحديثة": [
        "الأنظمة الحديثة", "التقنية", "التحول الرقمي", "الأنظمة الإلكترونية", "الأجهزة الحديثة", "الأدوات التقنية",
        "الأتمتة", "سهولة الأنظمة", "technology", "digital systems", "modern tools",
    ],
    "التقدير والتحفيز": [
        "التقدير", "تقدير الموظف", "تقدير الجهود", "التحفيز", "الشكر والتقدير", "تكريم الموظفين", "شهادات الشكر",
        "الاعتراف بالإنجاز", "recognition", "appreciation", "motivation",
    ],
    "فرص الترقي والنمو": [
        "الترقيات", "فرص الترقي", "الترقية", "المسار الوظيفي", "التطور الوظيفي", "النمو الوظيفي", "فرص النمو",
        "career growth", "promotion", "career path",
    ],
    "التأمين والخدمات": [
        "التأمين الطبي", "تأمين طبي", "التأمين", "الخدمات الصحية", "خدمات الموظفين", "العيادة", "الخصومات",
        "medical insurance", "health insurance",
    ],
    "الإجازات": [
        "الإجازات", "الاجازات", "الإجازات السنوية", "إجازات كافية", "رصيد الإجازات", "الاجازه", "vacation", "annual leave",
    ],
    "الوضوح والتنظيم": [
        "الوضوح", "وضوح المهام", "وضوح الأهداف", "التنظيم", "الإجراءات واضحة", "الأدوار واضحة", "الحوكمة",
        "clear roles", "organization", "clarity",
    ],
}

NEGATIVE_THEMES = {
    "كاميرات الكتف": [
        "كاميرات الكتف", "كاميرا الكتف", "كامريات الكتف", "كميرات الكتف", "الكاميرات", "الكاميرا", "كاميرات الجسم",
        "لبس الكاميرا", "إلزامية الكاميرات", "الكاميرات الكتفية", "تشغيل الكاميرا", "body camera", "body cam",
        "shoulder camera", "cameras",
    ],
    "ضعف المزايا المالية": [
        "مزايا مالية", "المزايا المالية", "ميزه ماليه", "ضعف المزايا", "قلة الرواتب", "الرواتب قليلة", "الراتب ضعيف",
        "الرواتب", "عدم وجود بدلات", "البدلات", "وقف البدلات", "قلة الحوافز", "لا يوجد مكافآت", "تأخر المكافآت",
        "عدم صرف الحوافز", "العلاوة", "low salary", "no bonus", "financial benefits", "allowances",
    ],
    "التحيز وعدم العدالة": [
        "تحيز", "التحيز", "التحيز في التقييم", "المحاباة", "الواسطة", "الواسطات", "المحسوبية", "عدم العدالة",
        "عدم المساواة", "التمييز بين الموظفين", "الكيل بمكيالين", "التفرقة", "تفضيل البعض", "عدم الإنصاف",
        "bias", "favoritism", "unfair", "nepotism",
    ],
    "مواقف السيارات": [
        "مواقف السيارات", "مافي مواقف", "مافي مواقف سيارات", "عدم وجود مواقف", "قلة المواقف", "المواقف", "المواقف بعيدة",
        "زحمة المواقف", "صعوبة إيجاد موقف", "المواقف غير كافية", "موقف", "parking", "no parking", "parking lot",
    ],
    "إتقان": [
        "اتقان", "إتقان", "منصة اتقان", "نظام اتقان", "برنامج اتقان", "مبادرة اتقان", "متطلبات اتقان", "ملفات اتقان",
        "itqan",
    ],
    "البيروقراطية وطول الإجراءات": [
        "البيروقراطية", "البيروقراطيه", "الروتين", "الروتين الممل", "طول الإجراءات", "الإجراءات الطويلة", "تعقيد الإجراءات",
        "كثرة الاعتمادات", "كثرة التواقيع", "التواقيع", "سلسلة الموافقات", "الموافقات", "الورقيات", "المعاملات الورقية",
        "الدورة المستندية", "bureaucracy", "red tape", "approvals", "paperwork",
    ],
    "كثرة التقارير والنماذج": [
        "التقارير", "كثرة التقارير", "التقارير اليومية", "التقارير الأسبوعية", "التقارير المكررة", "النماذج", "كثرة النماذج",
        "تعبئة النماذج", "الاستبيانات", "كثرة الاستبيانات", "الإحصائيات", "reports", "forms", "surveys",
    ],
    "كثرة الاجتماعات": [
        "الاجتماعات", "كثرة الاجتماعات", "اجتماعات بدون فائدة", "اجتماعات طويلة", "الاجتماعات المتكررة", "اجتماعات غير مجدية",
        "meetings", "too many meetings",
    ],
    "البصمة والحضور": [
        "البصمة", "البصمه", "بصمة الحضور", "بصمة الانصراف", "نظام البصمة", "التشديد على الحضور", "الحضور والانصراف",
        "الخصم على التأخير", "مراقبة الحضور", "attendance", "fingerprint",
    ],
    "الأنظمة التقنية": [
        "الأنظمة", "النظام بطيء", "بطء الأنظمة", "الأنظمة قديمة", "تعطل الأنظمة", "كثرة الأنظمة", "تعدد الأنظمة",
        "المنصات", "البرامج", "الشبكة ضعيفة", "الأجهزة قديمة", "system", "slow system", "systems",
    ],
    "ضعف التواصل": [
        "ضعف التواصل", "التواصل", "عدم وضوح التعاميم", "التعاميم", "غياب المعلومة", "قلة التواصل بين الإدارات",
        "عدم الشفافية", "الشفافية", "communication", "lack of communication",
    ],
    "تقييم الأداء": [
        "تقييم الأداء", "التقييم", "معايير التقييم", "التقييم غير عادل", "الأهداف", "مؤشرات الأداء", "التوزيع الإجباري",
        "منحنى التقييم", "performance evaluation", "kpis", "appraisal",
    ],
    "ضغط العمل ونقص الموظفين": [
        "ضغط العمل", "كثرة المهام", "عبء العمل", "نقص الموظفين", "قلة الموظفين", "المهام الإضافية", "العمل خارج الدوام",
        "ساعات إضافية", "workload", "understaffing", "overtime",
    ],
    "الترقيات": [
        "الترقيات", "تأخر الترقيات", "عدم وجود ترقيات", "الترقية", "ثبات المرتبة", "المسار الوظيفي", "promotion", "promotions",
    ],
    "ضعف التدريب": [
        "ضعف التدريب", "قلة الدورات", "عدم وجود تدريب", "الدورات غير مفيدة", "التدريب", "الدورات الإلزامية", "training",
    ],
    "المركزية وتأخر القرارات": [
        "المركزية", "تأخر القرارات", "بطء اتخاذ القرار", "عدم التفويض", "قلة الصلاحيات", "الصلاحيات", "centralization",
        "slow decisions",
    ],
    "تكرار الطلبات والبيانات": [
        "تكرار الطلبات", "تكرار طلب البيانات", "الطلبات المتكررة", "إعادة العمل", "تكرار العمل", "ازدواجية العمل",
        "ازدواجية المهام", "duplication", "rework",
    ],
    "الدوام والمناوبات": [
        "الدوام", "ساعات الدوام", "المناوبات", "الشفتات", "نظام الورديات", "الدوام الطويل", "عدم المرونة", "shifts",
        "working hours",
    ],
}

STOPWORDS = {
    "في", "من", "على", "الى", "الي", "عن", "مع", "او", "ان", "انه", "هذا", "هذه", "ذلك", "تلك", "التي", "الذي",
    "هو", "هي", "كل", "بعض", "غير", "لا", "ما", "لم", "لن", "قد", "كان", "يكون", "جدا", "ايضا", "عند", "بين",
    "عدم", "وجود", "يوجد", "مافي", "مافيه", "فيه", "فيها", "كثره", "قله", "نقص", "كثير", "قليل", "اكثر", "اقل",
    "ضعف", "شي", "شيء", "لبس", "بدون", "مره", "مرة", "كذا", "يعني", "حلو", "زين", "ممتاز", "جيد", "the", "and", "or", "of", "to", "in", "a", "an", "is", "are",
    "for", "with", "not", "no", "very", "too", "many", "lack", "good",
}

EMPTY_ANSWERS = {"", "لا", "لا يوجد", "لايوجد", "لا شيء", "لا شي", "مافي", "ما فيه", "-", ".", "na", "n/a", "none",
                 "nothing", "nan", "0", "لا اعلم", "ما ادري", "لا تعليق", "no comment"}

DIACRITICS = re.compile(r"[ؗ-ًؚ-ْٰـ]")
PUNCT = re.compile(r"[^\w\s]|_")
SPLITTER = re.compile(r"[\n\r،,;؛\.\-•\*/\|]+|\s\d+\s*[\)\-\.]|\sو(?=ال)|\sand\s")
PREFIXES = ["وبال", "وال", "بال", "كال", "فال", "ولل", "لل", "ال"]
SUFFIXES = ["اتهم", "اتنا", "اتي", "اته", "ات", "تهم", "تنا", "تي", "ته", "هم", "نا", "ها", "ين", "ون", "يه", "ه", "ي"]


def normalize(text):
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return ""
    t = str(text).lower()
    t = DIACRITICS.sub("", t)
    t = re.sub("[إأآٱ]", "ا", t)
    t = t.replace("ة", "ه").replace("ى", "ي").replace("ؤ", "و").replace("ئ", "ي")
    t = PUNCT.sub(" ", t)
    t = re.sub(r"(.)\1{2,}", r"\1", t)
    return re.sub(r"\s+", " ", t).strip()


def stem(word):
    w = word
    if re.fullmatch(r"[a-z0-9]+", w):
        for s in ["ies", "es", "s", "ing", "ed"]:
            if w.endswith(s) and len(w) - len(s) >= 3:
                return w[: -len(s)]
        return w
    for p in PREFIXES:
        if w.startswith(p) and len(w) - len(p) >= 3:
            w = w[len(p):]
            break
    if w.startswith("و") and len(w) >= 5:
        w = w[1:]
    for s in SUFFIXES:
        need = 4 if s in ("ون", "ين") else 3
        if w.endswith(s) and len(w) - len(s) >= need:
            w = w[: -len(s)]
            break
    return w


def stems(text):
    return [stem(w) for w in normalize(text).split() if w not in STOPWORDS and len(w) > 1]


def features(text):
    f = Counter()
    for s in stems(text):
        f["w:" + s] += 1
        padded = f"<{s}>"
        for n in (2, 3, 4):
            for i in range(len(padded) - n + 1):
                f[padded[i:i + n]] += 1
    return f


class TfIdf:
    def __init__(self, documents):
        df = Counter()
        for d in documents:
            df.update(set(features(d)))
        n = len(documents)
        self.idf = {k: math.log((1 + n) / (1 + v)) + 1 for k, v in df.items()}

    def vector(self, text):
        f = features(text)
        v = {k: (1 + math.log(c)) * self.idf.get(k, 1.0) for k, c in f.items()}
        norm = math.sqrt(sum(x * x for x in v.values())) or 1.0
        return {k: x / norm for k, x in v.items()}


def cosine(a, b):
    if len(a) > len(b):
        a, b = b, a
    return sum(x * b.get(k, 0.0) for k, x in a.items())


def is_empty(text):
    return normalize(text) in {normalize(x) for x in EMPTY_ANSWERS}


def split_segments(text):
    parts = [p.strip() for p in SPLITTER.split(" " + str(text) + " ") if p and p.strip()]
    return [p for p in parts if stems(p)] or ([str(text).strip()] if stems(str(text)) else [])


def analyze(series, themes):
    responses = []
    for idx, raw in series.items():
        if is_empty(raw):
            continue
        segs = split_segments(raw)
        if segs:
            responses.append((idx, segs))
    examples = [(theme, ex) for theme, exs in themes.items() for ex in exs]
    model = TfIdf([ex for _, ex in examples] + [s for _, segs in responses for s in segs])
    ex_vecs = [(theme, model.vector(ex)) for theme, ex in examples]
    hits = defaultdict(set)
    unmatched = []
    for idx, segs in responses:
        for seg in segs:
            v = model.vector(seg)
            best = {}
            for theme, ev in ex_vecs:
                sc = cosine(v, ev)
                if sc > best.get(theme, 0):
                    best[theme] = sc
            theme, score = max(best.items(), key=lambda x: x[1]) if best else (None, 0)
            if score >= MATCH_THRESHOLD:
                hits[theme].add(idx)
            else:
                unmatched.append((idx, seg, v))
    clusters = []
    for idx, seg, v in unmatched:
        target = next((c for c in clusters if cosine(v, c["vec"]) >= DISCOVERY_THRESHOLD), None)
        if target is None:
            clusters.append({"vec": v, "ids": {idx}, "texts": Counter([seg])})
        else:
            target["ids"].add(idx)
            target["texts"][seg] += 1
    for c in clusters:
        if len(c["ids"]) >= MIN_COUNT:
            label = min(c["texts"].most_common(3), key=lambda x: (-x[1], len(x[0])))[0]
            hits[label] |= c["ids"]
    total = len(responses)
    rows = sorted(((t, len(ids)) for t, ids in hits.items() if len(ids) >= MIN_COUNT), key=lambda x: -x[1])[:TOP_N]
    covered = set().union(*(hits[t] for t, _ in rows)) if rows else set()
    out = pd.DataFrame(
        [{"#": i, "الموضوع": t, "العدد": c, "النسبة": c / total if total else 0} for i, (t, c) in enumerate(rows, 1)]
    )
    others = total - len(covered)
    out.loc[len(out)] = ["", "ردود أخرى غير متكررة", others, others / total if total else 0]
    out.loc[len(out)] = ["", "إجمالي الردود", total, 1.0]
    return out


def style_sheet(ws):
    ws.sheet_view.rightToLeft = True
    navy = PatternFill("solid", fgColor="152446")
    light = PatternFill("solid", fgColor="E8F5F7")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF", name="Somar")
        cell.fill = navy
        cell.alignment = Alignment(horizontal="center", vertical="center")
    widths = {"A": 6, "B": 42, "C": 12, "D": 12}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name="Somar")
            c.alignment = Alignment(horizontal="center" if c.column != 2 else "right", vertical="center")
        row[3].number_format = "0.0%"
    for c in ws[ws.max_row]:
        c.font = Font(bold=True, name="Somar")
        c.fill = light


def main():
    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)
    lookup = {normalize(c): c for c in df.columns}
    jobs = [(POSITIVE_COLUMN, POSITIVE_THEMES, "ابرز المميزات"), (NEGATIVE_COLUMN, NEGATIVE_THEMES, "ابرز السلبيات")]
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        for col_name, themes, sheet in jobs:
            real = lookup.get(normalize(col_name))
            if real is None:
                print(f"العمود غير موجود: {col_name}")
                continue
            result = analyze(df[real], themes)
            print(f"\n{sheet}")
            print(result.to_string(index=False))
            result.to_excel(writer, sheet_name=sheet, index=False)
            style_sheet(writer.sheets[sheet])
    print(f"\nتم الحفظ: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""
quality_report.py — the full HR answer-quality analysis in one run.

Input
  1. Your tickets file (the clean file).
  2. Your Excel of questions and ideal answers (two columns: Question | Ideal Answer).

What it does
  - Assigns every ticket to the question it matches best (or uses your own label column).
  - Splits each ideal answer into its elements and checks which ones each actual answer covers
    → Completeness 0–4 and quality level (high 4 / medium 2–3 / low 0–1).
  - Reopened tickets, repeat inquiries within 14 days, service-level (SLA) comparison.
  - Most-missed elements, per-question results, findings written in formal Arabic.

Output (in OUTPUT_FOLDER)
  - quality_dashboard.html : the interactive dashboard, filled with your numbers
  - quality_results.xlsx   : Summary + one sheet per question with every ticket and its score

Needs: pandas, openpyxl, scikit-learn. Runs fully offline.
"""

import json
import math
import re
import sys
import time
from datetime import date
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd

# =============================================================================
# >>> 1. YOUR FILES <<<
# =============================================================================
TICKETS_FILE = r"C:\PUT\YOUR\PATH\HERE\clean\tickets_clean.csv"
QA_FILE = r"C:\PUT\YOUR\PATH\HERE\questions_and_answers.xlsx"      # Question | Ideal Answer
OUTPUT_FOLDER = r""          # empty = a "quality_report" folder next to the tickets file

# If your tickets file ALREADY says which question each ticket belongs to,
# write that column's name here (its values must be the same text as in the
# Question column). Leave "" to find the matching tickets automatically.
LABEL_COLUMN = ""

# Real examples (one low, one high) in the methodology section of the dashboard.
# Long numbers and emails are masked. Set to False if the dashboard will be shared widely.
INCLUDE_EXAMPLES = True

# Questions with fewer tickets than this are marked "limited" and are not used
# for the highest/lowest findings.
MIN_TICKETS_FOR_FINDINGS = 20

# Column holding the service-level result (Met / Missed). "" = not used.
SLA_COLUMN = "SLA Resolution Status Log"

# =============================================================================
# 2. SETTINGS
# =============================================================================
# Inquiry matching: share of the question's key words that must be in the inquiry.
MIN_MATCH = 0.75

# An ideal-answer element counts as covered if this share of its key words
# appears in the actual answer, or at least ELEMENT_MIN_WORDS of them
# (so long sentences are not judged too strictly).
ELEMENT_COVERED_AT = 0.25
ELEMENT_MIN_WORDS = 3

# Typo tolerance when comparing two words (0–1).
TYPO_TOLERANCE = 0.80

# Repeat inquiry window (days).
REPEAT_DAYS = 14

# Column names in your tickets file (matched ignoring case/spaces).
INQUIRY_COLUMN = "Inquiry"
ANSWER_COLUMN = "Public Logs"
REOPEN_COLUMN = "REOPEN COUNT"
SUBMIT_COLUMN = "Submit Date"
RESOLVED_COLUMN = "Last Resolved Date"
# The column that identifies the employee (for repeat inquiries).
# Leave "" to detect it automatically; otherwise write the exact column name,
# e.g. EMPLOYEE_COLUMN = "Employee Number"
EMPLOYEE_COLUMN = "Employee Name"
EMPLOYEE_ID_HINTS = ["employee id", "employee no", "employee number", "emp id", "emp no",
                     "staff id", "staff no", "badge", "person id", "personnel",
                     "رقم الموظف", "الرقم الوظيفي", "رقم وظيفي"]
# Used only if no ID column exists (less reliable: names can repeat).
EMPLOYEE_NAME_HINTS = ["employee name", "اسم الموظف", "requester name", "full name"]

# Column names in your Q&A Excel (if not found, the first two columns are used).
QUESTION_HEADERS = ["question", "السؤال", "سؤال"]
IDEAL_HEADERS = ["ideal answer", "ideal", "الإجابة النموذجية", "الاجابة النموذجية", "الإجابة المثالية", "الاجابة المثالية", "الإجابة"]

# Words ignored when matching (they don't carry the topic).
STOPWORDS = {
    "في", "من", "على", "عن", "الى", "ان", "او", "و", "يا", "هل", "مع", "ما", "لا", "هي", "هو",
    "انا", "ابي", "ابغى", "اريد", "ارغب", "ارجو", "اود", "لو", "كيف", "متى", "وش", "ايش",
    "طلب", "بخصوص", "لدي", "عندي", "الله", "السلام", "عليكم", "ورحمه", "وبركاته", "الذي",
    "التي", "ذلك", "هذا", "هذه", "تم", "يتم", "سيتم", "قد", "كل", "اي", "ثم", "حيث", "عند",
    "يرجي", "نامل", "نود", "نفيدكم", "بان", "بانه", "لكم", "لك", "منكم", "قبل", "بعد",
    "طريقه", "كيفيه", "اجراءات", "الاستفسار", "استفسار", "شكرا", "تحيه", "طيبه", "تحياتي",
}
# Sentences in the ideal answer that are only courtesy, not content.
# (stopwords are compared after normalization, e.g. الى → الي)
COURTESY = ["السلام عليكم", "تحية طيبة", "تحيه طيبه", "مرحبا", "حياك الله", "تحياتي",
            "مع التحية", "شاكرين", "شكرا", "دمتم", "نتمنى لك", "عزيزي", "عزيزتي"]

# =============================================================================
# Normalization (same as find_similar_inquiries.py)
# =============================================================================
AR_DIACRITICS = re.compile(r"[\u0610-\u061A\u064B-\u065F\u0670\u06D6-\u06ED\u0640]")
INVISIBLE = re.compile(r"[\u200b-\u200f\u202a-\u202e\u2066-\u2069\ufeff\xa0]")
DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
PREFIXES = ["وبال", "وكال", "وفال", "ولل", "وال", "بال", "كال", "فال", "لل", "ال",
            "و", "ب", "ل", "ف", "ك"]
SUFFIXES = ["تهم", "تها", "تكم", "تنا", "ات", "ون", "ين", "ان", "ها", "هم", "هن",
            "كم", "كن", "نا", "تي", "ته", "تك", "ي", "ه", "ك"]
TIMESTAMP = re.compile(
    r"\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}[ T,]+\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM|am|pm|ص|م)?")


def _norm_set(ws):
    return {normalize_text(w) for w in ws}


def normalize_text(t) -> str:
    if t is None or (isinstance(t, float) and pd.isna(t)):
        return ""
    t = INVISIBLE.sub(" ", str(t)).translate(DIGITS).lower()
    t = AR_DIACRITICS.sub("", t)
    t = re.sub("[إأآٱ]", "ا", t)
    t = t.replace("ة", "ه").replace("ى", "ي").replace("ؤ", "و").replace("ئ", "ي")
    t = re.sub(r"(.)\1{2,}", r"\1", t)
    t = re.sub(r"[^\w\s]", " ", t)
    return re.sub(r"\s+", " ", t).strip()


STOPWORDS_N = None


def stem(w: str) -> str:
    for p in PREFIXES:
        if w.startswith(p) and len(w) - len(p) >= (3 if len(p) > 1 else 4):
            w = w[len(p):]
            break
    for s in SUFFIXES:
        if w.endswith(s) and len(w) - len(s) >= 3:
            w = w[:-len(s)]
            break
    return w


def words(t, drop_stop=False):
    out = []
    for w in normalize_text(t).split():
        if drop_stop and (w in _stop() or len(w) < 2):
            continue
        out.append(stem(w))
    return out


def _stop():
    global STOPWORDS_N
    if STOPWORDS_N is None:
        STOPWORDS_N = _norm_set(STOPWORDS)
    return STOPWORDS_N


def same_word(a: str, b: str) -> bool:
    if a == b:
        return True
    if min(len(a), len(b)) >= 3 and (a.startswith(b) or b.startswith(a)):
        return True
    if min(len(a), len(b)) >= 4 and abs(len(a) - len(b)) <= 2:
        return SequenceMatcher(None, a, b).ratio() >= TYPO_TOLERANCE
    return False


def split_elements(ideal: str):
    """Split the ideal answer into content elements (sentences / clauses)."""
    text = TIMESTAMP.sub(" ", str(ideal))
    parts = re.split(r"[\n.؛;!؟?•\-–]+|،\s*(?=و?(?:مع|خلال|عبر|من خلال|ويتم|وتتم|علما|علماً)\b)", text)
    elements = []
    for p in parts:
        p = p.strip(" ،,:")
        if not p:
            continue
        np_ = normalize_text(p)
        if any(np_.startswith(normalize_text(c)) for c in COURTESY) and len(words(p, True)) <= 2:
            continue
        kw = list(dict.fromkeys(words(p, drop_stop=True)))
        if len(kw) >= 2:
            elements.append((p, kw))
    # Ideal answer written as one long sentence: split it further at commas,
    # so a partly complete answer can get 1–3 instead of all-or-nothing.
    if len(elements) == 1:
        pieces = [p.strip(" ،,:") for p in re.split(r"[،,]+", elements[0][0]) if p.strip(" ،,:")]
        # A condition ("في حال…", "إذا…") belongs with the clause that follows it
        merged = []
        for p in pieces:
            if merged and re.match(r"^(في حال|في حاله|اذا|إذا|عند|حال)\b", merged[-1]) and \
                    not merged[-1].endswith("|"):
                merged[-1] = merged[-1] + "، " + p + "|"
            else:
                merged.append(p)
        merged = [m.rstrip("|") for m in merged]
        finer = []
        for p in merged:
            kw = list(dict.fromkeys(words(p, drop_stop=True)))
            if len(kw) >= 2:
                finer.append((p, kw))
        if len(finer) > 1:
            elements = finer
    return elements


NUMBER_WORDS = {"واحد": "1", "اثنين": "2", "اثنان": "2", "يومين": "2", "يومان": "2",
                "ثلاث": "3", "اربع": "4", "خمس": "5", "ست": "6", "سته": "6", "سبع": "7",
                "ثمان": "8", "تسع": "9", "عشر": "10"}


def numbers_in(text: str) -> set:
    """Numbers in a text, as digits; Arabic number words are converted (خمسة → 5)."""
    t = normalize_text(text)
    nums = set(re.findall(r"\b\d{1,3}\b", t))
    for w in t.split():
        w2 = re.sub(r"^(?:وال|بال|لل|ال|و|ب|ل)", "", w) if len(w) > 4 else w
        for word, digit in NUMBER_WORDS.items():
            if w2 == word or w2 == word + "ه" or w2 == word + "ين" or w2 == word + "ه":
                nums.add(digit)
    return nums


def prepare_elements(elements, question):
    """Keep only each element's DISTINCTIVE words: drop words that appear in the
    question or in several elements (topic words like إجازة / وفاة / أيام)."""
    q_words = set(words(question, drop_stop=True))
    counts = {}
    for _, kw in elements:
        for w in set(kw):
            counts[w] = counts.get(w, 0) + 1
    shared_at = 2 if len(elements) >= 2 else 99
    prepared = []
    for text, kw in elements:
        dist = [w for w in kw
                if counts[w] < shared_at and not any(same_word(w, q) for q in q_words)
                and not w.isdigit()]
        prepared.append({"text": text, "words": dist or [w for w in kw if not w.isdigit()],
                         "numbers": numbers_in(text)})
    return prepared


def element_covered(el, answer_words, answer_numbers) -> bool:
    if el["numbers"]:
        # An element with a number (5 days, 3 days) is covered only when the
        # same number appears in the answer — as a digit or a word (خمسة = 5).
        return bool(el["numbers"] & answer_numbers)
    found = sum(any(same_word(w, a) for a in answer_words) for w in el["words"])
    need = min(math.ceil(len(el["words"]) * ELEMENT_COVERED_AT), ELEMENT_MIN_WORDS)
    return found >= max(need, 1)


def completeness_points(share: float) -> int:
    if share >= 0.999: return 4
    if share >= 0.75:  return 3
    if share >= 0.50:  return 2
    if share > 0:      return 1
    return 0


# =============================================================================
# Helpers
# =============================================================================

def load(path: Path):
    s = path.suffix.lower()
    if s == ".parquet":
        return pd.read_parquet(path)
    if s == ".csv":
        return pd.read_csv(path, dtype=str, encoding="utf-8-sig")
    return pd.read_excel(path, dtype=str)


def find_col(df, name):
    key = re.sub(r"\s+", "", name).lower()
    return next((c for c in df.columns if re.sub(r"\s+", "", str(c)).lower() == key), None)


def find_col_any(df, names):
    for n in names:
        c = find_col(df, n)
        if c:
            return c
    return None


def safe_sheet(name, used):
    name = re.sub(r"[\\/?*\[\]:]", "-", name).strip()[:31] or "Q"
    base, i = name, 2
    while name in used:
        name = f"{base[:28]}_{i}"; i += 1
    used.add(name)
    return name


def excel_safe(v):
    return v[:32000] if isinstance(v, str) and len(v) > 32000 else v



# =============================================================================
# Analysis
# =============================================================================

def clean_answer_for_display(text: str) -> str:
    lines = []
    for line in str(text or "").split("\n"):
        if re.search(r"view access|public\s*\d", line, re.I):
            continue
        line = TIMESTAMP.sub("", line).strip()
        if line:
            lines.append(line)
    t = "\n".join(lines)
    t = re.sub(r"[\w.+-]+@[\w-]+\.[\w.-]+", "[بريد]", t)
    t = re.sub(r"\d{6,}", "[رقم]", t)
    return t[:600]


def band_of(score):
    if score is None or pd.isna(score):
        return None
    return "good" if score >= 4 else "ok" if score >= 2 else "weak"


def pct(x):
    return None if x is None or (isinstance(x, float) and math.isnan(x)) else int(round(x * 100))


def main():
    try:
        from sklearn.feature_extraction.text import TfidfVectorizer
    except ImportError:
        sys.exit("scikit-learn is missing. Run:  pip install scikit-learn")
    if "PUT\\YOUR\\PATH" in TICKETS_FILE + QA_FILE:
        sys.exit("Please put your file paths in TICKETS_FILE and QA_FILE at the top of the script.")

    t0 = time.time()
    tickets_path, qa_path = Path(TICKETS_FILE), Path(QA_FILE)
    out_dir = Path(OUTPUT_FOLDER) if OUTPUT_FOLDER else tickets_path.parent / "quality_report"
    out_dir.mkdir(parents=True, exist_ok=True)

    # ---------------- Q&A ----------------
    qa = load(qa_path).dropna(how="all")
    qcol = find_col_any(qa, QUESTION_HEADERS) or qa.columns[0]
    acol_qa = find_col_any(qa, IDEAL_HEADERS) or qa.columns[1]
    qa = qa[[qcol, acol_qa]].dropna(subset=[qcol, acol_qa])
    qa.columns = ["Question", "Ideal Answer"]
    qa["Question"] = qa["Question"].astype(str).str.strip()
    questions = qa["Question"].tolist()
    ideals = qa["Ideal Answer"].astype(str).tolist()
    print(f"Questions: {len(questions)}")

    # ---------------- Tickets ----------------
    print(f"Reading {tickets_path.name}…")
    df = load(tickets_path).reset_index(drop=True)
    icol, anscol = find_col(df, INQUIRY_COLUMN), find_col(df, ANSWER_COLUMN)
    if not icol or not anscol:
        sys.exit(f"Need columns '{INQUIRY_COLUMN}' and '{ANSWER_COLUMN}'. Found: {list(df.columns)}")
    print(f"  {len(df):,} tickets")

    reopen_col = find_col(df, REOPEN_COLUMN)
    sub_col, res_col = find_col(df, SUBMIT_COLUMN), find_col(df, RESOLVED_COLUMN)
    sla_col = find_col(df, SLA_COLUMN) if SLA_COLUMN else None
    if EMPLOYEE_COLUMN:
        emp_col = find_col(df, EMPLOYEE_COLUMN)
    else:
        emp_col = next((c for c in df.columns if any(h in str(c).lower() for h in EMPLOYEE_ID_HINTS)), None) \
            or next((c for c in df.columns if any(h in str(c).lower() for h in EMPLOYEE_NAME_HINTS)), None)
    print(f"  Employee column: {emp_col or 'NOT FOUND (repeat inquiries skipped)'}")
    print(f"  SLA column: {sla_col or 'not found (SLA section hidden)'}")
    if not emp_col:
        print("    Your columns: " + ", ".join(map(str, df.columns)))

    sub_dt = pd.to_datetime(df[sub_col], errors="coerce", dayfirst=True, format="mixed") if sub_col else pd.Series(pd.NaT, index=df.index)
    res_dt = pd.to_datetime(df[res_col], errors="coerce", dayfirst=True, format="mixed") if res_col else pd.Series(pd.NaT, index=df.index)

    # ---------------- 1. assign each ticket to ONE question ----------------
    print("Assigning tickets to questions…")
    assigned = pd.Series(-1, index=df.index)
    match_score = pd.Series(0.0, index=df.index)
    lcol = find_col(df, LABEL_COLUMN) if LABEL_COLUMN else None
    if LABEL_COLUMN and not lcol:
        print(f"  ! LABEL_COLUMN '{LABEL_COLUMN}' not found — matching automatically instead.")
    if lcol:
        qnorm = {normalize_text(q): i for i, q in enumerate(questions)}
        assigned = df[lcol].map(lambda v: qnorm.get(normalize_text(v), -1))
        match_score[assigned >= 0] = 1.0
    else:
        inq_words = [set(words(x)) for x in df[icol]]
        vocab = set().union(*inq_words)
        cache = {}

        def variants(w):
            if w not in cache:
                cache[w] = {v for v in vocab if same_word(w, v)}
            return cache[w]

        for qi, q in enumerate(questions):
            qkw = list(dict.fromkeys(words(q, drop_stop=True)))
            if not qkw:
                continue
            var = [variants(w) for w in qkw]
            sc = pd.Series([sum(not v.isdisjoint(iw) for v in var) / len(qkw) for iw in inq_words], index=df.index)
            better = (sc >= MIN_MATCH) & (sc > match_score)
            assigned[better] = qi
            match_score[better] = sc[better]

    # ---------------- 2–3. completeness per ticket ----------------
    print("Scoring answers…")
    ans_clean = [TIMESTAMP.sub(" ", "" if pd.isna(x) else str(x)) for x in df[anscol]]
    ans_norm = [normalize_text(x) for x in ans_clean]
    rows = {}
    q_elements = {}
    for qi, (q, ideal) in enumerate(zip(questions, ideals)):
        idx = assigned[assigned == qi].index
        elements = split_elements(ideal)
        prepared = prepare_elements(elements, q)
        q_elements[qi] = prepared
        if not len(idx):
            continue
        vec = TfidfVectorizer(analyzer="char_wb", ngram_range=(3, 5), sublinear_tf=True)
        X = vec.fit_transform([normalize_text(ideal)] + [ans_norm[i] for i in idx])
        sim = (X[1:] @ X[0].T).toarray().ravel()
        for k, i in enumerate(idx):
            aw, an = set(words(ans_clean[i])), numbers_in(ans_clean[i])
            hits = [element_covered(el, aw, an) for el in prepared]
            share = sum(hits) / len(hits) if hits else 0
            score = completeness_points(share) if hits and ans_norm[i] else 0
            rows[i] = {
                "Question": q, "Match %": round(match_score[i] * 100),
                "Completeness (0-4)": score, "Quality level": score,
                "Coverage %": round(share * 100),
                "Elements covered": " | ".join(el["text"] for el, h in zip(prepared, hits) if h),
                "Elements missing": " | ".join(el["text"] for el, h in zip(prepared, hits) if not h),
                "_hits": hits,
                "Similarity to ideal %": round(sim[k] * 100),
            }
    if not rows:
        sys.exit("No tickets matched any question. Check the Question wording or lower MIN_MATCH.")
    res = pd.DataFrame.from_dict(rows, orient="index")
    res["band"] = res["Completeness (0-4)"].map(band_of)
    res["Quality level"] = res["band"].map({"good": "مرتفعة", "ok": "متوسطة", "weak": "منخفضة"})

    # ---------------- 4. outcome ----------------
    if reopen_col:
        res["Reopened"] = pd.to_numeric(df.loc[res.index, reopen_col], errors="coerce").fillna(0).gt(0)
    else:
        res["Reopened"] = False
    if sla_col:
        s = df.loc[res.index, sla_col].astype(str).str.lower()
        met = (s.str.contains("met") & ~s.str.contains("miss")).astype(object)
        met[~s.str.contains("met|miss")] = None
        res["SLA met"] = met
    res["Repeat 14d"] = None
    if emp_col:
        tmp = pd.DataFrame({"emp": df.loc[res.index, emp_col].map(normalize_text),
                            "q": res["Question"], "sub": sub_dt[res.index], "res": res_dt[res.index]})
        rep = pd.Series(False, index=res.index)
        for _, g in tmp[tmp.emp.isin(["", "nan", "none"]) == False].groupby(["emp", "q"]):
            if len(g) < 2:
                continue
            for i, r in g.iterrows():
                if pd.isna(r.res):
                    continue
                gap = (g.drop(index=i)["sub"] - r.res).dt.total_seconds() / 86400
                rep[i] = bool(gap.between(0, REPEAT_DAYS).any())
        res["Repeat 14d"] = rep

    # ---------------- 5. aggregate for the dashboard ----------------
    n = len(res)
    bands = res["band"].value_counts(normalize=True)
    data = {
        "isSample": False,
        "generatedAt": date.today().strftime("%Y-%m-%d"),
        "period": "—",
        "tickets": n,
        "questions": int(res["Question"].nunique()),
        "avgCompleteness": round(float(res["Completeness (0-4)"].mean()), 2),
        "scoreDist": {str(k): pct((res["Completeness (0-4)"] == k).mean()) for k in range(5)},
        "bands": {k: pct(bands.get(k, 0)) for k in ("good", "ok", "weak")},
        "reopenRate": pct(res["Reopened"].mean()),
    }
    sd = sub_dt[res.index].dropna()
    if len(sd):
        data["period"] = f"{sd.min():%Y-%m-%d} إلى {sd.max():%Y-%m-%d}"

    if sla_col and res["SLA met"].notna().any():
        met = res[res["SLA met"] == True]
        mb = met["band"].value_counts(normalize=True)
        data["sla"] = {"available": len(met) > 0, "metPct": pct(res["SLA met"].dropna().astype(bool).mean()),
                       "bands": {k: pct(mb.get(k, 0)) for k in ("good", "ok", "weak")}}
    else:
        data["sla"] = {"available": False}

    if emp_col:
        rb = res.groupby("band")["Repeat 14d"].mean()
        data["repeat"] = {"available": True, "overall": pct(res["Repeat 14d"].mean()),
                          "byBand": {k: pct(rb.get(k, 0)) for k in ("good", "ok", "weak")}}
    else:
        data["repeat"] = {"available": False}

    by_type = []
    for q in questions:
        g = res[res["Question"] == q]
        if not len(g):
            continue
        by_type.append({
            "name": q, "n": len(g),
            "completeness": round(float(g["Completeness (0-4)"].mean()), 2),
            "highPct": pct((g["band"] == "good").mean()),
            "reopen": pct(g["Reopened"].mean()),
            "repeat": pct(g["Repeat 14d"].mean()) if emp_col else None,
            "small": len(g) < MIN_TICKETS_FOR_FINDINGS,
        })
    data["byType"] = by_type

    missing = []
    for qi, q in enumerate(questions):
        g = res[res["Question"] == q]
        if len(g) < max(5, MIN_TICKETS_FOR_FINDINGS // 2) or not q_elements[qi]:
            continue
        hits = pd.DataFrame(g["_hits"].tolist())
        for j, el in enumerate(q_elements[qi]):
            missing.append({"question": q, "element": el["text"], "pct": pct(1 - hits[j].mean()), "n": len(g)})
    missing.sort(key=lambda m: (-m["pct"], -m["n"]))
    data["missing"] = [{k: m[k] for k in ("question", "element", "pct")} for m in missing[:8]]

    # ---------------- findings (formal Arabic) ----------------
    ins = []
    if data["sla"]["available"]:
        ins.append(f"صُنِّفت <b>{data['sla']['bands']['weak']}%</b> من الإجابات الملتزمة باتفاقية مستوى الخدمة ضمن "
                   f"مستوى الجودة المنخفضة؛ ما يؤكد أن الالتزام بالوقت لا يعكس بالضرورة اكتمال الإجابة.")
    if data["repeat"]["available"]:
        ins.append(f"لم تتجاوز نسبة إعادة فتح الاستفسارات <b>{data['reopenRate']}%</b>، في حين بلغت نسبة تكرار "
                   f"الاستفسار ذاته خلال أربعة عشر يوماً <b>{data['repeat']['overall']}%</b>؛ ما يشير إلى أن مؤشر "
                   f"إعادة الفتح لا يعكس حجم الاستفسارات غير المحلولة.")
        rb = data["repeat"]["byBand"]
        if rb["weak"] is not None and rb["good"] is not None:
            if rb["weak"] > rb["good"]:
                ins.append(f"بلغت نسبة تكرار الاستفسار <b>{rb['weak']}%</b> في الإجابات منخفضة الجودة، مقابل "
                           f"<b>{rb['good']}%</b> في الإجابات مرتفعة الجودة.")
            else:
                ins.append(f"لم يُظهر تكرار الاستفسار ارتباطاً واضحاً بمستوى الجودة "
                           f"(<b>{rb['weak']}%</b> للإجابات منخفضة الجودة مقابل <b>{rb['good']}%</b> للمرتفعة).")
    ins.append(f"لم تتضمن <b>{data['scoreDist']['0']}%</b> من الإجابات أي عنصر من عناصر الإجابة المرجعية، "
               f"في حين استوفت <b>{data['scoreDist']['4']}%</b> منها جميع العناصر.")
    big = [t for t in by_type if not t["small"]]
    if len(big) >= 2:
        lo, hi = min(big, key=lambda t: t["completeness"]), max(big, key=lambda t: t["completeness"])
        ins.append(f"سجّل استفسار «<b>{lo['name']}</b>» أدنى متوسط للاكتمال ({lo['completeness']} من 4)، "
                   f"في حين سجّل استفسار «<b>{hi['name']}</b>» أعلاه ({hi['completeness']} من 4).")
    if data["missing"]:
        m = data["missing"][0]
        ins.append(f"يُعد عنصر «<b>{m['element'][:80]}</b>» الأكثر غياباً عن إجابات استفسار «{m['question']}»، "
                   f"إذ لم يرد في <b>{m['pct']}%</b> منها.")
    data["insights"] = ins

    # ---------------- examples ----------------
    data["examples"] = None
    if INCLUDE_EXAMPLES and big:
        main_q = max(big, key=lambda t: t["n"])["name"]
        g = res[res["Question"] == main_q]
        lens = pd.Series([len(ans_clean[i]) for i in g.index], index=g.index)

        def pick(mask):
            cand = g[mask & lens.between(10, 500)]
            if not len(cand):
                return None
            i = lens[cand.index].sort_values().index[len(cand) // 2]
            r = res.loc[i]
            return {"question": main_q, "score": int(r["Completeness (0-4)"]),
                    "inquiry": clean_answer_for_display(df.at[i, icol])[:300],
                    "answer": clean_answer_for_display(df.at[i, anscol]),
                    "covered": [x for x in r["Elements covered"].split(" | ") if x],
                    "missing": [x for x in r["Elements missing"].split(" | ") if x]}
        low, high = pick(g["Completeness (0-4)"] <= 1), pick(g["Completeness (0-4)"] == 4)
        if low and high:
            data["examples"] = {"low": low, "high": high}

    # ---------------- write dashboard ----------------
    html = TEMPLATE.replace("/*__DATA__*/null", json.dumps(data, ensure_ascii=False))
    dash_path = out_dir / "quality_dashboard.html"
    dash_path.write_text(html, encoding="utf-8")

    # ---------------- write Excel ----------------
    xl_path = out_dir / "quality_results.xlsx"
    show = res.drop(columns=["_hits", "band"]).copy()
    for c in ("Reopened", "SLA met", "Repeat 14d"):
        if c in show:
            show[c] = show[c].map({True: "Yes", False: "No"})
    if not emp_col:
        show = show.drop(columns=["Repeat 14d"])
    full = pd.concat([show, df.loc[show.index]], axis=1).map(excel_safe)
    summary = pd.DataFrame([{
        "Question": t["name"], "Tickets": t["n"], "Avg completeness (0-4)": t["completeness"],
        "High quality %": t["highPct"], "Reopened %": t["reopen"],
        **({"Repeat 14d %": t["repeat"]} if emp_col else {}),
        "Limited sample": "Yes" if t["small"] else "No"} for t in by_type])
    with pd.ExcelWriter(xl_path, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="Summary", index=False)
        pd.DataFrame([{"Question": q, "Elements found in the ideal answer":
                       " | ".join(e["text"] for e in q_elements[i])} for i, q in enumerate(questions)]) \
            .to_excel(xw, sheet_name="Elements", index=False)
        used = set(["Summary", "Elements"])
        for i, q in enumerate(questions):
            part = full[full["Question"] == q].sort_values("Completeness (0-4)", ascending=False, kind="stable")
            if not len(part):
                continue
            sheet = safe_sheet(f"{i + 1}. {q}", used)
            part.to_excel(xw, sheet_name=sheet, index=False)
            ws = xw.sheets[sheet]
            ws.freeze_panes = "B2"
            ws.auto_filter.ref = ws.dimensions

    unmatched = int((assigned < 0).sum())
    print(f"\n  Tickets scored: {n:,}  (not matched to any question: {unmatched:,})")
    for t in by_type:
        print(f"  {t['name'][:45]:45s}  {t['n']:>6,} tickets  completeness {t['completeness']}"
              + ("   (limited sample)" if t["small"] else ""))
    print(f"\nSaved:\n  {dash_path}\n  {xl_path}\n({time.time() - t0:.0f}s)")



# =============================================================================
# Dashboard template (do not edit)
# =============================================================================
TEMPLATE = r'''<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>إجابتك أقرب</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Arabic:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<style>
  :root {
    --navy: #092751;
    --navy-2: #0d3366;
    --teal: #006358;
    --green: #69BD4B;
    --good: #69BD4B;
    --ok: #4BB8C4;
    --weak: #E8A33D;
    --review: #9AA9BC;
    --glass: rgba(255, 255, 255, 0.10);
    --glass-2: rgba(255, 255, 255, 0.16);
    --line: rgba(255, 255, 255, 0.22);
    --text: #ffffff;
    --soft: #D3E3EF;
    --muted: #A9C1D6;
  }
  * { box-sizing: border-box; }
  html { scroll-behavior: smooth; }
  html, body { margin: 0; }
  body {
    font-family: 'IBM Plex Sans Arabic', 'Segoe UI', Tahoma, sans-serif;
    color: var(--text);
    background: var(--navy);
  }

  /* ---------- Background (same as إجابتك أقرب) ---------- */
  .bg {
    position: fixed; inset: 0; z-index: 0;
    background: linear-gradient(160deg, #69BD4B 0%, #006358 48%, #092751 100%);
  }
  .snow { position: fixed; inset: 0; pointer-events: none; z-index: 1; }
  .orbit { position: absolute; width: 0; height: 0;
    animation-name: orbit; animation-timing-function: linear; animation-iteration-count: infinite; }
  .flake { display: block; border-radius: 50%; background: #fff; }
  @keyframes orbit { to { transform: rotate(360deg); } }
  .vignette { position: fixed; inset: 0; pointer-events: none; z-index: 1;
    background: radial-gradient(ellipse at 50% 40%, rgba(9,39,81,0) 30%, rgba(9,39,81,.45) 100%); }

  .wrap { position: relative; z-index: 2; }

  /* ---------- Hero ---------- */
  .hero {
    min-height: 100vh; display: flex; flex-direction: column;
    align-items: center; justify-content: center; text-align: center;
    padding: 48px clamp(16px, 6vw, 96px);
  }
  .badge {
    display: inline-flex; align-items: center; gap: 8px;
    padding: 8px 18px; border-radius: 999px;
    border: 1px solid var(--line); background: var(--glass);
    -webkit-backdrop-filter: blur(12px); backdrop-filter: blur(12px);
    font-size: 15px; font-weight: 500; color: #EAF5EE; margin-bottom: 26px;
  }
  .badge .dot { width: 8px; height: 8px; border-radius: 50%; background: var(--green); box-shadow: 0 0 0 4px rgba(105,189,75,.25); }
  h1 { margin: 0; font-size: clamp(56px, 9vw, 128px); font-weight: 600; line-height: 1.05; }
  .edition {
    margin: 18px 0 0; font-size: clamp(20px, 2.4vw, 32px); font-weight: 400;
    color: #EAF5EE; letter-spacing: .3px;
  }
  .slogan { margin: 16px 0 0; font-size: clamp(17px, 1.8vw, 24px); font-weight: 300; color: #EAF5EE; }

  .cta {
    margin-top: 54px; display: inline-flex; align-items: center; gap: 14px;
    min-height: 60px; padding: 0 40px; border-radius: 999px;
    border: 1px solid rgba(255,255,255,.45); background: rgba(255,255,255,.16);
    -webkit-backdrop-filter: blur(16px) saturate(140%); backdrop-filter: blur(16px) saturate(140%);
    box-shadow: inset 0 1px 0 rgba(255,255,255,.5), 0 10px 30px rgba(9,39,81,.25);
    color: #fff; font: inherit; font-size: 20px; font-weight: 500; text-decoration: none; cursor: pointer;
    transition: background .35s, border-color .35s, transform .35s, box-shadow .35s;
  }
  .cta:hover { background: rgba(255,255,255,.26); border-color: rgba(255,255,255,.7); transform: translateY(-2px); }
  .cta:focus-visible, button:focus-visible, input:focus-visible { outline: 2px solid #fff; outline-offset: 4px; }
  .cta svg { width: 22px; height: 22px; transition: transform .35s; }
  .cta:hover svg { transform: translateY(3px); }

  /* ---------- App ---------- */
  .app { max-width: 1200px; margin: 0 auto; padding: 32px 16px 80px; min-height: 100vh; }
  .app-head { text-align: center; margin-bottom: 26px; }
  .app-head h2 { margin: 0; font-size: clamp(26px, 3.2vw, 40px); font-weight: 600; }
  .app-head p { margin: 8px 0 0; color: var(--soft); font-size: 16px; }

  /* Matte navy surface with fine grain */
  .matte {
    background-color: #0A2347;
    background-image:
      linear-gradient(180deg, rgba(255,255,255,.06), rgba(255,255,255,0) 45%),
      url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='160' height='160'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='.9' numOctaves='2' stitchTiles='stitch'/%3E%3CfeColorMatrix values='0 0 0 0 1 0 0 0 0 1 0 0 0 0 1 0 0 0 .07 0'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)'/%3E%3C/svg%3E");
  }


  /* Ask bar: elliptical navy pill with a light running around it */
  .ask-ring {
    position: relative; max-width: 820px; margin: 0 auto; padding: 1.5px;
    border-radius: 999px; overflow: hidden; overflow: clip; isolation: isolate;
    box-shadow: 0 0 0 1px rgba(255,255,255,.10), 0 22px 55px rgba(4,18,40,.45), 0 0 40px rgba(75,184,196,.12);
  }
  .ask-ring::before {
    content: ""; position: absolute; z-index: -1;
    left: 50%; top: 50%; width: 140%; aspect-ratio: 1; translate: -50% -50%;
    background: conic-gradient(from 0deg, rgba(255,255,255,.08) 0deg, rgba(255,255,255,.08) 250deg,
      #4BB8C4 290deg, #ffffff 320deg, #69BD4B 345deg, rgba(255,255,255,.08) 360deg);
    animation: spin 6s linear infinite;
  }
  @keyframes spin { to { rotate: 360deg; } }
  .ask {
    display: flex; align-items: center; gap: 10px;
    min-height: 68px; padding: 8px; border-radius: 999px;
    box-shadow: inset 0 1px 0 rgba(255,255,255,.12), inset 0 -10px 24px rgba(0,0,0,.18);
  }
  .ask svg.lead { width: 22px; height: 22px; margin-inline-start: 18px; flex: none; opacity: .8; }
  .ask input {
    flex: 1; min-width: 0; border: 0; background: transparent; color: #fff;
    font: inherit; font-size: 18px; padding: 12px 6px;
  }
  .ask input::placeholder { color: rgba(255,255,255,.6); }
  .ask input:focus { outline: none; }
  .ask button {
    flex: none; display: inline-flex; align-items: center; justify-content: center; gap: 10px;
    height: 52px; padding: 0 26px; border-radius: 999px; border: 0; cursor: pointer;
    background: #fff; color: #0A2347; font: inherit; font-size: 16px; font-weight: 600;
    box-shadow: 0 6px 18px rgba(0,0,0,.25);
    transition: transform .2s, box-shadow .2s;
  }
  .ask button svg { width: 20px; height: 20px; }
  .ask button:hover { transform: translateY(-1px); box-shadow: 0 10px 24px rgba(0,0,0,.3); }

  .pick-label {
    display: flex; align-items: center; justify-content: center; gap: 14px;
    margin: 26px 0 14px; color: #fff; font-size: 16px;
  }
  .pick-label b { font-weight: 600; letter-spacing: .3px; }
  .pick-label span { width: 70px; height: 1px; background: linear-gradient(90deg, transparent, rgba(255,255,255,.55), transparent); }
  .chips { display: flex; flex-wrap: wrap; gap: 10px; justify-content: center; max-width: 960px; margin: 0 auto; }
  .chip {
    border: 1px solid rgba(255,255,255,.18); color: #fff;
    background-color: rgba(10,35,71,.72);
    box-shadow: inset 0 1px 0 rgba(255,255,255,.10), 0 6px 18px rgba(4,18,40,.25);
    -webkit-backdrop-filter: blur(12px); backdrop-filter: blur(12px);
    border-radius: 999px; padding: 11px 18px; font: inherit; font-size: 15px; cursor: pointer;
    transition: background .25s, border-color .25s, transform .25s;
  }
  .chip:hover { background-color: rgba(14,46,92,.85); border-color: rgba(255,255,255,.45); transform: translateY(-1px); }
  .chip.active { background: #fff; color: #0A2347; border-color: #fff; font-weight: 600;
    box-shadow: 0 0 0 4px rgba(255,255,255,.12), 0 8px 22px rgba(4,18,40,.3); }
  .chip.special { border-color: rgba(205,239,192,.6); }

  .meta-line { margin-top: 6px !important; font-size: 13px !important; color: var(--muted) !important; }
  .vs { display: grid; grid-template-columns: 1fr auto 1fr; gap: 16px; align-items: center; text-align: center; }
  .vs-item { padding: 18px 12px; border-radius: 18px; background: rgba(255,255,255,.05); border: 1px solid rgba(255,255,255,.12); }
  .vs-item .v { font-size: clamp(34px, 4vw, 48px); font-weight: 600; line-height: 1.1; }
  .vs-item .l { margin-top: 6px; color: var(--soft); font-size: 14px; }
  .vs-item.hl { border-color: rgba(232,163,61,.55); }
  .vs-item.hl .v { color: #FFD89C; }
  .vs-mid { color: var(--muted); font-size: 14px; }

  /* ---- Dashboard layout ---- */
  .bar-dock { position: sticky; top: 12px; z-index: 30; max-width: 820px; margin: 0 auto; }
  .ask-ring.typing::before { animation-duration: 1.4s; }
  .ask-ring.typing { box-shadow: 0 0 0 1px rgba(255,255,255,.18), 0 22px 55px rgba(4,18,40,.45), 0 0 46px rgba(75,184,196,.35); }
  .suggest-pop { margin-top: 10px; padding: 16px; border-radius: 22px; border: 1px solid var(--line);
    background-color: rgba(10,35,71,.95); box-shadow: 0 18px 40px rgba(4,18,40,.4); text-align: center; }
  .suggest-pop p { margin: 0 0 12px; font-size: 15px; color: var(--soft); }
  .suggest-pop div { display: flex; flex-wrap: wrap; gap: 8px; justify-content: center; }

  .kpi-strip { display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 14px; margin: 30px 0 18px; }
  .kpi-tile { border-radius: 20px; border: 1px solid rgba(255,255,255,.16); padding: 18px 20px;
    box-shadow: inset 0 1px 0 rgba(255,255,255,.08), 0 10px 26px rgba(4,18,40,.25); }
  .kpi-tile .v { font-size: clamp(28px, 3vw, 38px); font-weight: 600; line-height: 1.1; }
  .kpi-tile .v small { font-size: .5em; color: var(--muted); font-weight: 400; margin-inline-start: 2px; }
  .kpi-tile .l { margin-top: 6px; color: var(--soft); font-size: 13.5px; line-height: 1.5; }
  .kpi-tile.hl { border-color: rgba(232,163,61,.55); box-shadow: inset 0 0 0 1px rgba(232,163,61,.2), 0 10px 26px rgba(4,18,40,.25); }
  .kpi-tile.hl .v { color: #FFD89C; }
  @media (max-width: 900px) { .kpi-strip { grid-template-columns: repeat(2, minmax(0, 1fr)); } .kpi-tile.hl { grid-column: 1 / -1; } }

  .dash { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 18px; }
  .dash-card { position: relative; border-radius: 26px; border: 1px solid var(--line);
    background-color: rgba(10,35,71,.84); padding: clamp(18px, 2.4vw, 28px);
    box-shadow: 0 20px 50px rgba(4,18,40,.3); scroll-margin-top: 120px; min-width: 0; }
  .dash-card::before { content: ""; position: absolute; top: 0; left: 14%; right: 14%; height: 1px;
    background: linear-gradient(90deg, transparent, #4BB8C4, #69BD4B, transparent); opacity: .7; }
  .dash-card.span-full { grid-column: 1 / -1; }
  .dash-card header { display: flex; align-items: flex-start; gap: 12px; margin-bottom: 10px; }
  .dash-card header h3 { margin: 2px 0 0; font-size: clamp(17px, 1.7vw, 21px); font-weight: 600; line-height: 1.5; }
  .q-num { flex: none; width: 30px; height: 30px; border-radius: 50%; display: grid; place-items: center;
    background: rgba(255,255,255,.12); border: 1px solid rgba(255,255,255,.3); font-size: 13px; font-weight: 600; }
  .dash-card .lead-text { font-size: 15.5px; margin-bottom: 18px; }
  .dash-card:not(.span-full) .g2 { grid-template-columns: 1fr; }
  .dash-card.flash { animation: flash 1.8s ease; }
  @keyframes flash {
    0% { box-shadow: 0 0 0 0 rgba(75,184,196,0), 0 20px 50px rgba(4,18,40,.3); border-color: var(--line); }
    25% { box-shadow: 0 0 0 4px rgba(75,184,196,.45), 0 0 50px rgba(75,184,196,.35); border-color: #4BB8C4; }
    100% { box-shadow: 0 0 0 0 rgba(75,184,196,0), 0 20px 50px rgba(4,18,40,.3); border-color: var(--line); }
  }
  @media (max-width: 900px) { .dash { grid-template-columns: 1fr; } }

  /* Answer panel */
  .panel {
    margin: 30px auto 0; max-width: 1000px;
    border-radius: 28px; border: 1px solid var(--line);
    background-color: rgba(10, 35, 71, 0.82);
    -webkit-backdrop-filter: blur(20px) saturate(130%); backdrop-filter: blur(20px) saturate(130%);
    position: relative;
    box-shadow: 0 24px 60px rgba(4, 18, 40, .35);
    padding: clamp(20px, 3vw, 36px);
    animation: rise .45s ease both;
  }
  .panel::before { content: ""; position: absolute; top: 0; left: 12%; right: 12%; height: 1px;
    background: linear-gradient(90deg, transparent, #4BB8C4, #69BD4B, transparent); opacity: .8; }
  @keyframes rise { from { opacity: 0; transform: translateY(12px); } to { opacity: 1; transform: none; } }
  .panel .q { display: flex; align-items: flex-start; gap: 12px; margin-bottom: 18px; }
  .panel .q .ic { flex: none; width: 36px; height: 36px; border-radius: 50%; display: grid; place-items: center;
    background: rgba(105,189,75,.18); border: 1px solid rgba(105,189,75,.5); }
  .panel .q h3 { margin: 4px 0 0; font-size: clamp(20px, 2.2vw, 26px); font-weight: 600; line-height: 1.4; }
  .panel .lead-text { margin: 0 0 22px; color: var(--soft); font-size: 17px; line-height: 1.8; }
  .panel .lead-text strong { color: #fff; }
  .suggest-title { margin: 0 0 16px; text-align: center; font-size: 18px; font-weight: 500; color: #fff; }
  .suggest { display: flex; flex-wrap: wrap; gap: 10px; justify-content: center; }
  .empty { text-align: center; color: var(--soft); padding: 26px 8px; }

  .grid { display: grid; gap: 16px; }
  .g2 { grid-template-columns: repeat(2, minmax(0, 1fr)); }
  .g3 { grid-template-columns: repeat(3, minmax(0, 1fr)); }
  .g4 { grid-template-columns: repeat(4, minmax(0, 1fr)); }
  @media (max-width: 820px) { .g2, .g3, .g4 { grid-template-columns: 1fr; } }
  @media (max-width: 820px) { .g4 { grid-template-columns: repeat(2, minmax(0, 1fr)); } }

  .card {
    border-radius: 20px; border: 1px solid rgba(255,255,255,.14);
    background: rgba(255,255,255,.06); padding: 20px;
  }
  .card h4 { margin: 0 0 14px; font-size: 16px; font-weight: 600; color: #fff; }
  .card .sub { margin: -8px 0 14px; font-size: 13px; color: var(--muted); }

  .kpi .v { font-size: clamp(34px, 4vw, 46px); font-weight: 600; line-height: 1.1; }
  .kpi .v small { font-size: .45em; color: var(--muted); font-weight: 400; }
  .kpi .l { margin-top: 6px; color: var(--soft); font-size: 14px; }

  /* SVG chart text */
  svg text { font-family: 'IBM Plex Sans Arabic', 'Segoe UI', Tahoma, sans-serif; }
  .chart { width: 100%; height: auto; display: block; }

  .hbars { display: grid; gap: 14px; }
  .hb-row { display: grid; grid-template-columns: minmax(120px, 34%) 1fr 72px; gap: 14px; align-items: center; }
  .hb-label { font-size: 15px; color: #EAF5EE; line-height: 1.4; }
  .hb-track { height: 20px; border-radius: 999px; background: rgba(255,255,255,.08); overflow: hidden; }
  .hb-fill { height: 100%; border-radius: 999px; }
  .hb-val { font-weight: 600; font-size: 15px; direction: ltr; text-align: left; white-space: nowrap; }
  @media (max-width: 520px) {
    .hb-row { grid-template-columns: 1fr 64px; gap: 6px 10px; }
    .hb-label { grid-column: 1 / -1; }
  }

  /* ---- Unusual charts ---- */
  .draw { animation: draw 1.2s cubic-bezier(.2,.8,.2,1) both; }
  @keyframes draw { from { stroke-dashoffset: var(--len); } to { stroke-dashoffset: 0; } }
  .pop { transform-box: fill-box; transform-origin: center; animation: pop .6s cubic-bezier(.2,.9,.3,1.3) both; }
  @keyframes pop { from { transform: scale(0); opacity: 0; } to { transform: scale(1); } }

  .rings-wrap { display: grid; gap: 18px; }
  .ring-legend { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px 18px; }
  .ring-legend div { display: grid; grid-template-columns: 12px 1fr auto; gap: 10px; align-items: center; font-size: 15px; }
  @media (max-width: 520px) { .ring-legend { grid-template-columns: 1fr; } }
  .ring-legend i { width: 12px; height: 12px; border-radius: 50%; }
  .ring-legend b { font-weight: 600; white-space: nowrap; }
  .ring-legend small { color: var(--muted); font-weight: 400; }

  .waffle { display: grid; grid-template-columns: repeat(10, 1fr); gap: 6px; max-width: 300px; margin: 0 auto; direction: rtl; }
  .waffle span { aspect-ratio: 1; border-radius: 50%; animation: pop .4s ease both; }

  .lolli { display: grid; gap: 22px; padding: 6px 0; }
  .lo-row { display: grid; grid-template-columns: minmax(130px, 32%) 1fr; gap: 18px; align-items: center; }
  .lo-label { font-size: 15px; color: #EAF5EE; line-height: 1.4; }
  .lo-label small { display: block; color: var(--muted); font-size: 12.5px; }
  .lo-track { position: relative; height: 38px; margin-inline-start: 20px; }
  .lo-track::before { content: ""; position: absolute; inset: 50% 0 auto 0; height: 1px; background: rgba(255,255,255,.10); }
  .lo-stick { position: absolute; right: 0; top: calc(50% - 1.5px); height: 3px; border-radius: 3px;
    background: linear-gradient(270deg, color-mix(in srgb, var(--c) 20%, transparent), var(--c)); transform-origin: right; animation: grow .9s ease both; }
  @keyframes grow { from { transform: scaleX(0); } }
  .lo-dot { position: absolute; top: 0; width: 38px; height: 38px; border-radius: 50%; display: grid; place-items: center;
    background: var(--c); color: #0A2347; font-weight: 700; font-size: 13px;
    box-shadow: 0 0 0 5px color-mix(in srgb, var(--c) 25%, transparent), 0 6px 16px rgba(0,0,0,.3);
    animation: pop .5s .6s cubic-bezier(.2,.9,.3,1.3) both; }
  @media (max-width: 520px) { .lo-row { grid-template-columns: 1fr; gap: 6px; } .lo-track { margin-inline-start: 22px; } }
  .legend { display: flex; flex-wrap: wrap; gap: 8px 16px; margin-top: 12px; font-size: 13px; color: var(--soft); }
  .legend i { display: inline-block; width: 10px; height: 10px; border-radius: 3px; margin-inline-end: 6px; vertical-align: middle; }

  /* Insights */
  .insight { display: flex; gap: 14px; align-items: flex-start; }
  .insight .n { flex: none; width: 34px; height: 34px; border-radius: 10px; display: grid; place-items: center;
    background: rgba(255,255,255,.12); font-weight: 600; }
  .insight p { margin: 0; line-height: 1.75; color: var(--soft); font-size: 15px; }
  .insight p b { color: #fff; }

  /* Rubric */
  .rubric-row { display: grid; grid-template-columns: 1fr auto; gap: 8px 14px; align-items: center;
    padding: 14px 0; border-bottom: 1px solid rgba(255,255,255,.1); }
  .rubric-row:last-child { border-bottom: 0; }
  .rubric-row .name { font-weight: 600; }
  .rubric-row .name span { display: block; font-weight: 400; font-size: 13px; color: var(--muted); margin-top: 2px; }
  .pts { font-weight: 600; font-size: 18px; white-space: nowrap; }
  .pts small { color: var(--muted); font-weight: 400; font-size: 13px; }
  .bar-track { grid-column: 1 / -1; height: 8px; border-radius: 999px; background: rgba(255,255,255,.1); overflow: hidden; }
  .bar-fill { height: 100%; border-radius: 999px; }

  .rules { margin: 0; padding: 0; list-style: none; display: grid; gap: 10px; }
  .rules li { display: grid; grid-template-columns: 7px 1fr; gap: 10px; color: var(--soft); font-size: 14.5px; line-height: 1.7; }
  .rules li::before { content: ""; flex: none; width: 7px; height: 7px; border-radius: 50%; margin-top: 10px; background: var(--green); }

  /* Example tickets */
  .example { border-radius: 22px; padding: 22px; border: 1px solid; }
  .example.low { background: rgba(232,163,61,.08); border-color: rgba(232,163,61,.45); }
  .example.high { background: rgba(105,189,75,.10); border-color: rgba(105,189,75,.5); }
  .ex-head { display: flex; justify-content: space-between; align-items: center; gap: 12px; margin-bottom: 14px; }
  .ex-head .tag { font-size: 13px; font-weight: 600; padding: 5px 12px; border-radius: 999px; }
  .low .tag { background: rgba(232,163,61,.2); color: #FFD89C; }
  .high .tag { background: rgba(105,189,75,.22); color: #CDEFC0; }
  .score-big { font-size: 34px; font-weight: 600; line-height: 1; }
  .score-big small { font-size: 15px; color: var(--muted); font-weight: 400; }
  .bubble { border-radius: 14px; padding: 12px 14px; margin-bottom: 10px; font-size: 15px; line-height: 1.75; white-space: pre-line; }
  .bubble .who { display: block; font-size: 12px; color: var(--muted); margin-bottom: 4px; }
  .bubble.emp { background: rgba(255,255,255,.07); }
  .bubble.agent { background: rgba(255,255,255,.12); }
  .breakdown { margin: 14px 0 0; padding: 0; list-style: none; display: grid; gap: 8px; }

  .steps-flow { list-style: none; margin: 0; padding: 0; display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 12px; }
  .steps-flow li { position: relative; display: grid; gap: 10px; align-content: start;
    padding: 16px; border-radius: 16px; background: rgba(255,255,255,.05); border: 1px solid rgba(255,255,255,.1); }
  .steps-flow .num { width: 30px; height: 30px; border-radius: 50%; display: grid; place-items: center;
    background: #fff; color: #0A2347; font-weight: 700; font-size: 14px; }
  .steps-flow b { font-weight: 600; font-size: 15px; line-height: 1.6; }
  .steps-flow p { margin: 6px 0 0; color: var(--soft); font-size: 13.5px; line-height: 1.7; }
  @media (max-width: 980px) { .steps-flow { grid-template-columns: 1fr; }
    .steps-flow li { grid-template-columns: 30px 1fr; align-items: start; } }

  .breakdown li { display: grid; grid-template-columns: 1fr auto; gap: 2px 10px; font-size: 14px; }
  .breakdown li .why { grid-column: 1 / -1; color: var(--soft); font-size: 13.5px; line-height: 1.6; }
  .breakdown li .p { font-weight: 600; }
  .flag { display: inline-block; margin-top: 12px; font-size: 13px; padding: 6px 12px; border-radius: 10px;
    background: rgba(255,255,255,.1); color: #fff; }

  .note { margin-top: 18px; font-size: 12.5px; color: var(--muted); text-align: center; }

  .sample-flag {
    position: fixed; z-index: 5; bottom: 14px; inset-inline-start: 14px;
    font-size: 12px; padding: 6px 12px; border-radius: 999px;
    background: rgba(9,39,81,.7); border: 1px solid var(--line); color: var(--soft);
  }

  @media (max-width: 520px) {
    .ask { min-height: 60px; padding: 6px; }
    .ask svg.lead { display: none; }
    .ask input { padding: 10px 14px; font-size: 16px; }
    .ask button { width: 48px; height: 48px; padding: 0; }
    .ask .btn-text { display: none; }
    .chip { font-size: 14px; padding: 10px 14px; }
  }
  @media (prefers-reduced-motion: reduce) {
    .orbit { animation: none; } .panel { animation: none; } .dash-card.flash { animation: none; } .draw, .pop, .lo-stick, .lo-dot, .waffle span { animation: none !important; } .ask-ring::before { animation: none; } html { scroll-behavior: auto; }
  }
</style>
</head>
<body>
<div class="bg" aria-hidden="true"></div>
<div class="snow" id="snow" aria-hidden="true"></div>
<div class="vignette" aria-hidden="true"></div>

<div class="wrap">

  <!-- ===================== Hero ===================== -->
  <section class="hero">
    <h1>إجابتك أقرب</h1>
    <p class="edition">إصدار تقييم الإجابات</p>
    <a class="cta" href="#app">
      استعراض المؤشرات
      <svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 5v14"/><path d="M6 13l6 6 6-6"/></svg>
    </a>
  </section>

  <!-- ===================== App ===================== -->
  <section class="app" id="app">
    <div class="app-head">
      <h2>الاستعلام عن مؤشرات جودة الإجابات</h2>
      <p>يمكنكم كتابة الاستعلام، أو اختيار أحد الاستعلامات المتاحة أدناه</p>
      <p class="meta-line" id="metaLine"></p>
    </div>

    <div class="bar-dock" id="barDock">
      <div class="ask-ring" id="askRing">
        <form class="ask matte" id="askForm" autocomplete="off">
          <svg class="lead" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="1.8" stroke-linecap="round" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>
          <input id="askInput" type="text" placeholder="اكتب استعلامك هنا…" aria-label="اكتب استعلامك">
          <button type="submit" aria-label="استعلام"><span class="btn-text">استعلام</span>
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M19 12H5"/><path d="M11 6l-6 6 6 6"/></svg>
          </button>
        </form>
      </div>
      <div class="suggest-pop" id="suggestPop" hidden></div>
    </div>

    <div class="pick-label"><span></span><b>الاستعلامات المتاحة</b><span></span></div>
    <div class="chips" id="chips"></div>

    <div class="kpi-strip" id="kpiStrip"></div>
    <div class="dash" id="dash"></div>
  </section>
</div>

<div class="sample-flag" id="sampleFlag">بيانات استرشادية</div>

<script>
/* =====================================================================
   DATA — filled in automatically by quality_report.py
   ===================================================================== */
const DATA = /*__DATA__*/null;

const COLORS = { good: "#69BD4B", ok: "#4BB8C4", weak: "#E8A33D" };
const BAND_NAMES = { good: "مرتفعة الجودة (4 من 4)", ok: "متوسطة الجودة (2–3)", weak: "منخفضة الجودة (0–1)" };

/* ---------------------------- Questions ---------------------------- */
const QUESTIONS = [
  { id: "avg",      text: "ما المتوسط العام لدرجة اكتمال الإجابات؟", keys: ["متوسط", "درجه", "معدل", "اكتمال"] },
  { id: "bands",    text: "ما توزيع الإجابات حسب مستويات الجودة؟", keys: ["توزيع", "مستويات", "جيد", "ضعيف", "منخفض", "مرتفع"] },
  { id: "types",    text: "ما مستوى جودة الإجابات حسب نوع الاستفسار؟", keys: ["نوع", "سؤال", "اسئله", "تصنيف", "موضوع"] },
  { id: "missing",  text: "ما أكثر المعلومات غياباً عن الإجابات؟", keys: ["غياب", "ناقص", "نقص", "معلومات", "عناصر"] },
  { id: "sla",      text: "هل يعكس الالتزام بمستوى الخدمة جودة الإجابة؟", keys: ["sla", "التزام", "مستوي الخدمه", "سرعه", "وقت"] },
  { id: "repeat",   text: "ما نسبة تكرار الاستفسار خلال أربعة عشر يوماً؟", keys: ["تكرار", "يتكرر", "اعاده", "فتح", "14"] },
  { id: "insights", text: "ما أبرز النتائج والملاحظات؟", keys: ["ابرز", "مخرجات", "نتائج", "خلاصه", "ملخص", "ملاحظات"] },
  { id: "how",      text: "ما منهجية تقييم جودة الإجابات؟", keys: ["كيف", "تحديد", "معايير", "منهجيه", "تقييم", "طريقه"], special: true },
].filter(q => q.id !== "sla" || (DATA && DATA.sla && DATA.sla.available));

/* ---------------------------- Helpers ---------------------------- */
const fmt = (n, d = 1) => Number(n).toLocaleString("en-US", { maximumFractionDigits: d, minimumFractionDigits: 0 });
const esc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

function norm(t) {
  return t.toLowerCase()
    .replace(/[ً-ْـ]/g, "")
    .replace(/[إأآ]/g, "ا").replace(/ة/g, "ه").replace(/ى/g, "ي")
    .replace(/[^\w\s؀-ۿ]/g, " ");
}

/* Horizontal bars (HTML, responsive; RTL so bars grow from the right) */
function hBars(items, max, unit = "") {
  return `<div class="hbars">${items.map(it => {
    const m = it.max || max;
    const pct = Math.max(2, Math.min(100, (it.value / m) * 100));
    const val = `${fmt(it.value)}${it.max ? " / " + it.max : unit}`;
    return `<div class="hb-row" title="${esc(it.label)}: ${val}">
      <div class="hb-label">${esc(it.label)}</div>
      <div class="hb-track"><div class="hb-fill" style="width:${pct}%;background:${it.color}"></div></div>
      <div class="hb-val">${val}</div>
    </div>`;
  }).join("")}</div>`;
}

/* Donut */
function donut(parts, centerTop, centerBottom) {
  const r = 70, c = 2 * Math.PI * r, size = 200;
  let off = 0, s = `<svg class="chart" viewBox="0 0 ${size} ${size}" style="direction:ltr;max-width:240px;margin:0 auto" role="img">`;
  s += `<circle cx="100" cy="100" r="${r}" fill="none" stroke="rgba(255,255,255,.08)" stroke-width="26"/>`;
  parts.forEach(p => {
    const len = (p.value / 100) * c;
    s += `<circle cx="100" cy="100" r="${r}" fill="none" stroke="${p.color}" stroke-width="26"
           stroke-dasharray="${Math.max(len - 2, 0)} ${c}" stroke-dashoffset="${-off}"
           transform="rotate(-90 100 100)"><title>${esc(p.label)}: ${p.value}%</title></circle>`;
    off += len;
  });
  s += `<text x="100" y="98" text-anchor="middle" fill="#fff" font-size="30" font-weight="600">${centerTop}</text>`;
  s += `<text x="100" y="122" text-anchor="middle" fill="#A9C1D6" font-size="13">${centerBottom}</text>`;
  return s + "</svg>";
}

/* Vertical columns */
function columns(items, max, unit = "%") {
  const w = 480, h = 240, base = 200, colW = 74;
  const step = w / items.length;
  let s = `<svg class="chart" style="direction:ltr" viewBox="0 0 ${w} ${h}" role="img">`;
  s += `<line x1="0" x2="${w}" y1="${base}" y2="${base}" stroke="rgba(255,255,255,.18)"/>`;
  items.forEach((it, i) => {
    const x = w - (i + 0.5) * step;           // RTL order
    const ch = (it.value / max) * 160;
    s += `<rect x="${x - colW / 2}" y="${base - ch}" width="${colW}" height="${ch}" rx="10" fill="${it.color}"><title>${esc(it.label)}: ${it.value}${unit}</title></rect>`;
    s += `<text x="${x}" y="${base - ch - 10}" text-anchor="middle" fill="#fff" font-size="17" font-weight="600">${fmt(it.value)}${unit}</text>`;
    s += `<text x="${x}" y="${base + 24}" text-anchor="middle" fill="#EAF5EE" font-size="14">${esc(it.label)}</text>`;
  });
  return s + "</svg>";
}

/* 100% stacked bar */
function stacked(parts) {
  const w = 640, h = 46;
  let x = w, s = `<svg class="chart" style="direction:ltr" viewBox="0 0 ${w} ${h}" role="img">`;
  parts.forEach(p => {
    const pw = (p.value / 100) * w;
    x -= pw;
    s += `<rect x="${x}" y="0" width="${Math.max(pw - 2, 0)}" height="${h}" fill="${p.color}"><title>${esc(p.label)}: ${p.value}%</title></rect>`;
    if (p.value >= 7) s += `<text x="${x + pw / 2}" y="${h / 2 + 6}" text-anchor="middle" fill="#092751" font-size="16" font-weight="700">${p.value}%</text>`;
  });
  return `<div style="border-radius:14px;overflow:hidden">${s}</svg></div>`;
}


/* ===================== Unusual charts ===================== */

/* Gauge: half-circle dial */
function gauge(value, max, label) {
  const r = 90, cx = 120, cy = 120, a0 = Math.PI, frac = Math.min(value / max, 1);
  const pt = f => [cx + r * Math.cos(a0 - f * Math.PI), cy - r * Math.sin(a0 - f * Math.PI)];
  const arc = f => { const [x, y] = pt(f); return `M ${cx - r} ${cy} A ${r} ${r} 0 0 1 ${x.toFixed(1)} ${y.toFixed(1)}`; };
  const len = Math.PI * r;
  let ticks = "";
  for (let i = 0; i <= 10; i++) {
    const a = a0 - (i / 10) * Math.PI, r1 = r + 16, r2 = r + (i % 5 ? 21 : 26);
    ticks += `<line x1="${cx + r1 * Math.cos(a)}" y1="${cy - r1 * Math.sin(a)}" x2="${cx + r2 * Math.cos(a)}" y2="${cy - r2 * Math.sin(a)}" stroke="rgba(255,255,255,${i % 5 ? .3 : .6})" stroke-width="1.5"/>`;
  }
  const [dx, dy] = pt(frac);
  return `<svg class="chart" viewBox="0 0 240 150" style="direction:ltr;max-width:340px;margin:0 auto" role="img">
    <defs><linearGradient id="gg" x1="0" x2="1"><stop offset="0" stop-color="#E8A33D"/><stop offset=".5" stop-color="#4BB8C4"/><stop offset="1" stop-color="#69BD4B"/></linearGradient></defs>
    ${ticks}
    <path d="${arc(1)}" fill="none" stroke="rgba(255,255,255,.08)" stroke-width="16" stroke-linecap="round"/>
    <path class="draw" style="--len:${len}" d="${arc(frac)}" fill="none" stroke="url(#gg)" stroke-width="16" stroke-linecap="round"
      stroke-dasharray="${len}" stroke-dashoffset="0"/>
    <circle cx="${dx}" cy="${dy}" r="11" fill="#0A2347" stroke="#fff" stroke-width="3"/>
    <text x="${cx}" y="${cy - 8}" text-anchor="middle" fill="#fff" font-size="40" font-weight="600">${fmt(value)}</text>
    <text x="${cx}" y="${cy + 18}" text-anchor="middle" fill="#A9C1D6" font-size="13">${label}</text>
  </svg>`;
}

/* Radial rings: one concentric ring per criterion */
function rings(items) {
  const cx = 110, cy = 110, w = 13, gap = 7;
  let svg = `<svg class="chart" viewBox="0 0 220 220" style="direction:ltr;max-width:260px;margin:0 auto" role="img">`;
  items.forEach((it, i) => {
    const r = 96 - i * (w + gap), c = 2 * Math.PI * r, f = Math.min(it.value / it.max, 1);
    svg += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="rgba(255,255,255,.07)" stroke-width="${w}"/>`;
    svg += `<circle class="draw" style="--len:${c}" cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${it.color}" stroke-width="${w}" stroke-linecap="round"
      stroke-dasharray="${f * c} ${c}" transform="rotate(-90 ${cx} ${cy})"><title>${esc(it.label)}: ${it.value} / ${it.max}</title></circle>`;
  });
  svg += `</svg>`;
  const leg = `<div class="ring-legend">${items.map(it =>
    `<div><i style="background:${it.color}"></i><span>${esc(it.label)} ${it.src ? srcTag(it.src) : ""}</span><b>${fmt(it.value)} <small>من ${it.max}</small></b></div>`).join("")}</div>`;
  return `<div class="rings-wrap">${svg}${leg}</div>`;
}

/* Waffle: 100 dots, one per 1% */
function waffle(parts) {
  const counts = parts.map(p => Math.round(p.value));
  let diff = 100 - counts.reduce((a, b) => a + b, 0);
  counts[0] += diff;
  const cells = [];
  parts.forEach((p, i) => { for (let k = 0; k < counts[i]; k++) cells.push(p); });
  return `<div class="waffle">${cells.map((p, i) =>
    `<span style="background:${p.color};animation-delay:${i * 6}ms" title="${esc(p.label)}: ${p.value}%"></span>`).join("")}</div>`;
}

/* Lollipop: thin stick + dot with the value */
function lollipop(items, max) {
  return `<div class="lolli">${items.map(it => {
    const pct = Math.max(4, Math.min(100, (it.value / max) * 100));
    return `<div class="lo-row" title="${esc(it.label)}: ${it.value}">
      <div class="lo-label">${esc(it.label)}${it.note ? `<small>${it.note}</small>` : ""}</div>
      <div class="lo-track">
        <div class="lo-stick" style="width:${pct}%;--c:${it.color}"></div>
        <div class="lo-dot" style="right:calc(${pct}% - 19px);--c:${it.color}">${fmt(it.value)}</div>
      </div>
    </div>`;
  }).join("")}</div>`;
}

/* Nightingale rose: wedge length shows the share */
function rose(parts, centerText) {
  const cx = 130, cy = 130, R = 112, maxV = Math.max(...parts.map(p => p.value));
  const step = (2 * Math.PI) / parts.length;
  let svg = `<svg class="chart" viewBox="0 0 260 260" style="direction:ltr;max-width:300px;margin:0 auto" role="img">`;
  [0.33, 0.66, 1].forEach(f => svg += `<circle cx="${cx}" cy="${cy}" r="${R * f}" fill="none" stroke="rgba(255,255,255,.08)"/>`);
  parts.forEach((p, i) => {
    const r = Math.max(26, R * Math.sqrt(p.value / maxV));
    const a1 = -Math.PI / 2 + i * step + 0.03, a2 = a1 + step - 0.06;
    const x1 = cx + r * Math.cos(a1), y1 = cy + r * Math.sin(a1), x2 = cx + r * Math.cos(a2), y2 = cy + r * Math.sin(a2);
    const am = (a1 + a2) / 2, tr = r * 0.66;
    svg += `<path class="pop" style="animation-delay:${i * 90}ms" d="M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 0 1 ${x2} ${y2} Z" fill="${p.color}" opacity=".92"><title>${esc(p.label)}: ${p.value}%</title></path>`;
    svg += `<text x="${cx + tr * Math.cos(am)}" y="${cy + tr * Math.sin(am) + 6}" text-anchor="middle" fill="#0A2347" font-size="16" font-weight="700">${p.value}%</text>`;
  });
  svg += `<circle cx="${cx}" cy="${cy}" r="22" fill="#0A2347" stroke="rgba(255,255,255,.4)"/>
    <text x="${cx}" y="${cy + 5}" text-anchor="middle" fill="#fff" font-size="12" font-weight="600">${centerText}</text></svg>`;
  return svg;
}

/* Bubbles: circle area shows the rate */
function bubbles(items) {
  const maxV = Math.max(...items.map(i => i.value)), w = 480, h = 230, step = w / items.length;
  let svg = `<svg class="chart" viewBox="0 0 ${w} ${h}" style="direction:ltr" role="img">`;
  svg += `<line x1="20" x2="${w - 20}" y1="170" y2="170" stroke="rgba(255,255,255,.15)" stroke-dasharray="3 5"/>`;
  items.forEach((it, i) => {
    const x = w - (i + 0.5) * step, r = 18 + 62 * Math.sqrt(it.value / maxV), y = 170 - r;
    svg += `<circle class="pop" style="animation-delay:${i * 120}ms" cx="${x}" cy="${y}" r="${r}" fill="${it.color}" fill-opacity=".9"
      stroke="rgba(255,255,255,.5)" stroke-width="1.5"><title>${esc(it.label)}: ${it.value}%</title></circle>`;
    svg += `<text x="${x}" y="${y + 7}" text-anchor="middle" fill="#0A2347" font-size="${r > 30 ? 22 : 15}" font-weight="700">${it.value}%</text>`;
    svg += `<text x="${x}" y="200" text-anchor="middle" fill="#EAF5EE" font-size="15">${esc(it.label)}</text>`;
  });
  return svg + "</svg>";
}

const legend = keys => `<div class="legend">${keys.map(k => `<span><i style="background:${COLORS[k]}"></i>${BAND_NAMES[k]}</span>`).join("")}</div>`;
const bandParts = obj => ["good", "ok", "weak"].map(k => ({ label: BAND_NAMES[k], value: obj[k], color: COLORS[k] }));

/* ---------------------------- Answers ---------------------------- */
const pct = v => (v === null || v === undefined) ? "—" : fmt(v, 0) + "%";

const ANSWERS = {
  avg: () => ({
    lead: `بلغ المتوسط العام لدرجة اكتمال الإجابات <strong>${fmt(DATA.avgCompleteness, 2)} من 4</strong>، بناءً على تقييم ${fmt(DATA.tickets, 0)} استفساراً موزعة على ${fmt(DATA.questions, 0)} أنواع من الاستفسارات.`,
    body: `
      <div class="grid g2">
        <div class="card"><h4>المؤشر العام للاكتمال</h4>${gauge(DATA.avgCompleteness, 4, "من 4")}</div>
        <div class="card"><h4>توزيع درجات الاكتمال</h4>
          <div class="sub">نسبة الإجابات الحاصلة على كل درجة</div>
          ${lollipop([4, 3, 2, 1, 0].map(k => ({ label: `${k} من 4`, value: DATA.scoreDist[k] || 0,
              color: k === 4 ? COLORS.good : k >= 2 ? COLORS.ok : COLORS.weak })), 100)}
        </div>
      </div>`
  }),

  bands: () => ({
    lead: `صُنِّفت <strong>${pct(DATA.bands.good)}</strong> من الإجابات ضمن مستوى الجودة المرتفعة، و<strong>${pct(DATA.bands.weak)}</strong> ضمن مستوى الجودة المنخفضة.`,
    body: `
      <div class="grid g2">
        <div class="card"><h4>تمثّل كل نقطة 1% من الاستفسارات المُقيَّمة</h4>
          ${waffle(bandParts(DATA.bands))}
          ${legend(["good", "ok", "weak"])}
        </div>
        <div class="card"><h4>التوزيع النسبي</h4>
          ${lollipop(bandParts(DATA.bands).map(p => ({ label: p.label, value: p.value, color: p.color })), 100)}
        </div>
      </div>`
  }),

  types: () => {
    const rel = DATA.byType.filter(t => !t.small);
    const pool = rel.length ? rel : DATA.byType;
    const sorted = [...pool].sort((a, b) => b.completeness - a.completeness);
    const all = [...DATA.byType].sort((a, b) => b.completeness - a.completeness);
    return {
      lead: sorted.length > 1
        ? `سجّل استفسار <strong>${esc(sorted[0].name)}</strong> أعلى متوسط للاكتمال، في حين سجّل استفسار <strong>${esc(sorted[sorted.length - 1].name)}</strong> أدناه.`
        : `متوسط الاكتمال لكل نوع من أنواع الاستفسارات.`,
      body: `
        <div class="card"><h4>متوسط درجة الاكتمال حسب نوع الاستفسار (من 4)</h4>
          ${lollipop(all.map(t => ({ label: t.name,
              note: `${fmt(t.n, 0)} استفساراً${t.small ? " · عدد محدود لا يُعتد به إحصائياً" : ""}`,
              value: t.completeness,
              color: t.completeness >= 3 ? COLORS.good : t.completeness >= 2 ? COLORS.ok : COLORS.weak })), 4)}
        </div>`
    };
  },

  missing: () => {
    const m = DATA.missing;
    if (!m.length) return { lead: "لا تتوفر بيانات كافية لتحديد العناصر الغائبة.", body: "" };
    return {
      lead: `يُعد عنصر <strong>«${esc(m[0].element)}»</strong> الأكثر غياباً؛ إذ لم يرد في <strong>${pct(m[0].pct)}</strong> من إجابات استفسار «${esc(m[0].question)}».`,
      body: `
        <div class="card"><h4>نسبة الإجابات التي لم تتضمن العنصر (%)</h4>
          <div class="sub">العناصر مستخلصة من الإجابات المرجعية المعتمدة</div>
          ${lollipop(m.map(x => ({ label: x.element.length > 90 ? x.element.slice(0, 90) + "…" : x.element,
              note: x.question, value: x.pct, color: x.pct >= 70 ? COLORS.weak : x.pct >= 40 ? COLORS.ok : COLORS.good })), 100)}
        </div>`
    };
  },

  sla: () => ({
    lead: `لا يُعد الالتزام بمستوى الخدمة مؤشراً كافياً على جودة الإجابة؛ إذ صُنِّفت <strong>${pct(DATA.sla.bands.weak)}</strong> من الاستفسارات <strong>الملتزمة باتفاقية مستوى الخدمة</strong> ضمن مستوى الجودة المنخفضة.`,
    body: `
      <div class="card"><h4>مستويات جودة الإجابات في الاستفسارات الملتزمة بمستوى الخدمة</h4>
        <div class="sub">نسبة الالتزام بمستوى الخدمة: ${pct(DATA.sla.metPct)} · يعكس امتداد كل شريحة نسبتها من الإجمالي</div>
        ${rose(bandParts(DATA.sla.bands), "ملتزمة")}
        ${legend(["good", "ok", "weak"])}
      </div>
      <p class="lead-text" style="margin:18px 0 0">يقيس مستوى الخدمة <strong>سرعة الاستجابة</strong>، في حين تقيس الجودة <strong>مدى معالجة الإجابة لاحتياج المستفيد</strong>؛ ولذا يُعرض المؤشران جنباً إلى جنب.</p>`
  }),

  repeat: () => {
    const r = DATA.repeat;
    const compare = `
      <div class="vs">
        <div class="vs-item"><div class="v">${pct(DATA.reopenRate)}</div><div class="l">استفسارات أُعيد فتحها</div></div>
        <div class="vs-mid">مقابل</div>
        <div class="vs-item hl"><div class="v">${r.available ? pct(r.overall) : "—"}</div><div class="l">استفسارات تكررت خلال أربعة عشر يوماً</div></div>
      </div>`;
    if (!r.available) return {
      lead: `بلغت نسبة إعادة فتح الاستفسارات <strong>${pct(DATA.reopenRate)}</strong>. لم يُحتسب تكرار الاستفسار لعدم توفر بيانات المستفيد.`,
      body: `<div class="card">${compare}</div>`
    };
    return {
      lead: `لم تتجاوز نسبة إعادة الفتح <strong>${pct(DATA.reopenRate)}</strong>، في حين تكرر الاستفسار ذاته من المستفيد ذاته خلال أربعة عشر يوماً بنسبة <strong>${pct(r.overall)}</strong>.`,
      body: `
        <div class="card" style="margin-bottom:16px">${compare}</div>
        <div class="card"><h4>نسبة التكرار حسب مستوى جودة الإجابة</h4>
          <div class="sub">المستفيد ذاته · نوع الاستفسار ذاته · بعد إغلاق الاستفسار الأول</div>
          ${bubbles([
            { label: "مرتفعة الجودة", value: r.byBand.good, color: COLORS.good },
            { label: "متوسطة الجودة", value: r.byBand.ok,   color: COLORS.ok },
            { label: "منخفضة الجودة", value: r.byBand.weak, color: COLORS.weak },
          ])}
        </div>`
    };
  },

  insights: () => ({
    lead: `فيما يلي أبرز النتائج والملاحظات المستخلصة من تقييم ${fmt(DATA.tickets, 0)} استفساراً:`,
    body: `<div class="grid g2">${DATA.insights.map((t, i) =>
      `<div class="card insight"><div class="n">${fmt(i + 1, 0)}</div><p>${t}</p></div>`).join("")}</div>`
  }),

  how: () => ({
    lead: `تُقارَن كل إجابة بـ<strong>الإجابة المرجعية المعتمدة</strong> لنوع الاستفسار، وتُمنح <strong>درجة اكتمال من 4</strong> وفق نسبة العناصر الأساسية المستوفاة فيها.`,
    body: `
      <div class="card" style="margin-bottom:16px"><h4>مراحل التقييم</h4>
        <ol class="steps-flow">
          ${[
            ["تصنيف الاستفسارات", "حصر الاستفسارات المطابقة لكل نوع بالبحث في نصوصها بعد توحيد صيغ الكتابة."],
            ["تجزئة الإجابة المرجعية", "تقسيم كل إجابة مرجعية إلى عناصرها الأساسية من جمل ومقاطع."],
            ["التحقق من العناصر", "التحقق من ورود كل عنصر في الإجابة المقدَّمة، وبيان العناصر المستوفاة والناقصة."],
            ["قياس الأثر", "احتساب إعادة الفتح، وتكرار الاستفسار ذاته من المستفيد ذاته خلال أربعة عشر يوماً."],
            ["إصدار النتائج", "احتساب الدرجات وتصنيف مستويات الجودة واستخلاص الملاحظات."],
          ].map((st, i) => `<li><span class="num">${i + 1}</span><div><b>${st[0]}</b><p>${st[1]}</p></div></li>`).join("")}
        </ol>
      </div>

      <div class="grid g2" style="margin-bottom:16px">
        <div class="card"><h4>احتساب درجة الاكتمال</h4>
          ${[
            [4, "استيفاء جميع العناصر الأساسية", COLORS.good],
            [3, "استيفاء 75% من العناصر فأكثر", COLORS.ok],
            [2, "استيفاء 50% من العناصر فأكثر", COLORS.ok],
            [1, "استيفاء أقل من 50% من العناصر", COLORS.weak],
            [0, "عدم استيفاء أي عنصر", COLORS.weak],
          ].map(r => scoreRow(r[0], r[1], r[2])).join("")}
          <div class="legend" style="margin-top:16px">
            <span><i style="background:${COLORS.good}"></i>4 مرتفعة الجودة</span>
            <span><i style="background:${COLORS.ok}"></i>2–3 متوسطة الجودة</span>
            <span><i style="background:${COLORS.weak}"></i>0–1 منخفضة الجودة</span>
          </div>
        </div>
        <div class="card"><h4>ضوابط التقييم</h4>
          <ul class="rules">
            <li><span>يُعتد بالكلمات المميزة لكل عنصر، وتُستبعد الكلمات العامة المشتركة بين العناصر.</span></li>
            <li><span>لا يُعد العنصر المتضمن رقماً (كالمدة أو عدد الأيام) مستوفى إلا بورود الرقم ذاته.</span></li>
            <li><span>تُوحَّد صيغ الكتابة قبل المقارنة، بما يشمل الهمزات والتاء المربوطة والأخطاء الإملائية البسيطة.</span></li>
            <li><span>يُحتسب تكرار الاستفسار عند تقديم المستفيد ذاته استفساراً من النوع ذاته خلال أربعة عشر يوماً من إغلاق الاستفسار السابق.</span></li>
            <li><span>لا تشمل الدرجة محتوى المرفقات، ولا تقيس دقة المعلومات الواردة في الإجابة.</span></li>
          </ul>
        </div>
      </div>
      ${DATA.examples ? `
      <div class="grid g2">
        ${example("low", "نموذج لإجابة منخفضة الجودة", DATA.examples.low)}
        ${example("high", "نموذج لإجابة مرتفعة الجودة", DATA.examples.high)}
      </div>` : ""}`
  }),
};

function scoreRow(score, desc, color) {
  return `<div class="rubric-row">
    <div class="name">${desc}</div>
    <div class="pts">${fmt(score, 0)} <small>من 4</small></div>
    <div class="bar-track"><div class="bar-fill" style="width:${score * 25}%;background:${color}"></div></div>
  </div>`;
}
function example(kind, title, ex) {
  if (!ex) return "";
  return `<div class="example ${kind}">
    <div class="ex-head"><span class="tag">${title}</span><div class="score-big">${fmt(ex.score, 0)}<small> / 4</small></div></div>
    <div class="sub" style="margin:0 0 10px;color:#A9C1D6;font-size:13px">${esc(ex.question)}</div>
    <div class="bubble emp"><span class="who">استفسار المستفيد</span>${esc(ex.inquiry)}</div>
    <div class="bubble agent"><span class="who">الإجابة المقدَّمة</span>${esc(ex.answer)}</div>
    <ul class="breakdown">
      ${ex.covered.map(c => `<li><span>✓ ${esc(c)}</span><span class="p" style="color:#CDEFC0">مستوفى</span></li>`).join("")}
      ${ex.missing.map(c => `<li><span>✗ ${esc(c)}</span><span class="p" style="color:#FFD89C">غير مستوفى</span></li>`).join("")}
    </ul>
  </div>`;
}

/* ---------------------------- Dashboard ---------------------------- */
const chipsEl = document.getElementById("chips");
const dashEl = document.getElementById("dash");
const input = document.getElementById("askInput");
const ring = document.getElementById("askRing");
const pop = document.getElementById("suggestPop");
const FULL = new Set(["avg", "missing", "insights", "how"]);

/* Top strip: the numbers that matter most, always visible */
document.getElementById("kpiStrip").innerHTML = [
  [fmt(DATA.avgCompleteness, 2), "/4", "المتوسط العام لدرجة الاكتمال"],
  [fmt(DATA.tickets, 0), "", "إجمالي الاستفسارات المُقيَّمة"],
  [fmt(DATA.bands.good, 0), "%", "إجابات مرتفعة الجودة"],
  [fmt(DATA.reopenRate, 0), "%", "استفسارات أُعيد فتحها"],
  [DATA.repeat.available ? fmt(DATA.repeat.overall, 0) : "—", DATA.repeat.available ? "%" : "", "استفسارات تكررت خلال أربعة عشر يوماً", true],
].map(k => `<div class="kpi-tile matte${k[3] ? " hl" : ""}"><div class="v">${k[0]}<small>${k[1]}</small></div><div class="l">${k[2]}</div></div>`).join("");
document.getElementById("metaLine").textContent =
  `فترة البيانات: ${DATA.period} · تاريخ الإصدار: ${DATA.generatedAt}`;

function cardHTML(q) {
  const a = ANSWERS[q.id]();
  return `<header><span class="q-num">${fmt(QUESTIONS.indexOf(q) + 1, 0)}</span><h3>${esc(q.text)}</h3></header>
    <p class="lead-text">${a.lead}</p>${a.body}`;
}

QUESTIONS.forEach(q => {
  const card = document.createElement("article");
  card.className = "dash-card matte" + (FULL.has(q.id) ? " span-full" : "");
  card.id = "card-" + q.id;
  card.innerHTML = cardHTML(q);
  dashEl.appendChild(card);

  const b = document.createElement("button");
  b.type = "button";
  b.className = "chip" + (q.special ? " special" : "");
  b.textContent = q.text;
  b.dataset.id = q.id;
  b.addEventListener("click", () => pick(q.id));
  chipsEl.appendChild(b);
});

/* Picking a question: it is typed into the bar, then the dashboard jumps to it */
let typingTimer = null;
function pick(id) {
  const q = QUESTIONS.find(x => x.id === id);
  clearInterval(typingTimer);
  pop.hidden = true;
  input.value = "";
  ring.classList.add("typing");
  [...chipsEl.children].forEach(c => c.classList.toggle("active", c.dataset.id === id));
  const reduce = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  if (reduce) { input.value = q.text; ring.classList.remove("typing"); return go(id); }
  const step = Math.max(12, Math.min(35, 700 / q.text.length));
  let i = 0;
  typingTimer = setInterval(() => {
    input.value = q.text.slice(0, ++i);
    if (i >= q.text.length) {
      clearInterval(typingTimer);
      setTimeout(() => { ring.classList.remove("typing"); go(id); }, 220);
    }
  }, step);
}

function go(id) {
  const q = QUESTIONS.find(x => x.id === id);
  const card = document.getElementById("card-" + id);
  card.innerHTML = cardHTML(q);                 // replay the chart animation
  [...chipsEl.children].forEach(c => c.classList.toggle("active", c.dataset.id === id));
  card.scrollIntoView({ behavior: "smooth", block: "start" });
  card.classList.remove("flash"); void card.offsetWidth; card.classList.add("flash");
}

document.getElementById("askForm").addEventListener("submit", e => {
  e.preventDefault();
  const text = norm(input.value.trim());
  if (!text) return;
  let best = null, bestScore = 0;
  QUESTIONS.forEach(q => {
    const sc = q.keys.reduce((s, k) => s + (text.includes(norm(k)) ? 1 : 0), 0)
             + (norm(q.text) === text ? 5 : 0);
    if (sc > bestScore) { bestScore = sc; best = q; }
  });
  if (best) { pop.hidden = true; go(best.id); return; }
  suggest(text);
});

/* No direct match: suggest the closest questions under the bar */
function suggest(text) {
  const words = text.split(/\s+/).filter(w => w.length > 2);
  const ranked = QUESTIONS.map(q => {
    const qt = norm(q.text + " " + q.keys.join(" "));
    return { q, sc: words.reduce((a, w) => a + (qt.includes(w.slice(0, 3)) ? 1 : 0), 0) };
  }).sort((a, b) => b.sc - a.sc).slice(0, 3);
  pop.innerHTML = `<p>استعلامات مقترحة ذات صلة</p><div>${ranked.map(r =>
    `<button type="button" class="chip" data-go="${r.q.id}">${esc(r.q.text)}</button>`).join("")}</div>`;
  pop.hidden = false;
  pop.querySelectorAll("[data-go]").forEach(b => b.addEventListener("click", () => pick(b.dataset.go)));
}
input.addEventListener("input", () => { pop.hidden = true; });

if (!DATA.isSample) document.getElementById("sampleFlag").remove();

/* ---------------------------- Snow (same as إجابتك أقرب) ---------------------------- */
(function () {
  const snow = document.getElementById("snow");
  let seed = 7;
  const rnd = () => (seed = (seed * 9301 + 49297) % 233280) / 233280;
  for (let i = 0; i < 70; i++) {
    const size = 2 + rnd() * 4, radius = 18 + rnd() * 70, duration = 28 + rnd() * 44;
    const orbit = document.createElement("div");
    orbit.className = "orbit";
    orbit.style.left = (rnd() * 100).toFixed(2) + "%";
    orbit.style.top = (rnd() * 100).toFixed(2) + "%";
    orbit.style.animationDuration = duration.toFixed(1) + "s";
    orbit.style.animationDelay = (-rnd() * duration).toFixed(1) + "s";
    orbit.style.animationDirection = rnd() > 0.5 ? "reverse" : "normal";
    const flake = document.createElement("span");
    flake.className = "flake";
    flake.style.width = flake.style.height = size.toFixed(1) + "px";
    flake.style.opacity = (0.18 + rnd() * 0.55).toFixed(2);
    flake.style.transform = "translateX(" + radius.toFixed(0) + "px)";
    flake.style.filter = "blur(" + (size > 4.8 ? 1.2 : 0) + "px)";
    orbit.appendChild(flake);
    snow.appendChild(orbit);
  }
})();
</script>
</body>
</html>
'''


if __name__ == "__main__":
    main()

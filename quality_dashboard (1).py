# -*- coding: utf-8 -*-
"""
quality_dashboard.py

One run does everything:
  1. Reads the clean tickets file, the Questions file (Question, Answer, Match / Required /
     Excluded Keywords) and, optionally, the ratings file.
  2. Assigns each inquiry to one question type using the keyword columns.
  3. Scores each answer against the ideal answer (completeness of its elements).
     Answers saying the service is not possible are excluded from the score,
     unless the ideal answer itself says so.
  4. Calculates the indicators: quality, repeat inquiries over the year, reopen,
     satisfaction, dissatisfaction reasons, monthly trend + forecast, drivers of extra tickets.
  5. Writes the dashboard (one HTML file) and one Excel sheet with the scored inquiries.

Needs: pandas, numpy, openpyxl. Runs fully offline.
"""

import json
import re
import sys
import time
from pathlib import Path

import numpy as np
import pandas as pd

# =============================================================================
# >>> YOUR FILES <<<
# =============================================================================
TICKETS_FILE = r"C:\PUT\YOUR\PATH\HERE\clean\tickets_clean.csv"        # .csv / .parquet / .xlsx
QUESTIONS_FILE = r"C:\Users\rammutairi\Desktop\Ideal Answers.xlsx"
RATINGS_FILE = r"C:\PUT\YOUR\PATH\HERE\ratings.xlsx"      # leave "" if you don't have it
OUTPUT_FOLDER = r""     # empty = same folder as the tickets file

# ---- Column names: tickets file ----
INQUIRY_COLUMN = "Inquiry"
ANSWER_COLUMN = "Public Logs"
CATEGORY_COLUMN = "Assigned Group"          # shown in the dashboard as "الفئة"
TICKET_COLUMN = "Ticket Number"             # used to link the ratings file
EMPLOYEE_COLUMN = "Email Address"           # identifies the employee (repeat inquiries)
SUBMIT_COLUMN = "Submit Date"
RESOLVED_COLUMN = "Last Resolved Date"      # if missing, "Resolution Hours" is used
HOURS_COLUMN = "Resolution Hours"
REOPEN_COLUMN = "REOPEN COUNT"

# ---- Column names: questions file ----
QUESTION_COLUMN = "Question"
IDEAL_COLUMN = "Answer"
MATCH_COLUMN = "Match Keywords"
REQUIRED_COLUMN = "Required Keywords"
EXCLUDED_COLUMN = "Excluded Keywords"

# ---- Column names: ratings file ----
RATING_TICKET_COLUMN = "Ticket Number"
RATING_COLUMN = "Employee Satisfaction"     # غير راضي / راضي / راضي جدا (or numbers 1–5)
COMMENT_CATEGORY_COLUMN = "Employee Feedback Category"

# How the satisfaction answers map to the 1–5 scale shown in the dashboard.
RATING_MAP = {"غير راضي": 1, "راضي": 3, "راضي جدا": 5}

# =============================================================================
# SETTINGS (normally leave as they are)
# =============================================================================
DEPARTMENT_NAME = "إدارة أداء رأس المال البشري وضمان الجودة"
TARGET = 75               # quality target %
ACCEPTABLE = 65           # acceptable quality %
HIGH_LEVEL = 75           # answer quality >= this = high
LOW_LEVEL = 50            # answer quality < this = low
MIN_CATEGORY = 30         # categories with fewer inquiries are not ranked
ELEMENT_COVERED_AT = 0.6  # share of an element's key words that must appear in the answer

# Phrases meaning "this can't be done" (searched in the answers, after normalisation).
CANNOT_PHRASES = [
    "لا يمكن", "لا يمكننا", "لايمكن", "غير ممكن", "يتعذر", "تعذر", "نعتذر عن عدم", "نعتذر لعدم",
    "نأسف لعدم", "غير متاح", "غير متاحة", "غير متوفر", "غير متوفرة", "لا يتوفر", "لا تتوفر",
    "لا يوجد إمكانية", "لا توجد إمكانية", "ليس بالإمكان", "ليس من الممكن", "خارج صلاحية",
    "خارج الصلاحية", "خارج نطاق", "ليس من اختصاص", "ليست من اختصاص", "ليس من صلاحيات",
    "ليست من صلاحيات", "النظام لا يسمح", "لا يسمح النظام", "النظام ما يسمح", "ما يمكن", "مايمكن",
    "ما نقدر", "مانقدر", "ما نستطيع", "لا نستطيع", "تم إيقاف", "أوقفت", "موقوف", "غير مسموح",
    "لا يحق لنا",
]

STOPWORDS = set("""
في من على عن الى إلى ان أن او أو و يا هل مع ما لا هي هو انا أنا ابي ابغى اريد ارغب ارجو اود لو كيف متى
وش ايش طلب بخصوص لدي عندي الله السلام عليكم ورحمه وبركاته الذي التي ذلك هذا هذه تم يتم سيتم قد كل اي ثم
حيث عند يرجى يرجي نامل نود نفيدكم بان بانه لكم لك منكم قبل بعد طريقه كيفيه الاستفسار استفسار شكرا تحيه
طيبه تحياتي عزيزي الموظف المستفيد خلال ذات حال حاله وذلك كما يمكنكم يمكنك نرجو
""".split())

MONTHS_AR = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس",
             "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]

# =============================================================================

DIACRITICS = re.compile(r"[\u0610-\u061A\u064B-\u065F\u0670\u06D6-\u06ED\u0640]")
LETTERS = str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ٱ": "ا", "ة": "ه", "ى": "ي", "ؤ": "و",
                         "ئ": "ي", "٠": "0", "١": "1", "٢": "2", "٣": "3", "٤": "4", "٥": "5",
                         "٦": "6", "٧": "7", "٨": "8", "٩": "9"})
ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
PREFIXES = ("وال", "بال", "كال", "فال", "لل", "ال")
SUFFIXES = ("ات", "ون", "ين", "ها", "ه", "ي")


def stop(msg):
    print(f"\nSTOPPED - {msg}")
    sys.exit(1)


def normalize(text) -> str:
    if not isinstance(text, str):
        text = "" if pd.isna(text) else str(text)
    text = DIACRITICS.sub("", text).translate(LETTERS).lower()
    return re.sub(r"\s+", " ", text)


def stem(word):
    for p in PREFIXES:
        if word.startswith(p) and len(word) - len(p) >= 3:
            word = word[len(p):]
            break
    if len(word) > 4:
        for s in SUFFIXES:
            if word.endswith(s) and len(word) - len(s) >= 3:
                word = word[: -len(s)]
                break
    return word


def keywords_of(text):
    words = re.findall(r"[\u0600-\u06FFa-z0-9]+", normalize(text))
    out = []
    for w in words:
        if w in STOPWORDS or len(w) <= 2:
            continue
        s = stem(w)
        if s not in out:
            out.append(s)
    return out


def split_words(cell):
    if not isinstance(cell, str) or not cell.strip():
        return []
    out = []
    for item in re.split(r"[،,\n]", cell):
        parts = [normalize(p).strip() for p in item.split("+")]
        parts = [p for p in parts if p]
        if parts:
            out.append(parts)
    return out


def find_col(df, name, required=True, label="file"):
    for c in df.columns:
        if str(c).strip().lower() == name.strip().lower():
            return c
    if required:
        stop(f"Column '{name}' not found in the {label}.\n   Columns found: {list(df.columns)[:40]}")
    return None


def read_any(path, label):
    p = Path(path)
    if not p.exists():
        stop(f"{label} not found:\n   {path}")
    if p.suffix.lower() == ".parquet":
        return pd.read_parquet(p)
    if p.suffix.lower() == ".csv":
        return pd.read_csv(p, dtype=str, encoding="utf-8-sig")
    return pd.read_excel(p, dtype=str)


def read_questions(path):
    raw = pd.read_excel(path, header=None, dtype=str)
    for i in range(min(15, len(raw))):
        if any(str(v).strip().lower() == QUESTION_COLUMN.lower() for v in raw.iloc[i]):
            df = raw.iloc[i + 1:].copy()
            df.columns = [str(v).strip() for v in raw.iloc[i]]
            break
    else:
        stop(f"No '{QUESTION_COLUMN}' header found in the questions file.")
    cq = find_col(df, QUESTION_COLUMN, label="questions file")
    ci = find_col(df, IDEAL_COLUMN, label="questions file")
    cm = find_col(df, MATCH_COLUMN, label="questions file")
    cr = find_col(df, REQUIRED_COLUMN, label="questions file")
    ce = find_col(df, EXCLUDED_COLUMN, label="questions file")
    df = df[df[cq].fillna("").str.strip() != ""].reset_index(drop=True)
    qs = []
    for i, r in df.iterrows():
        ideal = "" if pd.isna(r[ci]) else str(r[ci]).strip()
        qs.append({"no": i + 1, "question": str(r[cq]).strip(), "ideal": ideal,
                   "match": split_words(r[cm]), "required": split_words(r[cr]),
                   "excluded": split_words(r[ce]), "elements": split_elements(ideal),
                   "ideal_cannot": any(normalize(p) in normalize(ideal) for p in CANNOT_PHRASES)})
    return qs


def split_elements(ideal):
    """Ideal answer -> list of elements {name, keys}. Short fragments merge into the previous one."""
    parts = [p.strip(" -•:") for p in re.split(r"[.!؟?؛\n]|،\s*(?=\S+\s+\S+\s+\S+)", ideal or "")]
    els = []
    for p in parts:
        keys = keywords_of(p)
        if not keys:
            continue
        if len(keys) < 2 and els:
            els[-1]["keys"] += [k for k in keys if k not in els[-1]["keys"]]
            continue
        words = p.split()
        name = " ".join(words[:6]) + ("…" if len(words) > 6 else "")
        els.append({"name": name, "keys": keys})
    return els


# =============================================================================
# MATCHING (same rules as find_inquiries.py)
# =============================================================================
def match_inquiries(text, questions):
    n = len(text)
    cache = {}

    def has(part):
        if part not in cache:
            cache[part] = text.str.contains(part, regex=False).to_numpy()
        return cache[part]

    def hits(phrases):
        if not phrases:
            return np.zeros((n, 0), dtype=bool)
        cols = []
        for parts in phrases:
            m = has(parts[0]).copy()
            for p in parts[1:]:
                m &= has(p)
            cols.append(m)
        return np.column_stack(cols)

    Q = len(questions)
    score = np.zeros((n, Q), dtype=np.int64)
    for j, q in enumerate(questions):
        M, R, E = hits(q["match"]), hits(q["required"]), hits(q["excluded"])
        if not M.shape[1]:
            continue
        ok = M.any(1)
        if R.shape[1]:
            ok &= R.any(1)
        if E.shape[1]:
            ok &= ~E.any(1)
        lm = np.array([sum(len(p) for p in ph) for ph in q["match"]])
        words = M.sum(1) + (R.sum(1) if R.shape[1] else 0)
        chars = M @ lm
        score[ok, j] = words[ok] * 100000 + chars[ok]
    best = score.argmax(1)
    return np.where(score.max(1) > 0, best, -1)


# =============================================================================
# SCORING
# =============================================================================
def score_answer(answer_norm, elements):
    """Returns (quality %, list of covered flags)."""
    if not elements:
        return np.nan, []
    flags = []
    for el in elements:
        found = sum(1 for k in el["keys"] if k in answer_norm)
        flags.append(found / len(el["keys"]) >= ELEMENT_COVERED_AT)
    return 100.0 * sum(flags) / len(flags), flags


def r1(x):
    return None if x is None or (isinstance(x, float) and np.isnan(x)) else round(float(x), 1)


def forecast(values, months_ahead):
    """Linear trend on the available months; returns (forecast, low, high) lists (80% band)."""
    y = np.array([v for v in values if v is not None], dtype=float)
    x = np.array([i for i, v in enumerate(values) if v is not None], dtype=float)
    if len(y) < 4 or months_ahead <= 0:
        return [], [], []
    b, a = np.polyfit(x, y, 1)
    resid = y - (a + b * x)
    sd = float(np.std(resid, ddof=2)) if len(y) > 2 else 0.0
    last = len(values) - 1
    fc, lo, hi = [], [], []
    for h in range(1, months_ahead + 1):
        v = float(np.clip(a + b * (last + h), 0, 100))
        w = 1.28 * sd * np.sqrt(1 + h / len(y))
        fc.append(round(v, 1)); lo.append(round(max(0, v - w), 1)); hi.append(round(min(100, v + w), 1))
    return fc, lo, hi


def main():
    t0 = time.time()
    print("Reading files ...")
    tk = read_any(TICKETS_FILE, "Tickets file")
    questions = read_questions(QUESTIONS_FILE)
    need = {"INQUIRY_COLUMN": INQUIRY_COLUMN, "ANSWER_COLUMN": ANSWER_COLUMN,
            "CATEGORY_COLUMN": CATEGORY_COLUMN, "SUBMIT_COLUMN": SUBMIT_COLUMN}
    missing_cols = [f"{k} = \"{v}\"" for k, v in need.items() if find_col(tk, v, required=False) is None]
    if missing_cols:
        stop("These columns were not found in the tickets file:\n   " + "\n   ".join(missing_cols)
             + "\n\nColumns in the file:\n   " + "\n   ".join(map(str, tk.columns)))
    c_inq = find_col(tk, INQUIRY_COLUMN)
    c_ans = find_col(tk, ANSWER_COLUMN)
    c_cat = find_col(tk, CATEGORY_COLUMN)
    c_sub = find_col(tk, SUBMIT_COLUMN)
    c_emp = find_col(tk, EMPLOYEE_COLUMN, required=False)
    if c_emp is None:
        for hint in ["رقم الموظف", "الرقم الوظيفي", "Employee ID", "Employee Number", "Employee No"]:
            c_emp = find_col(tk, hint, required=False)
            if c_emp:
                break
    c_hrs = find_col(tk, HOURS_COLUMN, required=False)
    c_tid = find_col(tk, TICKET_COLUMN, required=False)
    c_res = find_col(tk, RESOLVED_COLUMN, required=False)
    c_reo = find_col(tk, REOPEN_COLUMN, required=False)
    total_tickets = len(tk)
    print(f"   {total_tickets:,} tickets | {len(questions)} question types")
    if c_emp is None:
        print(f"   NOTE: '{EMPLOYEE_COLUMN}' not found - repeat inquiries cannot be measured.")

    # ---------------- matching ----------------
    print("Matching inquiries to question types ...")
    assigned = match_inquiries(tk[c_inq].map(normalize), questions)
    df = tk[assigned >= 0].copy()
    df["_q"] = assigned[assigned >= 0]
    if df.empty:
        stop("No inquiry matched any question type. Check the keyword columns.")
    df["_cat"] = df[c_cat].fillna("غير محدد").astype(str).str.strip().replace("", "غير محدد")
    df["_date"] = pd.to_datetime(df[c_sub], errors="coerce", dayfirst=False)
    df = df[df["_date"].notna()].copy()
    year = int(df["_date"].dt.year.mode().iloc[0])
    df = df[df["_date"].dt.year == year].copy()
    df["_m"] = df["_date"].dt.month - 1
    print(f"   {len(df):,} inquiries matched in {year} ({len(df) / max(total_tickets, 1):.0%} of all tickets)")

    # ---------------- scoring ----------------
    print("Scoring answers ...")
    ans_norm = df[c_ans].map(normalize)
    stemmed = ans_norm.map(lambda t: " ".join(stem(w) for w in re.findall(r"[\u0600-\u06FFa-z0-9]+", t)))
    cannot_norm = [normalize(p) for p in CANNOT_PHRASES]
    quality, cannot, cov = [], [], []
    for qi, a_raw, a_st in zip(df["_q"], ans_norm, stemmed):
        q = questions[qi]
        is_cannot = (not q["ideal_cannot"]) and any(p in a_raw for p in cannot_norm)
        cannot.append(is_cannot)
        val, flags = score_answer(a_raw + " " + a_st, q["elements"])
        quality.append(np.nan if is_cannot else val)
        cov.append(flags)
    df["_quality"] = quality
    df["_cannot"] = cannot
    df["_flags"] = cov
    df["_level"] = np.select([df["_quality"] >= HIGH_LEVEL, df["_quality"] >= LOW_LEVEL, df["_quality"].notna()],
                             ["high", "mid", "low"], default="")

    # reopen / closing time
    df["_reopen"] = pd.to_numeric(df[c_reo], errors="coerce").fillna(0) > 0 if c_reo else False
    if c_res:
        df["_hours"] = (pd.to_datetime(df[c_res], errors="coerce") - df["_date"]).dt.total_seconds() / 3600
    elif c_hrs:
        df["_hours"] = pd.to_numeric(df[c_hrs], errors="coerce")
    else:
        df["_hours"] = np.nan

    # ratings
    has_ratings = bool(RATINGS_FILE.strip())
    df["_rating"] = np.nan
    df["_reason"] = ""
    if has_ratings:
        rt = read_any(RATINGS_FILE, "Ratings file")
        r_id = find_col(rt, RATING_TICKET_COLUMN, label="ratings file")
        r_rt = find_col(rt, RATING_COLUMN, label="ratings file")
        r_cc = find_col(rt, COMMENT_CATEGORY_COLUMN, required=False)
        if c_tid is None:
            stop(f"'{TICKET_COLUMN}' not found in the tickets file - needed to link the ratings.")
        rt = rt.drop_duplicates(subset=[r_id], keep="last")
        key = lambda s: s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        rmap_norm = {normalize(k).strip(): v for k, v in RATING_MAP.items()}
        def to_score(v):
            t = normalize(v).strip()
            if t in rmap_norm:
                return rmap_norm[t]
            try:
                return float(t)
            except ValueError:
                return np.nan
        scores = rt[r_rt].map(to_score)
        unknown = rt[r_rt][scores.isna() & rt[r_rt].notna() & (rt[r_rt].astype(str).str.strip() != "")].astype(str).unique()
        if len(unknown):
            print(f"   NOTE: satisfaction values not recognised (ignored): {list(unknown)[:10]} - add them to RATING_MAP")
        rmap = dict(zip(key(rt[r_id]), scores))
        df["_rating"] = key(df[c_tid]).map(rmap)
        if r_cc:
            cmap = dict(zip(key(rt[r_id]), rt[r_cc].fillna("").astype(str).str.strip()))
            df["_reason"] = key(df[c_tid]).map(cmap).fillna("")
        print(f"   Ratings linked to {int(df['_rating'].notna().sum()):,} inquiries")

    # repeat over the year: per employee + question type
    if c_emp:
        df["_emp"] = df[c_emp].astype(str).str.strip()
        df = df.sort_values("_date")
        g = df.groupby(["_emp", "_q"])
        df["_pair_n"] = g["_q"].transform("size")
        df["_followups"] = df["_pair_n"] - g.cumcount() - 1
    else:
        df["_emp"] = ""
        df["_pair_n"] = 1
        df["_followups"] = 0

    # ---------------- metrics ----------------
    print("Calculating indicators ...")
    last_m = int(df["_m"].max())
    first_m = int(df["_m"].min())
    n_ahead = 11 - last_m
    month_idx = list(range(first_m, last_m + 1))

    def pairs_of(d):
        if not c_emp:
            return pd.Series(dtype=int)
        return d.groupby(["_emp", "_q"]).size()

    def scope_metrics(d, with_fc=True):
        scored = d[d["_quality"].notna()]
        qv = float(scored["_quality"].mean()) if len(scored) else None
        pr = pairs_of(d)
        rep = float((pr >= 2).mean() * 100) if len(pr) else None
        dist = [int((pr == 1).sum()), int((pr == 2).sum()), int((pr == 3).sum()), int((pr >= 4).sum())] if len(pr) else [0, 0, 0, 0]
        dist = [round(100 * x / max(sum(dist), 1), 1) for x in dist]
        rated = d[d["_rating"].notna()]
        sat = float(rated["_rating"].mean()) if len(rated) else None
        series = []
        for m in month_idx:
            s = scored[scored["_m"] == m]["_quality"]
            series.append(round(float(s.mean()), 1) if len(s) >= 5 else None)
        fc, lo, hi = forecast(series, n_ahead) if with_fc else ([], [], [])
        lv = scored["_level"].value_counts()
        tot = max(len(scored), 1)
        levels = [round(100 * lv.get(k, 0) / tot) for k in ("high", "mid", "low")]
        levels[1] = 100 - levels[0] - levels[2]
        sat_lv = []
        for k in ("low", "mid", "high"):
            s = rated[rated["_level"] == k]["_rating"]
            sat_lv.append(r1(s.mean()) if len(s) >= 5 else None)
        reasons = []
        rr = d[(d["_reason"] != "") & ((d["_rating"] <= 1) | d["_rating"].isna())]["_reason"]
        if len(rr):
            vc = rr.value_counts()
            reasons = [[k, round(100 * v / len(rr), 1)] for k, v in vc.head(5).items()]
        return {"q": r1(qv), "vol": int(len(d)), "scored": int(len(scored)), "rep": r1(rep), "dist": dist,
                "sat": r1(sat), "rated": int(len(rated)), "reopen": r1(d["_reopen"].mean() * 100),
                "series": series, "fc": fc, "lo": lo, "hi": hi, "levels": levels, "satLv": sat_lv,
                "reasons": reasons, "cannot": int(d["_cannot"].sum())}

    overall = scope_metrics(df)

    # categories
    cats = []
    for name, d in df.groupby("_cat"):
        m = scope_metrics(d)
        m["name"] = name
        m["ranked"] = len(d) >= MIN_CATEGORY
        # most missing elements in this category (elements of its question types)
        miss = {}
        for qi, dq in d[d["_quality"].notna()].groupby("_q"):
            els = questions[qi]["elements"]
            if not els or len(dq) < 5:
                continue
            arr = np.array([f for f in dq["_flags"] if len(f) == len(els)])
            if not len(arr):
                continue
            for k, el in enumerate(els):
                acc = miss.setdefault(el["name"], [0.0, 0])
                acc[0] += float((1 - arr[:, k]).sum()) * 100
                acc[1] += len(arr)
        miss = {k: v[0] / v[1] for k, v in miss.items() if v[1]}
        m["missing"] = [[k, round(v, 1)] for k, v in sorted(miss.items(), key=lambda x: -x[1])[:5]]
        # question types in this category
        qs_out = []
        for qi, dq in d.groupby("_q"):
            q = questions[qi]
            qm = scope_metrics(dq, with_fc=False)
            els = q["elements"]
            elements = []
            arr = np.array([f for f in dq[dq["_quality"].notna()]["_flags"] if len(f) == len(els)])
            if len(arr) and els:
                elements = sorted([[el["name"], round(100 * arr[:, k].mean(), 1)] for k, el in enumerate(els)], key=lambda x: x[1])
            qm.update({"t": q["question"], "no": q["no"], "elements": elements,
                       "m": elements[0][0] if elements else "—"})
            qs_out.append(qm)
        m["qs"] = sorted(qs_out, key=lambda x: -x["vol"])
        m["comp"] = m["q"]
        cats.append(m)
    cats.sort(key=lambda c: -c["vol"])

    # drivers: follow-up tickets per 100 tickets
    drivers = []
    base_mask = df["_level"] == "high"
    if has_ratings:
        base_mask &= df["_rating"] == 5
    def per100(mask):
        sub = df[mask]
        return (round(float(sub["_followups"].mean() * 100)), int(len(sub))) if len(sub) >= 30 else (None, len(sub))
    factors = [("تقييم منخفض", (df["_rating"] <= 2), "stars2"),
               ("إجابة غير مكتملة", df["_level"] == "low", ""),
               ("إعادة فتح الطلب", df["_reopen"].astype(bool), ""),
               ("إجابة بتعذر الخدمة نظاميًا", df["_cannot"], ""),
               ("إغلاق الطلب خلال أقل من يوم", df["_hours"] < 24, "")]
    if c_emp:
        for label, mask, kind in factors:
            v, n = per100(mask.fillna(False) if hasattr(mask, "fillna") else mask)
            if v is not None:
                drivers.append({"name": label, "v": v, "n": n, "kind": kind})
        drivers.sort(key=lambda x: -x["v"])
        bv, bn = per100(base_mask)
        baseline = {"name": "طلب مكتمل" if not has_ratings else "طلب مكتمل", "v": bv, "n": bn,
                    "kind": "stars5" if has_ratings else ""} if bv is not None else None
    else:
        baseline = None

    # ---------------- additional analyses (search) ----------------
    insights = {}
    ex = df[df["_cannot"]]
    if len(ex):
        top = ex["_cat"].value_counts(normalize=True).head(2)
        insights["cannot"] = {
            "label": "الإجابات المستبعدة لتعذر الخدمة نظاميًا", "sub": "نطاق التحليل",
            "title": f"استبعاد {len(ex) / len(df) * 100:.1f}% من الاستفسارات من مؤشر الجودة لتعذر الخدمة نظاميًا",
            "body": f"تضمنت {len(ex):,} إجابة إفادة بتعذر تقديم الخدمة نظاميًا، فاستُبعدت من احتساب مؤشر الجودة حتى لا تنعكس على أداء الفرق، باستثناء أنواع الاستفسارات التي تتضمن إجابتها النموذجية ذلك.",
            "chips": [["الإجابات المستبعدة", f"{len(ex):,}"], ["نسبتها", f"{len(ex) / len(df) * 100:.1f}%"]]
                     + [[f"فئة {k}", f"{v * 100:.0f}%"] for k, v in top.items()]}
    if df["_hours"].notna().sum() > 50:
        b = pd.cut(df["_hours"], [-1, 24, 48, 72, 1e9], labels=["أقل من يوم", "يوم إلى يومين", "يومان إلى ثلاثة", "أكثر من ثلاثة أيام"])
        qb = df[df["_quality"].notna()].groupby(b[df["_quality"].notna()], observed=False)["_quality"].mean()
        qb = qb.dropna()
        if len(qb) >= 2:
            insights["speed"] = {
                "label": "مدة الإغلاق مقابل جودة الإجابة", "sub": "علاقة",
                "title": f"مؤشر الجودة للطلبات المغلقة خلال أقل من يوم {qb.iloc[0]:.0f}%، مقابل {qb.max():.0f}% أعلى فئة زمنية",
                "body": "يقارن هذا التحليل متوسط مؤشر الجودة حسب المدة بين تقديم الطلب وإغلاقه، للتحقق من أن سرعة الإغلاق لا تكون على حساب اكتمال الإجابة.",
                "chips": [[str(k), f"{v:.0f}%"] for k, v in qb.items()][:4]}
    vol_m = df.groupby("_m").size()
    if len(vol_m) >= 3:
        topm = vol_m.sort_values(ascending=False)
        insights["peak"] = {
            "label": "أشهر الذروة في حجم الاستفسارات", "sub": "الحجم",
            "title": f"{MONTHS_AR[topm.index[0]]} و{MONTHS_AR[topm.index[1]]} الأعلى حجمًا للاستفسارات خلال الفترة",
            "body": f"بلغ المتوسط الشهري للاستفسارات المحللة {vol_m.mean():,.0f} استفسارًا، وسجل شهر {MONTHS_AR[topm.index[-1]]} أدنى حجم.",
            "chips": [[MONTHS_AR[topm.index[0]], f"{topm.iloc[0]:,}"], [MONTHS_AR[topm.index[1]], f"{topm.iloc[1]:,}"],
                      ["المتوسط الشهري", f"{vol_m.mean():,.0f}"], ["أدنى شهر", f"{MONTHS_AR[topm.index[-1]]} {topm.iloc[-1]:,}"]]}
    if overall["fc"]:
        insights["fc"] = {
            "label": f"منهجية توقع {'و'.join(MONTHS_AR[last_m + 1:])}" if n_ahead <= 2 else "منهجية التوقع", "sub": "المنهجية",
            "title": f"يُتوقع أن يبلغ مؤشر الجودة {overall['fc'][-1]:.0f}% في {MONTHS_AR[11]}",
            "body": f"بُني التوقع على الاتجاه الخطي لمؤشر الجودة خلال الأشهر الفعلية المتاحة، ويُعرض بنطاق ثقة 80%، ويُستبدل بالقيمة الفعلية عند ورود بيانات الشهر.",
            "chips": [[f"توقع {MONTHS_AR[last_m + 1 + i]}", f"{v:.0f}%"] for i, v in enumerate(overall["fc"][:2])]
                     + [[f"نطاق {MONTHS_AR[11]}", f"{overall['lo'][-1]:.0f}% – {overall['hi'][-1]:.0f}%"], ["أساس التوقع", f"{len(month_idx)} أشهر"]]}
    matched_types = sorted(set(int(x) for x in df["_q"]))
    unmatched = [q for i, q in enumerate(questions) if i not in matched_types]
    insights["gap"] = {
        "label": "أنواع الاستفسارات غير المشمولة في التحليل", "sub": "نطاق التحليل",
        "title": f"{len(unmatched)} نوعًا من أصل {len(questions)} لم تُطابق أي استفسار خلال الفترة",
        "body": ("أبرزها: " + "، ".join(q["question"].rstrip("؟?") for q in unmatched[:4]) + ". وتُراجع كلماتها المفتاحية دوريًا.") if unmatched else "طابقت جميع أنواع الاستفسارات استفسارات خلال الفترة.",
        "chips": [["الأنواع المحللة", str(len(matched_types))], ["غير المشمولة", str(len(unmatched))],
                  ["تغطية الدليل", f"{len(matched_types) / len(questions) * 100:.0f}%"], ["تغطية الاستفسارات", f"{len(df) / total_tickets * 100:.0f}%"]]}
    if has_ratings and all(v is not None for v in overall["satLv"]):
        insights["sat"] = {
            "label": "رضا المستفيد حسب مستوى جودة الإجابة", "sub": "علاقة",
            "title": f"بلغ متوسط تقييم المستفيدين {overall['satLv'][2]:.1f} من 5 للإجابات مرتفعة الجودة، مقابل {overall['satLv'][0]:.1f} للإجابات منخفضة الجودة",
            "body": "يقارن هذا التحليل متوسط تقييم المستفيدين حسب مستوى جودة الإجابة، للتحقق من اتساق مؤشر الجودة مع تجربة المستفيد.",
            "chips": [["جودة مرتفعة", f"{overall['satLv'][2]:.1f} / 5"], ["جودة متوسطة", f"{overall['satLv'][1]:.1f} / 5"],
                      ["جودة منخفضة", f"{overall['satLv'][0]:.1f} / 5"], ["الطلبات المقيَّمة", f"{overall['rated']:,}"]]}
        ro = df[df["_rating"].notna()].groupby("_reopen")["_rating"].mean()
        if len(ro) == 2:
            insights["reopen"] = {
                "label": "أثر إعادة فتح الطلب على رضا المستفيد", "sub": "علاقة",
                "title": f"متوسط رضا المستفيد {ro[True]:.1f} من 5 في الطلبات المعاد فتحها، مقابل {ro[False]:.1f} للطلبات المغلقة من المرة الأولى",
                "body": "يقارن هذا التحليل متوسط تقييم المستفيدين بين الطلبات التي أُعيد فتحها والطلبات المغلقة من المرة الأولى.",
                "chips": [["الطلبات المعاد فتحها", f"{overall['reopen']:.1f}%"], ["رضاها", f"{ro[True]:.1f}"], ["رضا المغلقة أولًا", f"{ro[False]:.1f}"]]}
    if c_emp:
        d = overall["dist"]
        insights["repdist"] = {
            "label": "عدد مرات تقديم الاستفسار نفسه خلال العام", "sub": "التكرار",
            "title": f"قدّم {overall['rep']:.0f}% من المستفيدين الاستفسار نفسه مرتين أو أكثر خلال العام",
            "body": "يُحتسب التكرار لكل مستفيد ونوع استفسار خلال العام.",
            "chips": [["مرة واحدة", f"{d[0]:.0f}%"], ["مرتان", f"{d[1]:.0f}%"], ["ثلاث مرات", f"{d[2]:.0f}%"], ["أربع مرات فأكثر", f"{d[3]:.0f}%"]]}

    # ---------------- build dashboard ----------------
    def ar_digits(n):
        return str(n).translate(str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩"))
    period = f"{MONTHS_AR[first_m]} {ar_digits(year)} – {MONTHS_AR[last_m]} {ar_digits(year)}"
    data = {
        "meta": {"period": period, "dept": DEPARTMENT_NAME, "total": total_tickets, "analyzed": int(len(df)),
                 "types": len(questions), "typesMatched": len(matched_types), "target": TARGET, "acceptable": ACCEPTABLE,
                 "months": [MONTHS_AR[m] for m in month_idx] + MONTHS_AR[last_m + 1:], "nActual": len(month_idx),
                 "hasRatings": has_ratings, "hasEmp": bool(c_emp), "minCat": MIN_CATEGORY,
                 "high": HIGH_LEVEL, "low": LOW_LEVEL},
        "overall": overall, "cats": cats, "drivers": drivers, "baseline": baseline, "insights": insights,
    }
    out_dir = Path(OUTPUT_FOLDER) if OUTPUT_FOLDER else Path(TICKETS_FILE).parent
    out_dir.mkdir(parents=True, exist_ok=True)
    html_path = out_dir / "answer_quality_dashboard.html"
    payload = json.dumps(data, ensure_ascii=False, default=lambda o: o.item() if hasattr(o, "item") else str(o)).replace("</", "<\\/")
    html_path.write_text(HTML.replace("__DATA__", payload), encoding="utf-8")

    # ---------------- scored inquiries (one sheet) ----------------
    xl = df.drop(columns=[c for c in df.columns if c.startswith("_")]).copy()
    xl.insert(0, "Question", [questions[i]["question"] for i in df["_q"]])
    xl.insert(1, "Quality %", df["_quality"].round(0).values)
    xl.insert(2, "Excluded (not possible)", np.where(df["_cannot"], "Yes", ""))
    xl.insert(3, "Missing elements", [
        "، ".join(el["name"] for el, f in zip(questions[qi]["elements"], fl) if not f) if fl else ""
        for qi, fl in zip(df["_q"], df["_flags"])])
    xl = xl.map(lambda v: ILLEGAL.sub("", v) if isinstance(v, str) else v)
    xl_path = out_dir / "scored_inquiries.xlsx"
    try:
        xl.to_excel(xl_path, sheet_name="Scored Inquiries", index=False)
    except PermissionError:
        stop(f"Cannot save. Is the file open in Excel?\n   {xl_path}")

    print("\n" + "=" * 54)
    print(f"  Inquiries analysed        : {len(df):,} of {total_tickets:,} ({len(df) / total_tickets:.0%})")
    print(f"  Question types matched    : {len(matched_types)} of {len(questions)}")
    print(f"  Quality index             : {overall['q']}%")
    print(f"  Excluded (not possible)   : {overall['cannot']:,}")
    if c_emp:
        print(f"  Repeat (same inquiry/yr)  : {overall['rep']}%")
    if has_ratings:
        print(f"  Satisfaction              : {overall['sat']} / 5 ({overall['rated']:,} rated)")
    print("=" * 54)
    print(f"Dashboard : {html_path}")
    print(f"Records   : {xl_path}   ({time.time() - t0:.0f}s)")


# =============================================================================
# DASHBOARD TEMPLATE
# =============================================================================
HTML = r"""<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>لوحة مؤشرات جودة الإجابات على استفسارات رأس المال البشري</title>
<style>
*{box-sizing:border-box}
body{margin:0;background:#F2F4F7;color:#152446;font-family:'Somar','IBM Plex Sans Arabic',Tahoma,sans-serif}
button{font-family:inherit}
@keyframes ring{0%{background-position:0% 50%}100%{background-position:200% 50%}}
.ring{background:linear-gradient(90deg,#009ADE,#4BB8C4,#89A563,#4BB8C4,#009ADE);background-size:200% 100%;animation:ring 7s linear infinite}
.sig{background:linear-gradient(270deg,#009ADE,#4BB8C4 50%,#89A563)}
header{position:relative;color:#fff;padding:24px 0 26px;z-index:30;background:radial-gradient(60% 120% at 10% 110%,rgba(137,165,99,.22) 0%,rgba(75,184,196,.12) 40%,transparent 70%),radial-gradient(120% 160% at 85% 0%,#1B3566 0%,#152446 45%,#0E1A35 100%)}
.wrap{max-width:1320px;margin:0 auto;padding:0 32px}
main.wrap{padding:16px 32px 32px;display:flex;flex-direction:column;gap:12px}
.tile{background:#fff;border:1px solid #E3E7EE;border-radius:14px;padding:14px 16px;display:flex;flex-direction:column;gap:6px;transition:box-shadow .2s;min-width:0}
.tile:hover{box-shadow:0 8px 24px rgba(21,36,70,.08)}
.t{margin:0;font-size:19px;font-weight:700;line-height:1.45}
.st{margin:0;font-size:13px;font-weight:500;line-height:1.6;color:#156342}
.note{font-size:11px;color:#5A6478}
.grid2{display:grid;gap:12px;grid-template-columns:repeat(2,minmax(0,1fr))}
.grid3{display:grid;gap:12px;grid-template-columns:repeat(3,minmax(0,1fr))}
.row{display:flex;align-items:center;gap:10px}
.track{flex:1;background:#EEF1F5;border-radius:4px;height:8px;display:block}
.fill{display:block;height:8px;border-radius:4px}
.hit{cursor:pointer;transition:opacity .2s,transform .2s;background:transparent;border:0;color:#152446;text-align:right}
.hit:hover{transform:translateY(-1px)}
.chip{background:#fff;border:1px solid #D5DBE5;border-radius:999px;padding:9px 16px;font-size:14px;color:#152446;cursor:pointer;min-height:40px}
.legend{display:flex;flex-wrap:wrap;gap:16px;font-size:11px;color:#4A5568}
.legend span{display:flex;align-items:center;gap:6px}
.kpi{position:relative;overflow:hidden;background:#fff;border:1px solid #E3E7EE;border-radius:14px;padding:12px 14px;display:flex;flex-direction:column;gap:2px}
.kpi .bar{position:absolute;top:0;right:0;left:0;height:3px}
input::placeholder{color:#9FB2CF}input:focus{outline:none}
bdi{unicode-bidi:isolate}
.empty{font-size:12px;color:#5A6478;padding:12px 0}
@media (max-width:900px){.grid2,.grid3,.kpis,.r1,.r2{grid-template-columns:minmax(0,1fr)!important}}
</style>
</head>
<body>
<header>
  <div class="sig" aria-hidden="true" style="position:absolute;right:0;bottom:0;left:0;height:4px"></div>
  <div class="wrap" style="position:relative;display:flex;flex-direction:column;gap:22px">
    <div style="display:flex;flex-direction:column;gap:6px">
      <p style="margin:0;font-size:13px;color:#8FD3F5">إجابتك أقرب | إصدار تقييم الإجابات</p>
      <h1 style="margin:0;font-size:30px;font-weight:700">لوحة مؤشرات جودة الإجابات على استفسارات رأس المال البشري</h1>
      <p id="dept" style="margin:0;font-size:16px;color:#E3E9F4"></p>
      <p id="period" style="margin:0;font-size:14px;color:#C9D3E6"></p>
    </div>
    <div style="position:relative;max-width:860px;width:100%;align-self:center">
      <div class="ring" style="border-radius:999px;padding:1.5px">
        <div style="display:flex;align-items:center;gap:12px;background:#13264A;border-radius:999px;padding:0 22px;min-height:54px">
          <svg aria-hidden="true" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#8FD3F5" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>
          <label for="q" style="position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0)">استعلام في بيانات اللوحة</label>
          <input id="q" type="search" autocomplete="off" placeholder="ابحث باسم الفئة أو الاستفسار أو التحليل المطلوب" style="flex:1;background:transparent;border:0;color:#fff;font-family:inherit;font-size:16px;min-height:50px">
        </div>
      </div>
      <div id="sug"></div>
    </div>
  </div>
</header>
<main class="wrap" id="app"></main>

<script>
const D = __DATA__;
const S = { cat: null, q: null, insight: null, query: '', open: false };
const GREEN = '#156342', LBLUE = '#009ADE', ORANGE = '#E07A2E', NAVY = '#152446', DBLUE = '#002F87';
const M = D.meta;
const fmt = n => (n == null ? '—' : Math.round(n).toLocaleString('en-US'));
const pct = n => (n == null ? '—' : Math.round(n) + '%');
const one = n => (n == null ? '—' : (Math.round(n * 10) / 10).toFixed(1));
const esc = t => String(t == null ? '' : t).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const col = v => v == null ? '#B8C3D6' : (v < M.acceptable ? ORANGE : (v < M.target ? LBLUE : GREEN));
const norm = t => (t || '').replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي').replace(/[ً-ْ]/g, '').toLowerCase();
const nm = t => '<bdi>' + esc(t) + '</bdi>';
const pts_ = n => n === 1 ? 'نقطة واحدة' : (n === 2 ? 'نقطتين' : (n <= 10 ? n + ' نقاط' : n + ' نقطة'));
const ORD = ['الأولى','الثانية','الثالثة','الرابعة','الخامسة','السادسة','السابعة','الثامنة','التاسعة','العاشرة'];
const set = o => { Object.assign(S, o); render(); };

document.getElementById('dept').textContent = M.dept;
document.getElementById('period').textContent = M.period;

function scope() {
  if (S.cat !== null && S.q !== null) return D.cats[S.cat].qs[S.q];
  if (S.cat !== null) return D.cats[S.cat];
  return D.overall;
}
const ranked = () => D.cats.map((c, i) => ({ c, i })).filter(x => x.c.ranked && x.c.q != null);

// ---------------- KPIs ----------------
function kpis(sc) {
  const cat = S.cat !== null;
  const diff = cat && sc.q != null && D.overall.q != null ? Math.round(sc.q - D.overall.q) : null;
  const k = [
    { l: 'مؤشر الجودة', v: pct(sc.q), s: cat ? (diff === 0 ? 'مماثل لمتوسط الهيئة' : (diff > 0 ? '▲ أعلى' : '▼ أدنى') + ' من متوسط الهيئة بـ' + pts_(Math.abs(diff))) : 'المستهدف ' + M.target + '%', c: cat && diff < 0 ? '#B4531F' : GREEN },
    { l: 'الاستفسارات المحللة', v: fmt(sc.vol), s: cat ? (Math.round(sc.vol / M.analyzed * 1000) / 10) + '% من إجمالي الاستفسارات المحللة' : Math.round(M.analyzed / M.total * 100) + '% من إجمالي ' + fmt(M.total) + ' استفسارًا، وتغطي ' + M.typesMatched + ' نوعًا من أصل ' + M.types, c: '#5A6478' },
    { l: 'رضا المستفيد', v: sc.sat == null ? '—' : one(sc.sat) + ' من 5', s: M.hasRatings ? 'متوسط ' + fmt(sc.rated) + ' تقييمًا' : 'لم تُربط بيانات التقييمات', c: '#5A6478' },
    { l: 'تكرار الاستفسار خلال العام', v: pct(sc.rep), s: 'من المستفيدين قدّموا الاستفسار نفسه مرتين أو أكثر', c: '#5A6478' },
    { l: 'إعادة فتح الطلب', v: sc.reopen == null ? '—' : one(sc.reopen) + '%', s: 'من الطلبات المغلقة', c: '#5A6478' }
  ];
  return '<div class="kpis" style="display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:12px">' + k.map(x =>
    `<div class="kpi"><div class="bar sig"></div><span style="font-size:13px;color:#5A6478">${x.l}</span><span style="font-size:24px;font-weight:700">${x.v}</span><span style="font-size:12px;color:${x.c}">${x.s}</span></div>`).join('') + '</div>';
}

// ---------------- Trend ----------------
function trend(sc) {
  const nA = M.nActual, months = M.months, N = months.length;
  const act = sc.series || [], fc = sc.fc || [], lo = sc.lo || [], hi = sc.hi || [];
  const all = act.concat(fc).concat(lo).concat(hi).concat(D.overall.series).filter(v => v != null);
  if (!all.length) return `<section class="tile"><h2 class="t">مؤشر الجودة الشهري</h2><p class="empty">لا تتوفر بيانات شهرية كافية.</p></section>`;
  let mn = Math.floor(Math.min(...all, M.target) / 10) * 10, mx = Math.ceil(Math.max(...all, M.target) / 10) * 10;
  if (mx - mn < 30) mn = Math.max(0, mx - 30);
  const step = N > 1 ? 760 / (N - 1) : 0;
  const X = i => 790 - i * step, Y = v => 190 - (v - mn) / (mx - mn) * 160;
  const line = (arr, off) => { let out = [], seg = []; arr.forEach((v, k) => { if (v == null) { if (seg.length) out.push(seg); seg = []; } else seg.push(X(off + k).toFixed(1) + ',' + Y(v).toFixed(1)); }); if (seg.length) out.push(seg); return out.map(s => `<polyline points="${s.join(' ')}" fill="none" stroke="%C" stroke-width="%W" stroke-linejoin="round"/>`).join(''); };
  let g = '';
  for (let k = 0; k <= 4; k++) { const v = mn + (mx - mn) * k / 4, y = Y(v); g += `<line x1="10" y1="${y}" x2="800" y2="${y}" stroke="#EEF1F5"/><text x="808" y="${y + 4}" font-size="11" fill="#5A6478">${Math.round(v)}%</text>`; }
  let zone = '';
  if (fc.length) {
    const zx = X(nA - 0.5);
    zone = `<defs><linearGradient id="fz" x1="1" x2="0" y1="0" y2="1"><stop offset="0" stop-color="#152446"/><stop offset=".6" stop-color="#1B3A63"/><stop offset="1" stop-color="#2A6A73"/></linearGradient><linearGradient id="sg" x1="1" x2="0"><stop offset="0" stop-color="#009ADE"/><stop offset=".5" stop-color="#4BB8C4"/><stop offset="1" stop-color="#89A563"/></linearGradient></defs>
      <rect x="8" y="14" width="${zx - 8}" height="186" rx="12" fill="url(#fz)"/><rect x="20" y="14" width="${zx - 32}" height="3" rx="1.5" fill="url(#sg)"/>
      <text x="${(zx + 8) / 2}" y="34" font-size="11" font-weight="600" fill="#8FD3F5" text-anchor="middle">متوقع</text>`;
  }
  const ty = Y(M.target);
  let svg = g + zone + `<line x1="${fc.length ? X(nA - 0.5) : 10}" y1="${ty}" x2="800" y2="${ty}" stroke="#9AA5B8" stroke-dasharray="5 5"/><text x="794" y="${ty - 6}" font-size="11" fill="#5A6478" direction="rtl" text-anchor="start">المستهدف ${M.target}%</text>`;
  if (S.cat !== null) svg += line(D.overall.series, 0).replace(/%C/g, '#9AA5B8').replace(/%W/g, '2');
  svg += line(act, 0).replace(/%C/g, NAVY).replace(/%W/g, '3');
  const lastI = act.map((v, i) => v == null ? -1 : i).filter(i => i >= 0).pop();
  if (lastI != null) svg += `<circle cx="${X(lastI)}" cy="${Y(act[lastI])}" r="5" fill="${NAVY}"/><text x="${X(lastI)}" y="${Y(act[lastI]) - 12}" font-size="11" font-weight="700" fill="${NAVY}" text-anchor="middle">${pct(act[lastI])}</text>`;
  if (fc.length && lastI === nA - 1) {
    svg += `<polyline points="${[X(nA - 1) + ',' + Y(act[nA - 1])].concat(fc.map((v, k) => X(nA + k) + ',' + Y(v))).join(' ')}" fill="none" stroke="#fff" stroke-width="2" stroke-opacity=".85"/>`;
    fc.forEach((v, k) => { const x = X(nA + k); svg += `<rect x="${x - 7}" y="${Y(hi[k])}" width="14" height="${Math.max(2, Y(lo[k]) - Y(hi[k]))}" rx="7" fill="#fff" opacity=".2"/><circle cx="${x}" cy="${Y(v)}" r="5" fill="#fff"/><text x="${x}" y="${Y(hi[k]) - 6}" font-size="11" font-weight="700" fill="#fff" text-anchor="middle">${pct(v)}</text>`; });
  }
  months.forEach((m, i) => { svg += `<text x="${X(i)}" y="214" font-size="11" text-anchor="middle" fill="${i >= nA ? NAVY : '#4A5568'}" font-weight="${i >= nA ? 700 : 400}">${m}</text>`; });
  const vals = act.filter(v => v != null);
  const firstI = act.findIndex(v => v != null);
  const subj = S.q !== null ? 'مؤشر الاستفسار' : (S.cat !== null ? 'مؤشر الفئة' : 'مؤشر الجودة');
  let title = '';
  if (vals.length >= 2) {
    title = (act[lastI] >= act[firstI] ? 'ارتفاع ' : 'انخفاض ') + subj + ' من ' + pct(act[firstI]) + ' في ' + months[firstI] + ' إلى ' + pct(act[lastI]) + ' في ' + months[lastI];
    if (fc.length && lastI === nA - 1) title += '، ويُتوقع أن يبلغ ' + pct(fc[fc.length - 1]) + ' في ' + months[N - 1] + ' ضمن نطاق يتراوح بين ' + pct(lo[lo.length - 1]) + ' و' + pct(hi[hi.length - 1]);
  } else title = 'لا تتوفر قيم شهرية كافية لهذا النطاق';
  const name = 'مؤشر الجودة الشهري: القيم الفعلية' + (fc.length ? ' والمتوقعة' : '') + (S.q !== null ? ' للاستفسار المحدد' : (S.cat !== null ? ' لفئة ' : ''));
  return `<section class="tile"><h2 class="t">${name}${S.cat !== null && S.q === null ? '<bdi>' + esc(D.cats[S.cat].name) + '</bdi>' : ''}</h2><p class="st">${esc(title)}</p>
    <svg role="img" aria-label="مؤشر الجودة الشهري" viewBox="0 0 840 222" style="width:100%;height:auto;direction:ltr;margin-top:4px">${svg}</svg>
    <div class="legend"><span><span style="width:18px;height:3px;background:${NAVY};border-radius:2px"></span>القيم الفعلية</span>${fc.length ? '<span><span style="width:14px;height:14px;border-radius:4px;background:linear-gradient(135deg,#152446,#2A6A73)"></span>القيم المتوقعة ونطاقها</span>' : ''}<span><span style="width:18px;border-top:2px dashed #9AA5B8"></span>المستهدف</span>${S.cat !== null ? '<span><span style="width:18px;height:2px;background:#9AA5B8"></span>متوسط الهيئة</span>' : ''}</div></section>`;
}

// ---------------- Ranking ----------------
function ranking() {
  const r = ranked().sort((a, b) => b.c.q - a.c.q);
  if (!r.length) return `<section class="tile"><h2 class="t">مؤشر الجودة حسب الفئة</h2><p class="empty">لا توجد فئات تتجاوز ${M.minCat} استفسارًا.</p></section>`;
  let sub;
  if (S.cat !== null) {
    const pos = r.findIndex(x => x.i === S.cat) + 1, c = D.cats[S.cat];
    sub = pos ? 'تحتل الفئة المرتبة ' + (ORD[pos - 1] || pos) + ' من أصل ' + r.length + (r.length <= 10 ? ' فئات' : ' فئة') + ' بمؤشر ' + pct(c.q) + (c.q < M.acceptable ? '، دون الحد المقبول البالغ ' + M.acceptable + '%' : '') : 'لا تُصنَّف الفئة لقلة استفساراتها';
  } else {
    const low = r.slice(-2).reverse(), below = r.filter(x => x.c.q < M.acceptable).length;
    sub = low.length === 2 ? 'سجّلت فئتا ' + nm(low[0].c.name) + ' (' + pct(low[0].c.q) + ') و' + nm(low[1].c.name) + ' (' + pct(low[1].c.q) + ') أدنى مؤشر' + (below >= 2 && low[1].c.q < M.acceptable ? (below === 2 ? '، وكلتاهما دون الحد المقبول البالغ ' + M.acceptable + '%' : '، وعدد الفئات دون الحد المقبول ' + below) : '') : 'مؤشر الجودة لكل فئة';
  }
  const rows = r.map(({ c, i }) => `<button class="hit row" data-cat="${i}" style="width:100%;padding:6px 8px;border-radius:8px;min-height:34px;background:${S.cat === i ? '#E7F0FA' : 'transparent'};opacity:${S.cat === null || S.cat === i ? 1 : .38}">
      <span style="width:140px;font-size:12px;line-height:1.3;text-align:right;overflow-wrap:anywhere"><bdi>${esc(c.name)}</bdi></span><span class="track"><span class="fill" style="width:${c.q}%;background:${col(c.q)}"></span></span><span style="width:36px;font-size:12px;font-weight:600;text-align:left">${pct(c.q)}</span></button>`).join('');
  return `<section class="tile"><h2 class="t">مؤشر الجودة حسب الفئة</h2><p class="st">${sub}</p><div style="display:flex;flex-direction:column;gap:2px;max-height:360px;overflow-y:auto">${rows}</div>
    <div class="legend" style="margin-top:auto"><span><span style="width:10px;height:10px;border-radius:3px;background:${GREEN}"></span>${M.target}% فأكثر</span><span><span style="width:10px;height:10px;border-radius:3px;background:${LBLUE}"></span>${M.acceptable}% – ${M.target - 1}%</span><span><span style="width:10px;height:10px;border-radius:3px;background:${ORANGE}"></span>أقل من ${M.acceptable}%</span></div></section>`;
}

// ---------------- Combo: completeness bars + repeat line ----------------
function combo() {
  const r = ranked().filter(x => x.c.rep != null).sort((a, b) => b.c.comp - a.c.comp).slice(0, 10);
  if (r.length < 2 || !M.hasEmp) return `<section class="tile"><h2 class="t">اكتمال الإجابة وتكرار الاستفسار حسب الفئة</h2><p class="empty">لا تتوفر بيانات كافية لهذا المخطط.</p></section>`;
  const n = r.length, rmax = Math.max(10, Math.ceil(Math.max(...r.map(x => x.c.rep)) / 5) * 5);
  const cx = k => 100 - (k + .5) / n * 100, ry = v => 100 - v / rmax * 100;
  const a = r[0].c, z = r[n - 1].c;
  const title = 'انخفاض اكتمال الإجابة من ' + pct(a.comp) + ' في فئة ' + nm(a.name) + ' إلى ' + pct(z.comp) + ' في فئة ' + nm(z.name) + ' يقابله ' + (z.rep >= a.rep ? 'ارتفاع' : 'انخفاض') + ' تكرار الاستفسار من ' + pct(a.rep) + ' إلى ' + pct(z.rep);
  const on = i => S.cat === null || S.cat === i;
  const bars = r.map(({ c, i }) => `<button class="hit" data-cat="${i}" aria-label="${esc(c.name)}" style="flex:1;height:100%;display:flex;align-items:flex-end;justify-content:center;padding:0"><span style="display:flex;align-items:flex-end;justify-content:center;width:30%;max-width:30px;min-width:16px;height:${c.comp}%;border-radius:6px 6px 0 0;background:${c.comp < M.acceptable ? DBLUE : NAVY};opacity:${on(i) ? 1 : .3}"><span style="font-size:10px;font-weight:700;color:#fff;padding-bottom:4px">${pct(c.comp)}</span></span></button>`).join('');
  const dots = r.map(({ c, i }, k) => `<span style="position:absolute;left:${cx(k)}%;top:${ry(c.rep)}%;width:9px;height:9px;margin-left:-4.5px;margin-top:-4.5px;border-radius:50%;background:#fff;border:2px solid ${ORANGE};opacity:${on(i) ? 1 : .35}"></span><span style="position:absolute;left:${cx(k)}%;top:${ry(c.rep)}%;transform:translate(-50%,-24px);font-size:11px;font-weight:700;color:#B4531F;white-space:nowrap;background:rgba(255,255,255,.85);padding:0 3px;border-radius:3px;opacity:${on(i) ? 1 : .35}">${pct(c.rep)}</span>`).join('');
  const names = r.map(({ c }) => `<span style="flex:1;text-align:center;font-size:10px;line-height:1.25;color:#4A5568;padding:0 2px;overflow-wrap:anywhere"><bdi>${esc(c.name)}</bdi></span>`).join('');
  return `<section class="tile"><h2 class="t">اكتمال الإجابة وتكرار الاستفسار حسب الفئة</h2><p class="st">${title}</p>
    <div style="display:flex;gap:8px;margin-top:6px">
      <div style="display:flex;flex-direction:column;justify-content:space-between;height:190px;font-size:11px;width:34px"><span>100%</span><span>75%</span><span>50%</span><span>25%</span><span>0%</span></div>
      <div style="flex:1;display:flex;flex-direction:column;gap:6px;min-width:0">
        <div style="position:relative;height:190px;background-image:linear-gradient(#EEF1F5 1px,transparent 1px);background-size:100% 25%;border-bottom:1px solid #D5DBE5">
          <div style="position:absolute;inset:0;display:flex;align-items:flex-end">${bars}</div>
          <svg aria-hidden="true" viewBox="0 0 100 100" preserveAspectRatio="none" style="position:absolute;inset:0;width:100%;height:100%;pointer-events:none"><polyline points="${r.map(({ c }, k) => cx(k).toFixed(2) + ',' + ry(c.rep).toFixed(2)).join(' ')}" fill="none" stroke="${ORANGE}" stroke-width="2" vector-effect="non-scaling-stroke"/></svg>
          <div style="position:absolute;inset:0;pointer-events:none;direction:ltr">${dots}</div>
        </div>
        <div style="display:flex">${names}</div>
      </div>
      <div style="display:flex;flex-direction:column;justify-content:space-between;height:190px;font-size:11px;color:#B4531F;width:34px;text-align:left"><span>${rmax}%</span><span>${rmax * .75}%</span><span>${rmax / 2}%</span><span>${rmax / 4}%</span><span>0%</span></div>
    </div>
    <div class="legend"><span><span style="width:14px;height:10px;border-radius:2px;background:${NAVY}"></span>اكتمال الإجابة (المحور الأيمن)</span><span><span style="width:18px;height:3px;background:${ORANGE}"></span>تكرار الاستفسار خلال العام (المحور الأيسر)</span></div></section>`;
}

// ---------------- Waffle ----------------
function waffle(sc) {
  const lv = sc.levels || [0, 0, 0];
  const items = [['جودة مرتفعة', lv[0], GREEN], ['جودة متوسطة', lv[1], LBLUE], ['جودة منخفضة', lv[2], ORANGE]];
  let dots = ''; items.forEach(([, p, c]) => { for (let k = 0; k < Math.round(p); k++) dots += `<span style="display:block;aspect-ratio:1;border-radius:50%;background:${c}"></span>`; });
  const cards = items.map(([nm, p, c]) => `<div style="background:#F5F7FA;border-radius:10px;padding:10px 12px;display:flex;flex-direction:column;gap:2px"><span style="display:flex;align-items:center;gap:6px;font-size:11px;color:#4A5568"><span style="width:10px;height:10px;border-radius:50%;background:${c}"></span>${nm}</span><span style="font-size:12px;font-weight:700">${Math.round(p)}%</span><span style="font-size:11px;color:#5A6478">${fmt(sc.scored * p / 100)} استفسارًا</span></div>`).join('');
  return `<section class="tile" style="justify-content:space-between"><h2 class="t">توزيع الإجابات حسب مستوى الجودة</h2><p class="st">بلغت نسبة الإجابات مرتفعة الجودة ${Math.round(lv[0])}%، ونسبة الإجابات منخفضة الجودة ${Math.round(lv[2])}%</p>
    <div style="display:grid;grid-template-columns:repeat(20,minmax(0,1fr));gap:3px;margin-top:10px">${dots}</div>
    <div class="grid3" style="gap:10px;margin-top:12px">${cards}</div>
    <span class="note" style="margin-top:10px">كل نقطة تمثل 1% من الاستفسارات المقيَّمة. الجودة المرتفعة ${M.high}% فأكثر، والمنخفضة أقل من ${M.low}%</span></section>`;
}

// ---------------- Drivers ----------------
function drivers() {
  const list = D.drivers || [];
  if (!list.length) return `<section class="tile"><h2 class="t">العوامل المرتبطة بارتفاع عدد الطلبات</h2><p class="empty">لا تتوفر بيانات كافية (يتطلب رقم الموظف).</p></section>`;
  const all = list.concat(D.baseline ? [D.baseline] : []), mx = Math.max(...all.map(x => x.v), 1);
  const stars = k => k === 'stars2' ? ' <span aria-label="نجمتان من خمس" style="font-size:13px;letter-spacing:1px;color:#E07A2E">★★☆☆☆</span>' : (k === 'stars5' ? ' <span aria-label="خمس نجوم" style="font-size:13px;letter-spacing:1px;color:#156342">★★★★★</span>' : '');
  const row = (x, base, top) => `<div style="display:grid;grid-template-columns:190px minmax(0,1fr) 40px;gap:10px;align-items:center;${base ? 'border-top:1px dashed #D5DBE5;padding-top:10px' : ''}"><span style="display:flex;align-items:center;gap:8px;font-size:12px;${base ? 'color:#4A5568' : 'font-weight:500'}">${esc(x.name)}${stars(x.kind)}</span><div style="height:10px;background:#F2F4F7;border-radius:5px"><div style="width:${Math.max(3, x.v / mx * 92)}%;height:10px;border-radius:5px;background:${base ? '#9FC9E8' : (top ? ORANGE : NAVY)}"></div></div><strong style="font-size:12px;${top ? 'color:#B4531F' : (base ? 'color:#4A5568' : '')}">+${x.v}</strong></div>`;
  const t = list[0];
  const sub = 'يرتبط عامل (' + t.name + ') بزيادة قدرها ' + t.v + ' لكل 100 طلب' + (D.baseline ? '، مقابل ' + D.baseline.v + ' للطلبات المكتملة' : '');
  return `<section class="tile"><h2 class="t">العوامل المرتبطة بارتفاع عدد الطلبات</h2><p class="st">${esc(sub)}</p><div style="display:flex;flex-direction:column;gap:10px;margin-top:8px">${list.map((x, k) => row(x, false, k === 0)).join('')}${D.baseline ? row(D.baseline, true, false) : ''}</div><p class="note" style="margin:4px 0 0">الطلبات اللاحقة من المستفيد نفسه لنوع الاستفسار نفسه خلال العام، لكل 100 طلب – علاقة ارتباطية</p></section>`;
}

// ---------------- Reasons ----------------
function reasons(sc) {
  const rs = sc.reasons || [];
  if (!rs.length) return `<section class="tile"><h2 class="t">أسباب عدم رضا المستفيد</h2><p class="empty">${M.hasRatings ? 'لا تتوفر ملاحظات مصنفة لهذا النطاق.' : 'لم تُربط بيانات التقييمات والملاحظات.'}</p></section>`;
  const pal = [NAVY, DBLUE, LBLUE, GREEN, '#9FC9E8'];
  const total = rs.reduce((a, r) => a + r[1], 0) || 1;
  return `<section class="tile"><h2 class="t">أسباب عدم رضا المستفيد</h2><p class="st">أعلى أسباب عدم الرضا: ${esc(rs[0][0])} بنسبة ${pct(rs[0][1])}${rs[1] ? '، ثم ' + esc(rs[1][0]) + ' بنسبة ' + pct(rs[1][1]) : ''}</p>
    <div style="display:flex;height:10px;border-radius:5px;overflow:hidden;margin-top:10px">${rs.map((r, k) => `<div style="width:${r[1] / total * 100}%;background:${pal[k]}"></div>`).join('')}</div>
    <div style="display:flex;flex-direction:column;gap:10px;margin-top:10px">${rs.map((r, k) => `<div class="row" style="font-size:12px"><span style="width:12px;height:12px;border-radius:3px;background:${pal[k]}"></span><span style="flex:1">${esc(r[0])}</span><strong>${pct(r[1])}</strong></div>`).join('')}</div></section>`;
}

// ---------------- Missing (drill-down) ----------------
function missing(sc) {
  if (S.q !== null) {
    const el = sc.elements || [];
    if (!el.length) return `<section class="tile"><h2 class="t">تغطية عناصر الإجابة النموذجية</h2><p class="empty">لا تتوفر إجابات مقيَّمة كافية.</p></section>`;
    return `<section class="tile"><h2 class="t">تغطية عناصر الإجابة النموذجية في إجابات الاستفسار</h2><p class="st">لم يرد عنصر (${esc(el[0][0])}) إلا في ${pct(el[0][1])} من الإجابات، وهو الأقل تغطية</p><div style="display:flex;flex-direction:column;gap:12px;margin-top:8px">${el.map(e => `<div style="display:flex;flex-direction:column;gap:5px"><div style="display:flex;justify-content:space-between;font-size:12px;gap:10px"><span>${esc(e[0])}</span><strong style="font-weight:600">${pct(e[1])}</strong></div><div style="background:#EEF1F5;border-radius:5px;height:8px"><div style="height:8px;border-radius:5px;width:${e[1]}%;background:${col(e[1])}"></div></div></div>`).join('')}</div></section>`;
  }
  const ms = sc.missing || [];
  if (!ms.length) return `<section class="tile"><h2 class="t">العناصر الأكثر غيابًا عن الإجابات</h2><p class="empty">لا تتوفر إجابات مقيَّمة كافية.</p></section>`;
  return `<section class="tile"><h2 class="t">العناصر الأكثر غيابًا عن إجابات الفئة</h2><p class="st">يغيب عنصر (${esc(ms[0][0])}) عن ${pct(ms[0][1])} من الإجابات${ms[1] ? '، يليه (' + esc(ms[1][0]) + ') بنسبة ' + pct(ms[1][1]) : ''}</p><div style="display:flex;flex-direction:column;gap:12px;margin-top:8px">${ms.map(e => `<div style="display:flex;flex-direction:column;gap:5px"><div style="display:flex;justify-content:space-between;font-size:12px;gap:10px"><span>${esc(e[0])}</span><strong style="font-weight:600">${pct(e[1])}</strong></div><div style="background:#EEF1F5;border-radius:5px;height:8px"><div style="height:8px;border-radius:5px;width:${e[1]}%;background:${DBLUE}"></div></div></div>`).join('')}</div></section>`;
}

// ---------------- Table (category view) ----------------
function table() {
  const c = D.cats[S.cat], rows = c.qs.slice(0, 8);
  const worst = rows.filter(r => r.q != null).sort((a, b) => a.q - b.q)[0];
  return `<section class="tile"><h2 class="t">أعلى أنواع الاستفسارات عددًا في الفئة</h2><p class="st">${worst ? 'سجّل الاستفسار (' + esc(worst.t) + ') أدنى مؤشر جودة بنسبة ' + pct(worst.q) + '، ويغيب عن إجاباته غالبًا عنصر (' + esc(worst.m) + ')' : ''}</p>
    <div style="overflow-x:auto"><div style="min-width:860px;display:flex;flex-direction:column">
      <div style="display:grid;grid-template-columns:minmax(0,3fr) 90px minmax(0,1.3fr) 170px;gap:12px;padding:10px 12px;font-size:11px;color:#5A6478;border-bottom:1px solid #E3E7EE"><span>الاستفسار</span><span>العدد</span><span>العنصر الأكثر غيابًا</span><span>مؤشر الجودة</span></div>
      ${rows.map((r, j) => `<button class="hit" data-q="${j}" style="display:grid;grid-template-columns:minmax(0,3fr) 90px minmax(0,1.3fr) 170px;gap:12px;align-items:center;width:100%;padding:12px;border-bottom:1px solid #EEF1F5;min-height:48px;background:${S.q === j ? '#E7F0FA' : 'transparent'}"><span style="font-size:12px;line-height:1.6">${esc(r.t)}</span><span style="font-size:12px">${fmt(r.vol)}</span><span style="font-size:12px;color:#B4531F">${esc(r.m)}</span><span class="row" style="gap:8px"><span class="track"><span class="fill" style="width:${r.q || 0}%;background:${col(r.q)}"></span></span><strong style="font-size:12px;font-weight:600;width:36px;text-align:left">${pct(r.q)}</strong></span></button>`).join('')}
    </div></div></section>`;
}

// ---------------- Insight panel ----------------
function insight() {
  const I = D.insights[S.insight]; if (!I) return '';
  return `<section aria-live="polite" style="position:relative;overflow:hidden;background:#0F1D3A;color:#fff;border-radius:18px;padding:24px 28px;display:grid;grid-template-columns:minmax(0,1.4fr) minmax(0,1fr);gap:24px">
    <div class="ring" aria-hidden="true" style="position:absolute;top:0;right:0;left:0;height:2px"></div>
    <div style="display:flex;flex-direction:column;gap:10px"><span style="font-size:12px;font-weight:700;color:#8FD3F5">تحليل إضافي</span><h2 style="margin:0;font-size:20px;line-height:1.5">${esc(I.title)}</h2><p style="margin:0;font-size:14px;line-height:1.9;color:#D7DFEE">${esc(I.body)}</p>
    <button id="closeIns" style="align-self:flex-start;background:transparent;color:#8FD3F5;border:1px solid rgba(143,211,245,.45);border-radius:999px;padding:8px 16px;font-size:13px;cursor:pointer">إغلاق التحليل</button></div>
    <div style="display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;align-content:center">${I.chips.map(c => `<div style="background:rgba(255,255,255,.06);border:1px solid rgba(143,211,245,.22);border-radius:14px;padding:14px 16px;display:flex;flex-direction:column;gap:4px"><span style="font-size:12px;color:#C9D3E6">${esc(c[0])}</span><span style="font-size:22px;font-weight:700">${esc(c[1])}</span></div>`).join('')}</div></section>`;
}

// ---------------- Scope bar ----------------
function scopeBar() {
  let h = `<button class="chip" id="all" style="background:${NAVY};color:#fff;border:0">جميع الفئات</button>`;
  if (S.cat !== null) h += `<span aria-hidden="true" style="color:#9AA5B8">‹</span><button class="chip" id="toCat" style="background:#E7F0FA;color:${DBLUE};border-color:#B9D2EE"><bdi>${esc(D.cats[S.cat].name)}</bdi></button>`;
  if (S.q !== null) h += `<span aria-hidden="true" style="color:#9AA5B8">‹</span><span class="chip" style="cursor:default;max-width:560px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(D.cats[S.cat].qs[S.q].t)}</span>`;
  return `<div style="display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:12px"><nav aria-label="نطاق العرض" style="display:flex;flex-wrap:wrap;align-items:center;gap:8px">${h}</nav><span style="font-size:13px;color:#5A6478">لتصفية اللوحة، اختر الفئة أو الاستفسار من المخططات</span></div>`;
}

// ---------------- Search ----------------
function renderSug() {
  const box = document.getElementById('sug');
  if (!S.open) { box.innerHTML = ''; return; }
  const qn = norm(S.query), match = t => !qn || norm(t).includes(qn);
  const ins = Object.keys(D.insights).filter(k => match(D.insights[k].label));
  const cats = D.cats.map((c, i) => ({ c, i })).filter(x => match(x.c.name));
  let qs = []; D.cats.forEach((c, i) => c.qs.forEach((q, j) => { if (match(q.t)) qs.push({ q, i, j, c }); }));
  qs = qs.slice(0, qn ? 10 : 5);
  const head = t => `<span style="font-size:12px;font-weight:700;color:${GREEN};padding:4px 10px">${t}</span>`;
  let h = '';
  if (ins.length) h += head('تحليلات إضافية') + ins.map(k => `<button data-ins="${k}" style="display:flex;justify-content:space-between;align-items:center;gap:12px;text-align:right;background:#F1F7FB;border:0;border-radius:12px;padding:12px 14px;font-size:15px;color:${NAVY};cursor:pointer;min-height:44px"><span>${esc(D.insights[k].label)}</span><span style="font-size:12px;color:#5A6478">${esc(D.insights[k].sub)}</span></button>`).join('');
  if (cats.length) h += head('الفئات') + `<div style="display:flex;flex-wrap:wrap;gap:8px;padding:0 6px">${cats.slice(0, 30).map(x => `<button class="chip" data-cat="${x.i}"><bdi>${esc(x.c.name)}</bdi></button>`).join('')}</div>`;
  if (qs.length) h += head('الاستفسارات') + qs.map(x => `<button data-cat="${x.i}" data-q="${x.j}" style="display:flex;justify-content:space-between;align-items:center;gap:12px;text-align:right;background:#fff;border:0;border-bottom:1px solid #EEF1F5;padding:12px 14px;font-size:14px;color:${NAVY};cursor:pointer;min-height:44px"><span>${esc(x.q.t)}</span><span style="font-size:12px;color:#5A6478;white-space:nowrap"><bdi>${esc(x.c.name)}</bdi></span></button>`).join('');
  if (!h) h = '<span style="font-size:14px;color:#5A6478;padding:10px">لا توجد نتائج مطابقة.</span>';
  box.innerHTML = `<button id="sugBack" aria-label="إغلاق قائمة الاقتراحات" style="position:fixed;inset:0;background:transparent;border:0;z-index:40;cursor:default"></button><div role="listbox" style="position:absolute;top:62px;right:0;left:0;z-index:50;background:#fff;color:${NAVY};border-radius:18px;box-shadow:0 24px 60px rgba(11,21,48,.35);padding:14px;display:flex;flex-direction:column;gap:10px;max-height:460px;overflow-y:auto">${h}</div>`;
}

// ---------------- Render ----------------
function render() {
  const sc = scope();
  let h = scopeBar();
  if (S.insight) h += insight();
  h += kpis(sc);
  h += `<div class="r1" style="display:grid;gap:12px;grid-template-columns:minmax(0,2fr) minmax(0,1fr)">${trend(sc)}${ranking()}</div>`;
  h += `<div class="r2" style="display:grid;gap:12px;grid-template-columns:minmax(0,2fr) minmax(0,1fr)">${combo()}${waffle(sc)}</div>`;
  h += `<div class="${S.cat !== null ? 'grid3' : 'grid2'}">${drivers()}${reasons(sc)}${S.cat !== null ? missing(sc) : ''}</div>`;
  if (S.cat !== null) h += table();
  h += `<p style="margin:4px 0 0;font-size:12px;color:#5A6478;line-height:1.8">نطاق التحليل: الاستفسارات المطابقة لأنواع الاستفسارات الواردة في دليل الأسئلة الأكثر شيوعًا. يُحتسب مؤشر الجودة بنسبة عناصر الإجابة النموذجية الواردة في إجابة الموظف، وتُستبعد منه الإجابات التي تتعذر فيها الخدمة نظاميًا ما لم تتضمن الإجابة النموذجية ذلك. لا تُصنَّف الفئة التي تقل استفساراتها عن ${M.minCat} استفسارًا.</p>`;
  document.getElementById('app').innerHTML = h;
  renderSug();
}

document.addEventListener('click', e => {
  const b = e.target.closest('button'); if (!b) return;
  if (b.id === 'all') return set({ cat: null, q: null, insight: null });
  if (b.id === 'toCat') return set({ q: null });
  if (b.id === 'closeIns') return set({ insight: null });
  if (b.id === 'sugBack') { S.open = false; return renderSug(); }
  if (b.dataset.ins) { document.getElementById('q').value = ''; return set({ insight: b.dataset.ins, open: false, query: '' }); }
  if (b.dataset.cat !== undefined) {
    const inSug = !!b.closest('#sug');
    if (inSug) document.getElementById('q').value = '';
    return set({ cat: +b.dataset.cat, q: b.dataset.q !== undefined ? +b.dataset.q : null, insight: null, open: inSug ? false : S.open, query: inSug ? '' : S.query });
  }
  if (b.dataset.q !== undefined && S.cat !== null) return set({ q: +b.dataset.q, insight: null });
});
const qi = document.getElementById('q');
qi.addEventListener('focus', () => { S.open = true; renderSug(); });
qi.addEventListener('input', e => { S.query = e.target.value; S.open = true; renderSug(); });
render();
</script>
</body>
</html>
"""

if __name__ == "__main__":
    main()

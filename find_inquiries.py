# -*- coding: utf-8 -*-
"""
find_inquiries.py

Finds the inquiries for each question using the keyword columns in your
Questions file (Match / Required / Excluded Keywords).

Output: one Excel file, two sheets:
  1. Summary    – each question: how many inquiries were found, and if none, why.
  2. Inquiries  – every matched inquiry (full record) + the Question it was assigned to.

Needs: pandas, openpyxl. Runs fully offline.
"""

import re
import sys
import time
from pathlib import Path

import numpy as np
import pandas as pd

# =============================================================================
# >>> YOUR FILES <<<
# =============================================================================
TICKETS_FILE = r"C:\PUT\YOUR\PATH\HERE\clean\tickets_clean.csv"      # .csv / .parquet / .xlsx
QUESTIONS_FILE = r"C:\PUT\YOUR\PATH\HERE\Questions and Answers.xlsx"
OUTPUT_FILE = r""   # empty = "inquiries_by_question.xlsx" next to the tickets file

# Column names
INQUIRY_COLUMN = "Inquiry"
QUESTION_COLUMN = "Question"
MATCH_COLUMN = "Match Keywords"
REQUIRED_COLUMN = "Required Keywords"
EXCLUDED_COLUMN = "Excluded Keywords"

# =============================================================================

DIACRITICS = re.compile(r"[\u0610-\u061A\u064B-\u065F\u0670\u06D6-\u06ED\u0640]")
LETTERS = str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ٱ": "ا", "ة": "ه", "ى": "ي",
                         "ؤ": "و", "ئ": "ي", "٠": "0", "١": "1", "٢": "2", "٣": "3", "٤": "4",
                         "٥": "5", "٦": "6", "٧": "7", "٨": "8", "٩": "9"})
ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def normalize(text) -> str:
    """Same rules for inquiries and keywords: إجازة = اجازه, case-insensitive, etc."""
    if not isinstance(text, str):
        text = "" if pd.isna(text) else str(text)
    text = DIACRITICS.sub("", text).translate(LETTERS).lower()
    return re.sub(r"\s+", " ", text)


def split_words(cell):
    """'كلمة، كلمة+كلمة' -> [['كلمة'], ['كلمة', 'كلمة']]  ('+' = all parts required)."""
    if not isinstance(cell, str) or not cell.strip():
        return []
    phrases = []
    for item in re.split(r"[،,\n]", cell):
        parts = [normalize(p).strip() for p in item.split("+")]
        parts = [p for p in parts if p]
        if parts:
            phrases.append(parts)
    return phrases


def find_col(df, name):
    for c in df.columns:
        if str(c).strip().lower() == name.strip().lower():
            return c
    return None


def stop(msg):
    print(f"\nSTOPPED - {msg}")
    sys.exit(1)


def read_tickets(path):
    p = Path(path)
    if not p.exists():
        stop(f"Tickets file not found:\n   {path}")
    if p.suffix.lower() == ".parquet":
        return pd.read_parquet(p)
    if p.suffix.lower() == ".csv":
        return pd.read_csv(p, dtype=str, encoding="utf-8-sig")
    return pd.read_excel(p, dtype=str)


def read_questions(path):
    p = Path(path)
    if not p.exists():
        stop(f"Questions file not found:\n   {path}")
    raw = pd.read_excel(p, header=None, dtype=str)
    # header = first row that contains "Question"
    for i in range(min(15, len(raw))):
        if any(str(v).strip().lower() == QUESTION_COLUMN.lower() for v in raw.iloc[i]):
            df = raw.iloc[i + 1:].copy()
            df.columns = [str(v).strip() for v in raw.iloc[i]]
            break
    else:
        stop(f"No '{QUESTION_COLUMN}' header found in the questions file.")
    cols = {k: find_col(df, n) for k, n in [("q", QUESTION_COLUMN), ("m", MATCH_COLUMN),
                                            ("r", REQUIRED_COLUMN), ("e", EXCLUDED_COLUMN)]}
    missing = [n for k, n in [("m", MATCH_COLUMN), ("r", REQUIRED_COLUMN), ("e", EXCLUDED_COLUMN)]
               if cols[k] is None]
    if missing:
        stop(f"Missing column(s) in the questions file: {missing}")
    df = df[df[cols["q"]].fillna("").str.strip() != ""].reset_index(drop=True)
    out = []
    for i, row in df.iterrows():
        out.append({
            "no": i + 1,
            "question": str(row[cols["q"]]).strip(),
            "match": split_words(row[cols["m"]]),
            "required": split_words(row[cols["r"]]),
            "excluded": split_words(row[cols["e"]]),
        })
    return out


def main():
    t0 = time.time()
    print("Reading files ...")
    tickets = read_tickets(TICKETS_FILE)
    inq_col = find_col(tickets, INQUIRY_COLUMN)
    if inq_col is None:
        stop(f"Column '{INQUIRY_COLUMN}' not found in the tickets file.")
    questions = read_questions(QUESTIONS_FILE)
    print(f"   {len(tickets):,} tickets | {len(questions)} questions")

    text = tickets[inq_col].map(normalize)
    n = len(text)
    cache = {}

    def has(part):
        if part not in cache:
            cache[part] = text.str.contains(part, regex=False).to_numpy()
        return cache[part]

    def phrase_hits(phrases):
        """Matrix: one column per phrase (True if all its parts are in the inquiry)."""
        if not phrases:
            return np.zeros((n, 0), dtype=bool)
        cols = []
        for parts in phrases:
            m = has(parts[0]).copy()
            for p in parts[1:]:
                m &= has(p)
            cols.append(m)
        return np.column_stack(cols)

    print("Matching ...")
    Q = len(questions)
    score = np.zeros((n, Q), dtype=np.int32)       # words matched (match + required)
    length = np.zeros((n, Q), dtype=np.int32)      # characters matched (tie-break)
    diag = []
    for j, q in enumerate(questions):
        if j % 20 == 0:
            print(f"   question {j + 1}/{Q}", end="\r")
        M, R, E = phrase_hits(q["match"]), phrase_hits(q["required"]), phrase_hits(q["excluded"])
        m_ok = M.any(axis=1) if M.shape[1] else np.zeros(n, bool)
        r_ok = R.any(axis=1) if R.shape[1] else np.ones(n, bool)
        e_hit = E.any(axis=1) if E.shape[1] else np.zeros(n, bool)
        ok = m_ok & r_ok & ~e_hit

        lens_m = np.array([sum(len(p) for p in ph) for ph in q["match"]] or [0])
        lens_r = np.array([sum(len(p) for p in ph) for ph in q["required"]] or [0])
        score[ok, j] = M[ok].sum(1) + (R[ok].sum(1) if R.shape[1] else 0)
        length[ok, j] = (M[ok] @ lens_m[:M.shape[1]]) + ((R[ok] @ lens_r[:R.shape[1]]) if R.shape[1] else 0)

        # which excluded words blocked the most inquiries (for the reason text)
        top_ex = []
        blocked = m_ok & r_ok & e_hit
        if blocked.any():
            counts = E[blocked].sum(0)
            order = np.argsort(-counts)[:3]
            top_ex = ["+".join(q["excluded"][k]) for k in order if counts[k]]
        diag.append({"m": int(m_ok.sum()), "mr": int((m_ok & r_ok).sum()),
                     "blocked": int(blocked.sum()), "top_ex": top_ex, "ok": ok})

    # Each inquiry goes to ONE question: most words matched, then longest match.
    combined = score.astype(np.int64) * 100000 + length
    best = combined.argmax(1)
    matched = score.max(1) > 0
    assigned = np.where(matched, best, -1)

    # ---------------- Summary ----------------
    rows = []
    for j, q in enumerate(questions):
        found = int((assigned == j).sum())
        d = diag[j]
        reason = ""
        if found == 0:
            if not q["match"]:
                reason = "لا توجد كلمات في عمود Match Keywords"
            elif d["m"] == 0:
                reason = "لا يوجد أي استفسار يحتوي على كلمات المطابقة"
            elif d["mr"] == 0:
                reason = (f"{d['m']:,} استفسار فيها كلمات المطابقة، "
                          f"لكن لا يوجد فيها أي كلمة إلزامية")
            elif d["ok"].sum() == 0:
                reason = (f"{d['blocked']:,} استفسار استُبعدت بسبب الكلمات المستبعدة: "
                          + "، ".join(d["top_ex"]))
            else:
                others = pd.Series(assigned[d["ok"]]).value_counts()
                others = [f"سؤال {k + 1} ({v:,})" for k, v in others.items() if k >= 0][:3]
                reason = (f"{int(d['ok'].sum()):,} استفسار طابقت السؤال، لكنها أقرب لأسئلة أخرى: "
                          + "، ".join(others))
        rows.append({"No": q["no"], "Question": q["question"], "Inquiries Found": found,
                     "Status": "تم العثور" if found else "لم يتم العثور", "Reason": reason})
    summary = pd.DataFrame(rows)

    # ---------------- Inquiries ----------------
    idx = np.where(assigned >= 0)[0]
    found_df = tickets.iloc[idx].copy()
    found_df.insert(0, "Question", [questions[k]["question"] for k in assigned[idx]])
    found_df.insert(0, "No", [questions[k]["no"] for k in assigned[idx]])
    found_df = found_df.sort_values("No", kind="stable")
    found_df = found_df.map(lambda v: ILLEGAL.sub("", v) if isinstance(v, str) else v)

    # ---------------- Save ----------------
    out = Path(OUTPUT_FILE) if OUTPUT_FILE else Path(TICKETS_FILE).with_name("inquiries_by_question.xlsx")
    print("\nSaving ...")
    try:
        with pd.ExcelWriter(out, engine="openpyxl") as xw:
            summary.to_excel(xw, sheet_name="Summary", index=False)
            found_df.to_excel(xw, sheet_name="Inquiries", index=False)
            style(xw.book["Summary"], {"A": 6, "B": 55, "C": 16, "D": 16, "E": 70})
            style(xw.book["Inquiries"], {"A": 6, "B": 45})
    except PermissionError:
        stop(f"Cannot save. Is the file open in Excel?\n   {out}")

    n_found = int((summary["Inquiries Found"] > 0).sum())
    print("\n" + "=" * 50)
    print(f"  Questions with inquiries   : {n_found} / {len(summary)}")
    print(f"  Questions without          : {len(summary) - n_found}")
    print(f"  Inquiries assigned         : {len(found_df):,} of {n:,} "
          f"({len(found_df) / max(n, 1):.0%})")
    print("=" * 50)
    print(f"Saved: {out}   ({time.time() - t0:.0f}s)")


def style(ws, widths):
    from openpyxl.styles import Alignment, Font, PatternFill
    ws.sheet_view.rightToLeft = True
    ws.freeze_panes = "A2"
    for c in ws[1]:
        c.font = Font(name="Somar", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="152446")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    if ws.title == "Summary":
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.font = Font(name="Somar")
                c.alignment = Alignment(wrap_text=True, vertical="top")
            if row[3].value == "لم يتم العثور":
                row[3].font = Font(name="Somar", bold=True, color="C0392B")
    ws.auto_filter.ref = ws.dimensions


if __name__ == "__main__":
    main()

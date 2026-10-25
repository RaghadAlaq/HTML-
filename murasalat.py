# -*- coding: utf-8 -*-
"""
متابعة المراسلات الخارجية — daily follow-up automation
=======================================================

Run it once a day after you drop the new system report into the "reports" folder:

    python murasalat.py

What it does, in order:
  1. Reads your choices from the tracker (dropdowns, close dates, emails, holidays, settings).
  2. Reads the newest report in /reports, keeps only «خارجية», finds who holds each
     transaction now (latest receipt; several people at once is fine).
  3. Checks Outlook (classic): emails you sent with a follow-up code, and replies to them.
     Replies are filed into Inbox > متابعة المراسلات > معاملة <number>.
  4. Counts working days (Friday + Saturday + holidays excluded).
  5. Creates one Outlook draft per person with all of their transactions that need follow-up.
  6. Rewrites Murasalat_Tracker.xlsx (keep it closed while the script runs).

Options:
  --today 2026-10-22   pretend today is another date (for testing)
  --report path.xlsx   use a specific report instead of the newest one
  --no-outlook         skip Outlook completely (emails are still written in «رسائل اليوم»)

Install once:  pip install pandas openpyxl pywin32 xlrd
"""
import argparse
import datetime as dt
import json
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ============================ SETTINGS YOU CAN CHANGE ============================
BASE = Path(__file__).resolve().parent
REPORTS_DIR = BASE / "reports"
TRACKER = BASE / "Murasalat_Tracker.xlsx"
STATE_FILE = BASE / "murasalat_state.json"

USE_OUTLOOK = True            # False = never touch Outlook
# Emails are NEVER sent by the script. Every reminder is saved in Outlook > Drafts for you to review and send.
ESCALATION_CC = ""            # e.g. "manager@company.sa" — CC'd when someone reaches (deadline - 1)
OUTLOOK_FOLDER = "متابعة المراسلات"
LOOKBACK_DAYS = 45            # how far back to look in Inbox / Sent Items
SIGNATURE = "الاسم\nالمسمى الوظيفي – الإدارة"

# ---- TEST MODE: while «وضع التجربة» = نعم in the tracker's «الإعدادات» sheet (default for a new setup),
# every email goes ONLY to these addresses, marked [تجربة]. Nothing reaches real people.
TEST_EMAILS = ["Raghadmusiar@gmail.com", "Relsashy@gmail.com", "Raghad4321r@gmail.com"]
REMIND_DAILY_UNTIL_CLOSED = True   # True = a reminder every working day from day 5 until it is closed,
                                   #        even if they replied (a reply is not a closure)
REPLIED_REMIND_AFTER = 2      # only used if the line above is False
SLA_FROM_FIRST = True         # the 10 days count from the transaction's FIRST receipt date
STOP_AFTER_DEADLINE = False   # True = stop reminding after day 10 (not recommended: overdue ones need it most)
# These two are only defaults; the tracker's «الإعدادات» sheet wins once it exists.
DEFAULT_THRESHOLD = 5         # start emailing at this many working days
DEFAULT_DEADLINE = 10         # must be closed within this many working days
DEFAULT_HOLIDAYS = ["2026-09-23", "2027-02-22"]
# ================================================================================

ST_NEW, ST_SENT, ST_REPLIED, ST_FWD, ST_CLOSED = (
    "لم يتم التواصل", "تم الإرسال", "تم الرد", "محولة لشخص آخر", "مغلقة")
STATUSES = [ST_NEW, ST_SENT, ST_REPLIED, ST_FWD, ST_CLOSED]
B_OVER, B_CRIT, B_WATCH, B_OK, B_DONE = "انتهت المهلة", "حرجة", "تحتاج متابعة", "ضمن المهلة", "مغلقة"

SH_TODAY, SH_OPEN, SH_PEOPLE, SH_DONE, SH_LOG = "اليوم", "المتابعة", "حسب الموظف", "المنجزة", "سجل المتابعة"
SH_CHANGES = "ما الذي تغيّر"
SH_EMAILS = "الإيميلات"
SH_PATHS = "المسارات"
SH_STATS, SH_MAILS, SH_CONTACTS, SH_HOL, SH_SET, SH_LISTS = (
    "إحصائيات", "رسائل اليوم", "الموظفين", "الإجازات", "الإعدادات", "_lists")

CODE_RE = re.compile(r"MR-\d{6}-\d{2}")


# ------------------------------------------------------------------ helpers
def norm(s):
    s = "" if s is None else str(s)
    s = re.sub(r"[\u064B-\u0652\u0640]", "", s)
    s = re.sub("[إأآ]", "ا", s).replace("ى", "ي").replace("ة", "ه")
    return re.sub(r"\s+", "", s)


def clean_name(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()


AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")


def parse_dt(v):
    """Return a datetime or None. Handles Excel dates, serials, '2026/09/12', '2026-09-21 19:45:06'."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, pd.Timestamp):
        return None if pd.isna(v) else v.to_pydatetime()
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, dt.date):
        return dt.datetime.combine(v, dt.time())
    if isinstance(v, (int, float)):
        if 20000 < v < 90000:
            return dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))
        return None
    s = str(v).translate(AR_DIGITS).strip()
    if not s:
        return None
    m = re.match(r"^(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?", s)
    if m:
        y, mo, d = int(m[1]), int(m[2]), int(m[3])
    else:
        m = re.match(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?", s)
        if not m:
            return None
        d, mo, y = int(m[1]), int(m[2]), int(m[3])
    if y < 1700:  # Hijri — not used for counting
        return None
    try:
        return dt.datetime(y, mo, d, int(m[4] or 0), int(m[5] or 0), int(m[6] or 0))
    except ValueError:
        return None


def to_date(v):
    x = parse_dt(v)
    return x.date() if x else None


def iso(d):
    return d.isoformat() if d else None


def from_iso(s):
    if not s:
        return None
    try:
        return dt.datetime.fromisoformat(s)
    except ValueError:
        return None


def d_iso(s):
    x = from_iso(s)
    return x.date() if x else None


class Calendar:
    def __init__(self, holidays):
        self.h = set(holidays)

    def is_work(self, d):
        return d.weekday() not in (4, 5) and d not in self.h   # Fri=4, Sat=5

    def days(self, start, end):
        """Working days AFTER the receipt day up to and including end (same rule as the Excel formula)."""
        if not start or not end or end <= start:
            return 0
        n, d = 0, start + dt.timedelta(days=1)
        while d <= end:
            n += self.is_work(d)
            d += dt.timedelta(days=1)
        return n


def days_ar(n):
    if n == 1:
        return "يوم عمل واحد"
    if n == 2:
        return "يوما عمل"
    if 3 <= n <= 10:
        return f"{n} أيام عمل"
    return f"{n} يوم عمل"


def count_ar(n):
    if n == 2:
        return "معاملتين خارجيتين"
    if 3 <= n <= 10:
        return f"{n} معاملات خارجية"
    return f"{n} معاملة خارجية"


def band(days, closed, threshold, deadline):
    if closed:
        return B_DONE
    if days >= deadline:
        return B_OVER
    if days >= deadline - 2:
        return B_CRIT
    if days >= threshold:
        return B_WATCH
    return B_OK


# ------------------------------------------------------------------ state
def load_state():
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    return {"rows": {}, "emails": {}, "batches": {}, "processed_mail": [], "log": [], "closed": {},
            "settings": {"threshold": DEFAULT_THRESHOLD, "deadline": DEFAULT_DEADLINE, "test_mode": True,
                         "holidays": DEFAULT_HOLIDAYS}}


def save_state(state):
    tmp = STATE_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(STATE_FILE)


def log(state, today, txn, person, action, details=""):
    state["log"].append({"date": iso(today), "txn": txn, "person": person, "action": action, "details": details})


def bpeople(b):
    return b.get("people") or [b["person"]]


def key(txn, person):
    return f"{txn}|{person}"


# ------------------------------------------------------------------ 1. read tracker edits
def read_tracker(state, today):
    if not TRACKER.exists():
        return
    wb = load_workbook(TRACKER)  # editable cells hold plain values, no need for data_only
    if SH_SET in wb.sheetnames:
        ws = wb[SH_SET]
        for r, k in ((2, "threshold"), (3, "deadline")):
            try:
                state["settings"][k] = max(1, int(ws.cell(r, 2).value))
            except (TypeError, ValueError):
                pass
        tm = clean_name(ws.cell(4, 2).value)
        if tm in ("نعم", "لا"):
            state["settings"]["test_mode"] = tm == "نعم"
    hol_src = (wb[SH_SET], 5) if SH_SET in wb.sheetnames and wb[SH_SET].cell(1, 5).value else \
        (wb[SH_HOL], 1) if SH_HOL in wb.sheetnames else None
    if hol_src:
        ws, c = hol_src
        hol = [to_date(r[0]) for r in ws.iter_rows(min_row=2, min_col=c, max_col=c, values_only=True)]
        state["settings"]["holidays"] = sorted({iso(h) for h in hol if h})
    ct_src = (wb[SH_SET], 7) if SH_SET in wb.sheetnames and wb[SH_SET].cell(1, 7).value else \
        (wb[SH_CONTACTS], 1) if SH_CONTACTS in wb.sheetnames else None
    if ct_src:
        ws, c = ct_src
        for name, email in ws.iter_rows(min_row=2, min_col=c, max_col=c + 1, values_only=True):
            name = clean_name(name)
            if name:
                email = clean_name(email)
                if email:
                    state["emails"][name] = email
                else:
                    state["emails"].pop(name, None)
    state["_selected"] = []
    if SH_EMAILS in wb.sheetnames:
        for row in wb[SH_EMAILS].iter_rows(min_row=5, max_col=2, values_only=True):
            if row[0] and str(row[0]).strip() and row[1]:
                state["_selected"].append(str(row[1]).strip())
    sheet = SH_OPEN if SH_OPEN in wb.sheetnames else "المفتوحة" if "المفتوحة" in wb.sheetnames else None
    if not sheet:
        read_paths_sheet(state, wb, today)
        return
    ws = wb[sheet]
    for row in ws.iter_rows(min_row=2, values_only=True):
        k = row[0]
        if not k or k not in state["rows"]:
            continue
        R = state["rows"][k]
        status = clean_name(row[8]) or R["status"]
        fwd_to = clean_name(row[9])
        close_d = to_date(row[10])
        notes = clean_name(row[11])
        R["notes"] = notes
        if close_d:                       # a close date alone means: closed
            status = ST_CLOSED
        elif fwd_to and status != ST_FWD and status != ST_CLOSED:
            status = ST_FWD
        if status != R["status"]:
            log(state, today, R["txn"], R["person"], "تغيير الحالة (يدوي)", f"{R['status']} ← {status}")
            if status == ST_SENT:
                made = [b["created"] for b in state["batches"].values()
                        if R["person"] in bpeople(b) and R["txn"] in b["txns"]]
                R["last_contact"] = max(made) if made else iso(today)
            elif status == ST_REPLIED:
                R["last_reply"] = iso(today)
            R["status"] = status
        if status == ST_FWD and fwd_to and fwd_to != R.get("fwd_to"):
            R["fwd_to"], R["fwd_date"] = fwd_to, iso(today)
        if status == ST_CLOSED:
            new_close = iso(close_d or today)
            if R.get("close_date") != new_close:
                R["close_date"] = new_close
                log(state, today, R["txn"], R["person"], "أُغلقت", f"بتاريخ {new_close}")
        elif R.get("close_date") and R.get("closed_by") != "system":
            R["close_date"] = None
    read_paths_sheet(state, wb, today)   # after «المتابعة», so a close date typed in «المسارات» wins


def read_paths_sheet(state, wb, today):
    """Manual edits in «المسارات»: close dates and notes on any step, and steps you add yourself."""
    if SH_PATHS not in wb.sheetnames:
        return
    rows = state["rows"]
    edits = state.setdefault("step_edits", {})
    manual = state.setdefault("manual_steps", {})
    for row in wb[SH_PATHS].iter_rows(min_row=3, max_col=10, values_only=True):
        k = row[0]
        if not k or not isinstance(k, str) or "|" not in k:
            continue
        frm, to = clean_name(row[2]), clean_name(row[3])
        recv_d, sent_d, note = to_date(row[5]), to_date(row[6]), clean_name(row[9])
        kind, rest = k.split("|", 1)
        if kind == "S":
            txn, person = rest.split("|")[0], rest.split("|")[1]
            sys_sent = next((st["sent"] for st in state.get("paths", {}).get(txn, {}).get("steps", [])
                             if st["k"] == k), None)
            old = edits.get(k, {})
            typed = iso(sent_d)
            new = {"sent": typed if typed and typed != sys_sent else None, "note": note}
            if not new["sent"] and not new["note"]:
                edits.pop(k, None)
            elif new != old:
                edits[k] = new
                if new["sent"] and not old.get("sent"):
                    log(state, today, txn, person, "تاريخ إغلاق يدوي (المسارات)", iso(sent_d))
                    close_holder(state, txn, person, sent_d, today)
        elif kind == "M":
            txn, mid = rest.split("|", 1)
            steps = manual.get(txn, [])
            st = next((x for x in steps if x["id"] == mid), None)
            if not st:
                continue
            if not to:                                   # cleared = delete your step
                steps.remove(st)
                log(state, today, txn, st["to"], "حذف خطوة يدوية", "")
                continue
            had_sent = st.get("sent")
            st.update({"from": frm or st["from"], "to": to, "recv": iso(recv_d) or st["recv"],
                       "sent": iso(sent_d), "note": note})
            if sent_d and not had_sent:
                close_holder(state, txn, to, sent_d, today)
        elif kind == "ADD" and to:
            txn = rest
            open_holders = [R for R in rows.values() if R["txn"] == txn and R.get("active")
                            and R["status"] not in (ST_CLOSED, ST_FWD)]
            src = next((R for R in open_holders if frm and norm(R["person"]) == norm(frm)), None)
            if not src and len(open_holders) == 1:
                src = open_holders[0]
            day = recv_d or today
            st = {"id": dt.datetime.now().strftime("%H%M%S%f"), "from": frm or (src["person"] if src else ""),
                  "to": to, "recv": iso(day), "sent": iso(sent_d), "note": note}
            manual.setdefault(txn, []).append(st)
            log(state, today, txn, to, "خطوة يدوية (المسارات)", f"من {st['from'] or '؟'} إلى {to}")
            if src and norm(src["person"]) != norm(to):  # works like «محولة إلى»
                src["status"], src["fwd_to"], src["fwd_date"] = ST_FWD, to, iso(day)
            elif not src and key(txn, to) not in rows:
                base = next((R for R in rows.values() if R["txn"] == txn), {})
                rows[key(txn, to)] = {**{kk: base.get(kk) for kk in ("txn", "subject", "first", "importance",
                                                                     "moves", "in_report", "created")},
                                      "txn": txn, "person": to, "status": ST_NEW, "source": "forward", "dept": "",
                                      "recv": base.get("recv") or iso(day), "fwd_to": None, "fwd_date": None,
                                      "close_date": None, "closed_by": None, "notes": "", "last_contact": None,
                                      "last_reply": None, "active": True}
            if sent_d:
                state.setdefault("_close_after_forward", []).append((txn, to, iso(sent_d)))


def close_holder(state, txn, person, day, today):
    R = state["rows"].get(key(txn, person))
    if R and R.get("active") and R["status"] != ST_CLOSED:
        R["status"], R["close_date"], R["closed_by"] = ST_CLOSED, iso(day), "me"


# ------------------------------------------------------------------ 2. read the report
COLS = {"txn": "رقمالمعامله", "subject": "الموضوع", "type": "نوعالمعامله", "recv": "تاريخالاستلام",
        "sent": "تاريخالارسال", "dept": "الادارات", "sender": "المرسل", "recipient": "المستلم",
        "importance": "درجهالاهميه", "created": "تاريخالانشاء", "from_h": "منتاريخ", "to_h": "اليتاريخ",
        "dept_name": "اسمالاداره"}


def newest_report():
    files = [p for p in REPORTS_DIR.glob("*") if p.suffix.lower() in (".xlsx", ".xls", ".csv")
             and not p.name.startswith("~$")]
    if not files:
        sys.exit(f"No report found. Put the system export in: {REPORTS_DIR}")
    return max(files, key=lambda p: p.stat().st_mtime)


ALIASES = {}   # name in the report -> name used everywhere (filled from the settings sheet)


def alias(name):
    name = clean_name(name)
    return ALIASES.get(norm(name), name) if name else name


def txn_id(v):
    t = clean_name(v).translate(AR_DIGITS)
    if t.endswith(".0"):
        t = t[:-2]
    return t.replace(" ", "")


def load_raw(path):
    if path.suffix.lower() == ".csv":
        try:
            return pd.read_csv(path, header=None, dtype=str, encoding="utf-8-sig")
        except UnicodeDecodeError:
            return pd.read_csv(path, header=None, dtype=str, encoding="cp1256")
    return pd.read_excel(path, header=None, dtype=object)


def read_report(path):
    return parse_raw(load_raw(path))


def parse_raw(raw):
    hdr = None
    for i in range(min(25, len(raw))):
        if COLS["txn"] in [norm(x) for x in raw.iloc[i].tolist()]:
            hdr = i
            break
    if hdr is None:
        raise RuntimeError("Could not find the «رقم المعاملة» column in the report.")
    names = [norm(x) for x in raw.iloc[hdr].tolist()]
    idx = {k: names.index(v) for k, v in COLS.items() if v in names}
    for need in ("txn", "type", "recv", "recipient"):
        if need not in idx:
            raise RuntimeError(f"Report is missing a column: {COLS[need]}")
    out, skipped, seen = [], 0, set()
    for _, r in raw.iloc[hdr + 1:].iterrows():
        g = lambda k: r.iloc[idx[k]] if k in idx else None
        txn = txn_id(g("txn"))
        if not txn:
            continue
        if "خارج" not in norm(g("type")):
            skipped += 1
            continue
        recv = parse_dt(g("recv"))
        if not recv:
            continue
        sent = parse_dt(g("sent"))
        rec = {"txn": txn, "subject": clean_name(g("subject")), "recv": recv,
               "sent": sent.date() if sent else None, "dept": clean_name(g("dept")),
               "sender": alias(g("sender")), "recipient": alias(g("recipient")),
               "importance": clean_name(g("importance")), "created": to_date(g("created")),
               "type": clean_name(g("type")), "from_h": clean_name(g("from_h")), "to_h": clean_name(g("to_h")),
               "dept_name": clean_name(g("dept_name"))}
        sig = (txn, rec["recipient"], recv, rec["sent"])
        if sig in seen:                     # the same movement listed twice (e.g. two paths in the report)
            continue
        seen.add(sig)
        out.append(rec)
    return out, skipped


def current_holders(rows):
    """Per transaction: everyone who received it on its latest receipt day and did not pass it on that day."""
    by = {}
    for r in rows:
        by.setdefault(r["txn"], []).append(r)
    result = {}
    for txn, rs in by.items():
        rs.sort(key=lambda r: r["recv"])
        last_day = max(r["recv"].date() for r in rs)
        latest = [r for r in rs if r["recv"].date() == last_day]
        holders = [r for r in latest
                   if not any(o["recv"] > r["recv"] and o["sender"] and o["sender"] == r["recipient"] for o in latest)]
        holders = holders or latest
        seen, uniq = set(), []
        for h in holders:
            if h["recipient"] and h["recipient"] not in seen:
                seen.add(h["recipient"])
                uniq.append(h)
        result[txn] = {"holders": uniq, "first": rs[0]["recv"].date(), "last_day": last_day,
                       "subject": next((x["subject"] for x in reversed(rs) if x["subject"]), ""),
                       "importance": rs[0]["importance"], "moves": len(rs),
                       "created": next((x["created"] for x in rs if x["created"]), None),
                       "type": rs[-1].get("type", ""), "from_h": rs[-1].get("from_h", ""),
                       "to_h": rs[-1].get("to_h", ""), "dept_name": rs[-1].get("dept_name", ""),
                       "steps": [{"k": f"S|{txn}|{x['recipient']}|{x['recv']:%Y%m%d%H%M}", "from": x["sender"],
                                  "to": x["recipient"], "dept": x["dept"], "recv": iso(x["recv"]),
                                  "sent": iso(x["sent"])} for x in rs]}
    return result


def report_snapshot(report_txns):
    """What the report itself says: who still has it open, who closed it (تاريخ الإرسال filled)."""
    return {txn: {"open": sorted(h["recipient"] for h in info["holders"] if not h["sent"]),
                  "closed": sorted(h["recipient"] for h in info["holders"] if h["sent"]),
                  "close_dates": {h["recipient"]: iso(h["sent"]) for h in info["holders"] if h["sent"]},
                  "subject": info["subject"], "first": iso(info["first"])}
            for txn, info in report_txns.items()}


def compare_reports(prev, cur):
    """Changes between the previous report and this one (Sunday vs Wednesday, etc.)."""
    if not prev:
        return []
    old, new = prev["txns"], cur
    changes = []
    for txn, n in new.items():
        o = old.get(txn)
        subj = n["subject"]
        if o is None:
            kind = "مغلقة" if not n["open"] else "جديدة"
            changes.append((kind, txn, subj, "", "، ".join(n["open"] or n["closed"]), "ظهرت لأول مرة"))
            continue
        if o["open"] and not n["open"]:
            when = max(n["close_dates"].values()) if n["close_dates"] else None
            changes.append(("أُغلقت بالكامل", txn, subj, "، ".join(o["open"]), "—",
                            f"أُغلقت بتاريخ {d_iso(when):%Y/%m/%d}" if when else "أُغلقت"))
            continue
        if not o["open"]:
            if n["open"]:
                changes.append(("أُعيد فتحها", txn, subj, "", "، ".join(n["open"]), ""))
            continue
        gone = [p for p in o["open"] if p not in n["open"]]
        came = [p for p in n["open"] if p not in o["open"]]
        closed_now = [p for p in gone if p in n["closed"]]
        moved_from = [p for p in gone if p not in n["closed"]]
        if closed_now:
            changes.append(("أُغلق جزء منها", txn, subj, "، ".join(o["open"]), "، ".join(n["open"]),
                            "أغلق: " + "، ".join(closed_now) + " · باقٍ: " + "، ".join(n["open"])))
        if came or moved_from:
            changes.append(("انتقلت", txn, subj, "، ".join(moved_from or o["open"]), "، ".join(came or n["open"]), ""))
        if not (closed_now or came or moved_from):
            changes.append(("بدون تغيير", txn, subj, "", "، ".join(n["open"]), ""))
    for txn, o in old.items():
        if txn not in new and o["open"]:
            changes.append(("اختفت من التقرير", txn, o["subject"], "، ".join(o["open"]), "",
                            "غالباً أُغلقت أو خرجت من فترة التقرير — تأكدي ثم اختاري «مغلقة»"))
    order = ["أُغلقت بالكامل", "أُغلق جزء منها", "انتقلت", "اختفت من التقرير", "جديدة", "أُعيد فتحها", "مغلقة", "بدون تغيير"]
    changes.sort(key=lambda c: (order.index(c[0]), c[1]))
    return changes


def merge_report(state, report_txns, today):
    rows = state["rows"]
    paths = state.setdefault("paths", {})
    locked = set(state.get("closed", {}))     # closed is final: the report never touches it again
    report_txns = {t: v for t, v in report_txns.items() if t not in locked}
    for txn, info in report_txns.items():      # the newest report always has the full path
        paths[txn] = {"subject": info["subject"], "steps": info["steps"]}
    for txn, info in report_txns.items():
        holder_keys = set()
        for h in info["holders"]:
            k = key(txn, h["recipient"])
            holder_keys.add(k)
            R = rows.get(k)
            if R is None:
                R = rows[k] = {"txn": txn, "person": h["recipient"], "status": ST_NEW, "source": "system",
                               "fwd_to": None, "fwd_date": None, "close_date": None, "closed_by": None,
                               "notes": "", "last_contact": None, "last_reply": None, "reply_text": "",
                               "suggestion": ""}
                log(state, today, txn, h["recipient"], "ظهرت في التقرير", f"استلام {h['recv']:%Y/%m/%d}")
            R.update({"subject": info["subject"], "dept": h["dept"], "recv": iso(h["recv"]),
                      "first": iso(info["first"]), "importance": info["importance"], "moves": info["moves"],
                      "created": iso(info["created"]), "type": info.get("type", ""),
                      "sender": h.get("sender", ""), "from_h": info.get("from_h", ""), "to_h": info.get("to_h", ""),
                      "dept_name": info.get("dept_name", ""),
                      "active": True, "source": "system", "in_report": True})
            if h["sent"] and R["status"] != ST_CLOSED:
                R["status"], R["close_date"], R["closed_by"] = ST_CLOSED, iso(h["sent"]), "system"
                log(state, today, txn, h["recipient"], "أُغلقت (حسب النظام)", f"تاريخ الإرسال {h['sent']}")
        # people who held it before in the system but no longer do
        for k, R in rows.items():
            if R["txn"] == txn and R.get("active") and k not in holder_keys and R["source"] == "system":
                R["active"] = False
                log(state, today, txn, R["person"], "انتقلت في النظام",
                    "إلى " + "، ".join(h["recipient"] for h in info["holders"]))
            if R["txn"] == txn and R["source"] == "forward" and d_iso(R.get("recv")) and \
                    d_iso(R["recv"]) < info["last_day"] and k not in holder_keys and R.get("active"):
                R["active"] = False  # system moved on past a manual forward
    for R in rows.values():
        if R["txn"] not in report_txns and R["txn"] not in locked:
            R["in_report"] = False


def apply_forwards(state, today):
    rows = state["rows"]
    for k, R in list(rows.items()):
        if R.get("active") and R["status"] == ST_FWD and R.get("fwd_to"):
            R["active"] = False
            nk = key(R["txn"], R["fwd_to"])
            if nk not in rows or not rows[nk].get("active"):
                base = {kk: R.get(kk) for kk in ("txn", "subject", "first", "importance", "moves", "in_report", "created")}
                rows[nk] = {**base, "person": R["fwd_to"], "status": ST_NEW, "source": "forward",
                            # the deadline keeps running from the original receipt (the system has not moved it)
                            "dept": "", "recv": R.get("recv") or iso(today), "fwd_to": None, "fwd_date": None,
                            "close_date": None, "closed_by": None, "notes": "", "last_contact": None,
                            "last_reply": None, "reply_text": "", "suggestion": "", "active": True}
                log(state, today, R["txn"], R["fwd_to"], "حُوّلت إليه (حسب المتابعة)", f"من {R['person']}")


# ------------------------------------------------------------------ 3. Outlook
class Outlook:
    def __init__(self):
        import win32com.client  # noqa: only on Windows with classic Outlook
        self.app = win32com.client.Dispatch("Outlook.Application")
        self.ns = self.app.GetNamespace("MAPI")
        self.inbox = self.ns.GetDefaultFolder(6)
        self.sent = self.ns.GetDefaultFolder(5)
        self.me = []
        try:
            self.me.append(self.ns.CurrentUser.Name)
            for acc in self.ns.Accounts:
                self.me.append(acc.SmtpAddress)
        except Exception:
            pass

    def recent(self, folder, field, since):
        items = folder.Items
        items.Sort(f"[{field}]", True)
        out = []
        for it in items:
            try:
                if it.Class != 43:          # 43 = MailItem
                    continue
                when = getattr(it, field)
                when = dt.datetime(when.year, when.month, when.day, when.hour, when.minute)
                if when.date() < since:
                    break
                out.append((it, when))
            except Exception:
                continue
        return out

    def folder(self, *path):
        f = self.inbox
        for name in path:
            try:
                f = f.Folders[name]
            except Exception:
                f = f.Folders.Add(name)
        return f


def strip_quoted(body):
    cut = re.split(r"\n\s*(?:From:|من:|-----Original Message|_{10,})", body or "", maxsplit=1)
    return cut[0].strip()


def sync_outlook(state, ol, today):
    since = today - dt.timedelta(days=LOOKBACK_DAYS)
    rows, batches = state["rows"], state["batches"]
    done_ids = set(state["processed_mail"])

    # emails you actually sent (a thread keeps one code, so every daily reply-all is recorded)
    for it, when in ol.recent(ol.sent, "SentOn", since):
        m = CODE_RE.search(it.Subject or "")
        if not m or m.group() not in batches:
            continue
        b = batches[m.group()]
        sends = b.setdefault("sends", [])
        if any(x["id"] == it.EntryID for x in sends):
            continue
        day = iso(when.date())
        sends.append({"id": it.EntryID, "date": day})
        sends.sort(key=lambda x: x["date"])
        b["sent"] = sends[0]["date"]
        b["sent_id"] = sends[-1]["id"]
        for txn in b["txns"]:
            for person in bpeople(b):
                R = rows.get(key(txn, person))
                if R and R.get("active") and R["status"] in (ST_NEW, ST_SENT, ST_REPLIED):
                    if not R.get("last_contact") or day >= R["last_contact"]:
                        R["status"], R["last_contact"] = ST_SENT, day
                    log(state, today, txn, person, "أُرسل بريد متابعة", m.group())

    # replies
    to_file = []
    for it, when in ol.recent(ol.inbox, "ReceivedTime", since):
        if it.EntryID in done_ids:
            continue
        m = CODE_RE.search(it.Subject or "")
        if not m or m.group() not in batches:
            continue
        b = batches[m.group()]
        sender = clean_name(it.SenderName)
        people = bpeople(b)
        targets = [p for p in people if sender and norm(p) == norm(sender)]
        rdate = iso(when.date())
        known = {norm(x) for x in people} | {norm(state["emails"].get(x, "")) for x in people} | \
                {norm(x) for x in TEST_EMAILS} | {norm(x) for x in getattr(ol, "me", [])} | {norm(sender)} | \
                {norm(x) for x in state["settings"].get("cc_list", [])}
        added = []
        try:
            for rc in it.Recipients:
                nm, ad = clean_name(rc.Name), clean_name(getattr(rc, "Address", ""))
                if nm and norm(nm) not in known and norm(ad) not in known:
                    added.append(nm)
        except Exception:
            pass
        subj = it.Subject or ""
        fw = bool(re.match(r"\s*(FW|Fwd|TR|إعادة توجيه|اعادة توجيه|تحويل)\s*[:：]", subj, re.I))
        try:
            body = strip_quoted(it.Body or "")
        except Exception:
            body = ""
        body = re.sub(r"\n\s*\n+", "\n", body).strip()[:1500]
        b.setdefault("replies", []).append({"from": sender or "؟", "date": rdate, "holder": bool(targets),
                                            "fw": fw, "added": added, "text": body})
        mentioned = b["txns"]
        for txn in mentioned:
            for person in targets or people:
                R = rows.get(key(txn, person))
                if not R:
                    continue
                if R["status"] in (ST_NEW, ST_SENT, ST_REPLIED):
                    R["status"] = ST_REPLIED
                R["last_reply"], R["reply_from"] = rdate, sender
                log(state, today, txn, person, "وصل رد", f"من {sender}" + ("" if targets else " (ليس من المرسَل لهم)"))
        to_file.append((it, mentioned))
        state["processed_mail"].append(it.EntryID)

    for it, txns in to_file:   # move after iterating, so the Inbox loop isn't disturbed
        try:
            for extra in txns[1:]:
                it.Copy().Move(ol.folder(OUTLOOK_FOLDER, f"معاملة {extra}"))
            it.Move(ol.folder(OUTLOOK_FOLDER, f"معاملة {txns[0]}"))
        except Exception as e:
            print("  (could not file an email:", e, ")")
    state["processed_mail"] = state["processed_mail"][-3000:]


# ------------------------------------------------------------------ 4. compute
def compute(state, today):
    S = state["settings"]
    cal = Calendar([d_iso(h) for h in S["holidays"]])
    th, dl = S["threshold"], S["deadline"]
    rows = state["rows"]
    txns = {}
    for k, R in rows.items():
        if not R.get("active"):
            continue
        recv = d_iso(R["recv"])
        end = d_iso(R["close_date"]) if R["status"] == ST_CLOSED else today
        R["_days"] = cal.days(d_iso(R.get("first")) or recv if SLA_FROM_FIRST else recv, end)
        R["_first_days"] = cal.days(d_iso(R.get("first")), end)
        R["_band"] = band(R["_days"], R["status"] == ST_CLOSED, th, dl)
        txns.setdefault(R["txn"], []).append(k)

    # whole transaction done only when EVERY active holder closed their part
    for txn, keys in txns.items():
        rs = [rows[k] for k in keys]
        if all(r["status"] == ST_CLOSED for r in rs):
            close = max(d_iso(r["close_date"]) or today for r in rs)
            first = d_iso(rs[0].get("first"))
            if txn not in state["closed"] or state["closed"][txn]["close"] != iso(close):
                total = cal.days(first, close)
                state["closed"][txn] = {"subject": rs[0].get("subject", ""), "people": [r["person"] for r in rs],
                                        "first": iso(first), "close": iso(close), "days": total,
                                        "on_time": total < dl or max(r["_days"] for r in rs) < dl,
                                        "max_person_days": max(r["_days"] for r in rs)}
                log(state, today, txn, "", "اكتملت المعاملة", f"أغلقها الجميع في {close}")
        else:
            state["closed"].pop(txn, None)
    for txn, keys in txns.items():
        n_closed = sum(rows[k]["status"] == ST_CLOSED for k in keys)
        for k in keys:
            rows[k]["_progress"] = f"{n_closed} من {len(keys)} أغلق" if len(keys) > 1 else ""
            rows[k]["_txn_open"] = txn not in state["closed"]
    return cal


def needs_email(R, cal, today, th, dl):
    if not R.get("active") or not R.get("_txn_open") or R["status"] in (ST_CLOSED, ST_FWD):
        return False
    # transactions missing from the latest report still get reminders until you close them
    if R["_days"] < th or R.get("last_contact") == iso(today):
        return False
    if STOP_AFTER_DEADLINE and R["_days"] > dl:
        return False
    if not cal.is_work(today):           # no reminders on Friday / Saturday / holidays
        return False
    if R["status"] == ST_REPLIED:
        # someone who answered today is not reminded again today
        if R.get("last_reply") == iso(today):
            return False
    if R["status"] == ST_REPLIED and not REMIND_DAILY_UNTIL_CLOSED:
        since = cal.days(d_iso(R["last_reply"]), today)
        return since >= REPLIED_REMIND_AFTER or R["_days"] >= dl - 2
    return True


def remaining_text(rem):
    """Unmistakable wording: days LEFT before the deadline, not days passed."""
    if rem > 1:
        return f"باقي {days_ar(rem)} فقط وتنتهي المهلة"
    if rem == 1:
        return "باقي يوم عمل واحد فقط — غداً آخر يوم في المهلة"
    if rem == 0:
        return "اليوم آخر يوم في المهلة"
    n = -rem
    since = "يوم عمل واحد" if n == 1 else "يومَي عمل" if n == 2 else days_ar(n)
    return f"انتهت المهلة منذ {since}"


def elapsed_text(days, recv):
    if days == 0:
        return "استُلمت اليوم"
    return f"مفتوحة منذ {days_ar(days)}"


TD = "padding:6px 8px;border:1px solid #c9ced8;vertical-align:top"
TH = TD + ";background:#1F5F8B;color:#ffffff;font-weight:bold;white-space:nowrap"


def build_emails(state, cal, today, only=None):
    """One e-mail per group of people (a transaction held by several people goes to all of them together)."""
    S = state["settings"]
    th, dl = S["threshold"], S["deadline"]
    rows = state["rows"]
    by_txn = {}
    for R in rows.values():
        if R.get("active") and R.get("_txn_open") and R["status"] not in (ST_CLOSED, ST_FWD):
            by_txn.setdefault(R["txn"], []).append(R)
    # one e-mail per connected group: if A, B, C share a transaction and A (or B) also has others,
    # all of those go in the same e-mail to A, B and C — nobody gets two e-mails on the same day
    need = [rs for txn, rs in by_txn.items() if any(needs_email(R, cal, today, th, dl) for R in rs)]
    parent = {}

    def find(x):
        parent.setdefault(x, x)
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x
    for rs in need:
        ps = [R["person"] for R in rs]
        for p in ps:
            find(p)
        for p in ps[1:]:
            parent[find(p)] = find(ps[0])
    comps = {}
    for rs in need:
        comps.setdefault(find(rs[0]["person"]), []).append(rs)
    groups = {}
    for root, lst in comps.items():
        people = tuple(sorted({R["person"] for rs in lst for R in rs}))
        groups[people] = lst
    order = {B_OVER: 0, B_CRIT: 1, B_WATCH: 2, B_OK: 3}
    F = "font-family:Somar,Arial,sans-serif"
    TDs = f"padding:6px 8px;border:1px solid #B8C3D1;vertical-align:top;{F};font-size:10.5pt;color:#000000"
    THs = f"padding:7px 8px;border:1px solid #152446;background:#152446;color:#FFFFFF;{F};font-size:10.5pt;font-weight:normal"
    mails = []
    for people, txn_groups in groups.items():
        items = []
        for rs in txn_groups:
            r = max(rs, key=lambda x: x["_days"])
            recv = (d_iso(r.get("first")) if SLA_FROM_FIRST else None) or d_iso(r["recv"])
            created = d_iso(r.get("created"))
            items.append({"txn": r["txn"], "subject": r.get("subject") or "",
                          "type": "وارد خارجي",
                          "created": f"{created:%Y/%m/%d}" if created else "",
                          "recv": f"{recv:%Y/%m/%d}",
                          "people": "، ".join(x["person"] for x in rs),
                          "depts": "، ".join(dict.fromkeys(x.get("dept") for x in rs if x.get("dept"))),
                          "dept_name": r.get("dept_name") or "",
                          "days": r["_days"], "elapsed": days_ar(r["_days"]) if r["_days"] else "أقل من يوم عمل",
                          "remaining": left_text(dl - r["_days"]), "band": r["_band"], "rows": rs})
        items.sort(key=lambda i: -i["days"])
        top = min((i["band"] for i in items), key=lambda b: order[b])
        cols = ["رقم المعاملة", "الموضوع", "نوع المعاملة", "تاريخ الإنشاء", "تاريخ الإستلام",
                "تاريخ الإرسال<br>(يرجى التعبئة)", "المستلم", "الإدارة", "اسم الإدارة", "المدة",
                "الأيام المتبقية من المدة المعتمدة"]
        ths = "".join(f"<th style='{THs}'>{c}</th>" for c in cols)
        trs = "".join(
            "<tr>" + "".join(f"<td style='{TDs}{extra}'>{v}</td>" for v, extra in (
                (i["txn"], ";white-space:nowrap"), (i["subject"], ""), (i["type"], ";white-space:nowrap"),
                (i["created"], ";white-space:nowrap"), (i["recv"], ";white-space:nowrap"),
                ("&nbsp;", ";background:#FFF8E1;min-width:90px"), (i["people"], ""), (i["depts"], ""),
                (i["dept_name"], ""), (i["elapsed"], ";white-space:nowrap"),
                (i["remaining"], ";color:#C0392B"))) + "</tr>" for i in items)
        last_day = [i["txn"] for i in items if dl - i["days"] == 1]
        alert = ""
        if last_day:
            alert = ("تبقى يوم واحد فقط على انتهاء المهلة، ويرجى إقفال المعاملة أو التحديث بحالتها"
                     + (f" (رقم {'، '.join(last_day)})" if len(items) > 1 else "") + ".")
        p1 = (f"لغرض احتساب مؤشر نسبة الرد على استفسارات الجهات الخارجية، نأمل التكرم بتزويدنا بتواريخ إغلاق "
              f"المعاملات المكتملة الموضحة في الجدول أدناه، والتي مضى عليها {th} أيام عمل فأكثر وفق الاتفاقية المعتمدة.")
        p2 = ("كما نأمل متابعة المعاملات السرية/ تحت الإجراء المرتبطة بمديري العموم، "
              "وتزويدنا بتواريخ إغلاقها فور الانتهاء منها.")
        p3 = f"المهام المفتوحة {th} أيام عمل فأعلى حسب الاتفاقية:"
        P = f"margin:0 0 10px 0;{F};font-size:11pt;color:#000000"
        html = (f"<div dir='rtl' style='text-align:right;{F};font-size:11pt;color:#000000'>"
                f"<p style='{P}'>تحية طيبة وبعد،</p><p style='{P}'>{p1}</p><p style='{P}'>{p2}</p>"
                + (f"<p style='{P};color:#C0392B;font-weight:bold;border:1px solid #C0392B;padding:6px 10px;"
                   f"display:inline-block'>{alert}</p>" if alert else "")
                + f"<p style='{P};font-weight:bold'>{p3}</p>"
                f"<table dir='rtl' style='border-collapse:collapse;margin:4px 0 14px 0'><tr>{ths}</tr>{trs}</table>"
                f"<p style='{P}'>شاكرين ومقدرين تعاونكم، مع خالص التحية،</p></div>")
        text = ("تحية طيبة وبعد،\n\n" + p1 + "\n" + p2 + "\n\n" + (alert + "\n\n" if alert else "") + p3 + "\n\n"
                + "\n\n".join(
                    f"رقم المعاملة: {i['txn']}\nالموضوع: {i['subject']}\nنوع المعاملة: {i['type']}\n"
                    f"تاريخ الإنشاء: {i['created']}\nتاريخ الإستلام: {i['recv']}\nتاريخ الإرسال (يرجى التعبئة): ______\n"
                    f"المستلم: {i['people']}\nالإدارة: {i['depts']}\nاسم الإدارة: {i['dept_name']}\n"
                    f"المدة: {i['elapsed']}\nالأيام المتبقية من المدة المعتمدة: {i['remaining']}" for i in items)
                + "\n\nشاكرين ومقدرين تعاونكم، مع خالص التحية،")
        title = f"متابعة {count_ar(len(items))}" if len(items) > 1 else f"متابعة المعاملة الخارجية رقم {items[0]['txn']}"
        to = [state["emails"].get(p) or "" for p in people]
        mails.append({"people": list(people), "to": "; ".join(x for x in to if x),
                      "missing": [p for p in people if not state["emails"].get(p)], "cc": "", "band": top,
                      "subject_core": f"إشعار مؤتمت: {title}", "text": text, "html": html,
                      "txns": [i["txn"] for i in items], "prev_id": None})
    mails.sort(key=lambda m: (order[m["band"]], -len(m["txns"])))
    return mails


def left_text(rem):
    if rem > 0:
        return days_ar(rem)
    if rem == 0:
        return "اليوم آخر يوم"
    return "تجاوزت المدة بـ " + ("يومَي عمل" if rem == -2 else days_ar(-rem))


def _insert_top(html_body, ours):
    """Put our content at the top of an Outlook body (keeps the signature and the quoted thread below it)."""
    mt = re.search(r"<body[^>]*>", html_body or "", re.I)
    if not mt:
        return ours + (html_body or "")
    return html_body[:mt.end()] + ours + html_body[mt.end():]


def create_drafts(state, ol, mails, today, mode="draft"):
    """Drafts only — nothing is ever sent. One thread per person/group:
    the first reminder starts it; later reminders are Reply-All in the same thread (same subject and code).
    If the people change, a new thread starts."""
    batches = state["batches"]
    stamp = today.strftime("%y%m%d")
    rows = state["rows"]

    def thread_open(b):
        return any((R := rows.get(key(t, p))) and R.get("active") and R["status"] not in (ST_CLOSED, ST_FWD)
                   and t not in state["closed"] for t in b["txns"] for p in bpeople(b))

    for m in mails:
        people = sorted(m["people"])
        thread_code = None
        for code, b in sorted(batches.items(), key=lambda kv: kv[1].get("created", ""), reverse=True):
            if sorted(bpeople(b)) == people and not b.get("ended") and thread_open(b):
                thread_code = code
                break
        if thread_code:
            b = batches[thread_code]
            b["txns"] = list(dict.fromkeys(b["txns"] + m["txns"]))
            code = thread_code
        else:
            n = 1 + sum(1 for c in batches if c.startswith(f"MR-{stamp}-"))
            code = f"MR-{stamp}-{n:02d}"
            b = batches[code] = {"people": m["people"], "txns": m["txns"], "created": iso(today), "sent": None,
                                 "draft_id": None, "subject": f"[{code}] {m['subject_core']}", "to": m["to"]}
        m["subject"] = b.get("subject") or f"[{code}] {m['subject_core']}"
        m["code"] = code
        m["draft"] = "لم تُنشأ (Outlook غير متاح)"
        if not ol:
            continue
        if b.get("draft_id"):                         # an unsent draft from before: replace it
            try:
                old = ol.ns.GetItemFromID(b["draft_id"])
                if not getattr(old, "Sent", False):
                    old.Delete()
            except Exception:
                pass
            b["draft_id"] = None
        try:
            mail = None
            if b.get("sent_id"):                      # continue the same thread
                try:
                    mail = ol.ns.GetItemFromID(b["sent_id"]).ReplyAll()
                except Exception:
                    mail = None
            if mail is None:
                mail = ol.app.CreateItem(0)
                mail.Subject = m["subject"]
            try:
                mail.GetInspector                      # makes Outlook add your usual signature
            except Exception:
                pass
            mail.To = m["to"]
            mail.CC = "; ".join(state["settings"].get("cc_list", []))
            mail.HTMLBody = _insert_top(mail.HTMLBody, m["html"] + "<br>")
            try:
                mail.Recipients.ResolveAll()
            except Exception:
                pass
            mail.Save()
            b["draft_id"] = mail.EntryID
            b["last_draft"] = iso(today)
            m["draft"] = "مسودة في Outlook" + (" (رد على نفس المراسلة)" if b.get("sent_id") else "")
        except Exception as e:
            m["draft"] = f"تعذر الإنشاء: {e}"
    for code, b in batches.items():                   # a thread whose transactions are all closed is finished
        if b.get("sent") and not thread_open(b):
            b["ended"] = True


def batch_status(state, b, cal, today):
    """(text, kind, waiting_days, still_open) for one e-mail thread."""
    rows = state["rows"]
    parts = [(t, p, rows.get(key(t, p))) for t in b["txns"] for p in bpeople(b)]
    still_open = sorted({t for t, p, R in parts if R and R.get("active") and R["status"] != ST_CLOSED
                         and t not in state["closed"]})
    if not b.get("sent"):
        return "draft", "draft", 0, still_open
    last = max([x["date"] for x in b.get("sends", [])] or [b["sent"]])
    n_closed = sum(1 for t, p, R in parts if R and R["status"] == ST_CLOSED)
    moved = [p for t, p, R in parts if R and not R.get("active") and R["status"] != ST_CLOSED]
    if n_closed == len(parts):
        return "closed", "closed", 0, still_open
    if any(x["date"] >= last for x in b.get("replies", [])):
        return "replied", "replied", 0, still_open
    if moved and not still_open:
        return "moved", "moved", 0, still_open
    return "waiting", "waiting", cal.days(d_iso(last), today), still_open


def write_paths(ws, state, cal, today, names):
    rows, edits, manual = state["rows"], state.get("step_edits", {}), state.get("manual_steps", {})
    S = state["settings"]
    setup(ws, [4, 15, 22, 22, 20, 13, 15, 11, 24, 30])
    ws.column_dimensions["A"].hidden = True
    ws["B1"] = ("مسار كل معاملة مرتب حسب الرقم. تقدرين تكتبين «تاريخ الإرسال/الإغلاق» أو ملاحظة لأي خطوة، "
                "أو تضيفين خطوة في صف «➕» (اختاري «إلى»)، ثم اضغطي «تحديث».")
    ws["B1"].font = f_(True, 11, C_HEAD)
    cols = ["مفتاح", "الخطوة", "من", "إلى", "الإدارة", "تاريخ الاستلام", "تاريخ الإرسال/الإغلاق ✎",
            "المدة", "الوضع", "ملاحظة ✎"]
    txns = set(state.get("paths", {})) | set(manual)
    def num(t):
        return (0, int(t)) if str(t).isdigit() else (1, str(t))
    r = 3
    yellow = PatternFill("solid", fgColor=C_EDIT)
    edit_rows, add_rows = [], []
    for txn in sorted(txns, key=num):
        sys_steps = state.get("paths", {}).get(txn, {}).get("steps", [])
        steps = []
        for st in sys_steps:
            e = edits.get(st["k"], {})
            steps.append({**st, "sent": e.get("sent") or st["sent"], "note": e.get("note", ""),
                          "manual_sent": bool(e.get("sent")), "src": "system"})
        for st in manual.get(txn, []):
            steps.append({**st, "k": f"M|{txn}|{st['id']}", "src": "manual", "dept": "", "manual_sent": bool(st.get("sent"))})
        if not steps:
            continue
        steps.sort(key=lambda x: x.get("recv") or "")
        holders = [R for R in rows.values() if R["txn"] == txn and R.get("active")]
        open_now = [R for R in holders if R["status"] != ST_CLOSED]
        done = txn in state["closed"]
        subject = state.get("paths", {}).get(txn, {}).get("subject") or (holders[0].get("subject") if holders else "")
        groups = []
        for st in steps:   # people who got it at the same moment from the same sender = one hop
            if groups and groups[-1]["from"] == st.get("from") and groups[-1]["recv"][:16] == (st.get("recv") or "")[:16]:
                groups[-1]["to"].append(st.get("to"))
            else:
                groups.append({"from": st.get("from"), "recv": st.get("recv") or "", "to": [st.get("to")]})
        chain_txt, last_names = "", []
        for g in groups:
            tos = "، ".join(x for x in g["to"] if x)
            tos = f"({tos})" if len(g["to"]) > 1 else tos
            if g["from"] and g["from"] in last_names:
                chain_txt += f"  ←  {tos}"
            else:
                chain_txt += (f"   |   " if chain_txt else "") + f"{g['from'] or '؟'}  ←  {tos}"
            last_names = g["to"]
        # title
        if done:
            status_txt = f"🔒 مقفلة — {d_iso(state['closed'][txn]['close']):%Y/%m/%d}"
            strong, light = "5D6880", "ECEEF2"
        else:
            worst = min((R.get("_band", B_OK) for R in open_now), default=B_OK,
                        key=lambda b: {B_OVER: 0, B_CRIT: 1, B_WATCH: 2, B_OK: 3, B_DONE: 4}[b])
            days = max((R.get("_days", 0) for R in open_now), default=0)
            status_txt = (f"🟢 مفتوحة — عند: {'، '.join(R['person'] for R in open_now)} · "
                          f"{days_ar(days) if days else 'استلمها اليوم'}"
                          if open_now else "مفتوحة")
            strong, light = FILL[worst]
        t = ws.cell(r, 2, f"معاملة {txn}")
        t.font, t.fill = Font(name=FONT, bold=True, size=13, color="FFFFFF"), PatternFill("solid", fgColor=strong)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        c = ws.cell(r, 3, subject or "")
        c.font, c.fill = f_(True, 12), PatternFill("solid", fgColor=light)
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=10)
        c = ws.cell(r, 7, status_txt)
        c.font = Font(name=FONT, bold=True, size=14 if done else 12, color=strong if done else C_INK)
        c.fill = PatternFill("solid", fgColor=light)
        ws.row_dimensions[r].height = 26
        r += 1
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
        c = ws.cell(r, 2, "المسار: " + chain_txt)
        c.font, c.alignment = f_(False, 11, C_MUTED), Alignment(horizontal="right", wrap_text=True)
        r += 1
        header(ws, r, cols)
        r += 1
        last_recv = max(st.get("recv") or "" for st in steps)
        for i, st in enumerate(steps, 1):
            recv, sent = d_iso(st.get("recv")), d_iso(st.get("sent"))
            R = rows.get(key(txn, st["to"]))
            is_holder = bool(R and R.get("active") and R["status"] != ST_CLOSED)
            if sent:
                state_txt = "✓ أغلق/أرسل" + (" (يدوي)" if st.get("manual_sent") else "")
                color = "3E8E6A"
            elif is_holder:
                state_txt, color = "⏳ عنده الآن", "C8901B"
            elif R and R["status"] == ST_CLOSED:
                state_txt, color = "✓ أغلق", "3E8E6A"
            elif (st.get("recv") or "") < last_recv:
                state_txt, color = "⚠ تركها مفتوحة", "B3261E"
            else:
                state_txt, color = "—", C_MUTED
            if st["src"] == "manual":
                state_txt += " · خطوة يدوية"
            end = sent or (today if not done else d_iso(state["closed"][txn]["close"]))
            vals = [st["k"], i, st.get("from"), st.get("to"), st.get("dept"), recv, sent,
                    (days_ar(cal.days(recv, end)) if cal.days(recv, end) else "نفس اليوم") if recv else "",
                    state_txt, st.get("note") or None]
            for j, v in enumerate(vals, 1):
                body_cell(ws.cell(r, j, v), "yyyy/mm/dd" if j in (6, 7) else None, wrap=j in (3, 4, 10))
            ws.cell(r, 9).font = Font(name=FONT, bold=True, color=color)
            for j in ([4, 6, 7, 10] if st["src"] == "manual" else [7, 10]):
                ws.cell(r, j).fill = yellow
            edit_rows.append(r)
            r += 1
        if done:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
            c = ws.cell(r, 2, f"🔒  مقفلة  —  أُغلقت بالكامل بتاريخ {d_iso(state['closed'][txn]['close']):%Y/%m/%d}")
            c.font = Font(name=FONT, bold=True, size=16, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="5D6880")
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[r].height = 30
            r += 1
        else:
            ws.cell(r, 1, f"ADD|{txn}")
            ws.cell(r, 2, "➕").font = f_(True, 12, C_HEAD)
            for j in (3, 4, 6, 7, 10):
                ws.cell(r, j).fill = yellow
                ws.cell(r, j).border = BORDER
            ws.cell(r, 9, "← لإضافة خطوة اختاري «إلى»").font = f_(False, 10, C_MUTED)
            add_rows.append(r)
            r += 1
        r += 1
    if r == 3:
        ws.cell(3, 2, "لا توجد مسارات بعد — اضغطي «تحديث» بعد وضع التقرير.").font = f_(False, 11, C_MUTED)
    dv_n = DataValidation(type="list", formula1=f"'{SH_LISTS}'!$A$2:$A${max(2, len(names) + 1)}", allow_blank=True)
    dv_d = DataValidation(type="date", operator="between", formula1="DATE(2020,1,1)", formula2="DATE(2040,12,31)",
                          allow_blank=True, showErrorMessage=True, errorTitle="تاريخ غير صحيح",
                          error="اكتبي التاريخ مثل 2026/10/21")
    ws.add_data_validation(dv_n)
    ws.add_data_validation(dv_d)
    for rr in add_rows:
        dv_n.add(f"C{rr}")
        dv_n.add(f"D{rr}")
        dv_d.add(f"F{rr}:G{rr}")
    for rr in edit_rows:
        dv_d.add(f"G{rr}")
    ws.freeze_panes = "A3"


KIND_FILL = {"waiting": "F8D7D5", "other": "FBF0D6", "moved": "FBF0D6", "replied": "E1F1E8",
             "closed": "ECEEF2", "draft": "E6EEF6"}


def write_workbook(state, cal, mails, today, report_name, skipped, changes=None, prev_info=None):
    S = state["settings"]
    th, dl = S["threshold"], S["deadline"]
    rows = state["rows"]
    wb = Workbook()
    ws_today = wb.active
    ws_today.title = SH_TODAY
    ws_open = wb.create_sheet(SH_OPEN)
    ws_done = wb.create_sheet(SH_DONE)
    ws_em = wb.create_sheet(SH_EMAILS)
    ws_paths = wb.create_sheet(SH_PATHS)
    ws_mail = wb.create_sheet(SH_MAILS)
    ws_set = wb.create_sheet(SH_SET)
    ws_log = wb.create_sheet(SH_LOG)
    ws_l = wb.create_sheet(SH_LISTS)
    ws_l.sheet_state = "hidden"
    ws_log.sheet_state = "hidden"

    # ================= الإعدادات (inputs) =================
    setup(ws_set, [34, 10, 44, 3, 16, 3, 24, 32])
    header(ws_set, 1, ["الإعداد", "القيمة", "ملاحظة"])
    for r, (lbl, val, note) in enumerate([
        ("أبدأ المتابعة بعد (أيام عمل)", th, "تعدّلينها هنا، والسكربت يقرؤها في التشغيل القادم"),
        ("المهلة القصوى (أيام عمل)", dl, "حرجة = آخر يومين قبل المهلة"),
        ("وضع التجربة (الإرسال لبريدك فقط)", "نعم" if S.get("test_mode", True) else "لا",
         "نعم = كل الإيميلات تذهب لبريدك التجريبي فقط. «لا» = الإرسال الحقيقي للموظفين"),
    ], 2):
        ws_set.cell(r, 1, lbl).font = f_()
        c = ws_set.cell(r, 2, val)
        c.font, c.fill = f_(True), PatternFill("solid", fgColor=C_EDIT)
        ws_set.cell(r, 3, note).font = f_(False, 10, C_MUTED)
    ws_set.cell(6, 1, "الجمعة والسبت لا تُحتسب أبداً. يوم الاستلام نفسه لا يُحتسب.").font = f_(False, 10, C_MUTED)
    ws_set.cell(7, 1, "الإجازات: اكتبي كل يوم إجازة في صف (العيدين وغيرها).").font = f_(False, 10, C_MUTED)
    ws_set.cell(8, 1, "البريد: اكتبي بريد كل موظف مرة واحدة. لإضافة شخص جديد اكتبي اسمه وبريده في آخر الجدول.").font = f_(False, 10, C_MUTED)
    dv_t = DataValidation(type="list", formula1='"نعم,لا"', allow_blank=False)
    ws_set.add_data_validation(dv_t)
    dv_t.add("B4")
    header(ws_set, 1, ["الإجازات الرسمية ✎"], col=5)
    for i, h in enumerate(sorted(d_iso(h) for h in S["holidays"]), 2):
        body_cell(ws_set.cell(i, 5, h), "yyyy/mm/dd")
    for i in range(2, 61):
        ws_set.cell(i, 5).fill = PatternFill("solid", fgColor=C_EDIT)
        ws_set.cell(i, 5).number_format = "yyyy/mm/dd"
    dv_h = DataValidation(type="date", operator="between", formula1="DATE(2020,1,1)", formula2="DATE(2040,12,31)",
                          showErrorMessage=True, errorTitle="تاريخ غير صحيح", error="اكتبي التاريخ مثل 2026/10/21")
    ws_set.add_data_validation(dv_h)
    dv_h.add("E2:E300")
    HOL_REF = f"'{SH_SET}'!$E$2:$E$300"
    TH_REF, DL_REF = f"'{SH_SET}'!$B$2", f"'{SH_SET}'!$B$3"
    names = sorted({R["person"] for R in rows.values() if R["person"]} | set(state["emails"]))
    header(ws_set, 1, ["الموظف", "البريد الإلكتروني ✎"], col=7)
    for i, n in enumerate(names, 2):
        ws_set.cell(i, 7, n).font = f_()
        c = ws_set.cell(i, 8, state["emails"].get(n))
        c.font, c.fill = f_(), PatternFill("solid", fgColor=C_EDIT)

    ws_l["A1"], ws_l["B1"] = "الأسماء", "الحالات"
    for i, n in enumerate(names, 2):
        ws_l.cell(i, 1, n)
    for i, s in enumerate(STATUSES, 2):
        ws_l.cell(i, 2, s)

    # ================= المتابعة: one table per person, worst first =================
    chg = {}
    for c in changes or []:
        if c[0] == "انتقلت":
            chg[c[1]] = f"انتقلت (كانت عند {c[3]})"
        elif c[0] != "بدون تغيير":
            chg[c[1]] = c[0]
    open_rows = [R for R in rows.values()
                 if R.get("active") and R.get("_txn_open") and R["status"] != ST_CLOSED]
    by_person = {}
    for R in open_rows:
        by_person.setdefault(R["person"], []).append(R)
    order = {B_OVER: 0, B_CRIT: 1, B_WATCH: 2, B_OK: 3, B_DONE: 4}
    people_order = sorted(by_person, key=lambda p: (-max(r["_days"] for r in by_person[p]),
                                                    -len(by_person[p]), p))
    cols = ["مفتاح", "رقم المعاملة", "الموضوع", "الموظف", "الإدارة", "تاريخ الاستلام", "أيام العمل", "التصنيف",
            "الحالة ▼", "محولة إلى ▼", "تاريخ الإغلاق ✎", "ملاحظاتك ✎", "آخر تواصل", "حالة الإيميل",
            "منذ آخر تقرير", "الماسكون"]
    setup(ws_open, [4, 12, 30, 18, 18, 12, 9, 13, 15, 18, 13, 22, 12, 34, 22, 12])
    ws_open.column_dimensions["A"].hidden = True
    ws_open["B1"] = ("اكتبي «تاريخ الإغلاق» = تُعتبر مغلقة (يتحول لونها رمادي فوراً) وتنتقل إلى «المنجزة» عند "
                     "الضغط على «تحديث». الجداول مرتبة: الأقدم أولاً.")
    ws_open["B1"].font = f_(True, 11, C_HEAD)
    r = 3
    data_rows = []
    for p in people_order:
        rs = sorted(by_person[p], key=lambda x: (-x["_days"], x["txn"]))
        worst = min((x["_band"] for x in rs), key=lambda b: order[b])
        lc = max((d_iso(x["last_contact"]) for x in rs if x.get("last_contact")), default=None)
        waiting = [x for x in rs if x.get("last_contact") and not (x.get("last_reply") and x["last_reply"] >= x["last_contact"])]
        replied = [x for x in rs if x.get("last_reply") and (not x.get("last_contact") or x["last_reply"] >= x["last_contact"])]
        fwd = [x for x in rs if x.get("reply_from") and norm(x["reply_from"]) != norm(p)]
        summary = (f"{len(rs)} مفتوحة · {sum(x['_band'] == B_OVER for x in rs)} انتهت مهلتها · "
                   f"{sum(x['_band'] == B_CRIT for x in rs)} حرجة · الأقدم: {days_ar(rs[0]['_days'])} · "
                   f"آخر تواصل: {lc:%Y/%m/%d}" if lc else
                   f"{len(rs)} مفتوحة · {sum(x['_band'] == B_OVER for x in rs)} انتهت مهلتها · "
                   f"{sum(x['_band'] == B_CRIT for x in rs)} حرجة · الأقدم: {days_ar(rs[0]['_days'])} · لم يُراسل بعد")
        if lc:
            summary += f" · ⏳ بلا رد: {len(waiting)} · ↩ ردّ: {len(replied)}" + (f" · ↪ ردّ غيره: {len(fwd)}" if fwd else "")
        t = ws_open.cell(r, 2, f"👤 {p}")
        t.font, t.fill = Font(name=FONT, bold=True, size=13, color="FFFFFF"), PatternFill("solid", fgColor=FILL[worst][0])
        ws_open.merge_cells(start_row=r, start_column=3, end_row=r, end_column=16)
        s = ws_open.cell(r, 3, summary)
        s.font, s.fill = f_(True, 11), PatternFill("solid", fgColor=FILL[worst][1])
        s.alignment = Alignment(horizontal="right", vertical="center")
        ws_open.row_dimensions[r].height = 24
        header(ws_open, r + 1, cols)
        r += 2
        for R in rs:
            vals = [key(R["txn"], R["person"]), R["txn"], R.get("subject"), R["person"], R.get("dept"),
                    d_iso(R["recv"]), None, None, R["status"], R.get("fwd_to") or None, None, R.get("notes") or None,
                    d_iso(R.get("last_contact")), row_email_status(R, cal, today), chg.get(R["txn"]),
                    R.get("_progress") or None]
            for j, v in enumerate(vals, 1):
                body_cell(ws_open.cell(r, j, v), "yyyy/mm/dd" if j in (6, 11, 13) else None, wrap=j in (3, 14, 15))
            ws_open.cell(r, 7).value = (f'=IF(F{r}="","",MAX(0,NETWORKDAYS.INTL(INT(F{r})+1,'
                                        f'IF(K{r}<>"",INT(K{r}),TODAY()),7,{HOL_REF})))')
            ws_open.cell(r, 8).value = (f'=IF(G{r}="","",IF(OR(I{r}="{ST_CLOSED}",K{r}<>""),"{B_DONE}",'
                                        f'IF(G{r}>={DL_REF},"{B_OVER}",IF(G{r}>={DL_REF}-2,"{B_CRIT}",'
                                        f'IF(G{r}>={TH_REF},"{B_WATCH}","{B_OK}")))))')
            ws_open.cell(r, 7).font = ws_open.cell(r, 8).font = f_(True)
            for j in (9, 10, 11, 12):
                ws_open.cell(r, j).fill = PatternFill("solid", fgColor=C_EDIT)
            data_rows.append(r)
            r += 1
        r += 1
    if not people_order:
        ws_open.cell(3, 2, "لا توجد معاملات مفتوحة 🎉").font = f_(True, 13)
    last = max(3, r)
    dv_s = DataValidation(type="list", formula1=f"'{SH_LISTS}'!$B$2:$B${len(STATUSES) + 1}", allow_blank=True,
                          showErrorMessage=True, errorTitle="اختاري من القائمة", error="اختاري الحالة من القائمة")
    dv_n = DataValidation(type="list", formula1=f"'{SH_LISTS}'!$A$2:$A${max(2, len(names) + 1)}", allow_blank=True,
                          showErrorMessage=True, errorTitle="اختاري من القائمة",
                          error="اختاري الاسم من القائمة. لإضافة اسم جديد أضيفيه في «الإعدادات».")
    dv_d = DataValidation(type="date", operator="between", formula1="DATE(2020,1,1)", formula2="DATE(2040,12,31)",
                          allow_blank=True, showErrorMessage=True, errorTitle="تاريخ غير صحيح",
                          error="اكتبي التاريخ مثل 2026/10/21")
    for dv, col in ((dv_s, "I"), (dv_n, "J"), (dv_d, "K")):
        ws_open.add_data_validation(dv)
        for rr in data_rows:
            dv.add(f"{col}{rr}")
    for b, (strong, light) in FILL.items():
        ws_open.conditional_formatting.add(
            f"H3:H{last}", FormulaRule(formula=[f'$H3="{b}"'], fill=PatternFill("solid", fgColor=strong),
                                       font=Font(name=FONT, bold=True, color="FFFFFF")))
        ws_open.conditional_formatting.add(
            f"B3:G{last}", FormulaRule(formula=[f'AND($A3<>"",$H3="{b}")'], fill=PatternFill("solid", fgColor=light)))

    # ================= المنجزة: every closed part, newest first =================
    setup(ws_done, [12, 30, 18, 18, 13, 13, 10, 10, 14, 22])
    title(ws_done, "المنجزة", "كل جزء أُغلق — منك (بكتابة تاريخ الإغلاق) أو من التقرير. الأحدث أولاً.")
    header(ws_done, 4, ["رقم المعاملة", "الموضوع", "الموظف", "الإدارة", "تاريخ الاستلام", "تاريخ الإغلاق",
                        "أيام العمل", "في الوقت؟", "أُغلقت بواسطة", "المعاملة كاملة؟"])
    closed_rows = [R for R in rows.values() if R.get("active") and R["status"] == ST_CLOSED]
    closed_rows.sort(key=lambda R: (R.get("close_date") or "", R["txn"]), reverse=True)
    for i, R in enumerate(closed_rows, 5):
        days = cal.days(d_iso(R["recv"]), d_iso(R.get("close_date")) or today)
        on_time = days < dl
        whole = "✅ مكتملة" if R["txn"] in state["closed"] else "⏳ بانتظار الآخرين"
        vals = [R["txn"], R.get("subject"), R["person"], R.get("dept"), d_iso(R["recv"]), d_iso(R.get("close_date")),
                days, "نعم" if on_time else "لا", "التقرير" if R.get("closed_by") == "system" else "أنا", whole]
        for j, v in enumerate(vals, 1):
            body_cell(ws_done.cell(i, j, v), "yyyy/mm/dd" if j in (5, 6) else None, wrap=j == 2)
        ws_done.cell(i, 8).fill = PatternFill("solid", fgColor=FILL[B_OK][1] if on_time else FILL[B_OVER][1])
    if closed_rows:
        ws_done.auto_filter.ref = f"A4:J{4 + len(closed_rows)}"
    ws_done.freeze_panes = "A5"

    # ================= المسارات: where every transaction went, by number =================
    write_paths(ws_paths, state, cal, today, names)

    # ================= الإيميلات: every email I sent and what happened =================
    setup(ws_em, [12, 15, 40, 26, 28, 12, 20, 40, 12, 12])
    title(ws_em, "الإيميلات المرسلة",
          "اختاري ✓ في العمود الأول لمن تريدين تذكيره، احفظي، ثم اضغطي «٤. إرسال تذكير للمحددين» في لوحة التحكم.")
    header(ws_em, 4, ["تذكير؟ ▼", "الكود", "عنوان الإيميل", "أُرسل إلى", "البريد", "تاريخ الإرسال",
                      "المعاملات", "الحالة", "أيام بلا رد", "ما زالت مفتوحة"])
    since = iso(today - dt.timedelta(days=LOOKBACK_DAYS))
    em = []
    for code, b in state["batches"].items():
        if (b.get("sent") or b["created"]) < since:
            continue
        text, kind, wait, still = batch_status(state, b, cal, today)
        em.append((code, b, text, kind, wait, still))
    korder = {"waiting": 0, "other": 1, "moved": 2, "replied": 3, "draft": 4, "closed": 5}
    em.sort(key=lambda x: (korder[x[3]], -x[4], x[0]))
    for i, (code, b, text, kind, wait, still) in enumerate(em, 5):
        vals = [None, code, b.get("subject", ""), "، ".join(bpeople(b)), b.get("to", ""), d_iso(b.get("sent")),
                "، ".join(b["txns"]), text, wait or None, len(still) or None]
        for j, v in enumerate(vals, 1):
            body_cell(ws_em.cell(i, j, v), "yyyy/mm/dd" if j == 6 else None, wrap=j in (3, 4, 5, 7, 8))
        ws_em.cell(i, 1).fill = PatternFill("solid", fgColor=C_EDIT)
        ws_em.cell(i, 1).alignment = Alignment(horizontal="center", vertical="top")
        ws_em.cell(i, 8).fill = PatternFill("solid", fgColor=KIND_FILL[kind])
    if not em:
        ws_em.cell(5, 2, "لم تُرسل إيميلات بعد.").font = f_(False, 10, C_MUTED)
    dv_x = DataValidation(type="list", formula1='"✓"', allow_blank=True)
    ws_em.add_data_validation(dv_x)
    dv_x.add(f"A5:A{max(5, 4 + len(em))}")
    if em:
        ws_em.auto_filter.ref = f"A4:J{4 + len(em)}"
    ws_em.freeze_panes = "A5"

    # ================= رسائل اليوم (only if Outlook could not make them) =================
    setup(ws_mail, [20, 26, 50, 90, 22])
    header(ws_mail, 1, ["المستلمون", "البريد", "الموضوع", "نص الرسالة", "الحالة"])
    for i, m in enumerate(mails, 2):
        to_txt = m["to"] + (f"\n(بلا بريد في «{SH_SET}»: {'، '.join(m['missing'])})" if m["missing"] else "")
        for j, v in enumerate(["، ".join(m["people"]), to_txt, m["subject"], m["text"], m["draft"]], 1):
            body_cell(ws_mail.cell(i, j, v), wrap=True)
        ws_mail.cell(i, 1).fill = PatternFill("solid", fgColor=FILL[m["band"]][1])
        ws_mail.row_dimensions[i].height = min(409, 140 + 130 * len(m["txns"]))

    # ================= hidden log =================
    setup(ws_log, [12, 12, 22, 24, 70])
    header(ws_log, 1, ["التاريخ", "المعاملة", "الموظف", "الإجراء", "التفاصيل"])
    for i, e in enumerate(reversed(state["log"][-5000:]), 2):
        for j, v in enumerate([d_iso(e["date"]), e["txn"], e["person"], e["action"], e["details"]], 1):
            body_cell(ws_log.cell(i, j, v), "yyyy/mm/dd" if j == 1 else None)

    # ================= اليوم =================
    setup(ws_today, [30, 16, 22, 16, 34, 14])
    title(ws_today, "متابعة المراسلات الخارجية",
          f"آخر تحديث: {today:%Y/%m/%d} · التقرير: {report_name} · استُبعدت {skipped} حركة داخلية أو فارغة")
    if S.get("test_mode", True):
        ws_today["C4"] = "🧪 وضع التجربة: الإيميلات تذهب فقط إلى " + "، ".join(TEST_EMAILS)
        ws_today["C4"].font = f_(True, 11, "B3261E")
    ws_today["A4"], ws_today["B4"] = "اليوم", "=TODAY()"
    ws_today["B4"].number_format = "yyyy/mm/dd"
    ws_today["A4"].font = ws_today["B4"].font = f_(True)
    header(ws_today, 6, ["التصنيف (موظف × معاملة)", "العدد"])
    for i, b in enumerate([B_OVER, B_CRIT, B_WATCH, B_OK], 7):
        ws_today.cell(i, 1, b).font = f_(True, 11, "FFFFFF")
        ws_today.cell(i, 1).fill = PatternFill("solid", fgColor=FILL[b][0])
        c = ws_today.cell(i, 2, f"=COUNTIF('{SH_OPEN}'!$H:$H,A{i})")
        body_cell(c)
        c.font = f_(True, 14)
    ws_today.cell(11, 1, "معاملات مفتوحة").font = f_()
    body_cell(ws_today.cell(11, 2, len({R["txn"] for R in open_rows})))

    def section(r, label, heads, data, fills=None):
        ws_today.cell(r, 1, label).font = f_(True, 13)
        header(ws_today, r + 1, heads)
        if not data:
            ws_today.cell(r + 2, 1, "لا شيء 👍").font = f_(False, 10, C_MUTED)
            return r + 4
        for i, rowv in enumerate(data, r + 2):
            for j, v in enumerate(rowv, 1):
                body_cell(ws_today.cell(i, j, v), "yyyy/mm/dd" if isinstance(v, dt.date) else None, wrap=True)
            if fills:
                ws_today.cell(i, 1).fill = PatternFill("solid", fgColor=fills[i - r - 2])
        return r + 3 + len(data)

    r = 13
    if changes:
        counts = {}
        for ch in changes:
            counts[ch[0]] = counts.get(ch[0], 0) + 1
        summ = [[k, counts[k]] for k in ("أُغلقت بالكامل", "أُغلق جزء منها", "انتقلت", "اختفت من التقرير",
                                         "جديدة", "أُعيد فتحها") if counts.get(k)]
        r = section(r, f"منذ التقرير السابق ({prev_info['name']})", ["التغيير", "عدد المعاملات"],
                    summ or [["لا تغيير", 0]])
    urgent = sorted([R for R in open_rows if R["_band"] in (B_OVER, B_CRIT)], key=lambda R: -R["_days"])
    r = section(r, "١. انتهت مهلتها أو حرجة", ["الموظف", "المعاملة", "أيام العمل", "حالة الإيميل", "الموضوع"],
                [[R["person"], R["txn"], R["_days"], row_email_status(R, cal, today), R.get("subject")] for R in urgent],
                [FILL[R["_band"]][1] for R in urgent])
    waiting = [x for x in em if x[3] == "waiting" and x[4] >= 1]
    r = section(r, f"٢. أرسلتِ ولم يردوا (اختاريهم في «{SH_EMAILS}» لإرسال تذكير)",
                ["أُرسل إلى", "الكود", "تاريخ الإرسال", "أيام بلا رد"],
                [["، ".join(bpeople(b)), code, d_iso(b["sent"]), w] for code, b, t, k, w, s in waiting])
    recent = today - dt.timedelta(days=4)
    got = [(code, b, x) for code, b in state["batches"].items() for x in b.get("replies", [])
           if d_iso(x["date"]) and d_iso(x["date"]) >= recent]
    got.sort(key=lambda g: g[2]["date"], reverse=True)
    r = section(r, "٣. ردود وصلت مؤخراً (اقرئيها في Outlook ← متابعة المراسلات)",
                ["من ردّ", "على إيميل", "التاريخ", "ملاحظة"],
                [[x["from"], code, d_iso(x["date"]), "" if x.get("holder") else "ليس من المرسَل لهم — ربما حُوّلت له"]
                 for code, b, x in got])
    r = section(r, "٤. تذكيرات اليوم", ["المستلمون", "عدد المعاملات", "الموضوع", "الحالة"],
                [["، ".join(m["people"]), len(m["txns"]), m["subject"], m["draft"]] for m in mails])
    gone = sorted({R["txn"] for R in open_rows if R.get("in_report") is False})
    r = section(r, "٥. اختفت من آخر تقرير (تأكدي منها ثم اكتبي تاريخ الإغلاق)", ["المعاملة"], [[t] for t in gone])

    if not any(m["draft"].startswith("لم يُنشأ") or m["draft"].startswith("تعذر") for m in mails):
        wb.remove(ws_mail)
    wb.active = 0
    for ws in wb.worksheets:
        ws.sheet_view.rightToLeft = True
        ws.sheet_view.tabSelected = ws.title == SH_TODAY
    wb.save(TRACKER)


# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--today")
    ap.add_argument("--report")
    ap.add_argument("--no-outlook", action="store_true")
    ap.add_argument("--refresh-only", action="store_true", help="update the tracker, don't touch Outlook drafts")
    ap.add_argument("--send", action="store_true", help="send today's reminders now")
    ap.add_argument("--send-selected", action="store_true", help="send a reminder for the emails ticked ✓")
    a = ap.parse_args()
    for stream in (sys.stdout, sys.stderr):   # so messages survive being written to a log file on Windows
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    REPORTS_DIR.mkdir(exist_ok=True)

    if TRACKER.exists():
        try:
            with open(TRACKER, "a"):
                pass
        except PermissionError:
            sys.exit("Close Murasalat_Tracker.xlsx in Excel first, then run again.")

    state = load_state()
    was_test = state["settings"].get("test_mode", True)
    read_tracker(state, today)
    if was_test and not state["settings"].get("test_mode", True):
        # switching from test to real: keep settings + emails, archive the practice data
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
        STATE_FILE.with_name(f"test_backup_{stamp}.json").write_text(
            json.dumps(state, ensure_ascii=False), encoding="utf-8")
        state = {"rows": {}, "batches": {}, "processed_mail": [], "log": [], "closed": {},
                 "emails": state["emails"], "settings": {**state["settings"], "test_mode": False}}
        print("Test mode switched OFF: practice data archived, starting fresh with real data.")

    report = Path(a.report) if a.report else newest_report()
    if not state["settings"].get("test_mode", True) and report.name.startswith("dummy_report"):
        sys.exit("Test mode is OFF but the newest report is the dummy one. Delete it from 'reports' "
                 "and put the real system export there.")
    print(f"Report: {report.name}")
    rows, skipped = read_report(report)
    report_txns = current_holders(rows)
    snap = report_snapshot(report_txns)
    prev = state.get("last_report")
    stamp = f"{report.name}|{report.stat().st_mtime:.0f}"
    if prev and prev.get("stamp") == stamp:
        # same file as last time: keep comparing against the report before it
        prev = state.get("prev_report")
    else:
        state["prev_report"] = state.get("last_report")
        state["last_report"] = {"name": report.name, "date": iso(today), "stamp": stamp, "txns": snap}
    changes = compare_reports(prev, snap)
    if prev:
        n_closed = sum(c[0] == "أُغلقت بالكامل" for c in changes)
        print(f"Compared with {prev['name']}: {n_closed} fully closed, "
              f"{sum(c[0] == 'أُغلق جزء منها' for c in changes)} partly closed, "
              f"{sum(c[0] == 'انتقلت' for c in changes)} moved, "
              f"{sum(c[0] == 'اختفت من التقرير' for c in changes)} gone from report.")
    merge_report(state, report_txns, today)
    apply_forwards(state, today)
    for txn, person, day in state.pop("_close_after_forward", []):
        close_holder(state, txn, person, d_iso(day), today)

    ol = None
    if USE_OUTLOOK and not a.no_outlook:
        try:
            ol = Outlook()
            sync_outlook(state, ol, today)
            print("Outlook: checked sent items and replies.")
        except Exception as e:
            ol = None
            print(f"Outlook not available ({e.__class__.__name__}). Email texts are in the tracker instead.")

    cal = compute(state, today)
    selected = state.pop("_selected", [])
    if a.send_selected:
        if not selected:
            print("Nothing ticked: put ✓ next to an email in the 'الإيميلات' sheet, save, and press the button again.")
        mails = build_emails(state, cal, today, only=selected)
    else:
        mails = build_emails(state, cal, today)
    mode = ("none" if a.refresh_only else
            "draft")
    if mode == "send" and not ol:
        print("Cannot send: Outlook is not available.")
        mode = "draft"
    create_drafts(state, ol, mails, today, mode)
    write_workbook(state, cal, mails, today, report.name, skipped, changes, prev)
    for R in state["rows"].values():   # drop calculated fields before saving
        for k in [k for k in R if k.startswith("_")]:
            del R[k]
    save_state(state)

    open_n = len({R["txn"] for R in state["rows"].values() if R.get("active") and R["txn"] not in state["closed"]})
    sent = sum(m.get("draft") == "أُرسل ✓" for m in mails)
    drafts = sum(str(m.get("draft", "")).startswith("مسودة") for m in mails)
    what = (f"{sent} reminder email(s) SENT, {drafts} left as drafts" if mode == "send" else
            f"{drafts} reminder draft(s) created in Outlook" if mode == "draft" and ol else
            f"{len(mails)} reminder(s) due today (not created yet)")
    print(f"Done. Open transactions: {open_n}. {what}.")


if __name__ == "__main__":
    main()

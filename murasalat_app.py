# -*- coding: utf-8 -*-
"""
Workbook engine for «متابعة_المراسلات.xlsm».

Actions (called by the buttons inside the workbook):
    refresh   تحديث المراسلات من ملف التقرير المختار
    preview   عرض قائمة تذكيرات اليوم قبل إنشائها
    send      إنشاء مسودات تذكيرات اليوم في Outlook (لا يتم إرسال أي رسالة)
    replies   متابعة الردود
    track     تحديث الجداول من التعديلات اليدوية
    move      نقل المعاملات المنجزة إلى صفحة المعاملات المقفلة
"""
import datetime as dt
import os
import sys
import time
import traceback
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink

import murasalat as M
from murasalat import (B_CRIT, B_DONE, B_OK, B_OVER, B_WATCH, ST_CLOSED, ST_FWD, bpeople, clean_name, d_iso,
                       iso, key, log, norm, to_date)

APP = M.BASE / "متابعة_المراسلات.xlsm"
ERR = M.BASE / "last_error.txt"

S_MAIN, S_OPEN, S_CLOSED = "الرئيسية", "المفتوحة", "المقفلة"
S_PREVIEW = S_PEOPLE = "المعاملات حسب الموظف"
S_TRACK, S_REPLIES = "المسارات", "تتبع الردود"
OLD_TRACK, OLD_PREVIEW = "تتبع الموظفين والإدارات", "تذكيرات اليوم"
S_ORIG, S_RUN, S_META, S_SET, S_LISTS = "الأصلية", "سجل التحديث", "البيانات الأساسية (Meta Data)", "الإعدادات", "_lists"
ORDER = [S_MAIN, S_PEOPLE, S_OPEN, S_CLOSED, S_TRACK, S_REPLIES, S_ORIG, S_RUN, S_META, S_SET, S_LISTS]
SEND_DRAFT = "مسودات فقط، ويتم الإرسال يدوياً من Outlook"
DEFAULT_FONT = "Somar"

# ------------------------------------------------------------ visual identity
NAVY, NAVY_2, TEAL, GREEN = "152446", "1F3560", "4FAFC4", "8BC34A"   # navy, Tiffany, apple green
INK, MUTED, LINE, PAPER = "000000", "000000", "B8C3D1", "F3F6FA"
HEAD_FILL, LIGHT_BAR = "D9DFEA", "DDF0F4"
EDIT_FILL = PatternFill("solid", fgColor="FFF8E1")
MULTI_FILL = PatternFill("solid", fgColor="FFF2B8")
RED, RED_LIGHT, AMBER_LIGHT, GREEN_LIGHT = "C0392B", "FAD4D0", "FCEBC8", "E4F2D3"
RED_MID = "F28B82"
HOL_REF, TH_REF, DL_REF = "'الإعدادات'!$E$2:$E$300", "'الإعدادات'!$B$2", "'الإعدادات'!$B$3"
FONT_NAME = DEFAULT_FONT
thin = Side(style="thin", color=LINE)
GRID = Border(top=thin, bottom=thin, left=thin, right=thin)


def T(size=10.5, bold=False, color=INK, underline=None):
    return Font(name=FONT_NAME, size=size, bold=bold, color=color, underline=underline)


def widths(ws, ws_widths):
    from openpyxl.utils import get_column_letter
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def head(ws, row, titles, col=1):
    for i, t in enumerate(titles, col):
        c = ws.cell(row, i, t)
        c.font = T(10.5, True)
        c.fill = PatternFill("solid", fgColor=HEAD_FILL)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(top=thin, left=thin, right=thin, bottom=Side(style="medium", color=NAVY))
    ws.row_dimensions[row].height = 30


def cell(c, fmt=None, wrap=False, center=False):
    c.font = T()
    c.border = GRID
    c.alignment = Alignment(vertical="center", wrap_text=wrap, horizontal="center" if center else "right")
    if fmt:
        c.number_format = fmt


def title(ws, text, sub=None, col=2):
    c = ws.cell(1, col, text)
    c.font = T(16, False)
    ws.row_dimensions[1].height = 30
    if sub:
        s = ws.cell(2, col, sub)
        s.font = T(10, False, MUTED)


def bar(ws, row, c1, c2, text, color=NAVY, size=12):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    c = ws.cell(row, c1, text)
    c.font = T(size, True)
    c.fill = PatternFill("solid", fgColor=color)
    c.border = Border(right=Side(style="thick", color=NAVY))
    c.alignment = Alignment(horizontal="right", vertical="center", indent=1)
    ws.row_dimensions[row].height = 26


def q(sheet):
    return "'" + sheet + "'"


def button(wb, ws, rng, text, act, color=NAVY):
    first = rng.split(":")[0]
    ws.merge_cells(rng)
    c = ws[first]
    c.value = text
    c.font = T(12, True)
    c.fill = PatternFill("solid", fgColor="CDECF2")
    c.border = Border(top=Side(style="medium", color=NAVY), bottom=Side(style="medium", color=NAVY),
                      left=Side(style="medium", color=NAVY), right=Side(style="medium", color=NAVY))
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.hyperlink = Hyperlink(ref=first, location=f"ACT_{act}", display=text)
    colL = "".join(ch for ch in first if ch.isalpha())
    rowN = "".join(ch for ch in first if ch.isdigit())
    name = f"ACT_{act}"
    if name in wb.defined_names:
        del wb.defined_names[name]
    wb.defined_names[name] = DefinedName(name, attr_text=f"{q(ws.title)}!${colL}${rowN}")


def nav(ws, rng, sheet, text, color=TEAL):
    first = rng.split(":")[0]
    if ":" in rng:
        ws.merge_cells(rng)
    c = ws[first]
    c.value = text
    c.font = T(11)
    c.fill = PatternFill("solid", fgColor="E4F2D3")
    c.border = GRID
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.hyperlink = Hyperlink(ref=first, location=f"{q(sheet)}!B1", display=text)


def link(c, sheet, ref, text):
    c.value = text
    c.hyperlink = Hyperlink(ref=c.coordinate, location=f"{q(sheet)}!{ref}", display=text)
    c.font = T(10, False, "000000", underline="single")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = GRID


def home(ws, ref):
    c = ws[ref]
    c.value = "الصفحة الرئيسية"
    c.hyperlink = Hyperlink(ref=ref, location=f"{q(S_MAIN)}!B1", display="الصفحة الرئيسية")
    c.font = T(10)
    c.fill = PatternFill("solid", fgColor="E4F2D3")
    c.border = GRID
    c.alignment = Alignment(horizontal="center", vertical="center")


def days_txt(n):
    if n == 0:
        return "أقل من يوم عمل"
    if n == 1:
        return "يوم عمل واحد"
    if n == 2:
        return "يوما عمل"
    if 3 <= n <= 10:
        return f"{n} أيام عمل"
    return f"{n} يوم عمل"


def status_formula(r, sent, days, lc, rep):
    return (f'=IF({sent}{r}<>"","تم الإغلاق",IF({days}{r}>={DL_REF}-2,"يجب الإغلاق",'
            f'IF(AND({lc}{r}<>"",{rep}{r}=""),IF(NETWORKDAYS.INTL(INT({lc}{r})+1,TODAY(),7,{HOL_REF})>=1,'
            f'"أُرسلت لآخر مستلم لم يتم الرد","أُرسلت لآخر مستلم بانتظار الرد"),"تحت الإجراء")))')


def color_rules(ws, days_rng, d, status_rng, st):
    black = Font(name=FONT_NAME, color="000000")
    for rng, f, fill in ((days_rng, f"AND(ISNUMBER({d}),{d}>={DL_REF})", RED_MID),
                         (days_rng, f"AND(ISNUMBER({d}),{d}>={DL_REF}-2)", RED_LIGHT),
                         (status_rng, f'{st}="يجب الإغلاق"', RED_MID),
                         (status_rng, f'{st}="تم الإغلاق"', GREEN_LIGHT),
                         (status_rng, f'{st}="أُرسلت لآخر مستلم لم يتم الرد"', AMBER_LIGHT)):
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f], fill=PatternFill("solid", fgColor=fill),
                                                       font=black, stopIfTrue=True))


def add_date_dv(ws, rng):
    dv = DataValidation(type="date", operator="between", formula1="DATE(2020,1,1)", formula2="DATE(2040,12,31)",
                        allow_blank=True, showErrorMessage=True, errorTitle="تاريخ غير صحيح",
                        error="يرجى إدخال التاريخ بالصيغة 2026/10/21")
    ws.add_data_validation(dv)
    dv.add(rng)


def add_name_dv(ws, rng, names):
    dv = DataValidation(type="list", formula1=f"{q(S_LISTS)}!$A$2:$A${max(2, len(names) + 1)}",
                        allow_blank=True, showErrorMessage=False)
    ws.add_data_validation(dv)
    dv.add(rng)


def band_of(days, th, dl):
    return M.band(days, False, th, dl)



def has_vba(path):
    import zipfile
    try:
        with zipfile.ZipFile(path) as z:
            return any(n.endswith("vbaProject.bin") for n in z.namelist())
    except Exception:
        return False


TEMPLATE = M.BASE / "_buttons_template.xlsm"


def ensure_template():
    """The clean file with only the buttons' code. Every rebuild starts from it, never from a file Excel saved."""
    if TEMPLATE.exists() and has_vba(TEMPLATE):
        return True
    if os.name != "nt":
        return False                                 # testing outside Windows
    import setup_app
    setup_app.create_workbook(TEMPLATE)
    return True


ensure_app = ensure_template   # kept for older callers


def wait_unlocked(path, timeout=120):
    """Excel needs a moment to let go of the file after closing it."""
    lock = path.with_name("~$" + path.name)
    end = time.time() + timeout
    while time.time() < end:
        try:
            if not lock.exists():
                with open(path, "r+b"):
                    return
        except (PermissionError, OSError):
            pass
        time.sleep(0.5)
    raise RuntimeError("The workbook is still open in Excel. Close it and try again.")


def fresh_sheet(wb, title):
    """Replace a sheet with an empty one, keeping its VBA code name so Excel stays happy."""
    code = None
    idx = len(wb.worksheets)
    if title in wb.sheetnames:
        old = wb[title]
        code = old.sheet_properties.codeName
        idx = wb.worksheets.index(old)
        wb.remove(old)
    ws = wb.create_sheet(title, idx)
    if code:
        ws.sheet_properties.codeName = code
    ws.sheet_view.rightToLeft = True
    return ws



def read_original(path):
    """All external records exactly as they are in the report (blank / internal removed)."""
    if path.suffix.lower() == ".csv":
        try:
            raw = pd.read_csv(path, header=None, dtype=str, encoding="utf-8-sig")
        except UnicodeDecodeError:
            raw = pd.read_csv(path, header=None, dtype=str, encoding="cp1256")
    else:
        raw = pd.read_excel(path, header=None, dtype=object)
    hdr = next(i for i in range(min(25, len(raw))) if "رقمالمعامله" in [norm(x) for x in raw.iloc[i]])
    head = [clean_name(x) or f"عمود{i}" for i, x in enumerate(raw.iloc[hdr])]
    tcol = next(i for i, h in enumerate(head) if norm(h) == "نوعالمعامله")
    ncol = next(i for i, h in enumerate(head) if norm(h) == "رقمالمعامله")
    out = []
    for _, r in raw.iloc[hdr + 1:].iterrows():
        vals = ["" if (v is None or (isinstance(v, float) and pd.isna(v))) else
                (v.strftime("%Y-%m-%d %H:%M:%S") if hasattr(v, "strftime") else str(v)) for v in r.tolist()]
        if not vals[ncol].strip() or "خارج" not in norm(vals[tcol]):
            continue
        vals = [M.ALIASES.get(norm(v), v) for v in vals]
        out.append(vals)
    uniq, seen = [], set()
    for v in out:                                  # identical lines in the report are kept once
        t = tuple(v)
        if t not in seen:
            seen.add(t)
            uniq.append(v)
    return head, uniq


def path_people(state, txn):
    """Ordered people for a transaction: the report's path, then every manual move in the order it happened."""
    steps = state.get("paths", {}).get(txn, {}).get("steps", [])
    lst = []
    names = lambda: [x[0] for x in lst]
    if steps:
        first = steps[0].get("from")
        if first:
            lst.append([first, ""])
        for st in sorted(steps, key=lambda x: x.get("recv") or ""):
            if st.get("to") and st["to"] not in names():
                lst.append([st["to"], st.get("dept") or ""])
    manual = [R for R in state["rows"].values() if R["txn"] == txn and R.get("source") == "forward"]
    manual.sort(key=lambda R: (R.get("moved_at") or "", R["person"]))
    for R in manual:
        if R["person"] not in names():
            lst.append([R["person"], R.get("dept") or ""])
    for R in state["rows"].values():               # anyone else holding it now
        if R["txn"] == txn and R.get("active") and R["person"] not in names():
            lst.append([R["person"], R.get("dept") or ""])
    return lst


def apply_path_edits(state, today):
    """Edits made in «المسارات» are applied ONCE, like a transfer — they are not a permanent rule.
    A name added at the end = the transaction moved to that person (from everyone holding it now).
    A current holder's name removed = that person no longer holds it."""
    rows = state["rows"]
    for txn, added, removed, depts in state.pop("_path_edits", []):
        holders = [R for R in rows.values() if R["txn"] == txn and R.get("active")
                   and R["status"] not in (ST_CLOSED, ST_FWD)]
        if added:
            to = added[-1]
            for R in holders:
                if norm(R["person"]) != norm(to):
                    R["status"], R["fwd_to"], R["fwd_date"] = ST_FWD, to, R.get("recv")
            state.setdefault("_new_dept", {})[key(txn, to)] = depts.get(to) or \
                state.get("meta", {}).get(to, {}).get("dept", "")
            log(state, today, txn, to, "نُقلت المعاملة إلى موظف (المسارات)", "")
        for name in removed:
            remaining = [R for R in holders if R.get("active") and norm(R["person"]) != norm(name)]
            if added or not remaining:          # never leave a transaction with nobody holding it
                continue
            for R in holders:
                if norm(R["person"]) == norm(name) and R.get("active"):
                    R["active"] = False
                    holders = remaining
                    log(state, today, txn, name, "أُزيل من مسار المعاملة", "")


# ============================================================ read edits from the workbook
def read_edits(state, wb, today):
    global FONT_NAME
    S = state["settings"]
    if S_SET in wb.sheetnames:
        ws = wb[S_SET]
        by_label = {clean_name(ws.cell(r, 1).value): ws.cell(r, 2).value for r in range(2, 15)}
        for lbl, k in (("بدء التذكير بعد (أيام عمل)", "threshold"), ("المهلة النظامية (أيام عمل)", "deadline")):
            try:
                S[k] = max(1, int(by_label.get(lbl)))
            except (TypeError, ValueError):
                pass
        rep = clean_name(by_label.get("ملف التقرير الأخير"))
        S["report_path"] = rep or S.get("report_path", "")
        fn = clean_name(by_label.get("نوع الخط"))
        if fn:
            S["font"] = fn
        hol = [to_date(r[0]) for r in ws.iter_rows(min_row=2, min_col=5, max_col=5, values_only=True)]
        S["holidays"] = sorted({iso(h) for h in hol if h})
        if clean_name(ws.cell(1, 7).value) == "نسخة دائمة (CC)":
            S["cc_list"] = [clean_name(r[0]) for r in ws.iter_rows(min_row=2, max_row=40, min_col=7, max_col=7,
                                                                   values_only=True) if clean_name(r[0])]
        if clean_name(ws.cell(1, 9).value) == "الاسم كما يرد في التقرير":
            S["aliases"] = [[clean_name(a), clean_name(b)] for a, b in
                            ws.iter_rows(min_row=2, max_row=40, min_col=9, max_col=10, values_only=True)
                            if clean_name(a) and clean_name(b)]
        # older layout kept e-mails in columns G:H of this sheet
        if clean_name(ws.cell(1, 7).value) == "الموظف" and not (S_META in wb.sheetnames):
            for name, email in ws.iter_rows(min_row=2, min_col=7, max_col=8, values_only=True):
                if clean_name(name) and clean_name(email):
                    state["emails"][clean_name(name)] = clean_name(email)
    FONT_NAME = S.get("font", DEFAULT_FONT)

    meta = state.setdefault("meta", {})
    if S_META in wb.sheetnames:
        listed = set()
        for name, dept, email, call, when in wb[S_META].iter_rows(min_row=5, min_col=2, max_col=6, values_only=True):
            name = clean_name(name)
            if not name:
                continue
            listed.add(name)
            meta[name] = {"dept": clean_name(dept), "call": clean_name(call), "when": clean_name(when)}
            email = clean_name(email)
            if email:
                state["emails"][name] = email
            else:
                state["emails"].pop(name, None)
        for gone in [n for n in list(meta) if n not in listed]:
            meta.pop(gone, None)
            state["emails"].pop(gone, None)

    rows = state["rows"]
    if S_OPEN in wb.sheetnames:
        hdr = [clean_name(c.value) for c in wb[S_OPEN][5]]
        sent_col = hdr.index("تاريخ الإرسال") if "تاريخ الإرسال" in hdr else 6
        for row in wb[S_OPEN].iter_rows(min_row=6, max_col=20, values_only=True):
            txn = row[0]
            if not txn or not isinstance(txn, str):
                continue
            typed = to_date(row[sent_col])
            for R in rows.values():
                if R["txn"] == txn and R.get("active") and R["status"] != ST_CLOSED:
                    R["pending_close"] = iso(typed) if typed else None

    force, transfers = [], []
    if OLD_PREVIEW in wb.sheetnames:
        for row in wb[OLD_PREVIEW].iter_rows(min_row=4, max_col=6, values_only=True):
            k, choice = row[0], clean_name(row[5])
            if isinstance(k, str) and k.startswith("X|") and choice == "نعم":
                force.append(k[2:])
    if S_PEOPLE in wb.sheetnames:
        # per person: G = تاريخ الإرسال, H = نقل إلى موظف, J = الإدراج في التذكير
        for row in wb[S_PEOPLE].iter_rows(min_row=5, max_col=10, values_only=True):
            k = row[0]
            if not isinstance(k, str) or not k.startswith("P|"):
                continue
            _, txn, person = k.split("|", 2)
            R = rows.get(key(txn, person))
            if not R or not R.get("active") or R["status"] == ST_CLOSED:
                continue
            typed = iso(to_date(row[6]))
            if typed != state.get("people_written", {}).get(k):      # only what you changed here
                for X in rows.values():
                    if X["txn"] == txn and X.get("active") and X["status"] != ST_CLOSED:
                        X["pending_close"] = typed
            to = clean_name(row[7])
            if to and norm(to) != norm(person):
                transfers.append((txn, person, to))
            if clean_name(row[9]) == "نعم":
                force.append(key(txn, person))
    old = state.get("force", {})
    keep = old.get("keys", []) if old.get("day") == iso(today) else []
    state["force"] = {"day": iso(today), "keys": sorted(set(keep) | set(force))}

    if S_CLOSED in wb.sheetnames:
        hdr = [clean_name(c.value) for c in wb[S_CLOSED][5]]
        if "إعادة الفتح" in hdr:
            col = hdr.index("إعادة الفتح")
            for row in wb[S_CLOSED].iter_rows(min_row=6, values_only=True):
                if row and row[0] and col < len(row) and clean_name(row[col]) == "نعم":
                    state.setdefault("_reopen", []).append(str(row[0]))
    track_sheet = S_TRACK if S_TRACK in wb.sheetnames else OLD_TRACK if OLD_TRACK in wb.sheetnames else None
    if track_sheet:
        lists = {}
        for row in wb[track_sheet].iter_rows(min_row=3, max_col=4, values_only=True):
            k = row[0]
            if not isinstance(k, str) or not (k.startswith("T|") or k.startswith("TADD|")):
                continue
            txn = k.split("|")[1]
            lists.setdefault(txn, [])
            name = clean_name(row[2])
            if name:
                lists[txn].append([name, clean_name(row[3])])
        written = state.get("track_written", {})
        for txn, lst in lists.items():
            now_names = [x[0] for x in lst]
            was = [x[0] for x in written.get(txn, [])]
            if txn not in written or [norm(x) for x in now_names] == [norm(x) for x in was]:
                continue                                # nothing changed (department text is ignored)
            added = [n for n in now_names if norm(n) not in {norm(w) for w in was}]
            removed = [w for w in was if norm(w) not in {norm(n) for n in now_names}]
            if added or removed:
                state.setdefault("_path_edits", []).append(
                    (txn, added, removed, {x[0]: x[1] for x in lst if x[1]}))
    # «نقل إلى موظف» from the per-person sheet: only that person's part moves to the new person
    for txn, person, to in transfers:
        R = rows.get(key(txn, person))
        if not R or R["status"] in (ST_CLOSED, ST_FWD):
            continue
        R["status"], R["fwd_to"], R["fwd_date"] = ST_FWD, to, R.get("recv")
        state.setdefault("_new_dept", {})[key(txn, to)] = state.get("meta", {}).get(to, {}).get("dept", "")
        log(state, today, txn, to, "نُقلت المعاملة إلى موظف", f"من {person}")


DEFAULT_ALIASES = [["حمد عبدالعزيز حمد الحميدي", "ملفي مرشد ناصر الدوسري"]]


def setup_aliases(state):
    S = state["settings"]
    if "aliases" not in S:
        S["aliases"] = [list(x) for x in DEFAULT_ALIASES]
    M.ALIASES = {norm(a): b for a, b in S["aliases"]}


def apply_aliases(state):
    """Replace a name everywhere (e.g. a manager whose work always goes to his delegate)."""
    if not M.ALIASES:
        return
    rows = state["rows"]
    for k in list(rows):
        R = rows[k]
        new = M.ALIASES.get(norm(R["person"]))
        if new:
            nk = key(R["txn"], new)
            R["person"] = new
            del rows[k]
            if nk not in rows:
                rows[nk] = R
        if R.get("fwd_to") and M.ALIASES.get(norm(R["fwd_to"])):
            R["fwd_to"] = M.ALIASES[norm(R["fwd_to"])]
        if R.get("sender") and M.ALIASES.get(norm(R["sender"])):
            R["sender"] = M.ALIASES[norm(R["sender"])]
    for info in state.get("paths", {}).values():
        for st in info.get("steps", []):
            for f in ("from", "to"):
                if st.get(f) and M.ALIASES.get(norm(st[f])):
                    st[f] = M.ALIASES[norm(st[f])]
    for lst in list(state.get("track_written", {}).values()):
        for x in lst:
            if M.ALIASES.get(norm(x[0])):
                x[0] = M.ALIASES[norm(x[0])]
    for b in state["batches"].values():
        if b.get("people"):
            b["people"] = [M.ALIASES.get(norm(p), p) for p in b["people"]]
    for d in (state.get("meta", {}), state["emails"]):
        for n in [n for n in list(d) if M.ALIASES.get(norm(n))]:
            d.pop(n, None)
    orig = state.get("orig")
    if orig:
        orig["rows"] = [[M.ALIASES.get(norm(v), v) if isinstance(v, str) else v for v in r] for r in orig["rows"]]


BACKUPS = M.BASE / "backups"


def backup_state(action):
    import shutil
    if not M.STATE_FILE.exists():
        return
    BACKUPS.mkdir(exist_ok=True)
    shutil.copy(M.STATE_FILE, BACKUPS / f"{dt.datetime.now():%Y%m%d_%H%M%S}_{action}.json")
    for old in sorted(BACKUPS.glob("*.json"))[:-20]:
        old.unlink()


def restore_last_backup():
    files = sorted(BACKUPS.glob("*.json")) if BACKUPS.exists() else []
    if not files:
        return None
    last = files[-1]
    import json
    state = json.loads(last.read_text(encoding="utf-8"))
    action = last.stem.split("_", 2)[-1]
    last.unlink()
    return state, action


def archive_closed(state, cal, today):
    """The closed sheet is a permanent record: each closed transaction is written once and never changes."""
    arc = state.setdefault("archive", {})
    dl = state["settings"]["deadline"]
    for txn, info in state["closed"].items():
        if txn in arc:
            continue
        rs = [R for R in state["rows"].values() if R["txn"] == txn and R.get("active")] or \
             [R for R in state["rows"].values() if R["txn"] == txn]
        if not rs:
            continue
        x = rs[0]
        first = min((d_iso(R.get("first")) or d_iso(R.get("recv")) for R in rs if R.get("recv")), default=None)
        close = d_iso(info.get("close")) or max((d_iso(R.get("close_date")) for R in rs if R.get("close_date")),
                                               default=today)
        days = cal.days(first, close) if first else None
        arc[txn] = {
            "txn": txn, "subject": x.get("subject"), "type": x.get("type") or "خارجية", "created": x.get("created"),
            "first": iso(first), "close": iso(close),
            "depts": "، ".join(dict.fromkeys(R.get("dept") for R in rs if R.get("dept"))),
            "sender": "، ".join(dict.fromkeys(R.get("sender") for R in rs if R.get("sender"))),
            "people": "، ".join(dict.fromkeys(R["person"] for R in rs)),
            "days": days, "on_time": "نعم" if days is not None and days < dl else "لا",
            "from_h": x.get("from_h"), "to_h": x.get("to_h"), "dept_name": x.get("dept_name"),
            "by": "يدوي" if any(R.get("closed_by") == "me" for R in rs) else "التقرير",
            "month": iso(close)[:7] if close else None, "archived_on": iso(today)}


def reopen_closed(state, today, summary):
    """«إعادة الفتح» in the closed sheet: a transaction closed by mistake comes back to the open ones."""
    done = []
    for txn in dict.fromkeys(state.pop("_reopen", [])):
        if txn not in state["closed"] and txn not in state.get("archive", {}):
            continue
        state["closed"].pop(txn, None)
        state.get("archive", {}).pop(txn, None)
        rs = [R for R in state["rows"].values() if R["txn"] == txn]
        best = [R for R in rs if R.get("active")] or rs
        for R in best:
            R["active"] = True
            if R["status"] == ST_CLOSED:
                R["status"] = M.ST_SENT if R.get("last_contact") else M.ST_NEW
                R["close_date"], R["closed_by"], R["pending_close"] = None, None, None
        log(state, today, txn, "", "أُعيد فتح المعاملة يدوياً", "")
        done.append(txn)
    if done:
        summary.append(f"أُعيد فتح المعاملات ({len(done)}): " + "، ".join(done))


def workbook_matches_state(wb, state):
    """True only if this workbook was built from this very data file."""
    try:
        mark = wb[S_LISTS]["Z1"].value if S_LISTS in wb.sheetnames else None
    except Exception:
        mark = None
    if mark:
        return mark == state.get("state_id")
    # older workbooks have no mark: compare the open transactions they show with the data file
    shown = set()
    if S_OPEN in wb.sheetnames:
        shown = {str(r[0]) for r in wb[S_OPEN].iter_rows(min_row=6, max_col=1, values_only=True)
                 if r[0] and isinstance(r[0], str)}
    have = open_set(state)
    if not have:
        return True
    return bool(shown & have)


def open_set(state):
    return {R["txn"] for R in state["rows"].values() if R.get("active") and R["txn"] not in state["closed"]}


def act_rebuild(state, today, summary, src_wb):
    """«اعتماد السجلات الأصلية»: the open transactions become exactly what the «الأصلية» sheet contains.
    Closed transactions and the closed archive are never touched."""
    if not src_wb or S_ORIG not in src_wb.sheetnames:
        summary.append("لا توجد سجلات أصلية لاعتمادها.")
        return
    ws = src_wb[S_ORIG]
    data = [list(r) for r in ws.iter_rows(min_row=4, values_only=True)]
    data = [r for r in data if any(v not in (None, "") for v in r)]
    if len(data) < 2:
        summary.append("صفحة السجلات الأصلية فارغة، ولم يتم أي تغيير.")
        return
    raw = pd.DataFrame(data)
    rows_in, skipped = M.parse_raw(raw)
    report_txns = M.current_holders(rows_in)
    locked = set(state["closed"])
    removed = []
    shown = set(state.get("orig_txns", []))            # what the sheet showed before your edits
    in_sheet = {M.txn_id(r[i]) for r in data[1:] for i, h in enumerate(data[0])
                if norm(h) == "رقمالمعامله" and i < len(r) and r[i] not in (None, "")}
    deliberately_deleted = shown - in_sheet
    for txn in sorted(deliberately_deleted - locked):
        for k in [k for k, R in state["rows"].items() if R["txn"] == txn]:
            del state["rows"][k]
        state.get("paths", {}).pop(txn, None)
        removed.append(txn)
    state["_removed"] = removed
    M.merge_report(state, report_txns, today)
    head_row = [clean_name(v) for v in data[0]]
    state["orig"] = {"name": (state.get("orig") or {}).get("name", "السجلات الأصلية"), "header": head_row,
                     "rows": [["" if v is None else (v.strftime("%Y-%m-%d %H:%M:%S") if hasattr(v, "strftime") else str(v))
                               for v in r] for r in data[1:]]}
    summary.append(f"تم اعتماد السجلات الأصلية: {len(report_txns)} معاملة.")
    if removed:
        summary.append(f"حُذفت من المعاملات المفتوحة لعدم ورودها في السجلات الأصلية ({len(removed)}): "
                       + "، ".join(removed[:60]) + (" ..." if len(removed) > 60 else ""))


# ============================================================ reminder list (who gets an e-mail today)
def install_force(state):
    """Rows you chose to include today even though they would normally be skipped."""
    keys = set(state.get("force", {}).get("keys", []))
    original = getattr(M, "_needs_email_original", M.needs_email)
    M._needs_email_original = original

    def needs(R, cal, today, th, dl):
        if key(R["txn"], R["person"]) in keys and R.get("active") and R.get("_txn_open") \
                and R["status"] not in (ST_CLOSED, ST_FWD):
            return True
        return original(R, cal, today, th, dl)
    M.needs_email = needs


def excluded_today(state, cal, today):
    S = state["settings"]
    th, dl = S["threshold"], S["deadline"]
    out = []
    for R in state["rows"].values():
        if not (R.get("active") and R.get("_txn_open") and R["status"] not in (ST_CLOSED, ST_FWD)):
            continue
        if R.get("_days", 0) < th or M.needs_email(R, cal, today, th, dl):
            continue
        if not cal.is_work(today):
            why = "اليوم ليس يوم عمل"
        elif R.get("in_report") is False:
            why = "المعاملة غير موجودة في آخر تقرير"
        elif R.get("last_contact") == iso(today):
            why = "تمت مراسلة الموظف اليوم"
        elif R.get("last_reply") == iso(today):
            why = "ورد رد من الموظف اليوم"
        else:
            why = "لا يستوفي شروط التذكير"
        out.append((R, why))
    out.sort(key=lambda x: (-x[0].get("_days", 0), x[0]["txn"]))
    return out


def formal_status(state, b, cal, today):
    text, kind, wait, still = M.batch_status(state, b, cal, today)
    rows = state["rows"]
    parts = [(t, p, rows.get(key(t, p))) for t in b["txns"] for p in bpeople(b)]
    n_closed = sum(1 for t, p, R in parts if R and R["status"] == ST_CLOSED)
    partial = f"، وأُغلق {n_closed} من {len(parts)}" if 0 < n_closed < len(parts) else ""
    words = {"draft": "مسودة لم تُرسل بعد", "closed": "أُغلقت جميع المعاملات",
             "replied": "ورد رد", "other": "ورد رد",
             "moved": "انتقلت المعاملة في النظام إلى موظف آخر",
             "waiting": f"لم يرد منذ {days_txt(wait)}" if wait else "أُرسلت اليوم"}
    return words[kind] + (partial if kind not in ("draft", "closed") else ""), kind, wait


# ============================================================ Hijri dates
HIJRI_MONTHS = ["محرم", "صفر", "ربيع الأول", "ربيع الآخر", "جمادى الأولى", "جمادى الآخرة", "رجب", "شعبان",
                "رمضان", "شوال", "ذو القعدة", "ذو الحجة"]
GREG_MONTHS = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر",
               "نوفمبر", "ديسمبر"]


def _hijri_tabular(d):
    jd = d.toordinal() + 1721425
    l = jd - 1948440 + 10632
    n = (l - 1) // 10631
    l = l - 10631 * n + 354
    j = ((10985 - l) // 5316) * ((50 * l) // 17719) + (l // 5670) * ((43 * l) // 15238)
    l = l - ((30 - j) // 15) * ((17719 * j) // 50) - (j // 16) * ((15238 * j) // 43) + 29
    m = (24 * l) // 709
    day = l - (709 * m) // 24
    return 30 * n + j - 30, m, day


def to_hijri(dates):
    """Umm al-Qura (official Saudi calendar) through Windows; arithmetic fallback elsewhere."""
    if os.name == "nt":
        try:
            import subprocess
            ps = "$c=New-Object System.Globalization.UmAlQuraCalendar; " + "; ".join(
                f"$d=[datetime]::ParseExact('{d:%Y-%m-%d}','yyyy-MM-dd',$null); "
                f"'{{0}}-{{1}}-{{2}}' -f $c.GetYear($d),$c.GetMonth($d),$c.GetDayOfMonth($d)" for d in dates)
            out = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True,
                                 timeout=30, creationflags=0x08000000).stdout.split()
            res = [tuple(int(x) for x in o.split("-")) for o in out]
            if len(res) == len(dates):
                return res
        except Exception:
            pass
    return [_hijri_tabular(d) for d in dates]


def hijri_to_greg(items):
    """[(y, m, d)] Umm al-Qura -> [date]."""
    if os.name == "nt":
        try:
            import subprocess
            ps = "$c=New-Object System.Globalization.UmAlQuraCalendar; " + "; ".join(
                f"try {{ $c.ToDateTime({y},{m},{d},0,0,0,0).ToString('yyyy-MM-dd') }} "
                f"catch {{ $c.ToDateTime({y},{m},{max(1, d - 1)},0,0,0,0).AddDays(1).ToString('yyyy-MM-dd') }}"
                for y, m, d in items)
            out = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True,
                                 timeout=30, creationflags=0x08000000).stdout.split()
            if len(out) == len(items):
                return [dt.date.fromisoformat(o) for o in out]
        except Exception:
            pass
    res = []
    for y, m, d in items:                       # arithmetic fallback (outside Windows only)
        g = dt.date(int((y - 1) * 0.970229 + 622), 1, 1)
        found = None
        for want, extra in (((y, m, d), 0), ((y, m, d - 1), 1)):
            for k in range(-400, 800):
                c = g + dt.timedelta(days=k)
                if _hijri_tabular(c) == want:
                    found = c + dt.timedelta(days=extra)
                    break
            if found:
                break
        res.append(found)
    return res


def _hnum(s_):
    import re as _re
    m = _re.match(r"^\s*(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})", str(s_ or "").translate(M.AR_DIGITS))
    if m:
        return int(m[1]), int(m[2]), int(m[3])
    m = _re.match(r"^\s*(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})", str(s_ or "").translate(M.AR_DIGITS))
    return (int(m[3]), int(m[2]), int(m[1])) if m else None


def report_period(hdr, orig):
    fh = lambda y, m, d: f"{d:02d}/{m:02d}/{y}"
    fg = lambda d: f"{d.day:02d}/{d.month:02d}/{d.year}" if d else ""
    now = dt.datetime.now()
    idx = {norm(h): i for i, h in enumerate(hdr)}
    hf, ht = None, None
    if "منتاريخ" in idx and "اليتاريخ" in idx:   # the period you chose when pulling the report
        froms = [_hnum(r[idx["منتاريخ"]]) for r in orig if idx["منتاريخ"] < len(r)]
        tos = [_hnum(r[idx["اليتاريخ"]]) for r in orig if idx["اليتاريخ"] < len(r)]
        froms, tos = [x for x in froms if x], [x for x in tos if x]
        if froms and tos:
            hf, ht = min(froms), max(tos)
    if hf and ht:
        gf, gt = hijri_to_greg([hf, ht])
    else:                                       # no period columns: use the first and last dates in the records
        cols = [i for i, h in enumerate(hdr) if norm(h) in ("تاريخالاستلام", "تاريخالارسال")]
        days = [M.parse_dt(r[i]) for r in orig for i in cols if i < len(r)]
        days = [x.date() for x in days if x]
        if not days:
            return None
        gf, gt = min(days), max(days)
        hf, ht = to_hijri([gf, gt])
    hn = to_hijri([now.date()])[0]
    hour = now.strftime("%I:%M").lstrip("0") + (" ص" if now.hour < 12 else " م")
    return {"from": iso(gf), "to": iso(gt), "from_g": fg(gf), "to_g": fg(gt), "from_h": fh(*hf), "to_h": fh(*ht),
            "at_g": fg(now.date()) + f"، الساعة {hour}", "at_h": fh(*hn)}


# ============================================================ actions
def act_refresh(state, today, summary):
    rp = state["settings"].get("report_path", "")
    path = Path(rp) if rp else None
    if not path or not path.exists():
        summary.append("لم يتم اختيار ملف تقرير صالح.")
        return False
    hdr, orig = read_original(path)
    state["orig"] = {"name": path.name, "header": hdr, "rows": orig}
    rows, skipped = M.read_report(path)
    report_txns = M.current_holders(rows)
    snap = M.report_snapshot(report_txns)
    prev = state.get("last_report")
    changes = M.compare_reports(prev, snap)
    state["prev_report"] = prev
    state["last_report"] = {"name": path.name, "date": iso(today), "stamp": f"{path.name}|{path.stat().st_mtime:.0f}",
                            "txns": snap}
    locked = set(state.get("closed", {}))
    changes = [c for c in changes if c[1] not in locked]
    M.merge_report(state, report_txns, today)
    state["period"] = report_period(hdr, orig)
    if state["period"]:
        p = state["period"]
        summary.append(f"فترة السحب: من {p['from_h']} إلى {p['to_h']}، الموافق من {p['from_g']} إلى {p['to_g']}")
    summary.append(f"اسم التقرير: {path.name}، عدد السجلات الخارجية: {len(orig)}، عدد المعاملات: {len(report_txns)}")
    if M.BAD_DATES:
        bad = sorted(set(M.BAD_DATES))
        summary.append(f"تنبيه: سجلات تعذّر قراءة تاريخ استلامها ولم تُضف ({len(bad)}): " + "، ".join(bad[:40]))
    groups = {}
    for c in changes:
        groups.setdefault(c[0], []).append(c[1])
    labels = {"جديدة": "المعاملات الجديدة", "أُغلقت بالكامل": "المعاملات المقفلة",
              "أُغلق جزء منها": "معاملات أُغلق جزء منها", "انتقلت": "معاملات انتقلت إلى موظف آخر",
              "اختفت من التقرير": "معاملات لم تعد واردة في التقرير", "أُعيد فتحها": "معاملات أُعيد فتحها"}
    for kind, label in labels.items():
        if groups.get(kind):
            summary.append(f"{label} ({len(groups[kind])}): " + "، ".join(groups[kind][:40]))
    if not prev:
        summary.append("هذا أول تقرير، وتبدأ المقارنة من التحديث القادم.")
    return True


def act_move(state, today, summary):
    moved = []
    for R in state["rows"].values():
        if R.get("active") and R.get("pending_close") and R["status"] != ST_CLOSED:
            R["status"], R["close_date"], R["closed_by"] = ST_CLOSED, R["pending_close"], "me"
            R["moved_on"] = iso(today)
            R["pending_close"] = None
            moved.append(R["txn"])
            log(state, today, R["txn"], R["person"], "نُقلت إلى المعاملات المقفلة", R["close_date"])
    moved = list(dict.fromkeys(moved))
    summary.append(f"عدد المعاملات المنقولة إلى المعاملات المقفلة: {len(moved)}"
                   + (f"، وهي: {'، '.join(moved[:40])}" if moved else ""))


# ============================================================ write the workbook
def write_all(state, cal, wb, today, mails, action, excluded):
    S = state["settings"]
    th, dl = S["threshold"], S["deadline"]
    rows = state["rows"]
    sheets = {t: fresh_sheet(wb, t) for t in ORDER}
    for t in list(wb.sheetnames):
        if t not in ORDER and len(wb.sheetnames) > 1:
            wb.remove(wb[t])
    wb._sheets.sort(key=lambda w: ORDER.index(w.title))
    for name in [n for n in list(wb.defined_names) if n.startswith("ACT_") or n == "CFG_REPORT"]:
        del wb.defined_names[name]
    HOL = f"{q(S_SET)}!$E$2:$E$300"
    TH, DL = f"{q(S_SET)}!$B$2", f"{q(S_SET)}!$B$3"

    # ---------------- people & e-mails ----------------
    meta = state.get("meta", {})
    dept_of = {}
    for R in rows.values():
        if R.get("person") and R.get("dept"):
            dept_of[R["person"]] = R["dept"]
    names = sorted({R["person"] for R in rows.values() if R.get("person")} | set(state["emails"]) | set(meta))

    def dept(name):
        return (meta.get(name, {}).get("dept") or dept_of.get(name) or "")

    # ---------------- open / closed ----------------
    open_txns, closed_txns = {}, {}
    for R in rows.values():
        if not R.get("active"):
            continue
        (closed_txns if R["txn"] in state["closed"] else open_txns).setdefault(R["txn"], []).append(R)
    txn_days = {t: max((R.get("_days", 0) for R in rs if R["status"] != ST_CLOSED), default=0)
                for t, rs in open_txns.items()}
    open_order = sorted(open_txns, key=lambda t: (-txn_days[t], t))
    holders_now = sorted({R["person"] for rs in open_txns.values() for R in rs if R["status"] != ST_CLOSED})
    missing = [p for p in holders_now if not state["emails"].get(p)]

    # ---------------- الإعدادات ----------------
    ws = sheets[S_SET]
    widths(ws, [34, 42, 50, 3, 18])
    head(ws, 1, ["الإعداد", "القيمة", "ملاحظات"])
    items = [("بدء التذكير بعد (أيام عمل)", th, "يبدأ إعداد رسائل التذكير بعد انقضاء هذا العدد من أيام العمل."),
             ("المهلة النظامية (أيام عمل)", dl, "تُضاف عبارة «تبقى يوم واحد فقط» في رسالة اليوم السابق لنهاية المهلة."),
             ("طريقة الإرسال", SEND_DRAFT, "لا يقوم النظام بإرسال أي رسالة تلقائياً."),
             ("ملف التقرير الأخير", S.get("report_path", ""), "يُحدَّث تلقائياً عند اختيار ملف التقرير."),
             ("نوع الخط", FONT_NAME, "يُطبَّق على جميع الصفحات عند التحديث التالي.")]
    for r, (lbl, val, note) in enumerate(items, 2):
        cell(ws.cell(r, 1, lbl))
        c = ws.cell(r, 2, val)
        cell(c, wrap=True)
        if r not in (4, 5):
            c.fill = EDIT_FILL
        n = ws.cell(r, 3, note)
        cell(n, wrap=True)
        n.font = T(10, False, MUTED)
    wb.defined_names["CFG_REPORT"] = DefinedName("CFG_REPORT", attr_text=f"{q(S_SET)}!$B$5")
    head(ws, 1, ["الإجازات الرسمية"], col=5)
    for i, h in enumerate(sorted(d_iso(h) for h in S["holidays"]), 2):
        ws.cell(i, 5, h)
    for i in range(2, 61):
        cell(ws.cell(i, 5), "yyyy/mm/dd", center=True)
        ws.cell(i, 5).fill = EDIT_FILL
    dvh = DataValidation(type="date", operator="between", formula1="DATE(2020,1,1)", formula2="DATE(2040,12,31)",
                         allow_blank=True, showErrorMessage=True, errorTitle="تاريخ غير صحيح",
                         error="يرجى إدخال التاريخ بالصيغة 2026/10/21")
    ws.add_data_validation(dvh)
    dvh.add("E2:E300")
    home(ws, "C10")
    head(ws, 1, ["نسخة دائمة (CC)"], col=7)
    ws.column_dimensions["G"].width = 34
    for i, e in enumerate(S.get("cc_list", []), 2):
        ws.cell(i, 7, e)
    for i in range(2, 21):
        cell(ws.cell(i, 7))
        ws.cell(i, 7).fill = EDIT_FILL
    head(ws, 1, ["الاسم كما يرد في التقرير", "يُستبدل في جميع الصفحات بـ"], col=9)
    ws.column_dimensions["H"].width = 3
    ws.column_dimensions["I"].width = 32
    ws.column_dimensions["J"].width = 32
    for i, (a_, b_) in enumerate(S.get("aliases", []), 2):
        ws.cell(i, 9, a_)
        ws.cell(i, 10, b_)
    for i in range(2, 21):
        for j in (9, 10):
            cell(ws.cell(i, j))
            ws.cell(i, j).fill = EDIT_FILL
    ws.cell(12, 1, "النسخة الدائمة: تُضاف هذه العناوين في خانة النسخة (CC) لجميع رسائل التذكير.").font = T(10)
    ws.cell(13, 1, "استبدال الأسماء: أي اسم في العمود الأول يُستبدل بالاسم المقابل في جميع الصفحات والرسائل.").font = T(10)

    lists = sheets[S_LISTS]
    lists.sheet_state = "hidden"
    lists["Z1"] = state.get("state_id")          # ties this workbook to its data file
    for i, n in enumerate(names, 2):
        lists.cell(i, 1, n)

    # ---------------- البيانات الأساسية ----------------
    ws = sheets[S_META]
    widths(ws, [3, 34, 30, 38, 26, 24])
    title(ws, "البيانات الأساسية للموظفين",
          "يُعتمد البريد الإلكتروني الوارد في هذه الصفحة عند إعداد رسائل التذكير. الحقول المظللة باللون الأحمر تستلزم إدخال البريد الإلكتروني.")
    home(ws, "D3")
    head(ws, 4, ["الاسم الكامل", "الإدارة", "البريد الإلكتروني", "اسم المخاطبة المفضل", "وقت التواصل المفضل"], col=2)
    r = 5
    for n in names:
        for j, v in enumerate([n, dept(n), state["emails"].get(n), meta.get(n, {}).get("call"),
                               meta.get(n, {}).get("when")], 2):
            cell(ws.cell(r, j, v))
            ws.cell(r, j).fill = EDIT_FILL
        r += 1
    for _ in range(30):
        for j in (2, 3, 4, 5, 6):
            cell(ws.cell(r, j))
            ws.cell(r, j).fill = EDIT_FILL
        r += 1
    ws.conditional_formatting.add(f"D5:D{r}", FormulaRule(formula=['AND($B5<>"",$D5="")'],
                                                            fill=PatternFill("solid", fgColor=RED_LIGHT)))

    # ---------------- المفتوحة / المقفلة ----------------
    open_row_of = {}

    def txn_table(ws, txn_list, closed):
        # A key | B المسار | C رقم | D الموضوع | E نوع المعاملة | F تاريخ الإنشاء | G تاريخ الإستلام | H تاريخ الإرسال
        # I الإدارات | J المرسل (مخفي) | K المستلم | L المدة | M أيام العمل المنقضية | N الحالة | O من تاريخ
        # P إلي تاريخ | Q إسم الإدارة (مخفي) | R آخر تواصل (مخفي) | S ورد الرد (مخفي)
        widths(ws, [4, 9, 12, 32, 13, 13, 13, 14, 22, 18, 24, 9, 12, 28, 12, 12, 16, 12, 8])
        head(ws, 5, ["مفتاح", "المسار", "رقم المعاملة", "الموضوع", "نوع المعاملة", "تاريخ الإنشاء", "تاريخ الإستلام",
                     "تاريخ الإرسال", "الإدارات", "المرسل", "المستلم", "المدة", "أيام العمل المنقضية", "الحالة",
                     "من تاريخ", "إلي تاريخ", "إسم الإدارة", "آخر تواصل", "ورد الرد"])
        r = 6
        for txn, rs in txn_list:
            holders = [R for R in rs if R["status"] != ST_CLOSED]
            first = min((d_iso(R.get("first")) or d_iso(R["recv"]) for R in rs), default=None)
            if closed:
                sent = max((d_iso(R.get("close_date")) for R in rs if R.get("close_date")), default=None)
            else:
                sent = next((d_iso(R["pending_close"]) for R in holders if R.get("pending_close")), None)
            last_people = holders or rs
            who = "، ".join(dict.fromkeys(R["person"] for R in last_people)) or None
            depts = "، ".join(dict.fromkeys(R.get("dept") or dept(R["person"]) for R in last_people
                                             if R.get("dept") or dept(R["person"]))) or None
            sender = "، ".join(dict.fromkeys(R.get("sender") for R in last_people if R.get("sender"))) or None
            contacted = [R for R in holders if R.get("last_contact")]
            lc = max((d_iso(R["last_contact"]) for R in contacted), default=None)
            replied = bool(contacted) and all(R.get("last_reply") and R["last_reply"] >= R["last_contact"]
                                              for R in contacted)
            x = rs[0]
            vals = [txn, None, txn, x.get("subject"), x.get("type") or "خارجية", d_iso(x.get("created")), first, sent,
                    depts, sender, who, None, None, None, x.get("from_h") or None, x.get("to_h") or None,
                    x.get("dept_name") or None, lc, "نعم" if replied else None]
            for j, v in enumerate(vals, 1):
                cell(ws.cell(r, j, v), "yyyy/mm/dd" if j in (6, 7, 8, 18) else None, wrap=j in (4, 9, 11),
                     center=j in (3, 5, 6, 7, 8, 12, 13, 15, 16))
            ws.cell(r, 12).value = (f'=IF(G{r}="","",MAX(0,NETWORKDAYS.INTL(INT(G{r})+1,'
                                    f'IF(H{r}<>"",INT(H{r}),TODAY()),7,{HOL})))')
            ws.cell(r, 13).value = f'=IF(H{r}<>"",0,L{r})'
            ws.cell(r, 14).value = status_formula(r, "H", "L", "R", "S")
            if not closed and len(holders) > 1:
                for j in range(2, 17):
                    ws.cell(r, j).fill = MULTI_FILL
            if not closed:
                ws.cell(r, 8).fill = EDIT_FILL
                open_row_of[txn] = r
            r += 1
        last = max(6, r - 1)
        for col in ("A", "J", "Q", "R", "S") + (("B",) if closed else ()):
            ws.column_dimensions[col].hidden = True
        color_rules(ws, f"L6:L{last}", "$L6", f"N6:N{last}", "$N6")
        if txn_list:
            ws.auto_filter.ref = f"B5:Q{last}"
        return last

    ws = sheets[S_OPEN]
    last_open = txn_table(ws, [(t, open_txns[t]) for t in open_order], closed=False)
    title(ws, "المعاملات المفتوحة",
          "تُرتب المعاملات من الأطول مدة إلى الأقصر. الصفوف المظللة باللون الأصفر معاملات محالة إلى أكثر من موظف في آن واحد.")
    button(wb, ws, "B3:D3", "نقل المعاملات المنجزة إلى المقفلة", "move")
    ws.row_dimensions[3].height = 28
    ws.merge_cells("F3:L3")
    ws["F3"] = "يُرجى إدخال تاريخ الإرسال للمعاملة المنجزة، ثم الضغط على زر النقل."
    ws["F3"].font = T(10)
    ws["F3"].alignment = Alignment(vertical="center", horizontal="right")
    home(ws, "N3")
    add_date_dv(ws, f"H6:H{last_open + 300}")
    if not open_order:
        ws.cell(6, 4, "لا توجد معاملات مفتوحة.").font = T(11)

    ws = sheets[S_CLOSED]
    widths(ws, [4, 12, 32, 13, 13, 13, 14, 22, 18, 24, 9, 12, 13, 12, 12, 16, 12, 12, 12])
    title(ws, "المعاملات المقفلة",
          "سجل دائم للمعاملات المقفلة، يُضاف إليه فقط ولا يتغير عند تحديث التقارير، ويصلح مصدراً للوحات المؤشرات الشهرية.")
    home(ws, "R3")
    head(ws, 5, ["مفتاح", "رقم المعاملة", "الموضوع", "نوع المعاملة", "تاريخ الإنشاء", "تاريخ الإستلام", "تاريخ الإرسال",
                 "الإدارات", "المرسل", "المستلم", "المدة", "الحالة", "في المدة المعتمدة", "من تاريخ", "إلي تاريخ",
                 "إسم الإدارة", "أُغلقت بواسطة", "شهر الإغلاق", "إعادة الفتح"])
    arc = sorted(state.get("archive", {}).values(), key=lambda a: (a.get("close") or "", a["txn"]), reverse=True)
    for i, a in enumerate(arc, 6):
        vals = [a["txn"], a["txn"], a.get("subject"), a.get("type"), d_iso(a.get("created")), d_iso(a.get("first")),
                d_iso(a.get("close")), a.get("depts") or None, a.get("sender") or None, a.get("people") or None,
                a.get("days"), "تم الإغلاق", a.get("on_time"), a.get("from_h") or None, a.get("to_h") or None,
                a.get("dept_name") or None, a.get("by"), a.get("month")]
        for j, v in enumerate(vals, 1):
            cell(ws.cell(i, j, v), "yyyy/mm/dd" if j in (5, 6, 7) else None, wrap=j in (3, 8, 10),
                 center=j in (2, 4, 5, 6, 7, 11, 12, 13, 14, 15, 17, 18))
        ws.cell(i, 13).fill = PatternFill("solid", fgColor=GREEN_LIGHT if a.get("on_time") == "نعم" else RED_LIGHT)
        cell(ws.cell(i, 19), center=True)
        ws.cell(i, 19).fill = EDIT_FILL
    for col in ("A", "I", "P"):
        ws.column_dimensions[col].hidden = True
    if arc:
        ws.auto_filter.ref = f"B5:R{5 + len(arc)}"
        dvr = DataValidation(type="list", formula1='"نعم"', allow_blank=True)
        ws.add_data_validation(dvr)
        dvr.add(f"S6:S{5 + len(arc)}")
        ws["F3"] = "لإعادة فتح معاملة أُغلقت بالخطأ: اختر «نعم» في عمود «إعادة الفتح» ثم اضغط «تحديث الجداول» في الصفحة الرئيسية."
        ws["F3"].font = T(10)
    else:
        ws.cell(6, 3, "لا توجد معاملات مقفلة.").font = T(11)

    # ---------------- المسارات ----------------
    ws = sheets[S_TRACK]
    widths(ws, [4, 12, 34, 32])
    ws.column_dimensions["A"].hidden = True
    title(ws, "المسارات",
          "لإضافة موظف إلى مسار معاملة: اختر اسمه في الصف الأصفر الفارغ أسفل الجدول، ثم اضغط زر «حفظ التعديلات وتحديث الصفحات». يُعد آخر اسم في الجدول هو الموظف الذي لديه المعاملة حالياً.")
    ws.row_dimensions[2].height = 32
    ws["B2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("B2:D2")
    button(wb, ws, "B3:C3", "حفظ التعديلات وتحديث الصفحات", "track")
    home(ws, "D3")
    ws.row_dimensions[3].height = 28
    r = 5
    written, track_row_of = {}, {}
    for txn in open_order:
        lst = path_people(state, txn)
        written[txn] = [list(x) for x in lst]
        bar(ws, r, 2, 4, f"المعاملة رقم {txn}، {open_txns[txn][0].get('subject') or ''}", LIGHT_BAR)
        track_row_of[txn] = r
        head(ws, r + 1, ["مفتاح", "الترتيب", "الاسم", "الإدارة"])
        r += 2
        for i, (name, dpt) in enumerate(lst, 1):
            for j, v in enumerate([f"T|{txn}|{i}", i, name, dpt or dept(name)], 1):
                cell(ws.cell(r, j, v), center=j == 2)
            r += 1
        ws.cell(r, 1, f"TADD|{txn}")
        cell(ws.cell(r, 2, len(lst) + 1), center=True)
        for j in (3, 4):
            cell(ws.cell(r, j))
            ws.cell(r, j).fill = EDIT_FILL
        add_name_dv(ws, f"C{r}", names)
        r += 2
    state["track_written"] = written
    if not open_order:
        ws.cell(5, 2, "لا توجد معاملات مفتوحة.").font = T(11)
    wso = sheets[S_OPEN]
    for txn, rr in open_row_of.items():
        if txn in track_row_of:
            link(wso.cell(rr, 2), S_TRACK, f"B{track_row_of[txn]}", "عرض")
            if len([R for R in open_txns[txn] if R["status"] != ST_CLOSED]) > 1:
                wso.cell(rr, 2).fill = MULTI_FILL

    # ---------------- المعاملات حسب الموظف + تذكيرات اليوم ----------------
    ws = sheets[S_PEOPLE]
    widths(ws, [4, 12, 32, 13, 9, 28, 14, 22, 30, 14, 10, 8])
    for col in ("A", "K", "L"):
        ws.column_dimensions[col].hidden = True
    title(ws, "المعاملات حسب الموظف وتذكيرات اليوم",
          f"بتاريخ {today:%Y/%m/%d}. يمكن إدخال تاريخ الإرسال أو نقل المعاملة إلى موظف آخر أو إدراجها في تذكير اليوم، ثم الضغط على زر الحفظ لتنعكس على جميع الصفحات.")
    ws.merge_cells("B2:J2")
    ws["B2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 32
    button(wb, ws, "B3:C3", "إنشاء مسودات التذكير في Outlook", "send")
    button(wb, ws, "D3:F3", "حفظ التعديلات وتحديث الصفحات", "track")
    home(ws, "J3")
    ws.row_dimensions[3].height = 28
    due = {}                                             # (txn, person) -> today's reminder state
    for m in mails:
        for t in m["txns"]:
            for p in m["people"]:
                due[(t, p)] = m.get("draft") or "مستحق اليوم، ولم تُنشأ المسودة بعد"
    made_today = {}
    for code, b in state["batches"].items():
        if b.get("created") == iso(today):
            for t in b["txns"]:
                for p in bpeople(b):
                    made_today[(t, p)] = "أُرسلت اليوم" if b.get("sent") else "أُنشئت مسودة اليوم"
    why = {(R["txn"], R["person"]): w for R, w in excluded}
    r = 5
    info = (f"عدد التذكيرات المستحقة اليوم: {len(mails)}، "
            f"عدد المسودات المنشأة اليوم: {len({c for c, b in state['batches'].items() if b.get('created') == iso(today)})}")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
    ws.cell(r, 2, info).font = T(10.5)
    r += 1
    due_missing = sorted({p for m in mails for p in m["people"] if not state["emails"].get(p)})
    if due_missing:
        ws.merge_cells(start_row=r, start_column=2, end_row=r + 1, end_column=10)
        c = ws.cell(r, 2, ("تنبيه: لن تُنشأ أي مسودة قبل إدخال البريد الإلكتروني للموظفين التالية أسماؤهم في صفحة "
                           "البيانات الأساسية: " + "، ".join(due_missing)))
        c.font = T(11, True)
        c.fill = PatternFill("solid", fgColor=RED_LIGHT)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="right")
        ws.row_dimensions[r].height = ws.row_dimensions[r + 1].height = 20
        r += 2
    r += 1
    by_person = {}
    for txn in open_order:
        for R in open_txns[txn]:
            if R["status"] != ST_CLOSED:
                by_person.setdefault(R["person"], []).append(R)
    recent = iso(today - dt.timedelta(days=30))
    done_people = sorted({R["person"] for R in rows.values() if R.get("active") and R["status"] == ST_CLOSED
                          and (R.get("close_date") or "") >= recent} - set(by_person))
    first_data = None
    state["people_written"] = {}
    for p in sorted(by_person, key=lambda p: (-max(R.get("_days", 0) for R in by_person[p]), p)):
        rs = sorted(by_person[p], key=lambda R: -R.get("_days", 0))
        top = rs[0].get("_days", 0)
        mail_state = [made_today.get((R["txn"], p)) or due.get((R["txn"], p)) for R in rs]
        if any(x and x.startswith("أُرسلت") for x in mail_state):
            today_txt = "أُرسل له تذكير اليوم"
        elif any(x and x.startswith("أُنشئت") for x in mail_state):
            today_txt = "أُنشئت مسودة تذكير اليوم"
        elif any(mail_state):
            today_txt = "مستحق له تذكير اليوم"
        else:
            today_txt = "لا يوجد تذكير اليوم"
        email = state["emails"].get(p)
        bar(ws, r, 2, 10, f"الموظف: {p}" + (f"، الإدارة: {dept(p)}" if dept(p) else "")
            + f"، عدد المعاملات: {len(rs)}، أطول مدة: {days_txt(top)}، {today_txt}"
            + ("" if email else "، البريد الإلكتروني غير مسجّل"), LIGHT_BAR)
        head(ws, r + 1, ["مفتاح", "رقم المعاملة", "الموضوع", "تاريخ الإستلام", "المدة", "الحالة",
                         "تاريخ الإرسال", "نقل إلى موظف", "تذكير اليوم", "الإدراج في التذكير"])
        r += 2
        for R in rs:
            k = (R["txn"], p)
            if k in made_today:
                rem = made_today[k]
            elif k in due:
                rem = due[k]
            elif k in why:
                rem = "غير مستحق: " + why[k]
            elif R.get("_days", 0) < th:
                rem = "لم يحن موعد التذكير"
            else:
                rem = ""
            state["people_written"][f"P|{R['txn']}|{p}"] = iso(d_iso(R.get("pending_close")))
            vals = [f"P|{R['txn']}|{p}", R["txn"], R.get("subject"), d_iso(R.get("first")) or d_iso(R["recv"]),
                    None, None, d_iso(R.get("pending_close")), None, rem, None,
                    d_iso(R.get("last_contact")),
                    "نعم" if R.get("last_reply") and R.get("last_contact") and R["last_reply"] >= R["last_contact"] else None]
            for j, v in enumerate(vals, 1):
                cell(ws.cell(r, j, v), "yyyy/mm/dd" if j in (4, 7, 11) else None, wrap=j in (3, 9),
                     center=j in (2, 4, 5, 7, 10))
            ws.cell(r, 5).value = (f'=IF(D{r}="","",MAX(0,NETWORKDAYS.INTL(INT(D{r})+1,'
                                   f'IF(G{r}<>"",INT(G{r}),TODAY()),7,{HOL})))')
            ws.cell(r, 6).value = status_formula(r, "G", "E", "K", "L")
            for j in (7, 8):
                ws.cell(r, j).fill = EDIT_FILL
            if k in why:
                ws.cell(r, 10).fill = EDIT_FILL
            if first_data is None:
                first_data = r
            r += 1
        r += 1
    for p in done_people:
        bar(ws, r, 2, 10, f"الموظف: {p}، تم إقفال جميع المعاملات", LIGHT_BAR)
        r += 2
    lastp = max(7, r)
    color_rules(ws, f"E7:E{lastp}", "$E7", f"F7:F{lastp}", "$F7")
    add_date_dv(ws, f"G7:G{lastp}")
    add_name_dv(ws, f"H7:H{lastp}", names)
    dvx = DataValidation(type="list", formula1='"نعم,لا"', allow_blank=True)
    ws.add_data_validation(dvx)
    dvx.add(f"J7:J{lastp}")
    if not by_person:
        ws.cell(7, 2, "لا توجد معاملات مفتوحة.").font = T(11)

    # ---------------- تتبع الردود: one block per open transaction ----------------
    ws = sheets[S_REPLIES]
    widths(ws, [3, 13, 14, 20, 26, 13, 22, 24, 12, 14, 20, 24, 11, 50])
    title(ws, "تتبع الردود",
          "سجل لكل معاملة مفتوحة يوضح أيام إرسال التذكير والردود الواردة والتحويلات بين الموظفين، بدءاً من اليوم الخامس.")
    home(ws, "H3")
    threads_of = {}
    for code, b in state["batches"].items():
        for t in b["txns"]:
            threads_of.setdefault(t, []).append((code, b))
    moves_of = {}
    for e in state.get("log", []):
        if e["action"] in ("نُقلت المعاملة إلى موظف", "حُوّلت إليه (حسب المتابعة)", "أصبح الماسك الحالي (تعديل يدوي)"):
            moves_of.setdefault((e["txn"], e["date"]), []).append(f"نُقلت إلى {e['person']}")
    waiting_threads = 0
    O = q(S_OPEN)
    r = 5
    for txn in open_order:
        rs = open_txns[txn]
        x = rs[0]
        holders = [R for R in rs if R["status"] != ST_CLOSED]
        first = min((d_iso(R.get("first")) or d_iso(R["recv"]) for R in rs), default=today)
        orow = open_row_of.get(txn)
        bar(ws, r, 2, 14, f"المعاملة رقم {txn}، {x.get('subject') or ''}", LIGHT_BAR)
        r += 1
        head(ws, r, ["رقم المعاملة", "نوع المعاملة", "تاريخ الإستلام", "المستلم", "الإدارات", "المدة", "الحالة"], col=2)
        r += 1
        rec = [txn, x.get("type") or "خارجية", first, "، ".join(R["person"] for R in holders),
               "، ".join(dict.fromkeys(R.get("dept") or dept(R["person"]) for R in holders
                                       if R.get("dept") or dept(R["person"]))),
               f"={O}!L{orow}" if orow else None, f"={O}!N{orow}" if orow else None]
        for j, v in enumerate(rec, 2):
            cell(ws.cell(r, j, v), "yyyy/mm/dd" if j == 4 else None, wrap=j in (5, 6, 8), center=j in (2, 3, 4, 7))
        color_rules(ws, f"G{r}:G{r}", f"$G{r}", f"H{r}:H{r}", f"$H{r}")
        r += 2
        # day-by-day log from day 5
        head(ws, r, ["يوم العمل", "التاريخ", "التذكير", "المرسل إليهم", "الرد", "مرسل الرد", "التحويل"], col=2)
        r += 1
        cur = txn_days.get(txn, 0)
        last_n = max(dl, cur)
        dates, d, n = {}, first, 0
        while n < last_n:
            d += dt.timedelta(days=1)
            if cal.is_work(d):
                n += 1
                dates[n] = d
        steps = state.get("paths", {}).get(txn, {}).get("steps", [])
        for n in range(th, last_n + 1):
            day = dates.get(n)
            if not day:
                continue
            ds = iso(day)
            sent_to, repl, made = [], [], False
            for code, b in threads_of.get(txn, []):
                if any(s["date"] == ds for s in b.get("sends", [])):
                    sent_to += bpeople(b)
                elif b.get("last_draft") == ds and not b.get("sends"):
                    made = True
                repl += [y["from"] for y in b.get("replies", []) if y["date"] == ds]
            moves = [f"انتقلت إلى {st['to']}" for st in steps[1:] if (st.get("recv") or "")[:10] == ds]
            moves += moves_of.get((txn, ds), [])
            future = day > today
            vals = [n, day,
                    "تم الإرسال" if sent_to else ("أُنشئت مسودة" if made else ("" if future else "لم يُرسل")),
                    "، ".join(dict.fromkeys(sent_to)) or None,
                    "ورد رد" if repl else ("" if future or not sent_to else "لم يرد"),
                    "، ".join(dict.fromkeys(repl)) or None, "، ".join(dict.fromkeys(moves)) or None]
            for j, v in enumerate(vals, 2):
                cell(ws.cell(r, j, v), "yyyy/mm/dd" if j == 3 else None, wrap=j in (5, 7, 8), center=j in (2, 3, 4, 6))
            if sent_to:
                ws.cell(r, 4).fill = PatternFill("solid", fgColor=GREEN_LIGHT)
            if repl:
                ws.cell(r, 6).fill = PatternFill("solid", fgColor=GREEN_LIGHT)
            if day == today:
                for j in range(2, 9):
                    ws.cell(r, j).border = Border(top=Side(style="medium", color=NAVY), bottom=Side(style="medium", color=NAVY),
                                                  left=thin, right=thin)
            if future:
                for j in range(2, 9):
                    ws.cell(r, j).fill = PatternFill("solid", fgColor=PAPER)
            r += 1
        r += 1
        # the e-mail threads for this transaction
        head(ws, r, ["عنوان الرسالة", "المرسل إليهم", "تاريخ الإرسال", "ورد رد", "عدد الردود", "مرسل الرد",
                     "تاريخ آخر رد", "إعادة توجيه", "أُضيف إلى المراسلة", "الحالة", "أيام دون رد", "نص آخر رد",
                     "عدد مرات التذكير"], col=2)
        r += 1
        ths = threads_of.get(txn, [])
        if not ths:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=14)
            ws.cell(r, 2, "لم تُرسل رسائل متابعة لهذه المعاملة بعد.").font = T(10)
            r += 1
        for code, b in sorted(ths, key=lambda kv: kv[1].get("created", "")):
            text, kind, wait = formal_status(state, b, cal, today)
            reps = [y for y in b.get("replies", []) if not b.get("sent") or y["date"] >= b["sent"]]
            if kind == "waiting":
                waiting_threads += 1
            fws = [y["from"] for y in reps if y.get("fw") or not y.get("holder")]
            added = sorted({a for y in reps for a in y.get("added", [])})
            lastrep = max(reps, key=lambda y: y["date"]) if reps else None
            vals = [b.get("subject", ""), "، ".join(bpeople(b)), d_iso(b.get("sent")),
                    ("نعم" if reps else "لا") if b.get("sent") else "مسودة", len(reps) or None,
                    "، ".join(dict.fromkeys(y["from"] for y in reps)) or None,
                    d_iso(lastrep["date"]) if lastrep else None,
                    ("نعم: " + "، ".join(dict.fromkeys(fws))) if fws else ("لا" if reps else None),
                    "، ".join(added) or None, text, wait if kind == "waiting" else None,
                    (lastrep.get("text") or None) if lastrep else None, len(b.get("sends", [])) or None]
            for j, v in enumerate(vals, 2):
                cell(ws.cell(r, j, v), "yyyy/mm/dd" if j in (4, 8) else None, wrap=j in (2, 3, 7, 9, 10, 11, 13),
                     center=j in (4, 5, 6, 8, 12, 14))
            ws.cell(r, 5).fill = PatternFill("solid", fgColor=GREEN_LIGHT if reps else (RED_LIGHT if b.get("sent") else PAPER))
            r += 1
        r += 2
    if not open_order:
        ws.cell(5, 2, "لا توجد معاملات مفتوحة.").font = T(11)

    # ---------------- الأصلية ----------------
    ws = sheets[S_ORIG]
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    orig = state.get("orig")
    ws["A1"] = "السجلات الأصلية"
    ws["A1"].font = T(16, False)
    home(ws, "B2")
    button(wb, ws, "C3:E3", "اعتماد السجلات الأصلية وتحديث جميع الصفحات", "rebuild")
    ws.row_dimensions[3].height = 28
    ws["F3"] = ("يمكن حذف السجلات غير المطلوبة أو تعديلها في هذه الصفحة ثم الضغط على الزر، فتصبح المعاملات المفتوحة "
                "مطابقة لهذه السجلات. لا يتأثر سجل المعاملات المقفلة.")
    ws["F3"].font = T(10)
    if orig:
        ws["C2"] = f"مصدر البيانات: {orig['name']}، ويقتصر على الوارد الخارجي بعد استبعاد السجلات الفارغة والداخلية."
        ws["C2"].font = T(10, False, MUTED)
        head(ws, 4, orig["header"])
        ni = next((i for i, h in enumerate(orig["header"]) if norm(h) == "رقمالمعامله"), None)
        state["orig_txns"] = sorted({M.txn_id(r[ni]) for r in orig["rows"] if ni is not None and ni < len(r) and r[ni]})
        for i, rv in enumerate(orig["rows"], 5):
            for j, v in enumerate(rv, 1):
                cell(ws.cell(i, j, v))
        for j, h in enumerate(orig["header"], 1):
            ws.column_dimensions[ws.cell(4, j).column_letter].width = min(40, max(12, len(str(h)) + 6))
        ws.auto_filter.ref = f"A4:{ws.cell(4, len(orig['header'])).column_letter}{4 + len(orig['rows'])}"
    else:
        ws["C2"] = "لم يتم تحميل تقرير بعد. يُرجى الضغط على زر «تحديث المراسلات» في الصفحة الرئيسية."
        ws["C2"].font = T(10, False, MUTED)

    # ---------------- سجل التحديث ----------------
    ws = sheets[S_RUN]
    widths(ws, [20, 30, 110])
    ws["A1"] = "سجل التحديثات"
    ws["A1"].font = T(16, False)
    home(ws, "B2")
    head(ws, 4, ["الوقت", "الإجراء", "التفاصيل"])
    for i, e in enumerate(state.get("runlog", [])[:300], 5):
        for j, v in enumerate([e["at"], e["action"], "\n".join(e["lines"])], 1):
            cell(ws.cell(i, j, v), wrap=j == 3, center=j == 1)

    # ---------------- الرئيسية ----------------
    ws = sheets[S_MAIN]
    widths(ws, [3, 24, 24, 24, 3, 24, 24, 24, 3])
    ws.merge_cells("B1:H1")
    ws["B1"] = "نظام متابعة المراسلات الخارجية"
    ws["B1"].font = T(20, False)
    ws["B1"].alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[1].height = 40
    last = state.get("runlog", [{}])[0] if state.get("runlog") else {}
    ws.merge_cells("B2:E2")
    button(wb, ws, "F2:H2", "التراجع عن آخر إجراء", "undo")
    ws.row_dimensions[2].height = 26
    ws["B2"] = f"آخر إجراء: {last.get('action', 'لا يوجد')}" + (f"، بتاريخ {last['at']}" if last.get("at") else "")
    ws["B2"].font = T(10.5, False, MUTED)
    for c in range(2, 9):                                   # thin brand line in the identity colours
        ws.cell(3, c).fill = PatternFill("solid", fgColor=TEAL if c < 5 else GREEN)
    ws.row_dimensions[3].height = 4

    ws["B5"] = "الإجراءات"
    ws["B5"].font = T(13, False)
    actions = [("B6:D7", "تحديث المراسلات", "refresh", "B8",
                "اختيار ملف التقرير وتحديث جميع الصفحات، وعرض سجل التحديث."),
               ("F6:H7", "قائمة تذكيرات اليوم", "preview", "F8",
                "عرض المستلمين والمعاملات قبل إنشاء المسودات."),
               ("B10:D11", "متابعة الردود", "replies", "B12",
                "مراجعة الرسائل المرسلة والردود الواردة عليها."),
               ("F10:H11", "تحديث الجداول", "track", "F12",
                "إعادة ترتيب الصفحات وفق التعديلات اليدوية.")]
    for rng, txt, act, drow, desc in actions:
        button(wb, ws, rng, txt, act)
        r0 = int(drow[1:])
        c0 = drow[0]
        ws.merge_cells(f"{c0}{r0}:{chr(ord(c0) + 2)}{r0}")
        ws[drow] = desc
        ws[drow].font = T(9.5, False, MUTED)
        ws[drow].alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
        ws.row_dimensions[r0].height = 30

    ws["B14"] = "الانتقال إلى الصفحات"
    ws["B14"].font = T(13, False)
    navs = [("B15", S_OPEN, "المعاملات المفتوحة"), ("C15", S_CLOSED, "المعاملات المقفلة"),
            ("D15", S_PEOPLE, "المعاملات حسب الموظف وتذكيرات اليوم"), ("F15", S_TRACK, "المسارات"),
            ("G15", S_REPLIES, "متابعة الردود"), ("H15", S_ORIG, "السجلات الأصلية"),
            ("B17", S_RUN, "سجل التحديثات"), ("C17", S_META, "البيانات الأساسية"), ("D17", S_SET, "الإعدادات")]
    for ref, sheet, txt in navs:
        nav(ws, ref, sheet, txt)
    ws.row_dimensions[15].height = ws.row_dimensions[17].height = 30

    ws["B19"] = "المؤشرات"
    ws["B19"].font = T(13, False)
    head(ws, 20, ["المؤشر", "العدد"], col=2)
    O = q(S_OPEN)
    stats = [("المعاملات المفتوحة", f'=COUNTIFS({O}!A6:A3000,"<>",{O}!N6:N3000,"<>تم الإغلاق")'),
             ("معاملات يجب إغلاقها", f'=COUNTIF({O}!N6:N3000,"يجب الإغلاق")'),
             ("معاملات تجاوزت المهلة", f'=COUNTIFS({O}!A6:A3000,"<>",{O}!N6:N3000,"<>تم الإغلاق",{O}!L6:L3000,">="&{DL})'),
             ("أُرسلت لآخر مستلم ولم يتم الرد", f'=COUNTIF({O}!N6:N3000,"أُرسلت لآخر مستلم لم يتم الرد")'),
             ("بانتظار النقل إلى المقفلة", f'=COUNTIF({O}!N6:N3000,"تم الإغلاق")'),
             ("المعاملات المقفلة", f'=COUNTA({q(S_CLOSED)}!A6:A5000)'),
             ("رسائل لم يرد عليها", waiting_threads)]
    for i, (lbl, fml) in enumerate(stats, 21):
        cell(ws.cell(i, 2, lbl))
        cell(ws.cell(i, 3, fml), center=True)

    # report period (Hijri and Gregorian)
    ws["F19"] = "بيانات التقرير"
    ws["F19"].font = T(13, False)
    head(ws, 20, ["البيان", "التاريخ الهجري", "التاريخ الميلادي"], col=6)
    per = state.get("period")
    if per:
        lines = [("بداية فترة السحب", per["from_h"], per["from_g"]),
                 ("نهاية فترة السحب", per["to_h"], per["to_g"]),
                 ("آخر تحديث للتقرير", per["at_h"], per["at_g"])]
        for i, (a1, a2, a3) in enumerate(lines, 21):
            for j, v in enumerate((a1, a2, a3), 6):
                cell(ws.cell(i, j, v), wrap=True, center=j > 6)
        ws.merge_cells("F24:H25")
        ws["F24"] = (f"تم سحب التقرير للفترة من {per['from_h']} إلى {per['to_h']}، "
                     f"الموافق من {per['from_g']} إلى {per['to_g']}، "
                     f"وآخر نسخة من التقرير حُدِّثت بتاريخ {per['at_h']}، الموافق {per['at_g']}.")
        ws["F24"].font = T(10.5)
        ws["F24"].alignment = Alignment(wrap_text=True, vertical="top", horizontal="right")
        ws.row_dimensions[24].height = ws.row_dimensions[25].height = 24
    else:
        ws.merge_cells("F21:H21")
        ws["F21"] = "لم يتم تحميل تقرير بعد."
        ws["F21"].font = T(10.5)
    notes = []
    if missing:
        notes.append(f"يوجد {len(missing)} من الموظفين دون بريد إلكتروني مسجّل في صفحة البيانات الأساسية: "
                     + "، ".join(missing))
    for i, txt in enumerate(notes):
        row0 = 27 + i * 2
        ws.merge_cells(start_row=row0, start_column=6, end_row=row0 + 1, end_column=8)
        c = ws.cell(row0, 6, txt)
        c.font = T(10.5, True)
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="right")
    if missing:
        nav(ws, f"F{27 + len(notes) * 2}", S_META, "إدخال البريد الإلكتروني")

    for w in wb.worksheets:
        w.freeze_panes = None
        w.sheet_view.rightToLeft = True
        w.sheet_view.showGridLines = False
    open_on = {"refresh": S_RUN, "rebuild": S_RUN, "undo": S_RUN, "move": S_CLOSED, "send": S_PEOPLE, "preview": S_PEOPLE,
               "replies": S_REPLIES}.get(action, S_MAIN)
    wb.active = wb.sheetnames.index(open_on)
    for w in wb.worksheets:
        w.sheet_view.tabSelected = w.title == open_on
    return missing


# ============================================================ main
LABELS = {"undo": "التراجع عن آخر إجراء", "rebuild": "اعتماد السجلات الأصلية", "send": "إنشاء مسودات التذكير", "preview": "عرض قائمة تذكيرات اليوم", "refresh": "تحديث المراسلات",
          "replies": "متابعة الردود", "track": "تحديث الجداول", "move": "نقل المعاملات إلى المقفلة"}


def run(action, today=None, use_outlook=True, reopen=True):
    today = today or dt.date.today()
    have_template = ensure_template()
    edits = M.BASE / "~edits.xlsm"
    if APP.exists():
        wait_unlocked(APP)
    state = M.load_state()
    state["settings"]["test_mode"] = False          # live mode only
    if not state.get("live_since"):                 # threads from the test period are never continued
        state["live_since"] = iso(today)
        for b in state["batches"].values():
            b["ended"] = True
    summary = []
    src_wb = None
    if action == "undo":
        got = restore_last_backup()
        if got:
            state, undone = got
            state["settings"]["test_mode"] = False
            summary.append(f"تم التراجع عن الإجراء الأخير ({LABELS.get(undone, undone)})، "
                           "وأُعيدت البيانات إلى ما كانت عليه قبله.")
        else:
            summary.append("لا يوجد إجراء سابق يمكن التراجع عنه.")
    else:
        if action != "preview":
            backup_state(action)                      # BEFORE anything changes, so undo returns to exactly this
        if not state.get("state_id"):
            import uuid
            state["state_id"] = uuid.uuid4().hex
        source = edits if edits.exists() else APP if APP.exists() else None
        if source:                                    # read your edits (values only)
            try:
                src_wb = load_workbook(source, data_only=False)
                if workbook_matches_state(src_wb, state):
                    read_edits(state, src_wb, today)
                else:
                    # the workbook was built from different data (e.g. you copied murasalat_state.json in):
                    # its contents are ignored and every sheet is rebuilt from the data file
                    summary.append("تم بناء جميع الصفحات من ملف البيانات murasalat_state.json، "
                                   "ولم تُقرأ الصفحات السابقة لأنها لا تخص هذه البيانات.")
            except Exception as e:
                print("Could not read edits from", source.name, "-", e)
    setup_aliases(state)
    apply_aliases(state)
    # always build a fresh workbook from the clean template
    wb = load_workbook(TEMPLATE, keep_vba=True) if have_template else Workbook()
    if state["settings"].get("signature"):
        M.SIGNATURE = state["settings"]["signature"]

    open_before = open_set(state)
    reopen_closed(state, today, summary)
    if action == "rebuild":
        act_rebuild(state, today, summary, src_wb)
    if action == "refresh":
        act_refresh(state, today, summary)
    if action == "move":
        act_move(state, today, summary)
    state.pop("track_override", None)               # retired: path edits are applied once, not forever
    state.pop("track_base", None)
    apply_path_edits(state, today)
    M.apply_forwards(state, today)
    for k, dpt in state.pop("_new_dept", {}).items():
        if k in state["rows"] and not state["rows"][k].get("dept"):
            state["rows"][k]["dept"] = dpt

    ol = None
    if use_outlook and M.USE_OUTLOOK:
        try:
            ol = M.Outlook()
            n_before = sum(len(b.get("replies", [])) for b in state["batches"].values())
            M.sync_outlook(state, ol, today)
            new = [x for b in state["batches"].values() for x in b.get("replies", [])][n_before:]
            if action == "replies" or new:
                summary.append(f"عدد الردود الجديدة: {len(new)}"
                               + (f"، من: {'، '.join(dict.fromkeys(x['from'] for x in new))}" if new else ""))
                fw = [x["from"] for x in new if x.get("fw") or not x.get("holder")]
                if fw:
                    summary.append("ردود من غير المرسل إليهم أو بإعادة توجيه: " + "، ".join(dict.fromkeys(fw)))
                added = sorted({a for x in new for a in x.get("added", [])})
                if added:
                    summary.append("أشخاص أُضيفوا إلى المراسلة: " + "، ".join(added))
        except Exception as e:
            ol = None
            if action in ("send", "replies"):
                summary.append(f"تعذر الاتصال ببرنامج Outlook ({e.__class__.__name__}).")

    install_force(state)
    cal = M.compute(state, today)
    archive_closed(state, cal, today)
    # safety net: an open transaction may only leave «المفتوحة» by being closed or deliberately removed
    vanished = sorted(open_before - open_set(state) - set(state["closed"]) - set(state.pop("_removed", [])))
    for txn in vanished:                              # bring it back rather than lose it
        rs = [R for R in state["rows"].values() if R["txn"] == txn]
        if rs:
            keep = max(rs, key=lambda R: R.get("recv") or "")
            keep["active"] = True
            if keep["status"] in (ST_CLOSED, ST_FWD):
                keep["status"] = M.ST_NEW
    if vanished:
        summary.append("تنبيه: أُعيدت معاملات كادت تختفي من المفتوحة دون إغلاق: " + "، ".join(vanished))
        cal = M.compute(state, today)
    mails = M.build_emails(state, cal, today)
    for m in mails:
        m["subject"] = m["subject_core"]
    test = False
    blocked = []
    if action == "send":
        missing_people = sorted({p for m in mails for p in m["people"] if not state["emails"].get(p)})
        if missing_people:
            # nothing is created until every recipient has an e-mail address, so nobody is forgotten
            blocked = missing_people
            for m in mails:
                m["draft"] = "لم تُنشأ: يوجد مستلمون دون بريد إلكتروني"
            summary.append("لم تُنشأ أي مسودة، لأن الموظفين التالية أسماؤهم دون بريد إلكتروني مسجّل في صفحة "
                           "البيانات الأساسية: " + "، ".join(missing_people)
                           + ". يُرجى إدخال بريدهم ثم الضغط على زر إنشاء المسودات مرة أخرى.")
        elif not ol:
            for m in mails:
                m["draft"] = "لم تُنشأ المسودة لتعذر الاتصال ببرنامج Outlook"
            summary.append("تعذر الاتصال ببرنامج Outlook، ولم تُنشأ أي مسودة.")
        else:
            M.create_drafts(state, ol, mails, today, "draft")
            drafts = sum(str(m.get("draft", "")).startswith("مسودة") for m in mails)
            summary.append(f"عدد التذكيرات المستحقة: {len(mails)}، وعدد المسودات المنشأة في Outlook: {drafts}، "
                           "ولم يتم إرسال أي رسالة.")
            for m in mails:
                summary.append(f"{'، '.join(m['people'])}: المعاملات {'، '.join(m['txns'])}، الحالة: {m['draft']}")
            for m in mails:                             # formal status words in the sheets
                if str(m.get("draft", "")).startswith("مسودة"):
                    m["draft"] = "أُنشئت المسودة في Outlook"
            state["force"] = {"day": iso(today), "keys": []}
            install_force(state)
            cal = M.compute(state, today)
    state["_blocked"] = blocked
    if action == "preview":
        summary.append(f"عدد التذكيرات المستحقة اليوم: {len(mails)}")
    if action == "track":
        summary.append("تم تحديث جميع الصفحات وفق التعديلات.")
    excluded = excluded_today(state, cal, today)

    state.setdefault("runlog", []).insert(0, {"at": dt.datetime.now().strftime("%Y/%m/%d %H:%M"),
                                              "action": LABELS.get(action, action),
                                              "lines": summary or ["لا يوجد تغيير."]})
    state["runlog"] = state["runlog"][:300]

    missing = write_all(state, cal, wb, today, mails, action, excluded)
    if missing and action in ("refresh", "track", "preview", "send"):
        state["runlog"][0]["lines"].append("موظفون دون بريد إلكتروني مسجّل: " + "، ".join(missing))
        write_all(state, cal, wb, today, mails, action, excluded)
    for k in [k for k in state if k.startswith("_")]:
        state.pop(k)
    for R in state["rows"].values():
        for k in [k for k in R if k.startswith("_")]:
            del R[k]
    if have_template and getattr(wb, "vba_archive", None) is None:
        raise RuntimeError("The workbook lost its button code; run: python setup_app.py")
    tmp = APP.with_name("~tmp_" + APP.stem + ".xlsm")
    wb.save(tmp)
    os.replace(tmp, APP)
    M.save_state(state)
    try:
        edits.unlink()
    except OSError:
        pass
    print(f"Done: {action}")
    for line in summary:
        print("  ", line)
    if reopen and os.name == "nt":
        os.startfile(str(APP))


def main():
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
    args = sys.argv[1:]
    action = args[0] if args else "track"
    today = None
    if "--today" in args:
        today = dt.date.fromisoformat(args[args.index("--today") + 1])
    lock = M.BASE / "murasalat.lock"
    try:
        if lock.exists() and time.time() - lock.stat().st_mtime < 300:
            print("Another action is already running; this one was skipped.")
            return
        lock.write_text(str(os.getpid()))
    except OSError:
        pass
    try:
        if ERR.exists():
            ERR.unlink()
        run(action, today, use_outlook="--no-outlook" not in args, reopen="--no-reopen" not in args)
    except Exception:
        traceback.print_exc()
        ERR.write_text(traceback.format_exc(), encoding="utf-8")
        if os.name == "nt" and APP.exists() and "--no-reopen" not in args:
            os.startfile(str(APP))
        sys.exit(1)
    finally:
        try:
            lock.unlink()
        except OSError:
            pass


if __name__ == "__main__":
    main()

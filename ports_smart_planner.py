# -*- coding: utf-8 -*-
"""
منسّق الإجازات: يقرأ نفس مصادر التقرير ويُنشئ ملف Excel بالمواعيد المقترحة.

التثبيت مرة واحدة: python -m pip install pandas openpyxl
التشغيل: python ports_smart_planner.py
عدّلي مساري المصادر فقط. الناتج: smart_planner.xlsx بجانب ملف الإجازات.

المستقبلية: حالات تبدأ بعد 18 نوفمبر 2026 (من 19 نوفمبر)، بكل النسب.
المجموعة: المدير + الموقع + المسار التنظيمي للمستويات الأول والثاني والثالث.
ينقل الإجازة كاملة دون تغيير مدتها. تثبت الإجازات الأخرى وتفحص كعوائق.
لا يكتب في ملفات المصادر. البحث محدود؛ الحالات غير المحلولة تظهر للمراجعة.
تعمل المعالجة محليًا دون روابط أو خدمات خارجية.
"""
from pathlib import Path
from datetime import date, datetime, timedelta
from collections import defaultdict, Counter
import hashlib
import json
import os
import re
import sys
import tempfile
import unicodedata
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

# ===================== عدّلي المسارات هنا =====================
MANAGERS_FILE = r'C:\Users\rammutairi\Downloads\Employee Master.xlsx'
LEAVES_FILE = r'C:\Users\rammutairi\Downloads\Ports Leaves.xlsx'
OUTPUT_XLSX = r''  # فارغ: smart_planner.xlsx بجانب ملف الإجازات
# ============================================================
MANAGERS_SHEET = 'مدراء المنافذ'
STAFF_SHEET = 'الموظفين تحت مدراء المنافذ'
LEAVES_SHEET = 'البيانات الخام'
PRESENTATION_DATE = '2026-11-18'
PERIOD_START = '2026-10-01'
PERIOD_END = '2026-12-31'
THRESHOLD = 50.0  # للعرض فقط؛ المنسّق يستهدف كل التقاطعات المستقبلية
END_DATE_INCLUSIVE = True
ALLOW_NO_APPROVAL_COLUMN = True
MAX_SEARCH_NODES_PER_GROUP = 30000
MAX_SEARCH_NODES_TOTAL = 300000

C = {
 'id': ['الرقم الوظيفي','Employee ID','(Employee) Person ID','Person ID','Employee Number','Emp ID'],
 'name': ['الاسم الرسمي بالعربي','اسم الموظف','Employee Name','(Employee) Full Name'],
 'supervisor': ['الرقم الوظيفي للمدير المباشر','Supervisor ID','Manager Employee ID'],
 'port': ['Location (Name)','(Employee) Location (Name)','المنفذ','اسم المنفذ','موقع العمل','Location'],
 'job': ['Job Title','المسمى الوظيفي'],
 'start':['Actual Start Date','تاريخ بداية الإجازة','تاريخ البداية','Start Date'],
 'end':['Actual End Date','تاريخ نهاية الإجازة','تاريخ النهاية','End Date'],
 'approval':['Approval Status','Approval Status (Name)','حالة الاعتماد','Approval Status (Label)'],
 'mgr':['المدير','اسم المدير','(Employee) Manager (Name)','Manager Name'],
 'lvl1':['المستوى الأول','Level 1','Level 1 (Name)','Level 1 Name'], 'lvl2':['المستوى الثاني','Level 2','Level 2 (Name)','Level 2 Name'],
 'lvl3':['المستوى الثالث','Level 3','Level 3 (Name)','Level 3 Name'], 'lvl4':['المستوى الرابع','Level 4','Level 4 (Name)','Level 4 Name'],
}
APPROVED = {'approved','approved request','معتمد','معتمده','موافق','موافق عليه','تم الاعتماد'}

def norm(v):
    s='' if pd.isna(v) else str(v).strip()
    s=unicodedata.normalize('NFKC',s)
    s=re.sub('[\u064b-\u065f\u0670\u0640\u200e\u200f\u200b]','',s)
    for a,b in [('أ','ا'),('إ','ا'),('آ','ا'),('ى','ي'),('ة','ه')]: s=s.replace(a,b)
    return re.sub(r'\s+',' ',s).strip().casefold()

def empid(x):
    if pd.isna(x):return ''
    x=str(x).strip()
    return re.sub(r'\.0$','',x) if re.fullmatch(r'\d+\.0',x) else x

def col(df,key,required=False):
    match={norm(c):c for c in df.columns}
    for a in C[key]:
        if norm(a) in match: return match[norm(a)]
    if required:
        raise ValueError(f'عمود {key} غير موجود. المطلوب أحد: {C[key]}\nالأعمدة المتوفرة: {list(df.columns)}')
    return None

def cell(r,k):
    c=k
    if c is None:return ''
    v=r[c]
    return '' if pd.isna(v) else str(v).strip()

def dt(v):
    if pd.isna(v) or str(v).strip()=='':return None
    if isinstance(v,(pd.Timestamp,datetime,date)):return pd.Timestamp(v).date()
    # Excel date serial; handles only genuine serials, not YYYYMMDD.
    if isinstance(v,(int,float)) and 25000 < float(v) < 90000:
        return (datetime(1899,12,30)+timedelta(days=int(v))).date()
    v=str(v).strip()
    if re.fullmatch(r'\d{5}(?:\.0)?',v) and 25000<int(float(v))<90000:
        return (datetime(1899,12,30)+timedelta(days=int(float(v)))).date()
    t=pd.to_datetime(v,errors='coerce',dayfirst=not bool(re.fullmatch(r'\d{4}-\d{1,2}-\d{1,2}',v)))
    return None if pd.isna(t) else t.date()

def excel_read(path,sheet):
    p=Path(path)
    if not p.is_file():raise FileNotFoundError(f'الملف غير موجود: {path}')
    df=pd.read_excel(p,sheet_name=sheet,dtype=object,engine='openpyxl')
    df.columns=[str(x).strip() for x in df.columns]
    return df.dropna(how='all')

def label_state(st,en,asof):
    return 'منتهية' if en < asof else ('جارية' if st<=asof else 'قادمة')

def calculate(asof_text=None):
    asof=date.fromisoformat(asof_text or PRESENTATION_DATE)
    print(f'تاريخ إعداد التقرير والتصنيف: {asof:%Y-%m-%d} (تاريخ التصنيف المعتمد في هذا التشغيل)', flush=True)
    pstart=date.fromisoformat(PERIOD_START);pend=date.fromisoformat(PERIOD_END)
    if pstart > pend:raise ValueError('تاريخ بداية الفترة بعد نهايتها.')
    print('1/5 قراءة ملفات المصادر...',flush=True)
    managers=excel_read(MANAGERS_FILE,MANAGERS_SHEET)
    staff=excel_read(MANAGERS_FILE,STAFF_SHEET)
    leaves=excel_read(LEAVES_FILE,LEAVES_SHEET)
    mid=col(managers,'id',True); mn=col(managers,'name',True)
    mj=col(managers,'job');ml=[col(managers,f'lvl{i}') for i in range(1,5)]
    si=col(staff,'id',True);ss=col(staff,'supervisor',True);sp=col(staff,'port',True)
    sn=col(staff,'name'); sl=[col(staff,f'lvl{i}') for i in range(1,5)]
    li=col(leaves,'id');ln=col(leaves,'name');lp=col(leaves,'port')
    ls=col(leaves,'start',True);le=col(leaves,'end',True)
    la=col(leaves,'approval');lm=col(leaves,'mgr')
    if not li and not ln:raise ValueError('ملف الإجازات لا يحتوي رقم وظيفي ولا اسم موظف.')
    if la is None and not ALLOW_NO_APPROVAL_COLUMN:
        raise ValueError('لا يوجد عمود اعتماد، تأكدي من مصدر البيانات.')
    print(f'  مدراء: {len(managers)} | موظفون: {len(staff)} | إجازات مصدرية: {len(leaves)}',flush=True)

    managers_by_id={}; managers_by_name={}
    for _,r in managers.iterrows():
        key=empid(cell(r,mid));name=cell(r,mn)
        if not key or not name:continue
        if key in managers_by_id and managers_by_id[key]['name']!=name:
            raise ValueError(f'تعارض رقم مدير في قائمة المدراء: {key}')
        entry={'id':key,'name':name,'job':cell(r,mj),'levels':[cell(r,k) for k in ml]}
        managers_by_id[key]=entry
        managers_by_name.setdefault(norm(name),set()).add(key)

    # Denominator: UNIQUE employees directly reporting to chosen manager AT a location.
    group_staff=defaultdict(set); staff_by_id={}; staff_by_name=defaultdict(set)
    display_port={}; display_levels={}; problems=Counter();ambiguous=[]
    employee_display_names={}
    # Separate team = supervisor + location + entire organizational path L1/L2/L3.
    # A manager can have multiple independent teams at the same port.
    # Never combine denominators across organizational units.
    for _,r in staff.iterrows():
        eid=empid(cell(r,si)); mgr_id=empid(cell(r,ss)); location=cell(r,sp)
        if mgr_id not in managers_by_id:continue
        if not eid or not location:
            problems['موظفون بلا رقم وظيفي أو موقع']+=1;continue
        levels=tuple((cell(r,k).strip() if k else '') for k in sl[:3])
        group=(mgr_id,norm(location),*(norm(v) for v in levels))
        if eid in staff_by_id and staff_by_id[eid]!=group:
            problems['موظفون بتبعيات متعددة']+=1
            ambiguous.append(('تبعيات متعددة',eid,group[0],location))
            continue
        staff_by_id[eid]=group
        group_staff[group].add(eid)
        display_port[group]=location
        display_levels[group]=levels
        if sn and cell(r,sn):
            staff_by_name[norm(cell(r,sn))].add(eid)
            employee_display_names[eid]=cell(r,sn)
    if not group_staff:raise ValueError('لم تتم مطابقة موظفين مع المدراء حسب رقم المدير المباشر والموقع.')
    print(f'2/5 مطابقة الإجازات مع {len(staff_by_id)} موظفًا مستهدفًا...',flush=True)
    daily=defaultdict(set); employee_leave_dates=defaultdict(list); record_signatures=set()
    accepted_records=[]
    matched=0;out_of_port=0
    for rowno,r in leaves.iterrows():
        st=dt(r[ls]);en=dt(r[le])
        original_start,original_end=st,en
        if not st or not en or st>en:
            problems['إجازات بتواريخ ناقصة أو غير صالحة']+=1;continue
        if en<pstart or st>pend:continue
        if la:
            approval=norm(cell(r,la))
            if approval and approval not in {norm(a) for a in APPROVED}:
                problems['إجازات غير معتمدة أو غير معروفة الحالة']+=1;continue
        eid=empid(cell(r,li)) if li else ''
        if not eid:
            candidates=staff_by_name.get(norm(cell(r,ln)),set()) if ln else set()
            if len(candidates)==1:eid=next(iter(candidates))
            else:
                problems['إجازات لا يمكن مطابقتها بموظف فريد']+=1;continue
        group=staff_by_id.get(eid)
        if not group:
            problems['إجازات لموظفين خارج قائمة القوى العاملة المستهدفة']+=1;continue
        location=cell(r,lp)
        if location and norm(location)!=group[1]:
            # Preserve denominator integrity: don't move the employee to a new
            # location silently or count a leave under the old location.
            out_of_port+=1;problems['إجازات بموقع مختلف بين المصدرين']+=1
            if len(ambiguous)<200:ambiguous.append(('اختلاف المنفذ',eid,display_port[group],location))
            continue
        if not END_DATE_INCLUSIVE:en=en-timedelta(days=1)
        st=max(st,pstart);en=min(en,pend)
        if st>en:continue
        signature=(eid,group,st,en)
        # Keep original source intervals before clipped-signature deduplication.
        # Reporting counts below remain identical to the supplied report.
        accepted_records.append({
            'employee_id':eid, 'group':group,
            'original_start':original_start, 'original_end':original_end,
            'start':original_start,
            'end':original_end if END_DATE_INCLUSIVE else original_end-timedelta(days=1),
            'clipped_start':st, 'clipped_end':en,
            'source_row':int(rowno)+2, 'source_name':cell(r,ln),
        })
        if signature in record_signatures:
            problems['صفوف إجازات مكررة تمامًا']+=1;continue
        record_signatures.add(signature);matched+=1
        employee_leave_dates[eid].append((st,en,original_start,original_end))
        if eid not in employee_display_names and ln and cell(r,ln):
            employee_display_names[eid]=cell(r,ln)
        for d in (st+timedelta(days=i) for i in range((en-st).days+1)):
            daily[(group,d)].add(eid)
    print(f'  سجلات إجازات مطابقة: {matched} | اختلاف منفذ: {out_of_port}',flush=True)

    print('3/5 حساب مجموعات الغياب المتطابقة عبر أيام متصلة...',flush=True)
    # No pseudo 9-day intersection: run splits whenever exact set changes.
    dates_by_group=defaultdict(list)
    for (group,d),emps in daily.items():
        dates_by_group[group].append((d,frozenset(emps)))
    cases=[];peak_group={};overlap_people=defaultdict(set);exact_people=defaultdict(set)
    for g,events in dates_by_group.items():
        events.sort()
        eligible=[(d,s) for d,s in events if len(s)>=2]
        if eligible:
            peak_day,peak_set=max(eligible,key=lambda a:(len(a[1]),-a[0].toordinal()))
            peak_group[g]=(peak_day,len(peak_set))
            for d,s in eligible:overlap_people[g].update(s)
        if not eligible:continue
        runs=[];run=[eligible[0]]
        for current in eligible[1:]:
            if current[0]==run[-1][0]+timedelta(days=1) and current[1]==run[-1][1]:
                run.append(current)
            else:runs.append(run);run=[current]
        runs.append(run)
        for run in runs:
            st=run[0][0];en=run[-1][0];cohort=run[0][1];size=len(cohort)
            workforce=len(group_staff[g]);ratio=100*size/workforce
            status=label_state(st,en,asof)
            days_to_start=(st-asof).days if st>asof else 0 if en>=asof else None
            days_left=(en-asof).days+1 if en>=asof else 0
            cases.append(dict(group=g,start=st,end=en,days=len(run),people=size,
                         percent=ratio,workforce=workforce,remaining=workforce-size,
                         status=status,eta=days_to_start,remaining_days=days_left,
                         employees=','.join(sorted(cohort)),employee_ids_list=sorted(cohort),flag=ratio>=THRESHOLD))

    all_cases = sorted(cases, key=lambda c: (
        {'جارية': 0, 'قادمة': 1, 'منتهية': 2}[c['status']],
        -c['percent'], -c['people'], c['start'], c['group']))
    priority = sorted((c for c in cases if c['flag']), key=lambda c: (
        {'جارية': 0, 'قادمة': 1, 'منتهية': 2}[c['status']],
        -c['percent'], -c['people'], c['start']))
    return {
        'asof': asof, 'pstart': pstart, 'pend': pend,
        'managers_by_id': managers_by_id, 'group_staff': group_staff,
        'staff_by_id': staff_by_id, 'display_port': display_port,
        'display_levels': display_levels, 'employee_display_names': employee_display_names,
        'employee_leave_dates': employee_leave_dates, 'record_signatures': record_signatures,
        'cases': cases, 'all_cases': all_cases, 'priority': priority,
        'matched': matched, 'out_of_port': out_of_port,
        'problems': problems, 'ambiguous': ambiguous, 'accepted_records': accepted_records,
        'source_rows': {'managers': len(managers), 'staff': len(staff), 'leaves': len(leaves)},
    }


"""Bounded, deterministic vacation scheduling; no workbook or source-file writes.

Only complete source leave intervals are moved. Every changed interval is
collision-free in the final calendar. Unresolved original collisions remain
explicit, including when a search budget prevents finding a better plan.
"""
from collections import defaultdict
from datetime import date, timedelta
import hashlib
import json
import math
from numbers import Integral, Real


def _date(value):
    return value if isinstance(value, date) else date.fromisoformat(str(value))


def _iso(ordinal):
    return date.fromordinal(ordinal).isoformat()


def _mask(start, end, lower, upper):
    start, end = max(start, lower), min(end, upper)
    return 0 if start > end else ((1 << (end - start + 1)) - 1) << (start - lower)


def _bits(mask):
    while mask:
        bit = mask & -mask
        yield bit.bit_length() - 1
        mask ^= bit


def _intersects(a_start, a_end, b_start, b_end):
    return a_start <= b_end and b_start <= a_end


def _daily(records, positions, lower, upper):
    days = defaultdict(set)
    for record in records:
        start, end = positions[record['id']]
        for day in range(max(lower, start), min(upper, end) + 1):
            days[(record['group_id'], day)].add(record['employee_id'])
    return days


def _pair_days(days):
    return sum(len(members) * (len(members) - 1) // 2 for members in days.values())


def _objective(records, positions, lower, upper):
    changed = [r for r in records if positions[r['id']][0] != r['_start']]
    return (_pair_days(_daily(records, positions, lower, upper)), len(changed),
            sum(abs(positions[r['id']][0] - r['_start']) for r in changed))


def _solve_group(records, lower, upper, node_limit, total_remaining):
    """Safe greedy incumbents followed by joint branch-and-bound search.

    Keeping an original interval may retain an original collision. Any moved
    interval must avoid fixed and assigned records, and later original choices
    must avoid moved intervals. Thus every completed assignment has no new
    collisions, even when not all original collisions could be eliminated.
    """
    original = {r['id']: (r['_start'], r['_end']) for r in records}
    best_positions = dict(original)
    best_score = _objective(records, original, lower, upper)
    baseline_score = best_score
    movable = [r for r in records if r['movable']]
    fixed = [r for r in records if not r['movable']]
    width = max(0, upper - lower + 1)
    cap = max(0, min(int(node_limit), int(total_remaining)))
    nodes = 0
    interrupted = False
    complete = False

    def consume():
        nonlocal nodes, interrupted
        if nodes >= cap:
            interrupted = True
            return False
        nodes += 1
        return True

    fixed_mask = 0
    for r in fixed:
        fixed_mask |= r['_mask']
    domains = {}
    for r in movable:
        length = r['_end'] - r['_start'] + 1
        candidates = []
        for start in range(lower, upper - length + 2):
            mask = _mask(start, start + length - 1, lower, upper)
            if start != r['_start'] and not mask & fixed_mask:
                candidates.append((start, start + length - 1, mask))
        candidates.sort(key=lambda v: (abs(v[0] - r['_start']), v[0] < r['_start'], v[0]))
        domains[r['id']] = candidates

    def keep_candidate(positions):
        nonlocal best_positions, best_score
        score = _objective(records, positions, lower, upper)
        if score < best_score:
            best_score, best_positions = score, dict(positions)

    # Each warm start preserves a valid complete calendar after every move.
    # They provide useful partial results even when joint search is interrupted.
    orders = [
        sorted(movable, key=lambda r: (len(domains[r['id']]), -r['_duration'], r['id'])),
        sorted(movable, key=lambda r: (-r['_duration'], -r['_start'], r['id'])),
        sorted(movable, key=lambda r: (r['_start'], r['id'])),
    ]
    seen_orders = set()
    for order in orders:
        order_ids = tuple(r['id'] for r in order)
        if order_ids in seen_orders:
            continue
        seen_orders.add(order_ids)
        positions = dict(original)
        masks = {r['id']: r['_mask'] for r in records}
        for r in order:
            if not consume():
                break
            rid = r['id']
            if not any(other['id'] != rid and other['employee_id'] != r['employee_id']
                       and masks[rid] & masks[other['id']] for other in records):
                continue
            obstacles = 0
            for other in records:
                if other['id'] != rid:
                    obstacles |= masks[other['id']]
            for start, end, mask in domains[rid]:
                if not consume():
                    break
                if not mask & obstacles:
                    positions[rid] = (start, end)
                    masks[rid] = mask
                    break
            if interrupted:
                break
        keep_candidate(positions)
        if interrupted:
            break

    # The fixed-day counts use unique employees, including any duplicate-source
    # intervals that were conservatively frozen by the input protection rules.
    counts = [0] * width
    for (_, day), employees in _daily(fixed, original, lower, upper).items():
        counts[day - lower] = len(employees)
    fixed_pair_days = sum(n * (n - 1) // 2 for n in counts)
    ordered = orders[0]
    partial = dict(original)
    memo = {}

    def search(depth, occupied, moved_mask, pair_days, changed_count, displacement):
        nonlocal best_positions, best_score
        if not consume():
            return
        lower_bound = (pair_days, changed_count, displacement)
        if lower_bound >= best_score:
            return
        if depth == len(ordered):
            best_score = lower_bound
            best_positions = dict(partial)
            return
        state_key = (depth, occupied, moved_mask, tuple(counts))
        cost = (changed_count, displacement)
        previous = memo.get(state_key)
        if previous is not None and previous <= cost:
            return
        memo[state_key] = cost
        r = ordered[depth]
        rid = r['id']
        choices = [(r['_start'], r['_end'], r['_mask'], False)]
        choices += [(s, e, mask, True) for s, e, mask in domains[rid]]
        for start, end, mask, changed in choices:
            if changed:
                if mask & occupied:
                    continue
            elif mask & moved_mask:
                continue
            # A movable record cannot originally overlap another record for the
            # same employee: both records would have been frozen. Changed
            # intervals avoid all records. Adding a movable record therefore
            # adds one distinct employee on each of these days.
            indices = tuple(_bits(mask))
            delta = sum(counts[index] for index in indices)
            for index in indices:
                counts[index] += 1
            partial[rid] = (start, end)
            search(depth + 1, occupied | mask, moved_mask | (mask if changed else 0),
                   pair_days + delta, changed_count + int(changed),
                   displacement + abs(start - r['_start']))
            partial[rid] = original[rid]
            for index in indices:
                counts[index] -= 1
            if interrupted:
                return

    # Avoid recursion depth failures on very large operational teams. Such
    # teams retain the valid bounded greedy result and disclose limited search.
    joint_eligible = len(ordered) <= 80
    if not movable:
        complete = True
        method = 'fixed_only'
    elif baseline_score[0] == fixed_pair_days:
        complete = True
        method = 'already_clear'
    elif not interrupted and joint_eligible:
        search(0, fixed_mask, 0, fixed_pair_days, 0, 0)
        complete = not interrupted
        method = 'joint_search'
    else:
        method = 'bounded_greedy'

    if not movable:
        status = 'no_movable_leaves'
        status_ar = 'لا توجد إجازات مستهدفة قابلة للنقل في هذه المجموعة'
    elif complete:
        status = 'search_complete'
        status_ar = 'اكتمل البحث ضمن البدائل والقيود المحددة'
    elif interrupted:
        status = 'total_node_limit' if total_remaining <= node_limit else 'node_limit'
        status_ar = 'توقف البحث عند الحد المحدد؛ احتُفظ بأفضل جدول آمن عُثر عليه'
    else:
        status = 'large_group_limited_search'
        status_ar = 'مجموعة كبيرة؛ استُخدم بحث محدود واحتُفظ بالتحسين الآمن المتاح'
    return best_positions, {
        'group_id': records[0]['group_id'] if records else '',
        'method': method, 'status': status, 'status_ar': status_ar,
        'search_complete': complete, 'nodes': nodes,
        'node_limit': cap, 'movable_leaves': len(movable),
        'pair_days_before': baseline_score[0], 'pair_days_after': best_score[0],
        'changed_leaves': best_score[1], 'total_shift_days': best_score[2],
        'no_fixed_free_alternative_ids': [r['id'] for r in movable if not domains[r['id']]],
    }


def _recalculate_cases(records, positions, groups, lower, upper, prefix):
    by_group = defaultdict(list)
    for (group_id, day), employees in _daily(records, positions, lower, upper).items():
        if len(employees) >= 2:
            by_group[group_id].append((day, frozenset(employees)))
    result = []
    for group_id in sorted(by_group):
        events = sorted(by_group[group_id])
        runs = []
        run = [events[0]]
        for event in events[1:]:
            if event[0] == run[-1][0] + 1 and event[1] == run[-1][1]:
                run.append(event)
            else:
                runs.append(run)
                run = [event]
        runs.append(run)
        group = groups[group_id]
        workforce = int(group['workforce'])
        for run in runs:
            start, end = run[0][0], run[-1][0]
            employees = sorted(run[0][1])
            key = json.dumps([group_id, start, end, employees], ensure_ascii=False, separators=(',', ':'))
            row = {
                'id': prefix + hashlib.sha256(key.encode('utf-8')).hexdigest()[:16],
                'group_id': group_id, 'start': _iso(start), 'end': _iso(end),
                'days': end - start + 1, 'people': len(employees),
                'workforce': workforce, 'remaining': workforce - len(employees),
                'percent': 100 * len(employees) / workforce,
                'percent_display': round(100 * len(employees) / workforce, 1),
                'employee_ids': employees, 'status': 'قادمة',
                'leave_ids': sorted(r['id'] for r in records
                    if r['group_id'] == group_id and r['employee_id'] in employees
                    and _intersects(*positions[r['id']], start, end)),
            }
            for field in ('manager_id', 'manager', 'title', 'port', 'level1', 'level2', 'level3'):
                row[field] = group.get(field, '')
            result.append(row)
    return sorted(result, key=lambda c: (c['start'], c['group_id'], c['end'], c['id']))


def _plain(value):
    """Return only JSON-native values, including when input numbers are scalars."""
    if value is None or isinstance(value, (str, bool)):
        return value
    if isinstance(value, Integral):
        return int(value)
    if isinstance(value, Real):
        result = float(value)
        if not math.isfinite(result):
            raise ValueError('قيمة عددية غير صالحة في بيانات المنسق.')
        return result
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, dict):
        return {str(k): _plain(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [_plain(v) for v in value]
    if isinstance(value, (set, frozenset)):
        return [_plain(v) for v in sorted(value)]
    raise TypeError('نوع غير مدعوم في بيانات المنسق: ' + type(value).__name__)


def build_plan(groups, leaves, cases, presentation_date='2026-11-18',
               period_end='2026-12-31', end_date_inclusive=True,
               max_nodes=30000, max_total_nodes=300000):
    """Return an auditable proposal without modifying any input objects.

    ``start``/``end`` in each input leave are full, inclusive effective absence
    dates. ``original_start``/``original_end`` are the untouched source dates.
    Final source end dates preserve ``end_date_inclusive`` exactly.
    """
    presentation = _date(presentation_date)
    lower = presentation.toordinal() + 1
    upper = _date(period_end).toordinal()
    if lower > upper:
        raise ValueError('لا توجد فترة جدولة بعد تاريخ العرض وقبل نهاية الفترة.')
    group_map = {}
    for original_group in groups:
        group = dict(original_group)
        group['id'] = str(group['id'])
        group['workforce'] = int(group['workforce'])
        if group['id'] in group_map:
            raise ValueError('معرف مجموعة مكرر: ' + group['id'])
        if group['workforce'] <= 0:
            raise ValueError('يجب أن يكون عدد موظفي المجموعة أكبر من صفر.')
        group_map[group['id']] = group
    records, ids = [], set()
    for raw in leaves:
        record = dict(raw)
        rid = str(record['id'])
        if rid in ids:
            raise ValueError('معرف إجازة مكرر: ' + rid)
        ids.add(rid)
        record['id'] = rid
        record['group_id'] = str(record['group_id'])
        record['employee_id'] = str(record['employee_id'])
        if record['group_id'] not in group_map:
            raise ValueError('مجموعة الإجازة غير موجودة: ' + rid)
        start, end = _date(record['start']), _date(record['end'])
        original_start, original_end = _date(record['original_start']), _date(record['original_end'])
        expected_end = original_end if end_date_inclusive else original_end - timedelta(days=1)
        if start != original_start or end != expected_end or start > end:
            raise ValueError('تواريخ الإجازة الفعلية لا تطابق تواريخ المصدر الكاملة: ' + rid)
        record['_start'], record['_end'] = start.toordinal(), end.toordinal()
        record['_duration'] = (end - start).days + 1
        record['_mask'] = _mask(record['_start'], record['_end'], lower, upper)
        record['target'] = False
        record['original_case_ids'] = []
        record['source_rows'] = [int(x) for x in record.get('source_rows', [])]
        record['ambiguous'] = bool(record.get('ambiguous', False))
        records.append(record)
    records.sort(key=lambda r: (r['group_id'], r['_start'], r['_end'], r['employee_id'], r['id']))
    record_map = {r['id']: r for r in records}
    group_records = defaultdict(list)
    employee_records = defaultdict(list)
    for record in records:
        group_records[record['group_id']].append(record)
        employee_records[(record['group_id'], record['employee_id'])].append(record)

    source_case_ids = set()
    future_cases = []
    for source in cases:
        start, end = _date(source['start']), _date(source['end'])
        if start <= presentation:
            continue
        case = dict(source)
        case_id = str(case['id'])
        if case_id in source_case_ids:
            raise ValueError('معرف حالة مستقبلية مكرر: ' + case_id)
        source_case_ids.add(case_id)
        case['id'], case['group_id'] = case_id, str(case['group_id'])
        case['start'], case['end'] = start.isoformat(), end.isoformat()
        case['employee_ids'] = sorted({str(eid) for eid in case['employee_ids']})
        linked = [r for r in group_records.get(case['group_id'], [])
                  if r['employee_id'] in case['employee_ids']
                  and _intersects(r['_start'], r['_end'], start.toordinal(), end.toordinal())]
        case['leave_ids'] = sorted(r['id'] for r in linked)
        case['missing_employee_ids'] = sorted(set(case['employee_ids']) - {r['employee_id'] for r in linked})
        for r in linked:
            r['target'] = True
            r['original_case_ids'].append(case_id)
        future_cases.append(case)

    # Overlapping source records for one employee are a source ambiguity, not
    # a report case. Freeze both records instead of moving part of that leave.
    same_employee_overlap = set()
    for items in employee_records.values():
        active = []
        for r in sorted(items, key=lambda row: (row['_start'], row['_end'], row['id'])):
            active = [old for old in active if old['_end'] >= r['_start']]
            for old in active:
                same_employee_overlap.add(old['id'])
                same_employee_overlap.add(r['id'])
            active.append(r)
    for record in records:
        reason = ''
        if not record['target']:
            reason = 'إجازة ثابتة خارج حالات التقاطع المستقبلية المستهدفة.'
        elif record['ambiguous']:
            reason = record.get('ambiguity_reason') or 'سجلات المصدر لهذه الإجازة تحتاج مراجعة قبل نقلها.'
        elif record['id'] in same_employee_overlap:
            reason = 'توجد سجلات إجازة متداخلة للموظف نفسه؛ ثُبتت للمراجعة.'
        elif record['_start'] < lower:
            reason = 'بدأت الإجازة بحلول تاريخ العرض؛ موعدها ثابت.'
        elif record['_end'] > upper:
            reason = 'تمتد الإجازة خارج نهاية فترة الجدولة؛ حُفظت مدتها وموعدها.'
        record['movable'] = bool(record['target'] and not reason)
        record['locked_reason'] = reason
        record['original_case_ids'] = sorted(set(record['original_case_ids']))

    original_positions = {r['id']: (r['_start'], r['_end']) for r in records}
    final_positions = dict(original_positions)
    solver_status = []
    total_cap = max(0, int(max_total_nodes))
    consumed = 0
    for group_id in sorted(group_records):
        items = group_records[group_id]
        if not any(r['target'] for r in items):
            continue
        positions, info = _solve_group(items, lower, upper, max(0, int(max_nodes)), total_cap - consumed)
        final_positions.update(positions)
        solver_status.append(info)
        consumed += info['nodes']
    solver_by_group = {s['group_id']: s for s in solver_status}

    baseline_daily = _daily(records, original_positions, lower, upper)
    final_daily = _daily(records, final_positions, lower, upper)
    pair_before, pair_after = _pair_days(baseline_daily), _pair_days(final_daily)
    # Validate the entire proposed calendar, independently of the search state.
    for record in records:
        rid = record['id']
        start, end = final_positions[rid]
        if end - start + 1 != record['_duration']:
            raise AssertionError('تغيرت مدة الإجازة: ' + rid)
        changed = (start, end) != original_positions[rid]
        if not record['movable'] and changed:
            raise AssertionError('تغير موعد إجازة ثابتة: ' + rid)
        if changed:
            if not lower <= start <= end <= upper:
                raise AssertionError('الموعد المقترح خارج فترة الجدولة: ' + rid)
            for other in group_records[record['group_id']]:
                if other['id'] != rid and _intersects(start, end, *final_positions[other['id']]):
                    raise AssertionError('الموعد المقترح يتقاطع مع إجازة أخرى: ' + rid)
    for key, employees in final_daily.items():
        if len(employees) >= 2 and not employees.issubset(baseline_daily.get(key, set())):
            raise AssertionError('ظهر تقاطع جديد بين موظفين بعد التصحيح.')
    if pair_after > pair_before:
        raise AssertionError('زاد إجمالي أيام تقاطع أزواج الموظفين بعد التصحيح.')

    exported = []
    for record in records:
        row = {key: value for key, value in record.items() if not key.startswith('_')}
        rid = record['id']
        start, end = final_positions[rid]
        row['original_start'], row['original_end'] = str(record['original_start']), str(record['original_end'])
        row['final_effective_start'], row['final_effective_end'] = _iso(start), _iso(end)
        row['final_start'] = _iso(start)
        row['final_end'] = _iso(end if end_date_inclusive else end + 1)
        row['changed'] = (start, end) != original_positions[rid]
        row['shift_days'] = start - record['_start']
        row['duration_days'] = record['_duration']
        row['conflict_leave_ids'] = sorted(other['id'] for other in group_records[record['group_id']]
            if other['id'] != rid and other['employee_id'] != record['employee_id']
            and max(start, final_positions[other['id']][0], lower)
                <= min(end, final_positions[other['id']][1], upper))
        row['remaining_conflict'] = bool(row['conflict_leave_ids'])
        if row['changed']:
            row['result_code'] = 'proposed'
            row['result'] = 'تم اقتراح موعد بديل دون تقاطع'
            row['reason'] = 'حُفظت مدة الإجازة، وفُحص الموعد مقابل جميع إجازات المجموعة.'
        elif record['target'] and not row['remaining_conflict']:
            row['result_code'] = 'resolved_unchanged'
            row['result'] = 'زالت المشكلة مع تثبيت الموعد'
            row['reason'] = 'أصبح الموعد دون تقاطع في فترة الجدولة بعد تعديل الإجازات الأخرى.'
        elif record['target']:
            row['result_code'] = 'unresolved'
            row['result'] = 'بقي تقاطع يحتاج معالجة'
            if record['locked_reason']:
                row['reason'] = record['locked_reason']
            else:
                info = solver_by_group.get(record['group_id'], {})
                if rid in info.get('no_fixed_free_alternative_ids', []):
                    row['reason'] = 'لا يوجد موعد بديل مختلف عن الموعد الأصلي يحفظ مدة الإجازة ويتجنب الإجازات الثابتة ضمن نطاق الجدولة.'
                else:
                    row['reason'] = 'لم يعثر البحث على موعد بديل آمن ضمن الفترة والقيود المحددة.'
        else:
            row['result_code'] = 'fixed'
            row['result'] = 'موعد ثابت'
            row['reason'] = record['locked_reason']
        exported.append(row)
    exported_by_id = {r['id']: r for r in exported}
    for case in future_cases:
        case['resolved'] = bool(case['leave_ids'] and not case['missing_employee_ids']
                                and all(not exported_by_id[rid]['remaining_conflict'] for rid in case['leave_ids']))
        case['changed_leave_ids'] = [rid for rid in case['leave_ids'] if exported_by_id[rid]['changed']]
        case['remaining_leave_ids'] = [rid for rid in case['leave_ids'] if exported_by_id[rid]['remaining_conflict']]
        case['result'] = 'تم فصل الإجازات المرتبطة بالحالة' if case['resolved'] else 'توجد إجازات مرتبطة بالحالة تحتاج معالجة'
        for field in ('manager_id', 'manager', 'title', 'port', 'level1', 'level2', 'level3', 'workforce'):
            case[field] = group_map.get(case['group_id'], {}).get(field, '')
    future_cases.sort(key=lambda c: (c['start'], c['group_id'], c['id']))
    before_window = _recalculate_cases(records, original_positions, group_map, lower, upper, 'B')
    after_window = _recalculate_cases(records, final_positions, group_map, lower, upper, 'A')
    target_rows = [r for r in exported if r['target']]
    stats = {
        'all_leaves': len(exported), 'target_leaves': len(target_rows),
        'movable_target_leaves': sum(r['movable'] for r in target_rows),
        'locked_target_leaves': sum(not r['movable'] for r in target_rows),
        'changed_leaves': sum(r['changed'] for r in exported),
        'resolved_target_leaves': sum(not r['remaining_conflict'] for r in target_rows),
        'unresolved_target_leaves': sum(r['remaining_conflict'] for r in target_rows),
        'future_case_count': len(future_cases),
        'resolved_future_cases': sum(c['resolved'] for c in future_cases),
        'unresolved_future_cases': sum(not c['resolved'] for c in future_cases),
        'window_cases_before': len(before_window), 'window_cases_after': len(after_window),
        'pair_days_before': pair_before, 'pair_days_after': pair_after,
        'pair_days_reduced': pair_before - pair_after,
        'nodes': consumed, 'max_total_nodes': total_cap,
        'limited_search_groups': sum(not s['search_complete'] for s in solver_status),
        'changed_leaves_conflict_free': True, 'fixed_leaves_unchanged': True,
        'durations_preserved': True, 'new_overlaps': 0,
    }
    return _plain({
        'meta': {'presentation_date': presentation.isoformat(), 'planning_start': _iso(lower),
                 'planning_end': _iso(upper), 'end_date_inclusive': bool(end_date_inclusive),
                 'scope': 'جميع الحالات الأصلية التي تبدأ بعد تاريخ العرض، بغض النظر عن النسبة',
                 'comparison_scope': 'تُقارن التقاطعات قبل وبعد داخل فترة الجدولة كاملة، بما فيها امتداد الحالات الجارية عند تاريخ العرض.'},
        'groups': [group_map[key] for key in sorted(group_map)],
        'leaves': exported, 'cases_before': future_cases,
        'cases_before_window': before_window, 'cases_after': after_window,
        'stats': stats, 'group_solver_status': solver_status,
    })

def stable_id(prefix, *parts):
    raw=json.dumps(parts,ensure_ascii=False,separators=(',',':'),default=str)
    return prefix+hashlib.sha256(raw.encode('utf-8')).hexdigest()[:24]


def planner_inputs(calculated):
    """Report grouping unchanged; retain all distinct original bookings as obstacles."""
    groups=[];group_ids={};leaves_by_key={};clipped_originals=defaultdict(set)
    uncertain_employees={str(row[1]) for row in calculated['ambiguous'] if row[0]=='تبعيات متعددة'}
    for group,employees in sorted(calculated['group_staff'].items()):
        gid=stable_id('G',*group);group_ids[group]=gid
        manager=calculated['managers_by_id'][group[0]]
        levels=calculated['display_levels'][group]
        groups.append(dict(id=gid,manager_id=group[0],manager=manager['name'],
                           title=manager['job'],port=calculated['display_port'][group],
                           level1=levels[0],level2=levels[1],level3=levels[2],
                           workforce=len(employees)))
    for record in calculated['accepted_records']:
        key=(record['group'],record['employee_id'],record['original_start'],record['original_end'])
        clipped=(record['group'],record['employee_id'],record['clipped_start'],record['clipped_end'])
        clipped_originals[clipped].add(key)
        if key not in leaves_by_key:
            gid=group_ids[record['group']];eid=record['employee_id']
            leaves_by_key[key]=dict(
                id=stable_id('L',gid,eid,record['original_start'],record['original_end']),
                group_id=gid,employee_id=eid,
                name=calculated['employee_display_names'].get(eid) or record['source_name'],
                original_start=record['original_start'].isoformat(),
                original_end=record['original_end'].isoformat(),
                start=record['start'].isoformat(),end=record['end'].isoformat(),
                source_rows=[],ambiguous=eid in uncertain_employees,
                ambiguity_reason='تعدد تبعية الموظف في مصدر القوى العاملة' if eid in uncertain_employees else '')
        leaves_by_key[key]['source_rows'].append(record['source_row'])
    for keys in clipped_originals.values():
        if len(keys)>1:
            for key in keys:
                row=leaves_by_key[key];row['ambiguous']=True
                row['ambiguity_reason']='نطاقات أصلية مختلفة تعطي التواريخ نفسها بعد قصها إلى الربع'
    leaves=sorted(leaves_by_key.values(),key=lambda x:(x['group_id'],x['start'],x['employee_id'],x['end'],x['id']))
    for row in leaves:row['source_rows']=sorted(set(row['source_rows']))
    cases=[]
    for case in calculated['all_cases']:
        gid=group_ids[case['group']];employees=case.get('employee_ids_list',case['employees'].split(','))
        cases.append(dict(id=stable_id('C',gid,case['start'],case['end'],sorted(employees)),
                          group_id=gid,start=case['start'].isoformat(),end=case['end'].isoformat(),
                          days=case['days'],people=case['people'],percent=case['percent'],
                          workforce=case['workforce'],employee_ids=employees))
    for values in (groups,leaves,cases):
        if len({x['id'] for x in values})!=len(values):
            raise ValueError('تعذر إنشاء معرفات فريدة للبيانات.')
    return groups,leaves,cases


def source_fingerprint(path):
    digest=hashlib.sha256()
    with Path(path).open('rb') as source:
        for block in iter(lambda:source.read(1024*1024),b''):digest.update(block)
    return digest.hexdigest()


def as_excel_date(value):
    return date.fromisoformat(value) if value else None


def workbook_data(calculated, groups, plan):
    """Single data specification shared by the local Excel writer and validation."""
    G={g['id']:g for g in groups};L={l['id']:l for l in plan['leaves']}
    target=[l for l in plan['leaves'] if l['target']]
    target.sort(key=lambda l:(not l['remaining_conflict'],l['original_start'],G[l['group_id']]['manager'],l['employee_id'],l['id']))
    future_start=(date.fromisoformat(PRESENTATION_DATE)+timedelta(days=1)).isoformat()
    rows=[]
    for l in target:
        g=G[l['group_id']]
        duration=(date.fromisoformat(l['end'])-date.fromisoformat(l['start'])).days+1
        rows.append([l['employee_id'],l['name'],g['manager'],g['port'],
                     as_excel_date(l['original_start']),as_excel_date(l['original_end']),
                     as_excel_date(l['final_start']),as_excel_date(l['final_end']),
                     duration,l['shift_days'],'نعم' if l['changed'] else 'لا',
                     l['result'],l['reason'],g['level1'],g['level2'],g['level3'],
                     g['workforce'],l['id'],g['id'],','.join(l['original_case_ids']),
                     '، '.join(str(x) for x in l['source_rows'])])
    plans=dict(name='خطة التصحيح',table='LeavePlan',header_row=8,
               title='خطة تصحيح الإجازات المستقبلية',
               notes=[
                   f'تاريخ العرض: {PRESENTATION_DATE}. بداية البدائل: {(date.fromisoformat(PRESENTATION_DATE)+timedelta(days=1)).isoformat()}. نهاية البدائل: {PERIOD_END}.',
                   'تستهدف كل حالات التقاطع التي تبدأ بعد تاريخ العرض، بكل النسب. الإجازة تنقل كاملة بنفس مدتها.',
                   'هذه نتائج تشغيل محفوظة. عند تغيير التواريخ في المصادر أعيدي تشغيل السكربت لتحديث الاقتراحات وفحصها.',
                   f"حالات مستقبلية أصلية: {len(plan['cases_before'])}   إجازات مرتبطة بها: {len(target)}   تغير موعدها: {sum(l['changed'] for l in target)}   بقي لها تقاطع: {sum(l['remaining_conflict'] for l in target)}",
                   'التحريك الموجب تأخير، والسالب تقديم. إذا بقي تقاطع فتظهر تفاصيله في ورقة التقاطعات المتبقية.',
               ],
               headers=['الرقم الوظيفي','اسم الموظف','المدير','المنفذ','البداية الأصلية','النهاية الأصلية',
                        'البداية المقترحة','النهاية المقترحة','المدة (يوم)','التحريك (يوم)','تغير الموعد',
                        'نتيجة الفحص','السبب','المستوى الأول','المستوى الثاني','المستوى الثالث',
                        'إجمالي المجموعة','معرف الإجازة','معرف المجموعة','معرفات الحالات الأصلية','صفوف المصدر'],
               rows=rows,widths=[18,30,30,30,18,18,18,18,14,14,14,29,62,25,25,25,18,30,30,40,22],
               date_cols=[5,6,7,8],percent_cols=[],wrap_cols=[2,3,4,12,13,14,15,16,20,21],
               empty='لا توجد إجازات مرتبطة بحالات تبدأ بعد تاريخ العرض.',
               result_col=12,result_warning='يحتاج',technical_cols=[18,19,20,21])
    cases_rows=[]
    for c in plan['cases_before']:
        g=G[c['group_id']];leave_ids=c.get('leave_ids',c.get('original_leave_ids',[]))
        if not leave_ids:
            leave_ids=[l['id'] for l in target if c['id'] in l['original_case_ids']]
        linked=[L[x] for x in leave_ids]
        resolved=c['resolved']
        cases_rows.append([c['id'],g['manager'],g['port'],g['level1'],g['level2'],g['level3'],
                           g['workforce'],c['people'],c['percent']/100,
                           as_excel_date(c['start']),as_excel_date(c['end']),c['days'],
                           ','.join(c['employee_ids']),
                           'حُل تقاطع الإجازات المرتبطة' if resolved else 'بقي تقاطع في إجازات مرتبطة',
                           ','.join(leave_ids),g['id']])
    case_headers=['معرف الحالة','المدير','المنفذ','المستوى الأول','المستوى الثاني','المستوى الثالث',
                  'إجمالي المجموعة','الموظفون المتقاطعون','نسبة الغياب','بداية التقاطع','نهاية التقاطع',
                  'مدة التقاطع (يوم)','أرقام الموظفين','النتيجة بعد التصحيح','معرفات الإجازات','معرف المجموعة']
    before=dict(name='الحالات المستقبلية',table='OriginalFutureCases',header_row=6,
                title='الحالات المستقبلية قبل التصحيح',
                notes=['الحالة هي نفس الموظفين اثنين فأكثر في أيام متصلة ضمن المجموعة نفسها.',
                       f'التصنيف بتاريخ العرض. تعرض هذه الورقة الحالات التي تبدأ من {future_start} فقط.',
                       'الحالة قد تضم أكثر من سجل إجازة. النتيجة تفحص الإجازات المرتبطة بعد إعادة الجدولة.'],
                headers=case_headers,rows=cases_rows,
                widths=[30,30,30,25,25,25,18,20,16,18,18,20,36,38,50,30],
                date_cols=[10,11],percent_cols=[9],wrap_cols=[2,3,4,5,6,13,14,15],
                empty='لا توجد حالات تقاطع تبدأ بعد تاريخ العرض.',technical_cols=[])
    remaining=[]
    for c in plan['cases_after']:
        g=G[c['group_id']]
        remaining.append([c['id'],g['manager'],g['port'],g['level1'],g['level2'],g['level3'],
                          g['workforce'],c['people'],c['percent']/100,as_excel_date(c['start']),
                          as_excel_date(c['end']),c['days'],','.join(c['employee_ids']),g['id']])
    after=dict(name='التقاطعات المتبقية',table='RemainingWindowCases',header_row=6,
               title=f'التقاطعات المتبقية من {future_start} إلى {PERIOD_END}',
               notes=['أعيد الحساب من المواعيد النهائية بجميع سجلات الإجازات، ومنها الإجازات المثبتة.',
                      'تتضمن هذه الورقة الجزء المستقبلي من التقاطعات بدأت قبل تاريخ العرض أو في يومه.',
                      'عددها لا يقارن مباشرة بعدد الحالات المستقبلية الأصلية لأن نطاقي التاريخ مختلفان.'],
               headers=case_headers[:13]+['معرف المجموعة'],rows=remaining,
               widths=[30,30,30,25,25,25,18,20,16,18,18,20,36,30],date_cols=[10,11],percent_cols=[9],
               wrap_cols=[2,3,4,5,6,13],empty='لا توجد تقاطعات متبقية خلال الفترة المستقبلية.',technical_cols=[])
    # This complete calendar is the sole future HTML input; it includes unchanged bookings.
    machine_headers=['leave_id','group_id','employee_id','employee_name','manager_id','manager','manager_title',
                     'port','level1','level2','level3','workforce','original_start','original_end',
                     'original_effective_start','original_effective_end','final_start','final_end',
                     'final_effective_start','final_effective_end','duration_days','shift_days',
                     'is_target','movable','changed','remaining_conflict','result_code','result','reason',
                     'locked_reason','original_case_ids','conflict_leave_ids','source_rows','ambiguous']
    machine=[]
    for l in plan['leaves']:
        g=G[l['group_id']]
        machine.append([l['id'],g['id'],l['employee_id'],l['name'],g['manager_id'],g['manager'],g['title'],
                        g['port'],g['level1'],g['level2'],g['level3'],g['workforce'],
                        *[as_excel_date(l[k]) for k in ['original_start','original_end','start','end','final_start','final_end','final_effective_start','final_effective_end']],
                        (date.fromisoformat(l['end'])-date.fromisoformat(l['start'])).days+1,l['shift_days'],
                        int(l['target']),int(l['movable']),int(l['changed']),int(l['remaining_conflict']),
                        l['result_code'],l['result'],l['reason'],l['locked_reason'],
                        ','.join(l['original_case_ids']),','.join(l['conflict_leave_ids']),
                        ','.join(str(x) for x in l['source_rows']),int(l['ambiguous'])])
    data=dict(name='بيانات الإجازات',table='PlannerLeaves',header_row=1,title='',notes=[],
              headers=machine_headers,rows=machine,widths=[30,30,18,30,18,30,30,30,25,25,25,14]+[19]*8+[16]*6+[24,30,62,52,40,40,22,14],
              date_cols=list(range(13,21)),percent_cols=[],wrap_cols=[4,6,7,8,9,10,11,28,29,30,31,32,33],
              empty='',technical_cols=[])
    meta=[
        ['schema_version','1'],['presentation_date',PRESENTATION_DATE],
        ['planning_start',(date.fromisoformat(PRESENTATION_DATE)+timedelta(days=1)).isoformat()],
        ['planning_end',PERIOD_END],['report_period_start',PERIOD_START],['report_period_end',PERIOD_END],
        ['future_definition','case_start > presentation_date'],['scope','all future overlap cases; at least 2 employees'],
        ['group_definition','manager_id + location + level1 + level2 + level3'],
        ['end_date_inclusive',int(END_DATE_INCLUSIVE)],
        ['generated_at',datetime.now().isoformat(timespec='seconds')],
        ['managers_source',Path(MANAGERS_FILE).name],['managers_sha256',source_fingerprint(MANAGERS_FILE)],
        ['staff_sheet',STAFF_SHEET],['leaves_source',Path(LEAVES_FILE).name],
        ['leaves_sha256',source_fingerprint(LEAVES_FILE)],['leaves_sheet',LEAVES_SHEET],
        ['source_leave_rows',calculated['source_rows']['leaves']],['report_matched_records',calculated['matched']],
        ['distinct_full_leave_records',len(plan['leaves'])],['report_all_overlap_cases',len(calculated['cases'])],
        ['original_future_cases',len(plan['cases_before'])],['target_leaves',len(target)],
        ['changed_target_leaves',sum(l['changed'] for l in target)],
        ['resolved_target_leaves',sum(not l['remaining_conflict'] for l in target)],
        ['unresolved_target_leaves',sum(l['remaining_conflict'] for l in target)],
        ['remaining_window_cases',len(plan['cases_after'])],
        ['max_search_nodes_per_group',MAX_SEARCH_NODES_PER_GROUP],['max_search_nodes_total',MAX_SEARCH_NODES_TOTAL],
    ]
    for key,value in sorted(plan.get('stats',{}).items()):
        if isinstance(value,(int,float,str,bool)):meta.append(['solver.'+key,value])
    for key,value in sorted(calculated['problems'].items()):meta.append(['ملاحظة مصدر: '+key,value])
    for status in plan.get('group_solver_status',[]):
        if isinstance(status,dict):
            gid=status.get('group_id','')
            meta.append(['بحث المجموعة '+gid,json.dumps(status,ensure_ascii=False)])
    metadata=dict(name='إعدادات وملاحظات',table='PlannerSettings',header_row=6,
                  title='إعدادات المنسّق وملاحظات المصادر',
                  notes=['المجموعة تعني المدير والموقع والمسار التنظيمي معًا، وفق منطق التقرير.',
                         'النتائج لقطة عند التشغيل. إعادة الجدولة تستخدم أيامًا تقويمية ومدة الإجازة الأصلية كاملة.',
                         'الإجازات غير المستهدفة مثبتة. البحث محدود؛ عدم العثور على بديل لا يثبت استحالة الحل.'],
                  headers=['البند','القيمة'],rows=meta,widths=[55,108],date_cols=[],percent_cols=[],wrap_cols=[1,2],
                  empty='',technical_cols=[])
    return [plans,before,after,data,metadata]


def write_planner_workbook(sheets,path):
    """User-side writer: uses the same openpyxl dependency as the supplied report."""
    wb=Workbook();wb.remove(wb.active)
    navy='152446';teal='00828E';muted='64748B';line='D9E0E6'
    for spec in sheets:
        ws=wb.create_sheet(spec['name']);ws.sheet_view.rightToLeft=True;ws.sheet_view.showGridLines=False
        ws.sheet_view.zoomScale=85;ws.sheet_properties.pageSetUpPr.fitToPage=True
        ws.sheet_properties.tabColor=navy if spec['name']=='خطة التصحيح' else teal if spec['name']=='بيانات الإجازات' else 'CBD5E1'
        header=spec['header_row'];ncols=len(spec['headers']);nrows=len(spec['rows'])
        def put(row,col,value):
            item=ws.cell(row,col,value)
            if isinstance(value,str):item.data_type='s'  # Source text never becomes an Excel formula.
            item.font=Font(name='Arial',size=11,color=navy)
            item.alignment=Alignment(horizontal='right',vertical='center',readingOrder=2)
            return item
        if spec['title']:
            put(2,1,spec['title']).font=Font(name='Arial',size=16,bold=True,color=navy)
            # Intentional presentation heading; no data cells are merged.
            ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=min(13,ncols))
            ws.row_dimensions[2].height=29
            for i,note in enumerate(spec['notes'],3):
                put(i,1,note).font=Font(name='Arial',size=11,color=muted)
                ws.merge_cells(start_row=i,start_column=1,end_row=i,end_column=min(13,ncols))
                ws.cell(i,1).alignment=Alignment(horizontal='right',vertical='center',wrap_text=True,readingOrder=2)
                ws.row_dimensions[i].height=31 if ncols==2 else 26
        for j,name in enumerate(spec['headers'],1):
            c=put(header,j,name);c.fill=PatternFill('solid',fgColor=navy)
            c.font=Font(name='Arial',size=11,color='FFFFFF',bold=True)
            c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True,readingOrder=2)
            c.border=Border(right=Side(style='thin',color='FFFFFF'))
        ws.row_dimensions[header].height=36
        for i,row in enumerate(spec['rows'],header+1):
            if len(row)!=ncols:raise ValueError(f"اختلاف عدد الأعمدة في {spec['name']}")
            ws.row_dimensions[i].height=48 if spec['name']=='خطة التصحيح' else 35
            for j,value in enumerate(row,1):
                c=put(i,j,value)
                if j in spec['wrap_cols']:c.alignment=Alignment(horizontal='right',vertical='center',wrap_text=True,readingOrder=2)
                if isinstance(value,(int,float)):
                    c.number_format='#,##0';c.alignment=Alignment(horizontal='center',vertical='center')
                if j in spec['date_cols']:
                    c.number_format='dd/mm/yyyy';c.alignment=Alignment(horizontal='center',vertical='center')
                if j in spec['percent_cols']:c.number_format='0.0%'
                if isinstance(value,str) and (j in (1,2) or 'id' in spec['headers'][j-1]):c.number_format='@'
            if spec['name']=='خطة التصحيح':
                for j in (7,8):ws.cell(i,j).font=Font(name='Arial',size=11,bold=True,color=teal)
        end=header+max(nrows,1)
        if nrows:
            table=Table(displayName=spec['table'],ref=f'A{header}:{get_column_letter(ncols)}{end}')
            table.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True,showColumnStripes=False)
            ws.add_table(table)
        elif spec['empty']:
            put(header+1,1,spec['empty']);ws.merge_cells(start_row=header+1,start_column=1,end_row=header+1,end_column=min(ncols,10))
        for j,width in enumerate(spec['widths'],1):ws.column_dimensions[get_column_letter(j)].width=width
        for j in spec['technical_cols']:
            ws.column_dimensions[get_column_letter(j)].outlineLevel=1
        ws.freeze_panes=f'C{header+1}' if ncols>4 else f'A{header+1}'
        ws.print_title_rows=f'1:{header}'
        ws.print_options.horizontalCentered=True
        ws.page_setup.orientation='landscape';ws.page_setup.paperSize=ws.PAPERSIZE_A3
        ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=0
        ws.print_area=f'A1:{get_column_letter(min(ncols,13))}{max(end,header+1)}'
        if spec['name']=='خطة التصحيح' and nrows:
            ws.conditional_formatting.add(f'L{header+1}:M{end}',FormulaRule(
                formula=[f'OR(ISNUMBER(SEARCH("لم يحل",$L{header+1})),ISNUMBER(SEARCH("يحتاج",$L{header+1})),ISNUMBER(SEARCH("متبق",$L{header+1})))'],
                fill=PatternFill('solid',fgColor='FFF0DB'),font=Font(color='8A4B0F')))
    out=Path(path);out.parent.mkdir(parents=True,exist_ok=True)
    if out.suffix.lower()!='.xlsx':raise ValueError('يجب أن ينتهي ملف الناتج بـ .xlsx')
    for source in (MANAGERS_FILE,LEAVES_FILE):
        if out.resolve()==Path(source).resolve():raise ValueError('اختاري اسم ناتج مختلفًا عن ملفات المصادر.')
    fd,temp=tempfile.mkstemp(prefix='smart_planner_',suffix='.xlsx',dir=out.parent);os.close(fd)
    try:
        wb.save(temp);wb.close()
        # Verify the saved calendar, including dates and literal IDs, before replacement.
        check=load_workbook(temp,read_only=True,data_only=False)
        try:
            actual=list(check['بيانات الإجازات'].values)
            expected=next(s for s in sheets if s['name']=='بيانات الإجازات')
            if list(actual[0])!=expected['headers'] or len(actual)-1!=len(expected['rows']):
                raise ValueError('فشل التحقق من جدول الإجازات المحفوظ.')
            normalize=lambda v:v.date() if isinstance(v,datetime) else v
            for saved,row in zip(actual[1:],expected['rows']):
                if [normalize(v) if v is not None else '' for v in saved]!=[v if v is not None else '' for v in row]:
                    raise ValueError('تغيّرت قيمة عند حفظ جدول الإجازات.')
        finally:check.close()
        os.replace(temp,out)
    finally:
        if os.path.exists(temp):os.unlink(temp)
    return out


def main():
    if date.fromisoformat(PRESENTATION_DATE)<date.fromisoformat(PERIOD_START) or date.fromisoformat(PRESENTATION_DATE)>=date.fromisoformat(PERIOD_END):
        raise ValueError('يجب أن يكون تاريخ العرض داخل فترة التقرير وقبل نهايتها.')
    calculated=calculate(PRESENTATION_DATE)
    groups,leaves,cases=planner_inputs(calculated)
    print('4/5 اقتراح المواعيد وفحصها مقابل جميع إجازات المجموعة...',flush=True)
    plan=build_plan(groups,leaves,cases,presentation_date=PRESENTATION_DATE,
                    period_end=PERIOD_END,end_date_inclusive=END_DATE_INCLUSIVE,
                    max_nodes=MAX_SEARCH_NODES_PER_GROUP,max_total_nodes=MAX_SEARCH_NODES_TOTAL)
    sheets=workbook_data(calculated,groups,plan)
    out=Path(OUTPUT_XLSX) if OUTPUT_XLSX else Path(LEAVES_FILE).resolve().parent/'smart_planner.xlsx'
    print('5/5 كتابة ملف Excel...',flush=True)
    write_planner_workbook(sheets,out)
    targeted=[l for l in plan['leaves'] if l['target']]
    print('تم إنشاء: '+str(out),flush=True)
    print('الحالات المستقبلية الأصلية: '+str(len(plan['cases_before'])))
    print('الإجازات المرتبطة بالحالات: '+str(len(targeted)))
    print('الإجازات التي تغير موعدها: '+str(sum(l['changed'] for l in targeted)))
    print('الإجازات المستهدفة التي بقي لها تقاطع: '+str(sum(l['remaining_conflict'] for l in targeted)))
    print('التفاصيل في ورقة «خطة التصحيح»، والجدول الكامل للربط لاحقًا في «بيانات الإجازات».')


if __name__=='__main__':
    try:main()
    except PermissionError:
        print('خطأ: أغلقي ملف Excel الناتج إذا كان مفتوحًا ثم أعيدي التشغيل.',file=sys.stderr);sys.exit(1)
    except Exception as exc:
        print('خطأ: '+str(exc),file=sys.stderr);sys.exit(1)

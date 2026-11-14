# -*- coding: utf-8 -*-
"""
داشبورد تقاطع الإجازات — نفس منطق سكربت التقرير المرفق.

التثبيت مرة واحدة:
    python -m pip install pandas openpyxl
التشغيل:
    python ports_dashboard_standalone.py

عدّلي مساري Employee Master وPorts Leaves أدناه، ثم شغّلي الملف.
الناتج: dashboard.html بجانب ملف الإجازات، ويُفتح بالمتصفح مباشرة.
ملف Python واحد، ويُنشئ HTML فقط. جميع مكونات العرض مضمنة وتعمل دون إنترنت.

الحالة = نفس مجموعة موظفين اثنين فأكثر في أيام تقويمية متصلة، ضمن
المدير نفسه + الموقع نفسه + المسار التنظيمي نفسه للمستويات الأول والثاني والثالث.
تغيّر الموظفين أو انقطاع الأيام يبدأ حالة جديدة. النسبة من موظفي المجموعة نفسها.
تعرض الأسماء والمسميات كما وردت في المصادر، وتُعرض جميع الحالات بغض النظر عن نسبتها.
المنسّق الذكي يعرض حالات التقرير التي تبدأ بعد تاريخ العرض، ويقترح مواعيد
داخل الصفحة مع إبقاء مدة الإجازة وإجازات الموظفين الآخرين. لا يكتب تغييرات إلى المصادر.
"""
from pathlib import Path
from datetime import date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError
from collections import defaultdict, Counter
import json
import os
import re
import sys
import tempfile
import unicodedata
import pandas as pd

SCRIPT_VERSION = '2026-11-14-standalone-v5'

# ===================== عدّلي المسارات هنا =====================
MANAGERS_FILE = r'C:\Users\rammutairi\Downloads\Employee Master.xlsx'
LEAVES_FILE = r'C:\Users\rammutairi\Downloads\Ports Leaves.xlsx'
OUTPUT_HTML = r''  # فارغ: dashboard.html بجانب ملف الإجازات
# ============================================================
MANAGERS_SHEET = 'مدراء المنافذ'
STAFF_SHEET = 'الموظفين تحت مدراء المنافذ'
LEAVES_SHEET = 'البيانات الخام'

try:
    RIYADH_TZ = ZoneInfo('Asia/Riyadh')
except ZoneInfoNotFoundError:
    # بعض أجهزة Windows لا تتضمن قاعدة المناطق الزمنية. الرياض UTC+3.
    RIYADH_TZ = timezone(timedelta(hours=3))
AS_OF = datetime.now(RIYADH_TZ).date().isoformat()  # لتثبيت التصنيف ضعي '2026-11-18' مثلًا
PERIOD_START = '2026-10-01'
PERIOD_END = '2026-12-31'
PRESENTATION_DATE = '2026-11-18'  # الحالات المستقبلية في المنسّق تبدأ بعد هذا التاريخ
THRESHOLD = 50.0
END_DATE_INCLUSIVE = True
ALLOW_NO_APPROVAL_COLUMN = True

REPORT_TITLE = 'تقرير تقاطع الإجازات حسب المدير وموقع العمل والوحدة التنظيمية'
DEPARTMENT = 'إدارة أداء وكالة رأس المال البشري'
PREPARED_BY = 'تنفيذ رغد المطيري'
SUPERVISED_BY = 'إشراف أمل السويلم'

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
    asof=date.fromisoformat(asof_text or AS_OF)
    print(f'تاريخ إعداد التقرير والتصنيف: {asof:%Y-%m-%d} (تاريخ التصنيف المعتمد في هذا التشغيل)', flush=True)
    pstart=date.fromisoformat(PERIOD_START);pend=date.fromisoformat(PERIOD_END)
    if pstart > pend:raise ValueError('تاريخ بداية الفترة بعد نهايتها.')
    print('1/4 قراءة ملفات المصادر...',flush=True)
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
    print(f'2/4 مطابقة الإجازات مع {len(staff_by_id)} موظفًا مستهدفًا...',flush=True)
    daily=defaultdict(set); employee_leave_dates=defaultdict(list); record_signatures=set()
    accepted_records=[]  # Full source intervals for scheduling; report deduplication is unchanged.
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

    print('3/4 حساب مجموعات الغياب المتطابقة عبر أيام متصلة...',flush=True)
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
                         employees=','.join(sorted(cohort)),flag=ratio>=THRESHOLD))

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


def short_action(case):
    if case['status'] == 'منتهية':
        return 'منتهية — للعلم'
    if case['status'] == 'جارية':
        return 'مراجعة التغطية الحالية'
    return 'مراجعة الجدولة قبل تاريخ البداية'


def extract_planner(result, group_ids, cases):
    """Retain complete accepted bookings; only eligible future-case leaves may move."""
    presentation=date.fromisoformat(PRESENTATION_DATE)
    first=max(result['pstart'], presentation+timedelta(days=1))
    last=result['pend']
    uncertain={str(row[1]) for row in result['ambiguous'] if row[0]=='تبعيات متعددة'}
    unique={}; clipped_originals=defaultdict(set)
    for record in result['accepted_records']:
        key=(record['group'],record['employee_id'],record['original_start'],record['original_end'])
        clipped=(record['group'],record['employee_id'],record['clipped_start'],record['clipped_end'])
        clipped_originals[clipped].add(key)
        if key not in unique:
            unique[key]={
                'group_id':group_ids[record['group']], 'employee_id':record['employee_id'],
                'name':result['employee_display_names'].get(record['employee_id'])
                    or record['source_name'] or 'اسم غير متوفر في المصدر',
                'start':record['start'].isoformat(), 'end':record['end'].isoformat(),
                'source_start':record['original_start'].isoformat(),
                'source_end':record['original_end'].isoformat(), 'source_rows':[],
                'ambiguous':record['employee_id'] in uncertain,
            }
        unique[key]['source_rows'].append(record['source_row'])
    # Different full intervals can collapse to the same clipped report interval.
    # Preserve both as obstacles, and do not choose a duration on the user's behalf.
    for keys in clipped_originals.values():
        if len(keys)>1:
            for key in keys:unique[key]['ambiguous']=True
    records=sorted(unique.values(),key=lambda r:(r['group_id'],r['employee_id'],r['start'],r['end']))
    by_employee=defaultdict(list)
    for index, record in enumerate(records,1):
        record['id']=f'L{index:06d}'
        record['source_rows']=sorted(set(record['source_rows']))
        by_employee[record['employee_id']].append(record)
    overlapping_source_ids=set()
    for items in by_employee.values():
        active=[]
        for record in sorted(items,key=lambda r:(r['start'],r['end'],r['id'])):
            active=[old for old in active if old['end']>=record['start']]
            for old in active:
                overlapping_source_ids.update((record['id'],old['id']))
            active.append(record)
    future_by_group=defaultdict(list)
    for case in cases:
        if case['start']>presentation.isoformat():future_by_group[case['group_id']].append(case)
    for record in records:
        linked=[case['id'] for case in future_by_group[record['group_id']]
                if record['employee_id'] in case['employee_ids']
                and record['start']<=case['end'] and record['end']>=case['start']]
        reason=''
        if not linked:reason='إجازة ثابتة خارج حالات التقاطع المستقبلية المستهدفة.'
        elif record['ambiguous']:reason='تحتاج بيانات الإجازة أو تبعية الموظف إلى مراجعة؛ موعدها ثابت.'
        elif record['id'] in overlapping_source_ids:
            reason='توجد سجلات إجازة متداخلة للموظف نفسه؛ ثُبتت للمراجعة.'
        elif record['start']<first.isoformat():reason='بدأت الإجازة بحلول تاريخ العرض؛ موعدها ثابت.'
        elif record['end']>last.isoformat():reason='تمتد الإجازة خارج فترة الجدولة؛ حُفظت مدتها وموعدها.'
        record['movable']=not bool(reason)
        record['locked_reason']=reason
        record['original_case_ids']=linked
    return {
        'presentation_date':presentation.isoformat(),
        'period_start':result['pstart'].isoformat(), 'period_end':last.isoformat(),
        'end_date_inclusive':END_DATE_INCLUSIVE, 'leaves':records,
    }


def extract_data(result):
    """Build display data from the same cases; display labels do not change grouping."""
    groups = []
    group_ids = {}
    for number, (key, members) in enumerate(sorted(result['group_staff'].items()), 1):
        group_id = f'G{number:04d}'
        group_ids[key] = group_id
        manager = result['managers_by_id'][key[0]]
        levels = result['display_levels'][key]
        groups.append({
            'id': group_id, 'manager_id': key[0], 'manager': manager['name'],
            'title': manager['job'], 'port': result['display_port'][key],
            'level1': levels[0], 'level2': levels[1], 'level3': levels[2],
            'workforce': len(members),
        })

    cases = []
    for number, case in enumerate(result['all_cases'], 1):
        employee_ids = case['employees'].split(',')
        employees = []
        for employee_id in sorted(employee_ids, key=lambda eid: (
                result['employee_display_names'].get(eid, eid), eid)):
            periods = sorted({(original_start, original_end)
                for start, end, original_start, original_end
                in result['employee_leave_dates'].get(employee_id, [])
                if start <= case['end'] and end >= case['start']})
            employees.append({
                'id': employee_id,
                'name': result['employee_display_names'].get(employee_id) or 'اسم غير متوفر في المصدر',
                'leave_periods': [{'start': start.isoformat(), 'end': end.isoformat()}
                                  for start, end in periods],
            })
        cases.append({
            'id': f'C{number:04d}', 'group_id': group_ids[case['group']],
            'start': case['start'].isoformat(), 'end': case['end'].isoformat(),
            'days': case['days'], 'people': case['people'], 'workforce': case['workforce'],
            'percent': case['percent'], 'percent_display': round(case['percent'], 1),
            'remaining': case['remaining'], 'status': case['status'],
            'action': short_action(case), 'flag': bool(case['flag']),
            'eta': case['eta'], 'remaining_days': case['remaining_days'],
            'employee_ids': employee_ids, 'employees': employees,
        })

    data = {
        'meta': {
            'title': REPORT_TITLE, 'department': DEPARTMENT,
            'prepared_by': PREPARED_BY, 'supervised_by': SUPERVISED_BY,
            'as_of': result['asof'].isoformat(),
            'period_start': result['pstart'].isoformat(), 'period_end': result['pend'].isoformat(),
            'threshold': THRESHOLD, 'end_date_inclusive': END_DATE_INCLUSIVE,
        },
        'groups': groups,
        'cases': cases,
        'planner': extract_planner(result, group_ids, cases),
        'qa': {
            'target_managers': len(result['managers_by_id']),
            'employees': len(result['staff_by_id']), 'matched_leaves': result['matched'],
            'groups': len(groups), 'all_cases': len(cases),
            'threshold_cases': len(result['priority']),
            'problems': [{'label': label, 'count': count}
                         for label, count in sorted(result['problems'].items())],
        },
    }
    # Gate only on consistency with this calculation, never on a different reference workbook.
    if sum(case['flag'] for case in cases) != len(result['priority']):
        raise ValueError('فشل اتساق عدد حالات النسبة المحددة مع الحساب الأصلي.')
    if len(cases) != len(result['cases']):
        raise ValueError('فشل نقل جميع حالات التقاطع إلى الداشبورد.')
    return data


def write_html(data, path):
    """Write one self-contained HTML file atomically; retain last good output on error."""
    out = Path(path).expanduser()
    if out.suffix.lower() != '.html':
        raise ValueError('يجب أن ينتهي اسم ملف الناتج بـ .html')
    if HTML_TEMPLATE.count('__REPORT_JSON__') != 1:
        raise ValueError('قالب العرض غير مكتمل؛ يجب وجود موضع واحد للبيانات.')
    payload = json.dumps(data, ensure_ascii=False, separators=(',', ':'), allow_nan=False)
    for character, escaped in (('<', '\\u003c'), ('>', '\\u003e'), ('&', '\\u0026'),
                               ('\u2028', '\\u2028'), ('\u2029', '\\u2029')):
        payload = payload.replace(character, escaped)
    document = HTML_TEMPLATE.replace('__REPORT_JSON__', payload)
    out.parent.mkdir(parents=True, exist_ok=True)
    handle, temp_path = tempfile.mkstemp(prefix='dashboard_', suffix='.tmp', dir=out.parent)
    try:
        with os.fdopen(handle, 'w', encoding='utf-8', newline='\n') as stream:
            stream.write(document)
        os.replace(temp_path, out)
    finally:
        if os.path.exists(temp_path):
            os.unlink(temp_path)
    return out.resolve()


def build():
    print('إصدار الداشبورد: ' + SCRIPT_VERSION, flush=True)
    print('الملف المشغّل: ' + str(Path(__file__).resolve()), flush=True)
    print('مصادر البيانات: ملف الموظفين وملف الإجازات فقط.', flush=True)
    result = calculate()
    data = extract_data(result)
    out = Path(OUTPUT_HTML) if OUTPUT_HTML else Path(LEAVES_FILE).resolve().parent / 'dashboard.html'
    print('4/4 إنشاء الداشبورد...', flush=True)
    out = write_html(data, out)
    ongoing = sum(case['status'] == 'جارية' for case in data['cases'])
    upcoming = sum(case['status'] == 'قادمة' for case in data['cases'])
    print(f"جميع حالات التقاطع: {len(data['cases'])}", flush=True)
    print(f"حالات {THRESHOLD:g}% فأكثر: {data['qa']['threshold_cases']}", flush=True)
    print(f'الحالات الجارية: {ongoing} | الحالات القادمة: {upcoming}', flush=True)
    future=[case for case in data['cases'] if case['start']>PRESENTATION_DATE]
    high=sum(case['percent']>=50.0 for case in future)
    print(f'المنسّق: {len(future)} حالات مستقبلية بعد {PRESENTATION_DATE}، منها {high} بنسبة 50% فأكثر.', flush=True)
    print('تم إنشاء الداشبورد: ' + str(out), flush=True)
    print('افتحي هذا الملف في المتصفح لعرض النتائج.', flush=True)
    if result['out_of_port']:
        print(f"سجلات الإجازات المستبعدة لاختلاف الموقع: {result['out_of_port']} (ضمن فحص البيانات).", flush=True)
    return data


HTML_TEMPLATE = r"""<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="color-scheme" content="light">
<title>لوحة تقاطع الإجازات</title>
<style>
:root{--ink:#17334a;--navy:#123752;--navy2:#214a62;--teal:#087f83;--teal-light:#e4f1ef;--amber:#a36120;--amber-light:#fbefdf;--paper:#fffefa;--page:#f3f4f1;--line:#dce2e0;--muted:#65757e;--shadow:0 7px 24px #17334a06}
*{box-sizing:border-box}html{scroll-behavior:smooth}body{margin:0;background:var(--page);color:var(--ink);font:14px Tahoma,Arial,sans-serif;line-height:1.65}button,input,select{font:inherit}button{cursor:pointer}button,input,select{min-height:44px;border:1px solid var(--line);border-radius:8px}button{padding:10px 16px;background:white;color:var(--navy);font-weight:700}button:hover{background:#edf4f3}button:disabled{cursor:default;opacity:.4}input,select{width:100%;padding:10px 12px;background:white;color:var(--ink)}input::placeholder{color:#7b898e}button:focus-visible,input:focus-visible,select:focus-visible,summary:focus-visible,a:focus-visible{outline:3px solid #bb792c;outline-offset:3px}bdi{unicode-bidi:isolate}a{color:var(--teal)}[hidden]{display:none!important}
.wrap{width:min(1440px,100%);margin-inline:auto;padding-inline:38px}.hero{background:var(--navy);color:#fff;position:relative;overflow:hidden}.hero::after{content:"";position:absolute;bottom:0;inset-inline:0;height:4px;background:linear-gradient(90deg,#b7bf9c,#56a6a2 50%,#087f83)}.hero-inner{padding-block:28px 31px;display:flex;justify-content:space-between;gap:30px;align-items:center}.department{font-size:13px;color:#c5dcdf;margin:0 0 10px}.hero h1{font-size:clamp(23px,2.5vw,33px);font-weight:700;line-height:1.6;margin:0;max-width:900px}.period{display:flex;align-items:center;gap:10px;margin-top:13px;color:#dae7eb;font-size:13px}.period::before{content:"";width:23px;height:2px;background:#78c3bb;flex-shrink:0}.hero-aside{flex-shrink:0;border-inline-start:1px solid #ffffff26;padding-inline-start:25px;text-align:start}.hero-aside span{display:block;color:#c5dcdf;font-size:12px}.hero-aside strong{display:block;font-size:16px;margin-top:5px}.hero-aside .print-button{margin-top:14px;border-color:#ffffff50;background:#ffffff10;color:#fff;padding:8px 17px;font-size:12px}.hero-aside .print-button:hover{background:#ffffff20}main{padding-block:26px 34px}.eyebrow{font-size:11px;letter-spacing:.3px;color:var(--teal);font-weight:700}.section-title{margin:0;font-size:20px;line-height:1.6;color:var(--navy)}.section-subtitle{font-size:12px;color:var(--muted);margin:4px 0 0}.section-header{display:flex;align-items:center;justify-content:space-between;gap:18px;margin-bottom:18px}.scope-label{padding:5px 11px;border:1px solid #c8dcd7;background:var(--teal-light);color:#17665f;border-radius:30px;font-size:11px;white-space:nowrap}.filters{display:grid;grid-template-columns:1fr 1fr 1.4fr auto;gap:12px;align-items:end;background:var(--paper);padding:17px 20px;border:1px solid var(--line);border-radius:12px;margin-bottom:20px;box-shadow:var(--shadow)}.field label{font-size:11px;color:var(--muted);display:block;margin-bottom:6px}.reset-button{font-size:12px;white-space:nowrap;font-weight:400}.kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px;margin-bottom:18px}.kpi{background:var(--paper);border:1px solid var(--line);border-radius:12px;padding:18px 22px 16px;position:relative;overflow:hidden;box-shadow:var(--shadow)}.kpi::before{content:"";position:absolute;top:0;inset-inline:0;height:3px;background:var(--navy)}.kpi:nth-child(2)::before{background:var(--amber)}.kpi:nth-child(3)::before{background:var(--teal)}.kpi:nth-child(4)::before{background:#729085}.kpi-head{display:flex;justify-content:space-between;gap:10px;align-items:center;color:var(--muted);font-size:12px}.kpi-index{color:#aebbbf;font-size:10px;letter-spacing:1px;font-weight:700}.kpi-value{font-size:42px;line-height:1.4;letter-spacing:-1px;color:var(--navy);margin:9px 0 2px;font-weight:700;font-variant-numeric:tabular-nums}.kpi:nth-child(2) .kpi-value{color:var(--amber)}.kpi:nth-child(3) .kpi-value{color:var(--teal)}.kpi-note{color:var(--muted);font-size:11px}.highest{border:1px solid #c8dcda;border-radius:12px;overflow:hidden;display:grid;grid-template-columns:220px minmax(0,1fr);background:var(--paper);margin-bottom:26px;box-shadow:var(--shadow)}.highest-score{padding:24px;background:var(--teal-light);display:flex;flex-direction:column;justify-content:center}.highest-score h3{font-size:12px;color:#2b6d69;font-weight:400;margin:0}.highest-number{font-size:49px;line-height:1.4;letter-spacing:-1.5px;color:#146963;font-weight:700;font-variant-numeric:tabular-nums}.highest-score p{font-size:11px;color:#4f756f;margin:5px 0 0}.highest-main{padding:20px 25px;min-width:0}.highest-top{display:flex;justify-content:space-between;align-items:flex-start;gap:18px}.highest-name{font-size:17px;font-weight:700;line-height:1.6}.highest-job{font-size:11px;color:var(--muted);margin-top:2px}.highest-meta{display:grid;grid-template-columns:1fr 1fr;gap:10px 28px;margin-top:13px}.info-label{display:block;font-size:10px;color:var(--muted);margin-bottom:3px}.info-value{font-size:13px;font-weight:700;overflow-wrap:anywhere}.org-path{font-size:11px;color:var(--muted);line-height:1.8;overflow-wrap:anywhere}.org-path bdi:not(:last-child)::after{content:" / ";color:#9eabae;white-space:pre}.highest-bottom{margin-top:13px;padding-top:12px;border-top:1px solid var(--line);display:flex;gap:12px;justify-content:space-between;align-items:center}.highest-bottom .subtle{font-size:11px;margin:0}.link-button{font-size:11px;min-height:36px;padding:6px 10px;background:transparent;color:var(--teal);border-color:#c9dcd8}.badge{display:inline-block;padding:3px 9px;border-radius:6px;font-size:10px;font-weight:700;white-space:nowrap;background:#edf0f1;color:#65757e}.badge.running{background:#e2f1ed;color:#21735e}.badge.future{background:#e5edf7;color:#335c85}.badge.ended{background:#edf0ee;color:#737e78}.badge.threshold{background:var(--amber-light);color:var(--amber)}.subtle{color:var(--muted)}.panel{background:var(--paper);border:1px solid var(--line);border-radius:12px;padding:22px;box-shadow:var(--shadow)}.register{margin-top:0}.table-heading{display:flex;gap:12px;align-items:center}.count-label{font-size:12px;color:var(--muted);background:#eff3f1;border-radius:20px;padding:4px 12px;white-space:nowrap}.table-controls{display:flex;justify-content:space-between;gap:18px;align-items:end;margin-bottom:15px}.table-filters{display:flex;gap:12px;align-items:end;flex-wrap:wrap}.table-filters .field{min-width:165px}.table-scope-note{font-size:11px;color:var(--muted);margin:0;max-width:370px}.table-scroll{overflow:auto;border:1px solid var(--line);border-radius:8px}table{width:100%;border-collapse:collapse;font-size:12px;text-align:start}th{padding:13px 14px;background:#edf2f1;color:#4d6571;white-space:nowrap;font-size:11px;font-weight:700;border-bottom:1px solid #cbd8d5;text-align:start}td{padding:15px 14px;border-bottom:1px solid #e5eae7;vertical-align:top}tbody tr:last-child td{border-bottom:0}tbody tr:hover{background:#f2f8f6}.manager-cell{min-width:175px;max-width:250px}.port-cell{min-width:210px;max-width:310px}.date-cell{min-width:185px}.cell-title{font-weight:700;color:var(--navy)}.cell-note{font-size:10px;color:var(--muted);margin-top:4px;line-height:1.7}.ratio-cell{min-width:105px}.ratio-value{font-size:15px;font-weight:700;color:var(--navy)}.ratio-value.over{color:var(--amber)}.ratio-track{height:4px;background:#e8edeb;width:76px;margin-top:8px;border-radius:4px;overflow:hidden}.ratio-track span{height:100%;display:block;background:var(--teal);border-radius:4px}.ratio-track span.over{background:#bb864d}.people-cell{white-space:nowrap}.people-number{font-size:15px;font-weight:700}.detail-button{min-height:34px;padding:5px 11px;font-size:11px;font-weight:400;white-space:nowrap}.case-id{display:block;color:#93a2a7;font-size:9px;margin-top:4px;direction:ltr;text-align:start}.ended-row .cell-title{color:#5c6f78}.pager{display:flex;align-items:center;justify-content:space-between;gap:15px;margin-top:16px;font-size:11px;color:var(--muted)}.pager-actions{display:flex;gap:8px}.pager button{font-size:11px;min-height:36px;padding:6px 14px;font-weight:400}.empty-cell{padding:44px 20px;text-align:center;color:var(--muted)}.empty-highest{padding:30px;color:var(--muted);align-self:center}.methodology{margin-top:20px;background:#eef2ee;border:1px solid var(--line);border-radius:10px;padding:13px 18px}.methodology summary{cursor:pointer;font-size:12px;font-weight:700;color:var(--navy)}.method-body{font-size:12px;color:#60727b;max-width:1100px;padding:5px 4px 2px}.method-body p{margin:8px 0}.qa-tiles{display:flex;flex-wrap:wrap;gap:9px 22px;border-top:1px solid #d6e0da;margin-top:12px;padding-top:12px;font-size:11px}.qa-tiles b{color:var(--navy);margin-inline-end:4px}.qa-issues{margin-top:13px;font-size:11px}.qa-issues summary{font-size:11px;color:#785b35;font-weight:400}.qa-issues ul{margin-block:8px;padding-inline-start:22px}.footer{display:flex;justify-content:space-between;gap:15px;padding-top:21px;font-size:11px;color:var(--muted)}.footer-credit{display:flex;gap:17px;flex-wrap:wrap}.footer-credit strong{color:#526972;font-weight:400}.noscript{padding:20px;border:1px solid #dfc5a2;background:#fbefdf;color:#765426;margin-block:20px}dialog{width:min(920px,94vw);max-height:88vh;border:1px solid var(--line);border-radius:14px;padding:0;background:var(--paper);color:var(--ink);box-shadow:0 18px 70px #12375230}dialog::backdrop{background:#102d4390;backdrop-filter:blur(2px)}.dialog-head{display:flex;justify-content:space-between;align-items:center;gap:20px;border-bottom:1px solid var(--line);padding:18px 24px;position:sticky;top:0;background:var(--paper);z-index:2}.dialog-head h2{font-size:18px;margin:0}.close-button{min-height:36px;width:36px;padding:0;font-size:23px;font-weight:400;line-height:1;border-radius:50%}.dialog-content{padding:22px 24px 28px}.detail-summary{display:grid;grid-template-columns:1fr 1fr;gap:16px 24px;border-bottom:1px solid var(--line);padding-bottom:19px;margin-bottom:18px}.detail-summary .full{grid-column:1/-1}.detail-summary strong{font-size:14px;overflow-wrap:anywhere}.detail-summary .info-label{font-size:11px}.detail-metrics{display:flex;flex-wrap:wrap;gap:9px;margin:14px 0}.detail-metrics span{background:#edf3f0;padding:6px 10px;border-radius:6px;font-size:11px}.employee-title{font-size:15px;margin:20px 0 5px}.employee-note{font-size:11px;color:var(--muted);margin:0 0 12px}.employee-list{margin:0;padding:0;list-style:none}.employee-item{display:grid;grid-template-columns:29px minmax(0,1fr);gap:10px;padding:14px 0;border-bottom:1px solid var(--line)}.employee-item:last-child{border-bottom:0}.employee-index{font-size:10px;font-weight:700;line-height:28px;text-align:center;width:28px;height:28px;border-radius:50%;background:#e5efeb;color:#417563}.employee-name{font-size:13px;font-weight:700}.employee-id{font-size:10px;color:var(--muted);font-weight:400;margin-inline-start:8px}.leave-period{font-size:12px;color:#536b78;margin-top:4px}.tie-intro{font-size:12px;color:var(--muted);margin:0 0 16px}.tie-card{border:1px solid var(--line);border-radius:9px;padding:15px 17px;margin-top:10px;display:flex;align-items:start;justify-content:space-between;gap:18px}.tie-card .cell-note{font-size:11px}.tie-card .detail-button{flex-shrink:0}.dialog-fallback{position:fixed;inset:6vh 0 auto;z-index:1000;overflow:auto}.fallback-shade{position:fixed;inset:0;background:#102d4390;z-index:999}.print-context{display:none}
@media(max-width:1080px){.wrap{padding-inline:24px}.filters{grid-template-columns:1fr 1fr}.filters .search-field{grid-column:1}.reset-button{justify-self:stretch}.hero-aside{max-width:200px}.highest{grid-template-columns:195px 1fr}.highest-number{font-size:43px}.table-controls{align-items:start;flex-direction:column;gap:10px}.kpi{padding:17px}.kpi-head{font-size:11px}}
@media(max-width:700px){.wrap{padding-inline:16px}.hero-inner{display:block;padding-block:22px 24px}.hero h1{font-size:23px}.department{font-size:11px}.hero-aside{border:0;padding:0;margin-top:20px;display:flex;align-items:center;justify-content:space-between;max-width:none;gap:12px}.hero-aside strong{font-size:13px;margin:0}.hero-aside .print-button{margin:0}.hero-aside span{font-size:10px}.period{font-size:11px}.section-header{gap:12px;align-items:flex-start}.section-title{font-size:18px}.scope-label{font-size:10px}.filters{gap:10px;padding:14px}.filters .search-field{grid-column:1/-1}.reset-button{grid-column:1/-1}.kpis{grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}.kpi{padding:15px 14px}.kpi-value{font-size:34px}.kpi-index{display:none}.kpi-note{font-size:10px}.highest{grid-template-columns:1fr}.highest-score{padding:17px 20px;display:grid;grid-template-columns:1fr auto;align-items:center;gap:0 14px}.highest-score h3{font-size:12px}.highest-number{grid-column:2;grid-row:1/3;font-size:38px}.highest-score p{margin:0}.highest-main{padding:18px}.highest-name{font-size:15px}.highest-meta{grid-template-columns:1fr;gap:11px}.highest-bottom{align-items:flex-start;flex-direction:column}.panel{padding:16px}.table-heading{gap:7px}.count-label{font-size:10px;padding:3px 8px}.table-filters{width:100%;gap:10px}.table-filters .field{flex:1;min-width:130px}.table-scope-note{font-size:10px}.pager{align-items:flex-start;gap:10px}.pager-actions{flex-shrink:0}.pager button{padding:5px 10px}.footer{flex-direction:column;font-size:10px}.dialog-head{padding:15px 18px}.dialog-content{padding:18px}.detail-summary{grid-template-columns:1fr}.employee-id{display:block;margin:3px 0 0}.tie-card{padding:13px;gap:10px}.tie-card .detail-button{font-size:10px}.method-body{font-size:11px}}
@media(prefers-reduced-motion:reduce){html{scroll-behavior:auto}}
@media print{@page{size:A4 landscape;margin:11mm}body{background:white;font-size:10px;color:#17334a;-webkit-print-color-adjust:exact;print-color-adjust:exact}.wrap{width:100%;padding:0}.hero{background:#123752!important}.hero-inner{padding:16px 20px;gap:16px}.hero h1{font-size:21px}.hero-aside{display:block;border:0;margin:0;padding:0;max-width:none}.hero-aside span{font-size:9px}.hero-aside strong{font-size:11px}.period{font-size:10px;margin-top:5px}.department{margin-bottom:3px;font-size:10px}.print-button,.filters,.table-controls,.pager,.detail-button,.link-button,dialog,.fallback-shade,.methodology{display:none!important}main{padding-block:16px}.section-header{margin-bottom:11px}.section-title{font-size:16px}.section-subtitle{font-size:10px}.kpis{gap:11px;margin-bottom:12px}.kpi{padding:11px 14px;box-shadow:none;break-inside:avoid}.kpi-value{font-size:29px;margin-top:3px}.kpi-head,.kpi-note{font-size:9px}.highest{grid-template-columns:170px 1fr;margin-bottom:17px;box-shadow:none;break-inside:avoid}.highest-score{padding:15px;display:flex;align-items:stretch}.highest-number{font-size:36px}.highest-main{padding:13px 18px}.highest-name{font-size:13px}.highest-meta{margin-top:8px;gap:6px 18px;grid-template-columns:1fr 1fr}.highest-bottom{margin-top:7px;padding-top:6px}.highest-bottom .subtle{font-size:9px}.info-value{font-size:10px}.highest-job,.org-path{font-size:9px}.panel{padding:13px;box-shadow:none}.table-scroll{overflow:visible;border:0;border-radius:0}table{font-size:9px;table-layout:fixed;overflow-wrap:anywhere}th{font-size:8px;white-space:normal;padding:8px 6px}td{padding:9px 6px}.manager-cell,.port-cell,.date-cell,.ratio-cell{min-width:0;max-width:none}.manager-cell{width:21%}.port-cell{width:26%}.date-cell{width:18%}.people-number,.ratio-value{font-size:10px}.cell-note{font-size:8px}.ratio-track{width:48px;height:3px}.badge{font-size:8px;padding:2px 5px}.people-cell{white-space:normal}.case-id{font-size:7px}tr{break-inside:avoid}thead{display:table-header-group}.screen-detail-col{display:none}.footer{font-size:9px;margin-top:10px}.print-context{display:block;font-size:10px;color:#526b77;margin:0 0 10px}.scope-label{font-size:9px}}

#smart-leave-calendar {
  color-scheme: light;
  --sp-bg: #f8f6f2;
  --sp-surface: #ffffff;
  --sp-ink: #152446;
  --sp-muted: #666b74;
  --sp-line: #dedbd4;
  --sp-soft: #f7f5f1;
  --sp-past: #faf9f6;
  --sp-weekend: #e9e8e5;
  --sp-teal: #156342;
  --sp-teal-soft: #e3eee8;
  --sp-alert: #794025;
  --sp-alert-soft: #edb08b;
  --sp-alert-line: #de9c75;
  --sp-focus: #00828e;
  direction: rtl;
  background: var(--sp-bg);
  color: var(--sp-ink);
  font-family: Tahoma, Arial, sans-serif;
  font-size: 14px;
  line-height: 1.65;
  padding: 24px;
  margin-top: 24px;
  border: 1px solid var(--sp-line);
  border-top: 3px solid #00828e;
  border-radius: 12px;
  scroll-margin-top: 20px;
  text-align: right;
  isolation: isolate;
}
#smart-leave-calendar *, #smart-leave-calendar *::before, #smart-leave-calendar *::after { box-sizing: border-box; }
#smart-leave-calendar [hidden] { display: none !important; }
#smart-leave-calendar button { font: inherit; color: inherit; cursor: pointer; }
#smart-leave-calendar button:focus-visible { outline: 3px solid var(--sp-focus); outline-offset: 3px; }
#smart-leave-calendar button:disabled { cursor: default; opacity: .4; }
#smart-leave-calendar h1, #smart-leave-calendar h2, #smart-leave-calendar h3, #smart-leave-calendar p { margin: 0; }
#smart-leave-calendar .sp-shell { max-width: 1440px; margin: 0 auto; }
#smart-leave-calendar .sp-header { display: flex; align-items: center; justify-content: space-between; gap: 20px; margin: 0 0 25px; }
#smart-leave-calendar .sp-heading { display: flex; align-items: flex-start; gap: 13px; }
#smart-leave-calendar .sp-mark { width: 39px; height: 40px; border: 2px solid var(--sp-teal); border-radius: 9px; position: relative; margin-top: 4px; flex: 0 0 auto; }
#smart-leave-calendar .sp-mark::before { content: ''; position: absolute; top: 9px; left: 0; right: 0; border-top: 2px solid var(--sp-teal); }
#smart-leave-calendar .sp-mark::after { content: ''; position: absolute; width: 8px; height: 8px; border-radius: 2px; background: var(--sp-teal); top: 18px; right: 7px; box-shadow: -13px 0 0 var(--sp-teal); }
#smart-leave-calendar .sp-heading h2 { font-size: 28px; letter-spacing: -.6px; font-weight: 700; line-height: 1.5; }
#smart-leave-calendar .sp-subtitle { color: var(--sp-muted); margin-top: 3px; font-size: 14px; }
#smart-leave-calendar .sp-demo { color: var(--sp-muted); background: var(--sp-surface); border: 1px solid var(--sp-line); font-size: 12px; padding: 7px 12px; border-radius: 7px; white-space: nowrap; }
#smart-leave-calendar .sp-workspace { display: grid; grid-template-columns: minmax(0, 1fr); gap: 18px; align-items: start; }
#smart-leave-calendar .sp-workspace.sp-open { grid-template-columns: minmax(0, 1fr) 350px; }
#smart-leave-calendar .sp-calendar { min-width: 0; background: var(--sp-surface); border: 1px solid var(--sp-line); border-radius: 12px; overflow: hidden; }
#smart-leave-calendar .sp-toolbar { display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 14px; padding: 20px 22px; }
#smart-leave-calendar .sp-month-nav { display: flex; align-items: center; gap: 12px; }
#smart-leave-calendar .sp-month { font-size: 22px; font-weight: 700; line-height: 1.4; min-width: 140px; }
#smart-leave-calendar .sp-nav-buttons { display: flex; gap: 5px; }
#smart-leave-calendar .sp-icon-button { background: var(--sp-surface); border: 1px solid var(--sp-line); border-radius: 7px; min-width: 36px; min-height: 36px; display: inline-grid; place-items: center; padding: 6px; }
#smart-leave-calendar .sp-icon-button:hover:not(:disabled) { background: var(--sp-soft); }
#smart-leave-calendar .sp-chevron { width: 8px; height: 8px; border-top: 1.7px solid currentColor; border-right: 1.7px solid currentColor; transform: rotate(45deg); }
#smart-leave-calendar .sp-chevron.sp-left { transform: rotate(-135deg); }
#smart-leave-calendar .sp-date-note { color: var(--sp-muted); font-size: 12px; }
#smart-leave-calendar .sp-compare { display: flex; background: var(--sp-soft); border: 1px solid var(--sp-line); padding: 3px; border-radius: 7px; gap: 3px; }
#smart-leave-calendar .sp-compare button { padding: 6px 12px; min-height: 33px; border: 0; background: transparent; border-radius: 4px; font-size: 12px; color: var(--sp-muted); }
#smart-leave-calendar .sp-compare button[aria-pressed="true"] { color: var(--sp-ink); background: var(--sp-surface); box-shadow: 0 1px 3px #0000000a; }
#smart-leave-calendar .sp-weekdays { display: grid; grid-template-columns: repeat(7, minmax(0, 1fr)); padding: 0; border-block: 1px solid var(--sp-line); color: var(--sp-muted); background: var(--sp-soft); font-size: 13px; text-align: center; }
#smart-leave-calendar .sp-week { position: relative; min-height: 105px; }
#smart-leave-calendar .sp-week + .sp-week { border-top: 1px solid var(--sp-line); }
#smart-leave-calendar .sp-day-grid { display: grid; grid-template-columns: repeat(7, minmax(0, 1fr)); position: absolute; inset: 0; }
#smart-leave-calendar .sp-day { padding: 9px 12px; background: var(--sp-surface); min-width: 0; }
#smart-leave-calendar .sp-day + .sp-day { border-right: 1px solid var(--sp-line); }
#smart-leave-calendar .sp-day.sp-past { background: var(--sp-past); }
#smart-leave-calendar .sp-day.sp-outside { background: var(--sp-past); color: var(--sp-muted); }
#smart-leave-calendar .sp-day.sp-past .sp-number { color: var(--sp-muted); }
#smart-leave-calendar .sp-day.sp-outside .sp-number { opacity: .5; }
#smart-leave-calendar .sp-number { font-family: Arial, Tahoma, sans-serif; font-size: 22px; font-weight: 400; display: inline-flex; align-items: center; justify-content: center; width: 30px; min-height: 30px; line-height: 1; }
#smart-leave-calendar .sp-asof .sp-number { color: var(--sp-teal) !important; border: 1.5px solid var(--sp-teal); border-radius: 50%; }
#smart-leave-calendar .sp-asof-label { color: var(--sp-teal); font-size: 11px; margin-inline-start: 3px; white-space: nowrap; }
#smart-leave-calendar .sp-events { display: grid; position: relative; grid-template-columns: repeat(7, minmax(0, 1fr)); grid-auto-rows: 43px; gap: 6px 0; padding: 46px 0 12px; min-height: 105px; }
#smart-leave-calendar .sp-event { display: flex; align-items: center; gap: 8px; text-align: right; border: 1px solid var(--sp-alert-line); background: var(--sp-alert-soft); color: var(--sp-alert); border-radius: 6px; margin: 0 5px; padding: 7px 10px; min-width: 0; transition: filter .16s, border-color .16s; }
#smart-leave-calendar .sp-event:hover { filter: brightness(.97); }
#smart-leave-calendar .sp-event.sp-selected { outline: 2px solid var(--sp-alert); outline-offset: -1px; }
#smart-leave-calendar .sp-event.sp-good { background: var(--sp-teal-soft); border-color: transparent; color: var(--sp-teal); }
#smart-leave-calendar .sp-event.sp-good.sp-selected { outline-color: var(--sp-teal); }
#smart-leave-calendar .sp-event-name { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-weight: 700; font-size: 15px; flex: 1; }
#smart-leave-calendar .sp-event-count { font-size: 12px; white-space: nowrap; font-weight: 400; }
#smart-leave-calendar .sp-event .sp-mobile-period { display: none; font-size: 11px; font-weight: 400; }
#smart-leave-calendar .sp-calendar-bottom { display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap; border-top: 1px solid var(--sp-line); padding: 13px 20px; color: var(--sp-muted); font-size: 12px; }
#smart-leave-calendar .sp-legend { display: flex; gap: 17px; flex-wrap: wrap; }
#smart-leave-calendar .sp-legend span { display: flex; align-items: center; gap: 6px; }
#smart-leave-calendar .sp-dot { display: inline-block; width: 7px; height: 7px; border-radius: 50%; background: var(--sp-alert); }
#smart-leave-calendar .sp-dot.sp-green { background: var(--sp-teal); }
#smart-leave-calendar .sp-zero { text-align: center; padding: 12px 20px; background: var(--sp-teal-soft); color: var(--sp-teal); font-size: 13px; }
#smart-leave-calendar .sp-panel { background: var(--sp-surface); border: 1px solid var(--sp-line); border-radius: 12px; min-width: 0; overflow: hidden; }
#smart-leave-calendar .sp-panel-top { padding: 22px 22px 17px; border-bottom: 1px solid var(--sp-line); }
#smart-leave-calendar .sp-panel-label-row { display: flex; align-items: center; justify-content: space-between; color: var(--sp-muted); font-size: 12px; margin-bottom: 11px; }
#smart-leave-calendar .sp-close { border: 0; background: transparent; font-size: 24px; line-height: 1; padding: 3px; min-width: 28px; min-height: 28px; border-radius: 5px; }
#smart-leave-calendar .sp-close:hover { background: var(--sp-soft); }
#smart-leave-calendar .sp-panel h2 { font-size: 22px; line-height: 1.6; }
#smart-leave-calendar .sp-port { color: var(--sp-muted); font-size: 12px; margin-top: 3px; }
#smart-leave-calendar .sp-panel-note { display: flex; align-items: center; gap: 7px; color: var(--sp-alert); font-size: 12px; padding-top: 13px; }
#smart-leave-calendar .sp-panel-note.sp-success { color: var(--sp-teal); }
#smart-leave-calendar .sp-panel-content { padding: 18px 22px 0; }
#smart-leave-calendar .sp-panel h3 { font-size: 13px; font-weight: 400; color: var(--sp-muted); }
#smart-leave-calendar .sp-employee { padding: 16px 0; border-bottom: 1px solid var(--sp-line); }
#smart-leave-calendar .sp-employee:last-child { border-bottom: none; }
#smart-leave-calendar .sp-employee-head { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin-bottom: 7px; }
#smart-leave-calendar .sp-employee-name { font-size: 16px; font-weight: 700; }
#smart-leave-calendar .sp-duration { color: var(--sp-muted); font-size: 12px; white-space: nowrap; }
#smart-leave-calendar .sp-period { font-size: 14px; font-variant-numeric: tabular-nums; }
#smart-leave-calendar .sp-old-period { color: var(--sp-muted); font-size: 12px; text-decoration: line-through; margin-bottom: 4px; }
#smart-leave-calendar .sp-new-period { color: var(--sp-teal); font-weight: 700; display: flex; align-items: center; justify-content: space-between; gap: 8px; }
#smart-leave-calendar .sp-tiny-label { color: var(--sp-teal); background: var(--sp-teal-soft); padding: 2px 6px; border-radius: 4px; font-size: 11px; font-weight: 400; white-space: nowrap; }
#smart-leave-calendar .sp-unchanged { color: var(--sp-muted); font-size: 11px; margin-top: 5px; }
#smart-leave-calendar .sp-fixed { margin-top: 13px; border-top: 1px solid var(--sp-line); padding: 13px 0; color: var(--sp-muted); font-size: 12px; }
#smart-leave-calendar .sp-fixed summary { cursor: pointer; }
#smart-leave-calendar .sp-fixed-list { padding: 8px 0 0; }
#smart-leave-calendar .sp-fixed-item { display: flex; align-items: start; justify-content: space-between; flex-wrap: wrap; gap: 4px 10px; padding-top: 8px; }
#smart-leave-calendar .sp-fixed-item b { font-weight: 400; color: var(--sp-ink); }
#smart-leave-calendar .sp-result { background: var(--sp-teal-soft); color: var(--sp-teal); border-radius: 7px; padding: 13px 14px; margin: 0 0 16px; font-size: 13px; }
#smart-leave-calendar .sp-result.sp-partial { background: var(--sp-alert-soft); color: var(--sp-alert); }
#smart-leave-calendar .sp-result b { display: block; font-size: 14px; margin-bottom: 4px; }
#smart-leave-calendar .sp-panel-footer { padding: 16px 22px 22px; }
#smart-leave-calendar .sp-solve { display: flex; align-items: center; justify-content: center; gap: 10px; width: 100%; background: #00828e; color: #fff; border: 0; border-radius: 7px; font-size: 14px; font-weight: 700; padding: 12px 15px; min-height: 47px; }
#smart-leave-calendar .sp-solve:hover:not(:disabled) { filter: brightness(.96); }
#smart-leave-calendar .sp-rule { text-align: center; font-size: 12px; color: var(--sp-muted); margin-top: 9px; line-height: 1.8; }
#smart-leave-calendar .sp-undo { display: block; background: transparent; border: 1px solid var(--sp-line); border-radius: 7px; min-height: 42px; width: 100%; padding: 9px 12px; font-size: 12px; }
#smart-leave-calendar .sp-undo:hover { background: var(--sp-soft); }
#smart-leave-calendar .sp-footer { padding: 14px 2px 0; color: var(--sp-muted); font-size: 11px; display: flex; justify-content: space-between; gap: 12px; flex-wrap: wrap; }
#smart-leave-calendar .sp-live { position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip: rect(0,0,0,0); white-space: nowrap; border: 0; }
#smart-leave-calendar .sp-event:only-child .sp-event-name { flex: 0 1 auto; }
@media (max-width: 1120px) {
  #smart-leave-calendar { padding: 24px 20px 20px; }
  #smart-leave-calendar .sp-workspace.sp-open { grid-template-columns: minmax(0, 1fr) 310px; gap: 12px; }
  #smart-leave-calendar .sp-panel-top, #smart-leave-calendar .sp-panel-content { padding-inline: 18px; }
  #smart-leave-calendar .sp-panel-footer { padding-inline: 18px; }
  #smart-leave-calendar .sp-open .sp-asof-label { display: none; }
  #smart-leave-calendar .sp-open .sp-event-count { display: none; }
  #smart-leave-calendar .sp-open .sp-event { margin-inline: 3px; padding-inline: 8px; }
  #smart-leave-calendar .sp-month { min-width: 118px; font-size: 20px; }
  #smart-leave-calendar .sp-toolbar { padding: 16px; }
}
@media (max-width: 840px) {
  #smart-leave-calendar .sp-workspace.sp-open { grid-template-columns: minmax(0, 1fr); }
  #smart-leave-calendar .sp-panel { grid-row: 1; }
  #smart-leave-calendar .sp-open .sp-event-count { display: inline; }
  #smart-leave-calendar .sp-panel-top { padding-block: 16px; }
  #smart-leave-calendar .sp-employee { padding-block: 12px; }
  #smart-leave-calendar .sp-header { align-items: start; }
}
@media (max-width: 540px) {
  #smart-leave-calendar { padding: 18px 10px 16px; }
  #smart-leave-calendar .sp-heading h2 { font-size: 23px; }
  #smart-leave-calendar .sp-mark { display: none; }
  #smart-leave-calendar .sp-header { gap: 12px; flex-wrap: wrap; margin-bottom: 16px; }
  #smart-leave-calendar .sp-demo { font-size: 11px; padding: 4px 7px; }
  #smart-leave-calendar .sp-subtitle { font-size: 12px; }
  #smart-leave-calendar .sp-toolbar { padding: 14px 12px; gap: 10px; }
  #smart-leave-calendar .sp-month { min-width: 107px; font-size: 19px; }
  #smart-leave-calendar .sp-month-nav { gap: 8px; }
  #smart-leave-calendar .sp-date-note { font-size: 11px; }
  #smart-leave-calendar .sp-compare { width: 100%; }
  #smart-leave-calendar .sp-compare button { flex: 1; }
  #smart-leave-calendar .sp-weekdays { font-size: 11px; padding-block: 9px; }
  #smart-leave-calendar .sp-day { padding: 7px 2px; text-align: center; }
  #smart-leave-calendar .sp-number { font-size: 18px; width: 27px; min-height: 27px; }
  #smart-leave-calendar .sp-asof-label { display: none; }
  #smart-leave-calendar .sp-events { grid-auto-rows: 44px; min-height: 72px; padding-top: 42px; }
  #smart-leave-calendar .sp-week { min-height: 72px; }
  #smart-leave-calendar .sp-event { grid-column: 1 / -1 !important; grid-row: auto !important; margin-inline: 7px; padding-inline: 11px; }
  #smart-leave-calendar .sp-event-name { font-size: 13px; }
  #smart-leave-calendar .sp-event .sp-mobile-period { display: inline; }
  #smart-leave-calendar .sp-event-count, #smart-leave-calendar .sp-open .sp-event-count { display: none; }
  #smart-leave-calendar .sp-calendar-bottom { padding: 12px; font-size: 11px; gap: 7px; }
  #smart-leave-calendar .sp-footer { font-size: 11px; }
}
@media (pointer: coarse) {
  #smart-leave-calendar .sp-icon-button, #smart-leave-calendar .sp-close { min-width: 44px; min-height: 44px; }
  #smart-leave-calendar .sp-compare button { min-height: 40px; }
}
@media (prefers-reduced-motion: reduce) {
  #smart-leave-calendar *, #smart-leave-calendar *::before, #smart-leave-calendar *::after { transition: none !important; scroll-behavior: auto !important; }
}

/* The weekend remains gray in every state; overlap is shown on the ribbon. */
#smart-leave-calendar .sp-weekdays > span { padding: 11px 2px; }
#smart-leave-calendar .sp-weekdays > span.sp-weekend,
#smart-leave-calendar .sp-day.sp-weekend,
#smart-leave-calendar .sp-day.sp-weekend.sp-past,
#smart-leave-calendar .sp-day.sp-weekend.sp-outside { background: var(--sp-weekend); }
#smart-leave-calendar .sp-header { background: transparent; color: var(--sp-ink); }
#smart-leave-calendar .sp-header::after { display: none; }
#smart-leave-calendar .sp-filterbar { display: flex; gap: 12px; flex-wrap: wrap; align-items: end; margin: 0 0 13px; }
#smart-leave-calendar .sp-field { display: grid; gap: 5px; flex: 1; min-width: 200px; }
#smart-leave-calendar .sp-field:first-child { flex: 1.5; }
#smart-leave-calendar .sp-field > span { font-size: 12px; color: var(--sp-muted); }
#smart-leave-calendar .sp-field input, #smart-leave-calendar .sp-field select { font: inherit; color: var(--sp-ink); background: white; border: 1px solid #cfc8bc; padding: 10px 12px; width: 100%; min-height: 44px; border-radius: 5px; }
#smart-leave-calendar input:focus-visible, #smart-leave-calendar select:focus-visible { outline: 3px solid var(--sp-focus); outline-offset: 3px; }
#smart-leave-calendar .sp-reset { background: #fff; border: 1px solid #cfc8bc; padding: 9px 14px; min-height: 44px; border-radius: 5px; }
#smart-leave-calendar .sp-countline { display: flex; align-items: center; gap: 10px 20px; flex-wrap: wrap; margin-bottom: 16px; font-size: 12px; color: var(--sp-muted); }
#smart-leave-calendar .sp-countline strong { color: var(--sp-ink); font-size: 21px; margin-inline: 4px; }
#smart-leave-calendar .sp-countline .sp-count-note { flex: 1 1 100%; font-size: 11px; }
#smart-leave-calendar .sp-org { margin-top: 5px; color: var(--sp-muted); font-size: 11px; line-height: 1.7; overflow-wrap: anywhere; }
#smart-leave-calendar .sp-original-cases { margin: 14px 0 2px; font-size: 11px; color: var(--sp-muted); border-top: 1px solid var(--sp-line); padding-top: 10px; }
#smart-leave-calendar .sp-original-cases summary { cursor: pointer; color: var(--sp-ink); line-height: 1.9; }
#smart-leave-calendar .sp-case-item { padding: 8px 0; border-bottom: 1px solid var(--sp-line); }
#smart-leave-calendar .sp-case-item:last-child { border: 0; }
#smart-leave-calendar .sp-case-item strong { color: var(--sp-ink); }
#smart-leave-calendar .sp-locked { color: #87643e; font-size: 11px; margin-top: 5px; }
#smart-leave-calendar .sp-employee-id { display: block; font-size: 10px; color: var(--sp-muted); font-weight: 400; }
#smart-leave-calendar .sp-weekdays .sp-weekend-label { background: var(--sp-weekend); }
#smart-leave-calendar .sp-dot.sp-gray { background: #a5a4a1; }
#smart-leave-calendar .sp-event { cursor: pointer; }
#smart-leave-calendar .sp-panel { max-height: 900px; overflow-y: auto; }
#smart-leave-calendar #sp-manager-title { font-size: 21px; font-weight: 700; color: var(--sp-ink); line-height: 1.6; }
#smart-leave-calendar .sp-panel-footer { background: #fff; }
#smart-leave-calendar .sp-solve + .sp-undo { margin-top: 8px; }
#smart-leave-calendar .sp-footer { background: none; border: 0; margin: 0; }
#smart-leave-calendar .sp-footer::before { display: none; }
@media (max-width: 540px) {
  #smart-leave-calendar .sp-heading h2 { font-size: 23px; }
  #smart-leave-calendar .sp-field { min-width: 100%; }
  #smart-leave-calendar .sp-reset { min-width: 100%; }
  #smart-leave-calendar .sp-weekdays { padding: 0; }
  #smart-leave-calendar .sp-weekdays > span { padding-block: 9px; font-size: 10px; }
  #smart-leave-calendar .sp-countline strong { font-size: 18px; }
  #smart-leave-calendar .sp-panel { max-height: none; }
}
@media print {
  #smart-leave-calendar { break-before: page; border-radius: 0; margin-top: 15px; padding: 14px; }
  #smart-leave-calendar .sp-filterbar, #smart-leave-calendar .sp-nav-buttons,
  #smart-leave-calendar .sp-compare, #smart-leave-calendar .sp-panel-footer,
  #smart-leave-calendar .sp-close { display: none !important; }
  #smart-leave-calendar .sp-workspace.sp-open { grid-template-columns: minmax(0,1fr); }
  #smart-leave-calendar .sp-panel { max-height: none; overflow: visible; }
  #smart-leave-calendar .sp-week { break-inside: avoid; }
}

.integration-demo-notice{padding:12px 18px;margin:0 0 18px;border:1px solid #ddbd87;background:#fff5e3;color:#70552e;border-radius:8px;font-size:13px}.planner-jump{display:inline-block;margin-top:10px;color:#d5eeeb;font-size:12px;text-decoration:underline;text-underline-offset:4px}
</style>
</head>
<body>
<header class="hero">
  <div class="wrap hero-inner">
    <div><p class="department" id="department"></p><h1 id="report-title">لوحة تقاطع الإجازات</h1><div class="period" id="report-period"></div></div>
    <aside class="hero-aside"><div><span>تاريخ تصنيف الحالات</span><strong id="report-asof"></strong></div><button type="button" class="print-button" id="print-report">طباعة العرض الحالي</button><br><a class="planner-jump" href="#smart-leave-calendar">المنسّق الذكي</a></aside>
  </div>
</header>
<main class="wrap">
  <div class="integration-demo-notice" id="integration-demo-notice" hidden></div>
  <noscript><div class="noscript">يحتاج عرض المؤشرات والجدول إلى تفعيل JavaScript في المتصفح. جميع بيانات التقرير مضمّنة في هذا الملف.</div></noscript>
  <section aria-labelledby="overview-title">
    <div class="section-header"><div><h2 class="section-title" id="overview-title">نظرة عامة على تقاطع الإجازات</h2><p class="section-subtitle">جميع فترات التقاطع، مع إبراز الحالات التي تصل إلى حد النسبة المحدد.</p></div><span class="scope-label" id="scope-label">كامل التقرير</span></div>
    <div class="filters" aria-label="تصفية التقرير حسب المجموعة">
      <div class="field"><label for="manager-filter">المدير</label><select id="manager-filter"></select></div>
      <div class="field"><label for="port-filter">موقع العمل</label><select id="port-filter"></select></div>
      <div class="field search-field"><label for="group-search">البحث في المجموعات</label><input type="search" id="group-search" placeholder="اسم المدير، الموقع، أو أحد المستويات التنظيمية"></div>
      <button class="reset-button" id="reset-filters" type="button">إعادة ضبط</button>
    </div>
    <div class="kpis" aria-label="مؤشرات حالات التقاطع" aria-live="polite">
      <article class="kpi"><div class="kpi-head"><span>جميع حالات التقاطع</span><span class="kpi-index" aria-hidden="true">01</span></div><div class="kpi-value" id="kpi-all">—</div><div class="kpi-note">بغض النظر عن نسبة الغياب</div></article>
      <article class="kpi"><div class="kpi-head"><span id="threshold-kpi-label">تقاطع 50% فأكثر</span><span class="kpi-index" aria-hidden="true">02</span></div><div class="kpi-value" id="kpi-threshold">—</div><div class="kpi-note">جزء من جميع حالات التقاطع</div></article>
      <article class="kpi"><div class="kpi-head"><span>الحالات الجارية</span><span class="kpi-index" aria-hidden="true">03</span></div><div class="kpi-value" id="kpi-ongoing">—</div><div class="kpi-note">من جميع الحالات، حسب تاريخ التصنيف</div></article>
      <article class="kpi"><div class="kpi-head"><span>الحالات القادمة</span><span class="kpi-index" aria-hidden="true">04</span></div><div class="kpi-value" id="kpi-upcoming">—</div><div class="kpi-note">من جميع الحالات، بعد تاريخ التصنيف</div></article>
    </div>
    <section class="highest" aria-labelledby="highest-title"><div class="highest-score"><h3 id="highest-title">أعلى نسبة غياب مجدول</h3><div class="highest-number" id="highest-number">—</div><p id="highest-ratio"></p></div><div id="highest-details" class="highest-main"></div></section>
  </section>
  <section class="panel register" aria-labelledby="register-title">
    <div class="section-header"><div class="table-heading"><h2 class="section-title" id="register-title">سجل حالات التقاطع</h2><span id="table-count" class="count-label"></span></div></div>
    <div class="table-controls">
      <div class="table-filters"><div class="field"><label for="table-threshold">نسبة التقاطع</label><select id="table-threshold"><option value="all">جميع النسب</option><option value="threshold" id="threshold-option">50% فأكثر فقط</option></select></div><div class="field"><label for="table-status">حالة الفترة</label><select id="table-status"><option value="all">جميع الحالات</option><option value="جارية">جارية</option><option value="قادمة">قادمة</option><option value="منتهية">منتهية</option></select></div></div>
      <p class="table-scope-note">فلتر النسبة والحالة يغيّر السجل فقط. المؤشرات أعلاه تشمل جميع الحالات ضمن تصفية المجموعة.</p>
    </div>
    <p class="print-context" id="print-context"></p>
    <div class="table-scroll" tabindex="0" aria-label="جدول حالات التقاطع؛ يمكن التمرير أفقياً">
      <table><thead><tr><th class="manager-cell" scope="col">المدير / المسمى الوظيفي</th><th class="port-cell" scope="col">الموقع / المسار التنظيمي</th><th scope="col">المجازون / المجموعة</th><th class="ratio-cell" scope="col">نسبة الغياب</th><th class="date-cell" scope="col">فترة التقاطع</th><th scope="col">الحالة</th><th scope="col" class="screen-detail-col">الموظفون</th></tr></thead><tbody id="case-rows"></tbody></table>
    </div>
    <div class="pager"><span id="page-info" aria-live="polite"></span><div class="pager-actions"><button id="previous-page" type="button">السابق</button><button id="next-page" type="button">التالي</button></div></div>
  </section>
  <section id="smart-leave-calendar" lang="ar" dir="rtl" aria-labelledby="sp-title">
  <div class="sp-shell">
    <header class="sp-header">
      <div class="sp-heading"><span class="sp-mark" aria-hidden="true"></span><div><h2 id="sp-title">المنسّق الذكي</h2><p class="sp-subtitle">اضغطي اسم المدير لعرض الموظفين وتنسيق إجازاتهم.</p></div></div>
      <span class="sp-demo" data-ui="demo-label" hidden></span>
    </header>
    <div class="sp-filterbar" data-ui="filterbar" aria-label="تصفية المنسّق الذكي">
      <label class="sp-field"><span>اسم المدير</span><input data-ui="manager-search" id="sp-manager-search" type="search" placeholder="ابحثي باسم المدير" list="sp-manager-names" autocomplete="off"></label>
      <datalist id="sp-manager-names" data-ui="manager-names"></datalist>
      <label class="sp-field"><span>نسبة التقاطع</span><select data-ui="threshold" id="sp-threshold"><option value="0">جميع حالات التقاطع</option><option value="50">تقاطع 50% فأكثر</option></select></label>
      <button type="button" class="sp-reset" data-action="reset">إعادة ضبط الفلتر</button>
    </div>
    <div class="sp-countline" aria-live="polite"><span>الحالات المستقبلية الأصلية: <strong data-ui="original-count">0</strong></span><span data-ui="scope-groups"></span><span data-ui="remaining-count" hidden></span><span class="sp-count-note" data-ui="scope-note"></span></div>
    <div class="sp-workspace" data-ui="workspace">
      <section class="sp-calendar" aria-label="تقويم الإجازات">
        <div class="sp-toolbar">
          <div class="sp-month-nav"><h3 class="sp-month" data-ui="month-label"></h3><div class="sp-nav-buttons"><button type="button" class="sp-icon-button" data-action="previous" aria-label="الشهر السابق"><span class="sp-chevron" aria-hidden="true"></span></button><button type="button" class="sp-icon-button" data-action="next" aria-label="الشهر التالي"><span class="sp-chevron sp-left" aria-hidden="true"></span></button></div></div>
          <span class="sp-date-note" data-ui="date-note"></span>
          <div class="sp-compare" data-ui="compare" aria-label="مقارنة المواعيد" hidden><button type="button" data-action="before" aria-pressed="false">قبل التعديل</button><button type="button" data-action="after" aria-pressed="true">بعد التعديل</button></div>
        </div>
        <div class="sp-weekdays" aria-hidden="true"><span>الأحد</span><span>الإثنين</span><span>الثلاثاء</span><span>الأربعاء</span><span>الخميس</span><span class="sp-weekend">الجمعة</span><span class="sp-weekend">السبت</span></div>
        <div data-ui="weeks"></div>
        <div class="sp-zero" data-ui="empty" role="status" hidden></div>
        <div class="sp-calendar-bottom"><div class="sp-legend"><span><i class="sp-dot" aria-hidden="true"></i> تقاطع إجازات</span><span><i class="sp-dot sp-green" aria-hidden="true"></i> موعد بعد التنسيق</span><span><i class="sp-dot sp-gray" aria-hidden="true"></i> نهاية الأسبوع</span></div><span data-ui="month-count"></span></div>
      </section>
      <aside class="sp-panel" data-ui="panel" aria-labelledby="sp-manager-title" hidden>
        <div class="sp-panel-top"><div class="sp-panel-label-row"><span>إجازات مجموعة المدير والموقع</span><button class="sp-close" type="button" data-action="close" aria-label="إغلاق تفاصيل المدير">×</button></div><h3 id="sp-manager-title" data-ui="manager-title"></h3><p class="sp-port" data-ui="manager-port"></p><p class="sp-org" data-ui="manager-org"></p><div class="sp-panel-note" data-ui="panel-note"></div><details class="sp-original-cases"><summary data-ui="case-summary"></summary><div data-ui="case-list"></div></details></div>
        <div class="sp-panel-content"><div class="sp-result" data-ui="result" role="status" hidden></div><h3 data-ui="employee-heading">الموظفون في التقاطع</h3><div data-ui="employees"></div><details class="sp-fixed" data-ui="fixed"><summary data-ui="fixed-label"></summary><div class="sp-fixed-list" data-ui="fixed-list"></div></details></div>
        <div class="sp-panel-footer"><button type="button" class="sp-solve" data-action="solve">تنسيق التقاطع</button><button type="button" class="sp-undo" data-action="undo" hidden>استعادة المواعيد الأصلية</button><p class="sp-rule" data-ui="rule">نفس مدة الإجازة، مع مراعاة إجازات باقي الموظفين.</p></div>
      </aside>
    </div>
    <footer class="sp-footer"><span data-ui="footer-period"></span><span>الجمعة والسبت ضمن مدة الإجازة. المقترحات داخل الصفحة ولا تعدّل ملفات Excel.</span></footer>
    <div class="sp-live" data-ui="live" aria-live="polite" aria-atomic="true"></div>
    <noscript>يلزم تفعيل JavaScript لعرض المنسّق الذكي.</noscript>
  </div>
</section>

  <details class="methodology" id="methodology"><summary>كيف تُحتسب الحالة والنسبة؟</summary><div class="method-body"><p>الحالة هي إجازة <strong>الموظفين أنفسهم، وعددهم اثنان فأكثر، في أيام تقويمية متصلة</strong>. عند تغيّر أي موظف في المجموعة المتقاطعة، أو انقطاع الأيام، تبدأ حالة جديدة.</p><p>تُحسب كل مجموعة بصورة مستقلة حسب <strong>المدير المباشر + موقع العمل + المستويات التنظيمية الأول والثاني والثالث</strong>. نسبة الغياب = عدد المجازين معاً ÷ إجمالي موظفي المجموعة نفسها × 100.</p><p id="method-dates"></p><p>الأسماء والمسميات التنظيمية معروضة كما وردت في المصادر. النسبة المعروضة مقربة إلى منزلة عشرية، ويُطبق حد النسبة على القيمة قبل التقريب.</p><div class="qa-tiles" id="qa-tiles"></div><details class="qa-issues" id="qa-issues" hidden><summary id="qa-summary">ملاحظات معالجة البيانات</summary><ul id="qa-list"></ul></details></div></details>
  <footer class="footer"><div class="footer-credit"><strong id="prepared-by"></strong><strong id="supervised-by"></strong></div><span id="footer-note"></span></footer>
</main>
<div class="fallback-shade" id="fallback-shade" hidden></div>
<dialog id="details-dialog" aria-labelledby="dialog-title"><div class="dialog-head"><h2 id="dialog-title">تفاصيل حالة التقاطع</h2><button type="button" class="close-button" id="close-dialog" aria-label="إغلاق التفاصيل">×</button></div><div class="dialog-content" id="dialog-content"></div></dialog>
<script id="report-data" type="application/json">__REPORT_JSON__</script>
<script>
'use strict';
const DATA = JSON.parse(document.getElementById('report-data').textContent);
const $ = id => document.getElementById(id);
const G = new Map(DATA.groups.map(g => [String(g.id), g]));
const CASES = new Map(DATA.cases.map(c => [String(c.id), c]));
const PAGE_SIZE = 50;
let page = 0, printing = false, previousFocus = null;
const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const raw = value => `<bdi>${esc(value || 'غير محدد')}</bdi>`;
const norm = value => String(value ?? '').normalize('NFKC').replace(/[\u064b-\u065f\u0670\u0640\u200e\u200f\u200b]/g,'').replace(/[أإآ]/g,'ا').replace(/ى/g,'ي').replace(/ة/g,'ه').replace(/\s+/g,' ').trim().toLocaleLowerCase();
const number = value => new Intl.NumberFormat('en-US', {maximumFractionDigits:1}).format(Number(value));
const pct = c => number(c.percent_display ?? c.percent) + '%';
const dateFormatter = new Intl.DateTimeFormat('ar-SA-u-ca-gregory-nu-latn', {day:'numeric',month:'short',year:'numeric',timeZone:'UTC'});
const fmt = value => value ? dateFormatter.format(new Date(value+'T00:00:00Z')) : 'غير متاح';
const dateRange = c => c.start === c.end ? fmt(c.start) : `${fmt(c.start)} — ${fmt(c.end)}`;
const groupOf = c => G.get(String(c.group_id));
const stateClass = status => ({'جارية':'running','قادمة':'future','منتهية':'ended'}[status] || '');
const badge = status => `<span class="badge ${stateClass(status)}">${esc(status)}</span>`;
const pathHtml = g => `<span class="org-path">${[g.level1,g.level2,g.level3].map(raw).join('')}</span>`;
const threshold = number(DATA.meta.threshold);

function scopeActive(){return Boolean($('manager-filter').value || $('port-filter').value || $('group-search').value.trim());}
function groupMatches(g){
  if($('manager-filter').value && String(g.manager_id) !== $('manager-filter').value)return false;
  if($('port-filter').value && g.port !== $('port-filter').value)return false;
  const query=norm($('group-search').value);
  return !query || norm([g.manager,g.port,g.level1,g.level2,g.level3].join(' ')).includes(query);
}
function getScopeCases(){return DATA.cases.filter(c => groupMatches(groupOf(c)));}
function getTableCases(){return getScopeCases().filter(c => ($('table-threshold').value !== 'threshold' || c.flag) && ($('table-status').value === 'all' || c.status === $('table-status').value));}
function getHighestCases(rows){
  if(!rows.length)return [];
  let best=rows[0];
  for(const c of rows){if(c.people*best.workforce > best.people*c.workforce)best=c;}
  return rows.filter(c => c.people*best.workforce === best.people*c.workforce).sort((a,b) => a.start.localeCompare(b.start) || a.end.localeCompare(b.end) || String(groupOf(a).manager).localeCompare(String(groupOf(b).manager),'ar') || String(a.id).localeCompare(String(b.id)));
}
function renderHighest(rows){
  const ties=getHighestCases(rows),c=ties[0];
  if(!c){$('highest-number').textContent='—';$('highest-ratio').textContent='لا توجد حالة لاحتساب النسبة';$('highest-details').innerHTML='<div class="empty-highest">لا توجد حالات تقاطع ضمن التصفية الحالية.</div>';return;}
  const g=groupOf(c);
  $('highest-number').textContent=pct(c);
  $('highest-ratio').textContent=`${number(c.people)} مجازين من ${number(c.workforce)} موظفاً`;
  $('highest-details').innerHTML=`<div class="highest-top"><div><div class="highest-name">${raw(g.manager)}</div><div class="highest-job">${raw(g.title)}</div></div>${badge(c.status)}</div><div class="highest-meta"><div><span class="info-label">موقع العمل / المجموعة التنظيمية</span><div class="info-value">${raw(g.port)}</div>${pathHtml(g)}</div><div><span class="info-label">فترة التقاطع</span><div class="info-value">${esc(dateRange(c))}</div><div class="highest-job">${number(c.days)} أيام متصلة بنفس الموظفين</div></div></div><div class="highest-bottom"><p class="subtle">${ties.length>1?`تكررت أعلى نسبة في ${number(ties.length)} حالات؛ المعروضة هي الأسبق تاريخاً.`:'هذه أعلى نسبة بين جميع حالات التقاطع ضمن النطاق المعروض.'}</p><button type="button" class="link-button" ${ties.length>1?'id="show-highest-ties"':`data-case-detail="${esc(c.id)}"`}>${ties.length>1?`عرض جميع الحالات الأعلى (${number(ties.length)})`:'عرض الموظفين والتفاصيل'}</button></div>`;
  if(ties.length>1)$('show-highest-ties').onclick=()=>showHighest(ties);
}
function renderTable(){
  const rows=getTableCases(),scopeRows=getScopeCases();
  page=Math.min(page,Math.max(0,Math.ceil(rows.length/PAGE_SIZE)-1));
  const shown=printing?rows:rows.slice(page*PAGE_SIZE,(page+1)*PAGE_SIZE);
  $('table-count').textContent=number(rows.length)+' حالة';
  $('case-rows').innerHTML=shown.map(c=>{
    const g=groupOf(c);
    return `<tr class="${c.status==='منتهية'?'ended-row':''}"><td class="manager-cell"><div class="cell-title">${raw(g.manager)}</div><div class="cell-note">${raw(g.title)}</div><span class="case-id">${esc(c.id)}</span></td><td class="port-cell"><div class="cell-title">${raw(g.port)}</div><div class="cell-note">${pathHtml(g)}</div></td><td class="people-cell"><span class="people-number" dir="ltr">${number(c.people)} / ${number(c.workforce)}</span><div class="cell-note">المتبقون: ${number(c.remaining)}</div></td><td class="ratio-cell"><span class="ratio-value ${c.flag?'over':''}" dir="ltr">${pct(c)}</span><div class="ratio-track" aria-hidden="true"><span class="${c.flag?'over':''}" style="width:${Math.max(0,Math.min(100,Number(c.percent)))}%"></span></div></td><td class="date-cell"><div>${esc(dateRange(c))}</div><div class="cell-note">${number(c.days)} أيام متصلة</div></td><td>${badge(c.status)}</td><td class="screen-detail-col"><button type="button" class="detail-button" data-case-detail="${esc(c.id)}" aria-label="تفاصيل الحالة ${esc(c.id)} للمدير ${esc(g.manager)}">التفاصيل</button></td></tr>`;
  }).join('') || '<tr><td colspan="7" class="empty-cell">لا توجد حالات مطابقة للتصفية الحالية.</td></tr>';
  $('page-info').textContent=rows.length ? `${number(page*PAGE_SIZE+1)}–${number(Math.min(rows.length,(page+1)*PAGE_SIZE))} من ${number(rows.length)} حالة` : 'لا توجد حالات لعرضها';
  $('previous-page').disabled=page===0;
  $('next-page').disabled=(page+1)*PAGE_SIZE>=rows.length;
  const parts=[];
  if($('manager-filter').value)parts.push('المدير: '+$('manager-filter').selectedOptions[0].textContent);
  if($('port-filter').value)parts.push('الموقع: '+$('port-filter').value);
  if($('group-search').value.trim())parts.push('بحث المجموعة: '+$('group-search').value.trim());
  if($('table-threshold').value==='threshold')parts.push(threshold+'% فأكثر');
  if($('table-status').value!=='all')parts.push('الحالة: '+$('table-status').value);
  $('print-context').textContent=`السجل المعروض: ${number(rows.length)} من ${number(scopeRows.length)} حالة ضمن نطاق المجموعة.`+(parts.length?' '+parts.join(' · '):' جميع النسب والحالات.');
}
function render(){
  const rows=getScopeCases();
  $('scope-label').textContent=scopeActive()?'ضمن التصفية':'كامل التقرير';
  $('kpi-all').textContent=number(rows.length);
  $('kpi-threshold').textContent=number(rows.filter(c=>c.flag).length);
  $('kpi-ongoing').textContent=number(rows.filter(c=>c.status==='جارية').length);
  $('kpi-upcoming').textContent=number(rows.filter(c=>c.status==='قادمة').length);
  renderHighest(rows);renderTable();
}
function openDialog(title,html){
  const dialog=$('details-dialog');
  if(!dialog.hasAttribute('open')){previousFocus=document.activeElement;}
  $('dialog-title').textContent=title;$('dialog-content').innerHTML=html;
  if(!dialog.hasAttribute('open')){if(typeof dialog.showModal==='function')dialog.showModal();else{dialog.setAttribute('open','');dialog.classList.add('dialog-fallback');$('fallback-shade').hidden=false;}}
  dialog.scrollTop=0;$('close-dialog').focus();
}
function closeDialog(){
  const dialog=$('details-dialog');
  if(typeof dialog.close==='function')dialog.close();else dialog.removeAttribute('open');
  dialog.classList.remove('dialog-fallback');$('fallback-shade').hidden=true;
  if(previousFocus && document.contains(previousFocus))previousFocus.focus();
}
function showCase(id){
  const c=CASES.get(String(id));if(!c)return;const g=groupOf(c);
  const employees=c.employees||[];
  const summary=`<div class="detail-summary"><div><span class="info-label">اسم المدير</span><strong>${raw(g.manager)}</strong><div class="cell-note">${raw(g.title)}</div></div><div><span class="info-label">موقع العمل</span><strong>${raw(g.port)}</strong></div><div class="full"><span class="info-label">المستوى الأول / الثاني / الثالث</span>${pathHtml(g)}</div><div><span class="info-label">فترة التقاطع</span><strong>${esc(dateRange(c))}</strong><div class="cell-note">${number(c.days)} أيام متصلة بنفس الموظفين</div></div><div><span class="info-label">حالة الفترة بتاريخ ${esc(fmt(DATA.meta.as_of))}</span>${badge(c.status)}<div class="cell-note">${esc(c.action||'')}</div></div></div><div class="detail-metrics"><span>المجازون: <b>${number(c.people)}</b></span><span>إجمالي المجموعة: <b>${number(c.workforce)}</b></span><span>نسبة الغياب: <b dir="ltr">${pct(c)}</b></span><span>المتبقون: <b>${number(c.remaining)}</b></span></div>`;
  const list=employees.map((e,i)=>`<li class="employee-item"><span class="employee-index">${i+1}</span><div><div class="employee-name">${raw(e.name||'اسم غير متوفر في المصدر')}<span class="employee-id">الرقم الوظيفي: <bdi>${esc(e.id)}</bdi></span></div>${(e.leave_periods||[]).length?(e.leave_periods||[]).map(p=>`<div class="leave-period">من ${esc(fmt(p.start))} إلى ${esc(fmt(p.end))}</div>`).join(''):'<div class="leave-period">تواريخ الإجازة الأصلية غير متاحة.</div>'}</div></li>`).join('');
  openDialog('تفاصيل حالة التقاطع · '+c.id,summary+`<h3 class="employee-title">الموظفون المشاركون في الحالة (${number(c.people)})</h3><p class="employee-note">الفترات أدناه هي تواريخ الإجازات الأصلية في المصدر، وقد تكون أطول من فترة التقاطع.</p><ol class="employee-list">${list||'<li class="subtle">تفاصيل الموظفين غير متاحة.</li>'}</ol>`);
}
function showHighest(ties){
  openDialog('الحالات ذات أعلى نسبة غياب',`<p class="tie-intro">${number(ties.length)} حالات تشترك في أعلى نسبة (${pct(ties[0])}) ضمن نطاق المجموعة المعروض. المقارنة على النسبة قبل التقريب.</p>`+ties.map(c=>{const g=groupOf(c);return `<article class="tie-card"><div><div class="cell-title">${raw(g.manager)}</div><div class="cell-note">${raw(g.port)}</div><div>${pathHtml(g)}</div><div class="cell-note">${esc(dateRange(c))} · ${number(c.people)} من ${number(c.workforce)} · ${number(c.days)} أيام</div><div style="margin-top:7px">${badge(c.status)}</div></div><button type="button" class="detail-button" data-case-detail="${esc(c.id)}">عرض الموظفين</button></article>`}).join(''));
}
function resetFilters(){
  $('manager-filter').value='';$('port-filter').value='';$('group-search').value='';$('table-threshold').value='all';$('table-status').value='all';page=0;render();
}
function initialize(){
  document.title='لوحة تقاطع الإجازات';
  $('report-title').textContent=String(DATA.meta.title||'لوحة تقاطع الإجازات').replace(/^تقرير /,'لوحة ');
  $('department').textContent=DATA.meta.department||'';
  $('report-period').textContent='فترة التقرير: '+fmt(DATA.meta.period_start)+' — '+fmt(DATA.meta.period_end);
  $('report-asof').textContent=fmt(DATA.meta.as_of);
  $('threshold-kpi-label').textContent='تقاطع '+threshold+'% فأكثر';
  $('threshold-option').textContent=threshold+'% فأكثر فقط';
  $('prepared-by').textContent=DATA.meta.prepared_by||'';$('supervised-by').textContent=DATA.meta.supervised_by||'';
  $('footer-note').textContent='يعرض '+number(DATA.cases.length)+' حالة تقاطع في فترة التقرير';
  $('method-dates').textContent='تُصنف الحالات بتاريخ '+fmt(DATA.meta.as_of)+': جارية عندما يقع هذا التاريخ داخل فترة التقاطع، وقادمة عندما تبدأ بعده، ومنتهية عندما تنتهي قبله. '+(DATA.meta.end_date_inclusive?'يُحتسب يوم نهاية الإجازة ضمن أيام الغياب.':'لا يُحتسب يوم نهاية الإجازة ضمن أيام الغياب.')+' تُحصر حالات التقاطع في فترة التقرير.';
  const managers=new Map();for(const g of DATA.groups){if(!managers.has(String(g.manager_id)))managers.set(String(g.manager_id),g.manager);}
  $('manager-filter').innerHTML='<option value="">جميع المدراء</option>'+[...managers].sort((a,b)=>String(a[1]).localeCompare(String(b[1]),'ar')||a[0].localeCompare(b[0])).map(([id,name])=>`<option value="${esc(id)}">${esc(name)} · ${esc(id)}</option>`).join('');
  $('port-filter').innerHTML='<option value="">جميع مواقع العمل</option>'+[...new Set(DATA.groups.map(g=>g.port))].sort((a,b)=>String(a).localeCompare(String(b),'ar')).map(port=>`<option value="${esc(port)}">${esc(port)}</option>`).join('');
  const qa=DATA.qa||{};
  $('qa-tiles').innerHTML=[['موظفون ضمن المجموعات',qa.employees],['مجموعات تنظيمية',qa.groups],['سجلات إجازات مطابقة',qa.matched_leaves]].filter(([,v])=>v!==undefined).map(([label,value])=>`<span><b>${number(value)}</b>${esc(label)}</span>`).join('');
  const problems=(qa.problems||[]).filter(p=>Number(p.count)>0);
  $('qa-issues').hidden=!problems.length;
  $('qa-summary').textContent='ملاحظات معالجة البيانات ('+number(problems.length)+' أنواع)';
  $('qa-list').innerHTML=problems.map(p=>`<li>${esc(p.label)}: <b>${number(p.count)}</b></li>`).join('');
  for(const id of ['manager-filter','port-filter'])$(id).addEventListener('change',()=>{page=0;render();});
  $('group-search').addEventListener('input',()=>{page=0;render();});
  for(const id of ['table-threshold','table-status'])$(id).addEventListener('change',()=>{page=0;renderTable();});
  $('reset-filters').addEventListener('click',resetFilters);
  $('previous-page').addEventListener('click',()=>{if(page>0){page--;renderTable();}});
  $('next-page').addEventListener('click',()=>{if((page+1)*PAGE_SIZE<getTableCases().length){page++;renderTable();}});
  $('print-report').addEventListener('click',()=>window.print());
  window.addEventListener('beforeprint',()=>{printing=true;renderTable();});
  window.addEventListener('afterprint',()=>{printing=false;renderTable();});
  $('close-dialog').addEventListener('click',closeDialog);$('fallback-shade').addEventListener('click',closeDialog);
  $('details-dialog').addEventListener('cancel',()=>{$('fallback-shade').hidden=true;});
  document.addEventListener('click',event=>{const button=event.target.closest('[data-case-detail]');if(button)showCase(button.dataset.caseDetail);});
  document.addEventListener('keydown',event=>{if(event.key==='Escape' && $('details-dialog').hasAttribute('open')){closeDialog();}});
  render();
}
window.Dashboard={render,renderTable,getScopeCases,getTableCases,getHighestCases,showCase,resetFilters,getPage:()=>page,getData:()=>DATA};
initialize();
</script>
<script>
(function (root) {
  'use strict';
  const DAY = 86400000;
  const copy = value => JSON.parse(JSON.stringify(value));
  const day = value => Date.parse(value + 'T00:00:00Z') / DAY;
  const iso = value => new Date(value * DAY).toISOString().slice(0, 10);
  const validDate = value => typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value) &&
    Number.isFinite(day(value)) && iso(day(value)) === value;
  const deepFreeze = value => {
    if (value && typeof value === 'object' && !Object.isFrozen(value)) {
      Object.freeze(value);
      Object.values(value).forEach(deepFreeze);
    }
    return value;
  };

  function createPlannerModel(input) {
    const initial = copy(input || {});
    for (const field of ['presentation_date', 'period_start', 'period_end']) {
      if (!validDate(initial[field])) throw new Error('تاريخ إعداد المنسّق غير صالح: ' + field);
    }
    if (initial.period_start > initial.period_end) throw new Error('فترة التقرير غير صالحة.');
    if (!Array.isArray(initial.groups) || !Array.isArray(initial.leaves) || !Array.isArray(initial.cases)) {
      throw new Error('بيانات المنسّق غير مكتملة.');
    }
    const PRESENTATION = initial.presentation_date;
    const FIRST = iso(Math.max(day(initial.period_start), day(PRESENTATION) + 1));
    const LAST = initial.period_end;
    const groupsById = new Map();
    for (const g of initial.groups) {
      if (!g.id || groupsById.has(g.id) || !Number.isFinite(+g.workforce) || +g.workforce <= 0) {
        throw new Error('بيانات إحدى مجموعات القوى العاملة غير صالحة.');
      }
      groupsById.set(g.id, g);
    }
    const originals = new Map();
    for (const l of initial.leaves) {
      if (!l.id || originals.has(l.id) || !l.employee_id || !groupsById.has(l.group_id) ||
          !validDate(l.start) || !validDate(l.end) || l.start > l.end) {
        throw new Error('بيانات إحدى الإجازات غير صالحة أو معرفها مكرر.');
      }
      originals.set(l.id, l);
    }
    const caseIds = new Set();
    for (const c of initial.cases) {
      if (!c.id || caseIds.has(c.id) || !groupsById.has(c.group_id) || !validDate(c.start) ||
          !validDate(c.end) || c.start > c.end || !Number.isFinite(+c.percent)) {
        throw new Error('بيانات إحدى حالات التقرير غير صالحة.');
      }
      caseIds.add(c.id);
      c.employee_ids = Array.isArray(c.employee_ids) ? c.employee_ids.map(String) : String(c.employee_ids || '').split(',').filter(Boolean);
    }
    const initialData = deepFreeze(initial);
    const baselineFuture = initialData.cases.filter(c => c.start > PRESENTATION);
    const thresholdValue = value => Number.isFinite(+value) ? Math.max(0, +value) : 0;
    const originalCases = (threshold = 0) => baselineFuture.filter(c => +c.percent >= thresholdValue(threshold));

    // All identifiers and percentage denominators come from the report's exact group.
    // Display names never determine grouping, and original counts never count pairs.
    const targetCache = new Map();
    function targetIds(threshold = 0, groupId) {
      const value = thresholdValue(threshold);
      if (!targetCache.has(value)) {
        const byGroup = new Map();
        for (const c of originalCases(value)) {
          if (!byGroup.has(c.group_id)) byGroup.set(c.group_id, []);
          byGroup.get(c.group_id).push(c);
        }
        targetCache.set(value, new Set(initialData.leaves.filter(l =>
          (byGroup.get(l.group_id) || []).some(c => c.employee_ids.includes(String(l.employee_id)) && l.start <= c.end && l.end >= c.start)
        ).map(l => l.id)));
      }
      return new Set([...targetCache.get(value)].filter(id => !groupId || originals.get(id).group_id === groupId));
    }
    const allTargets = targetIds();
    const canMove = l => l.movable === true && allTargets.has(l.id) && l.start > PRESENTATION &&
      l.start >= initialData.period_start && l.end <= LAST;

    // This deliberately also blocks simultaneous records for the same employee.
    // Distinct-person overlap cases are counted separately by conflicts().
    const overlap = (a, b) => a.id !== b.id && a.group_id === b.group_id && a.start <= b.end && b.start <= a.end;
    const changedFrom = (a, b) => a.start !== b.start || a.end !== b.end;
    const movedFromOriginal = l => changedFrom(l, originals.get(l.id));

    function calendar(rows, start, end) {
      const lo = day(start), hi = day(end), dates = new Map();
      if (lo > hi) return dates;
      for (const l of rows) {
        for (let n = Math.max(lo, day(l.start)); n <= Math.min(hi, day(l.end)); n++) {
          const key = l.group_id + '|' + n;
          if (!dates.has(key)) dates.set(key, {group_id: l.group_id, n, employees: new Set(), leaves: new Set()});
          dates.get(key).employees.add(String(l.employee_id));
          dates.get(key).leaves.add(l.id);
        }
      }
      return dates;
    }

    function conflicts(rows, options) {
      const opts = options || {};
      const start = opts.start && opts.start > initialData.period_start ? opts.start : initialData.period_start;
      const end = opts.end && opts.end < LAST ? opts.end : LAST;
      const eligible = [...calendar(rows, start, end).values()].filter(item => item.employees.size >= 2)
        .sort((a, b) => a.group_id.localeCompare(b.group_id) || a.n - b.n);
      const runs = [];
      for (const item of eligible) {
        const employees = [...item.employees].sort(), key = JSON.stringify(employees);
        const prior = runs[runs.length - 1];
        if (prior && prior.group_id === item.group_id && day(prior.end) + 1 === item.n && prior.key === key) {
          prior.end = iso(item.n);
          item.leaves.forEach(id => prior.leaveSet.add(id));
        } else {
          runs.push({group_id: item.group_id, start: iso(item.n), end: iso(item.n), key,
            employee_ids: employees, people: employees.length, leaveSet: new Set(item.leaves)});
        }
      }
      return runs.map(({key, leaveSet, ...run}) => {
        const workforce = +groupsById.get(run.group_id).workforce;
        return {...run,
          id: 'PC-' + run.group_id + '-' + run.start + '-' + run.end + '-' + run.employee_ids.join('-'),
          days: day(run.end) - day(run.start) + 1, workforce,
          percent: 100 * run.people / workforce,
          percent_display: Math.round(1000 * run.people / workforce) / 10,
          leave_ids: [...leaveSet].sort()
        };
      });
    }

    // The threshold chooses the ORIGINAL report cases to coordinate. After proposals,
    // every remaining overlap involving their leave records is shown, even below 50%.
    // Recompute in the future planning window so a continuing case is counted once.
    // This count is not subtracted from the original case count: cohorts can split.
    function remainingCases(rows, threshold = 0, groupId) {
      const ids = targetIds(threshold, groupId);
      if (!ids.size || FIRST > LAST) return [];
      const selectedGroups = new Set([...ids].map(id => originals.get(id).group_id));
      return conflicts(rows.filter(l => selectedGroups.has(l.group_id)), {start: FIRST, end: LAST})
        .filter(c => c.leave_ids.some(id => ids.has(id)));
    }

    function pairDays(rows) {
      let total = 0;
      for (const d of calendar(rows, FIRST, LAST).values()) total += d.employees.size * (d.employees.size - 1) / 2;
      return total;
    }

    function validate(original, proposed, options) {
      const errors = [], opts = options || {};
      const previous = new Map(original.map(l => [l.id, l]));
      const allowed = opts.allowedIds ? new Set(opts.allowedIds) : null;
      if (original.length !== proposed.length || new Set(proposed.map(l => l.id)).size !== original.length) {
        errors.push('تغيّر عدد سجلات الإجازات أو تكررت معرفاتها.');
      }
      const proposedByGroup = new Map();
      for (const l of proposed) {
        if (!proposedByGroup.has(l.group_id)) proposedByGroup.set(l.group_id, []);
        proposedByGroup.get(l.group_id).push(l);
      }
      for (const l of proposed) {
        const o = previous.get(l.id), baseline = originals.get(l.id);
        if (!o || !baseline) { errors.push('ظهرت إجازة غير موجودة في الأصل.'); continue; }
        if (l.employee_id !== baseline.employee_id || l.group_id !== baseline.group_id || l.name !== baseline.name ||
            l.movable !== baseline.movable || l.locked_reason !== baseline.locked_reason ||
            l.source_start !== baseline.source_start || l.source_end !== baseline.source_end) {
          errors.push('تغيّرت هوية الموظف أو خصائص سجل الإجازة.');
        }
        if (!validDate(l.start) || !validDate(l.end) || l.start > l.end ||
            day(l.end) - day(l.start) !== day(baseline.end) - day(baseline.start)) {
          errors.push('تغيّرت مدة الإجازة أو أصبح تاريخها غير صالح.'); continue;
        }
        if (changedFrom(l, baseline)) {
          if (!canMove(baseline)) errors.push('تغيّرت إجازة ثابتة أو بدأت بالفعل.');
          if (l.start < FIRST || l.end > LAST) errors.push('الموعد المقترح خارج فترة التنسيق.');
          if ((proposedByGroup.get(l.group_id) || []).some(other => overlap(l, other))) {
            errors.push('الموعد المعدّل يتقاطع مع إجازة أخرى.');
          }
        }
        if (allowed && changedFrom(l, o) && !allowed.has(l.id)) errors.push('تغيّرت إجازة خارج الحالات المختارة.');
      }
      if (!errors.length) {
        const oldCalendar = calendar(original, FIRST, LAST), newCalendar = calendar(proposed, FIRST, LAST);
        for (const [key, entry] of newCalendar) {
          if (entry.employees.size < 2) continue;
          const before = oldCalendar.get(key)?.employees || new Set();
          if ([...entry.employees].some(id => !before.has(id))) { errors.push('ظهر تقاطع جديد.'); break; }
        }
      }
      return {ok: errors.length === 0, errors: [...new Set(errors)]};
    }

    function proposedSourcePeriod(l) {
      const o = originals.get(l.id), shift = day(l.start) - day(o.start);
      return {
        start: iso(day(validDate(o.source_start) ? o.source_start : o.start) + shift),
        end: iso(day(validDate(o.source_end) ? o.source_end : o.end) + shift)
      };
    }

    function solveGroup(groupId, current, options) {
      if (!groupsById.has(groupId)) throw new Error('المجموعة غير موجودة.');
      if (!sourceConsistency.ok) throw new Error('تعذّر اقتراح المواعيد لأن تفاصيل الإجازات لا تطابق حالات التقرير. أعيدي إنشاء الداشبورد من المصادر نفسها.');
      const settings = options || {}, threshold = thresholdValue(settings.threshold);
      const limit = Number.isInteger(settings.maxNodes) && settings.maxNodes >= 0 ? settings.maxNodes : 20000;
      const maxMillis = Number.isFinite(settings.maxMillis) && settings.maxMillis >= 0 ? settings.maxMillis : 350;
      const now = () => typeof performance !== 'undefined' && performance.now ? performance.now() : Date.now();
      const source = copy(current || initialData.leaves);
      const inputCheck = validate(initialData.leaves, source);
      if (!inputCheck.ok) throw new Error(inputCheck.errors.join(' '));
      const scoped = source.filter(l => l.group_id === groupId);
      const selected = targetIds(threshold, groupId);
      const variableIds = new Set([...selected].filter(id => canMove(originals.get(id))));
      const fixed = scoped.filter(l => !variableIds.has(l.id));
      const variables = scoped.filter(l => variableIds.has(l.id));
      const before = remainingCases(source, threshold, groupId);
      const score = rows => [pairDays(rows), rows.filter(movedFromOriginal).length,
        rows.reduce((sum, l) => sum + Math.abs(day(l.start) - day(originals.get(l.id).start)), 0)];
      const better = (a, b) => a[0] < b[0] || (a[0] === b[0] && (a[1] < b[1] || (a[1] === b[1] && a[2] < b[2])));
      let best = copy(scoped), bestScore = score(best), nodes = 0, exhausted = false;
      const domains = new Map();
      if (before.length && variables.length && limit > 0 && maxMillis > 0) {
        const deadline = now() + maxMillis;
        for (const l of variables) {
          const o = originals.get(l.id), duration = day(o.end) - day(o.start), choices = [];
          for (let n = day(FIRST); n + duration <= day(LAST); n++) {
            const candidate = {...l, start: iso(n), end: iso(n + duration)};
            const changed = movedFromOriginal(candidate);
            if (fixed.some(other => overlap(candidate, other) && (changed || movedFromOriginal(other)))) continue;
            choices.push({leave: candidate, changed, shift: Math.abs(n - day(o.start))});
          }
          choices.sort((a, b) => Number(a.changed) - Number(b.changed) || a.shift - b.shift || b.leave.start.localeCompare(a.leave.start));
          domains.set(l.id, choices);
        }

        // Establish a valid incumbent quickly before bounded branch-and-bound.
        // Fixed and out-of-filter bookings remain blockers throughout this pass.
        const greedy = copy(scoped);
        const greedyOrder = variables.slice().sort((a, b) => b.start.localeCompare(a.start) || a.id.localeCompare(b.id));
        for (let pass = 0; pass < 2; pass++) {
          let moved = false;
          for (const l of greedyOrder) {
            const index = greedy.findIndex(row => row.id === l.id), active = greedy[index];
            const hasOverlap = greedy.some(other => other.employee_id !== active.employee_id && overlap(active, other));
            if (!hasOverlap && !movedFromOriginal(active)) continue;
            for (const choice of domains.get(l.id)) {
              if (greedy.some(other => overlap(choice.leave, other))) continue;
              if (!hasOverlap && choice.changed && choice.shift >= Math.abs(day(active.start) - day(originals.get(active.id).start))) break;
              if (changedFrom(choice.leave, active)) { greedy[index] = copy(choice.leave); moved = true; }
              break;
            }
          }
          if (!moved) break;
        }
        const greedyScore = score(greedy);
        if (better(greedyScore, bestScore)) { best = copy(greedy); bestScore = greedyScore; }

        variables.sort((a, b) => domains.get(a.id).length - domains.get(b.id).length ||
          (day(b.end) - day(b.start)) - (day(a.end) - day(a.start)) || a.id.localeCompare(b.id));
        const assigned = [], attendance = new Map();
        let partialPairs = 0;
        function changeAttendance(l, amount) {
          for (let n = Math.max(day(l.start), day(FIRST)); n <= Math.min(day(l.end), day(LAST)); n++) {
            if (!attendance.has(n)) attendance.set(n, new Map());
            const employees = attendance.get(n), old = employees.get(l.employee_id) || 0;
            if (amount > 0) {
              if (old === 0) partialPairs += employees.size;
              employees.set(l.employee_id, old + 1);
            } else if (old === 1) {
              partialPairs -= employees.size - 1;
              employees.delete(l.employee_id);
            } else employees.set(l.employee_id, old - 1);
          }
        }
        fixed.forEach(l => changeAttendance(l, 1));
        const fixedMoved = fixed.filter(movedFromOriginal).length;
        const fixedShift = fixed.reduce((sum, l) => sum + Math.abs(day(l.start) - day(originals.get(l.id).start)), 0);
        function search(index, moved, shift) {
          if (nodes >= limit || (nodes % 128 === 0 && now() >= deadline)) { exhausted = true; return; }
          nodes++;
          if (partialPairs > bestScore[0] || (partialPairs === bestScore[0] &&
              (moved > bestScore[1] || (moved === bestScore[1] && shift >= bestScore[2])))) return;
          if (index === variables.length) {
            const candidateScore = [partialPairs, moved, shift];
            if (better(candidateScore, bestScore)) {
              best = copy(fixed.concat(assigned.map(choice => choice.leave)));
              bestScore = candidateScore;
            }
            return;
          }
          for (const choice of domains.get(variables[index].id)) {
            if (assigned.some(other => overlap(choice.leave, other.leave) && (choice.changed || other.changed))) continue;
            assigned.push(choice); changeAttendance(choice.leave, 1);
            search(index + 1, moved + Number(choice.changed), shift + choice.shift);
            changeAttendance(choice.leave, -1); assigned.pop();
            if (exhausted) break;
          }
        }
        if (now() >= deadline) exhausted = true;
        else search(0, fixedMoved, fixedShift);
      } else if (before.length && variables.length && (limit === 0 || maxMillis === 0)) exhausted = true;

      const replacement = new Map(best.map(l => [l.id, l]));
      const finalRows = source.map(l => replacement.get(l.id) || l);
      const checked = validate(initialData.leaves, finalRows);
      const scopedCheck = validate(source, finalRows, {allowedIds: variableIds});
      if (!checked.ok || !scopedCheck.ok) {
        throw new Error('فشل فحص المواعيد المقترحة: ' + [...checked.errors, ...scopedCheck.errors].join(' '));
      }
      const changed = finalRows.filter(l => l.group_id === groupId && movedFromOriginal(l)).map(l => {
        const original = originals.get(l.id), proposedSource = proposedSourcePeriod(l);
        return {...l, original_start: original.start, original_end: original.end,
          proposed_source_start: proposedSource.start, proposed_source_end: proposedSource.end,
          shift_days: day(l.start) - day(original.start), duration_days: day(l.end) - day(l.start) + 1};
      });
      const after = remainingCases(finalRows, threshold, groupId);
      const sourceById = new Map(source.map(l => [l.id, l]));
      const changedThisRun = finalRows.filter(l => changedFrom(l, sourceById.get(l.id))).map(l => l.id);
      return {
        leaves: finalRows, changed, changed_this_run: changedThisRun,
        before, after, unresolved: after, resolved: after.length === 0, checked,
        searched_nodes: nodes, search_limit_reached: exhausted,
        pair_days_before: pairDays(scoped), pair_days_after: pairDays(best),
        result_code: !before.length ? 'not_needed' : !after.length ? 'solved' :
          !variables.length ? 'no_movable_leaves' : changedThisRun.length ? 'partial' : 'no_alternative_found'
      };
    }

    // Compare cohort identities and dates, never replace the supplied report's IDs
    // or counts. A missing leave can otherwise make a false "resolved" result.
    const signature = c => JSON.stringify([c.group_id, c.start, c.end, c.employee_ids.map(String).slice().sort()]);
    const sourceSignatures = initialData.cases.map(signature).sort();
    const recomputedSignatures = conflicts(initialData.leaves).map(signature).sort();
    const sourceConsistency = Object.freeze({
      ok: JSON.stringify(sourceSignatures) === JSON.stringify(recomputedSignatures),
      report_cases: sourceSignatures.length, leave_cases: recomputedSignatures.length
    });
    const isTarget = (id, threshold = 0) => {
      const value = thresholdValue(threshold);
      if (!targetCache.has(value)) targetIds(value);
      return targetCache.get(value).has(id);
    };
    return {initialData, groups: initialData.groups, leaves: initialData.leaves, sourceConsistency,
      presentationDate: PRESENTATION, firstDate: FIRST, lastDate: LAST,
      originalCases, targetIds, conflicts, remainingCases, solveGroup, validate, overlap,
      proposedSourcePeriod, isTarget,
      cloneInitial: () => copy(initialData.leaves)};
  }

  root.createPlannerModel = createPlannerModel;
  if (typeof module !== 'undefined' && module.exports) module.exports = createPlannerModel;
})(typeof globalThis !== 'undefined' ? globalThis : window);

(function () {
  'use strict';
  const root = document.getElementById('smart-leave-calendar');
  if (!root) return;
  const ui = name => root.querySelector('[data-ui="' + name + '"]');
  const action = name => root.querySelector('[data-action="' + name + '"]');
  const input = typeof DATA === 'undefined' ? null : DATA;
  if (!input || !input.planner || typeof globalThis.createPlannerModel !== 'function') {
    ui('empty').hidden = false;
    ui('empty').textContent = 'بيانات المنسّق غير متاحة في هذا الملف.';
    ui('filterbar')?.setAttribute('hidden', '');
    return;
  }
  const model = globalThis.createPlannerModel({...input.planner, groups: input.groups, cases: input.cases});
  const original = model.cloneInitial();
  const originalById = new Map(original.map(leave => [leave.id, leave]));
  const groups = new Map(model.groups.map(group => [group.id, group]));
  const DAY = 86400000;
  const asof = model.presentationDate;
  const toDay = value => Date.parse(value + 'T00:00:00Z') / DAY;
  const toIso = value => new Date(value * DAY).toISOString().slice(0, 10);
  const number = value => new Intl.NumberFormat('en-US', {maximumFractionDigits: 1}).format(Number(value));
  const dateFormatter = new Intl.DateTimeFormat('ar-SA-u-ca-gregory-nu-latn', {day: 'numeric', month: 'long', timeZone: 'UTC'});
  const fullFormatter = new Intl.DateTimeFormat('ar-SA-u-ca-gregory-nu-latn', {day: 'numeric', month: 'long', year: 'numeric', weekday: 'long', timeZone: 'UTC'});
  const monthFormatter = new Intl.DateTimeFormat('ar-SA-u-ca-gregory-nu-latn', {month: 'long', year: 'numeric', timeZone: 'UTC'});
  const dateText = value => dateFormatter.format(new Date(value + 'T00:00:00Z'));
  const period = (start, end) => start === end ? dateText(start) : dateText(start) + ' – ' + dateText(end);
  const leavePeriod = leave => {
    const source = model.proposedSourcePeriod(leave);
    return period(source.start, source.end);
  };
  const escape = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[ch]));
  const norm = value => String(value ?? '').normalize('NFKC').replace(/[\u064b-\u065f\u0670\u0640\u200e\u200f\u200b]/g, '').replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي').replace(/\s+/g, ' ').trim().toLocaleLowerCase();
  const durationText = days => days === 1 ? 'يوم واحد' : days === 2 ? 'يومان' : number(days) + ' أيام';
  const months = [];
  const firstMonth = new Date(Math.max(toDay(input.planner.period_start), toDay(asof)) * DAY);
  const finalMonth = model.lastDate.slice(0, 7);
  let cursor = Date.UTC(firstMonth.getUTCFullYear(), firstMonth.getUTCMonth(), 1);
  while (new Date(cursor).toISOString().slice(0, 7) <= finalMonth) {
    months.push(cursor / DAY);
    const d = new Date(cursor);
    cursor = Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 1);
  }
  if (!months.length) months.push(Date.UTC(firstMonth.getUTCFullYear(), firstMonth.getUTCMonth(), 1) / DAY);
  let current = model.cloneInitial();
  let monthIndex = 0, selected = null, view = 'after', busy = false;
  const results = new Map(), errors = new Map();
  const threshold = () => Number(ui('threshold').value) || 0;
  const resultKey = id => id + '|' + threshold();
  const groupMatches = id => !norm(ui('manager-search').value) || norm(groups.get(id)?.manager).includes(norm(ui('manager-search').value));
  const scopedCases = () => model.originalCases(threshold()).filter(item => groupMatches(item.group_id));
  const scopedGroups = () => new Set(scopedCases().map(item => item.group_id));
  const isChanged = leave => {
    const old = originalById.get(leave.id);
    return leave.start !== old.start || leave.end !== old.end;
  };
  const changedInGroup = id => current.filter(leave => leave.group_id === id && isChanged(leave));
  const hasAttempt = id => [...results.keys()].some(key => key.startsWith(id + '|'));
  const groupRemaining = id => model.remainingCases(current, threshold(), id);
  const caseLeaves = item => item.leave_ids || original.filter(leave => leave.group_id === item.group_id &&
    item.employee_ids.includes(leave.employee_id) && leave.start <= item.end && leave.end >= item.start).map(leave => leave.id);
  const caseEvent = item => ({...item, id: 'case-' + item.id, kind: 'overlap', leave_ids: caseLeaves(item)});

  function eventsForView() {
    const base = scopedCases();
    if (view === 'before') return base.map(caseEvent);
    const events = [];
    for (const id of new Set(base.map(item => item.group_id))) {
      if (!hasAttempt(id) && !changedInGroup(id).length) {
        events.push(...base.filter(item => item.group_id === id).map(caseEvent));
        continue;
      }
      const remaining = groupRemaining(id);
      events.push(...remaining.map(caseEvent));
      const pendingIds = new Set(remaining.flatMap(caseLeaves));
      const targets = model.targetIds(threshold(), id);
      for (const leave of current) {
        if (leave.group_id !== id || !targets.has(leave.id) || pendingIds.has(leave.id)) continue;
        if (leave.end < model.firstDate || leave.start > model.lastDate) continue;
        events.push({id: 'planned-' + leave.id, group_id: id, start: leave.start < model.firstDate ? model.firstDate : leave.start,
          end: leave.end > model.lastDate ? model.lastDate : leave.end, kind: 'planned',
          employee_ids: [leave.employee_id], leave_ids: [leave.id]});
      }
    }
    return events.sort((a, b) => a.start.localeCompare(b.start) || a.group_id.localeCompare(b.group_id) || a.id.localeCompare(b.id));
  }

  function renderCounts() {
    const base = scopedCases();
    const ids = new Set(base.map(item => item.group_id));
    ui('original-count').textContent = number(base.length);
    ui('scope-groups').textContent = number(ids.size) + ' مجموعات مدير وموقع';
    ui('scope-note').textContent = 'حالات تبدأ بعد ' + dateText(asof) + ' ' + asof.slice(0, 4) +
      '، حسب الحالات الأصلية في التقرير' + (threshold() ? ' وبنسبة 50% فأكثر.' : '، بجميع النسب.');
    const showRemaining = [...ids].some(id => hasAttempt(id) || changedInGroup(id).length);
    ui('remaining-count').hidden = !showRemaining;
    if (showRemaining) {
      const remaining = [...ids].reduce((sum, id) => sum + groupRemaining(id).length, 0);
      ui('remaining-count').textContent = 'التقاطعات المتبقية في فترة التنسيق بعد الاقتراح: ' + number(remaining);
    }
    ui('manager-search').disabled = busy;
    ui('threshold').disabled = busy;
    action('reset').disabled = busy;
  }

  function renderCalendar() {
    const low = months[monthIndex];
    const monthDate = new Date(low * DAY);
    const high = Date.UTC(monthDate.getUTCFullYear(), monthDate.getUTCMonth() + 1, 0) / DAY;
    const first = low - new Date(low * DAY).getUTCDay();
    const last = high + 6 - new Date(high * DAY).getUTCDay();
    const changed = current.some(isChanged);
    const events = eventsForView();
    const monthEvents = events.filter(event => toDay(event.start) <= high && toDay(event.end) >= low);
    ui('month-label').textContent = monthFormatter.format(new Date(low * DAY));
    action('previous').disabled = busy || monthIndex === 0;
    action('next').disabled = busy || monthIndex === months.length - 1;
    ui('compare').hidden = !changed;
    ui('date-note').textContent = 'تاريخ العرض: ' + dateText(asof);
    ui('date-note').hidden = changed;
    action('before').disabled = action('after').disabled = busy;
    action('before').setAttribute('aria-pressed', String(view === 'before'));
    action('after').setAttribute('aria-pressed', String(view === 'after'));
    let html = '';
    for (let week = first; week <= last; week += 7) {
      let cells = '';
      for (let offset = 0; offset < 7; offset++) {
        const n = week + offset, iso = toIso(n), outside = n < low || n > high;
        const isAsOf = iso === asof, weekend = offset >= 5;
        const classes = ['sp-day', weekend ? 'sp-weekend' : '', iso < asof ? 'sp-past' : '', outside ? 'sp-outside' : '', isAsOf ? 'sp-asof' : ''].join(' ');
        cells += `<div class="${classes}" data-date="${iso}" role="group" aria-label="${escape(fullFormatter.format(new Date(n * DAY)))}${weekend ? '، نهاية الأسبوع' : ''}${isAsOf ? '، تاريخ العرض' : ''}"><span class="sp-number">${number(new Date(n * DAY).getUTCDate())}</span>${isAsOf ? '<span class="sp-asof-label">العرض</span>' : ''}</div>`;
      }
      const lanes = [];
      let bars = '';
      for (const event of monthEvents.filter(item => toDay(item.start) <= week + 6 && toDay(item.end) >= week)) {
        const start = Math.max(week, low, toDay(event.start)), end = Math.min(week + 6, high, toDay(event.end));
        if (start > end) continue;
        let lane = lanes.findIndex(tail => tail < start);
        if (lane < 0) { lane = lanes.length; lanes.push(end); } else lanes[lane] = end;
        const group = groups.get(event.group_id), good = event.kind === 'planned';
        const ratio = event.percent ?? (100 * event.employee_ids.length / group.workforce);
        const tail = good ? 'بعد التنسيق' : number(ratio) + '%';
        const org = [group.level1, group.level2, group.level3].filter(Boolean).join(' / ');
        const label = group.manager + '، ' + group.port + (org ? '، ' + org : '') + '، ' + (good ? 'موعد بعد التنسيق' : 'تقاطع ' + number(ratio) + '%') + '، ' + period(event.start, event.end) + '، عرض الموظفين';
        bars += `<button type="button" class="sp-event${good ? ' sp-good' : ''}${selected === event.group_id ? ' sp-selected' : ''}" data-group="${escape(event.group_id)}" data-event="${escape(event.id)}" aria-label="${escape(label)}" title="${escape(label)}" aria-expanded="${selected === event.group_id}" ${busy ? 'disabled' : ''} style="grid-column:${start - week + 1} / span ${end - start + 1};grid-row:${lane + 1}"><span class="sp-event-name">${escape(group.manager)}</span><span class="sp-event-count">${tail}</span><span class="sp-mobile-period">${escape(period(toIso(start), toIso(end)))}</span></button>`;
      }
      html += `<div class="sp-week"><div class="sp-day-grid">${cells}</div><div class="sp-events">${bars}</div></div>`;
    }
    ui('weeks').innerHTML = html;
    ui('month-count').textContent = number(new Set(monthEvents.map(event => event.group_id)).size) + ' مجموعات معروضة في هذا الشهر';
    ui('empty').hidden = monthEvents.length > 0;
    ui('empty').textContent = scopedCases().length
      ? 'لا توجد فترات لعرضها في هذا الشهر ضمن الفلتر. يمكنك الانتقال إلى الشهر الآخر.'
      : 'لا توجد حالات مستقبلية أصلية ضمن الفلتر المختار.';
  }

  function renderPanel() {
    ui('workspace').classList.toggle('sp-open', Boolean(selected));
    ui('panel').hidden = !selected;
    if (!selected) return;
    const group = groups.get(selected), targets = model.targetIds(threshold(), selected);
    const base = original.filter(leave => leave.group_id === selected && targets.has(leave.id))
      .sort((a, b) => a.start.localeCompare(b.start) || a.name.localeCompare(b.name) || a.id.localeCompare(b.id));
    const baseCases = scopedCases().filter(item => item.group_id === selected);
    const isBefore = view === 'before', shown = isBefore ? original : current;
    const shownById = new Map(shown.map(leave => [leave.id, leave]));
    const fixed = shown.filter(leave => leave.group_id === selected && !targets.has(leave.id) && leave.end >= model.firstDate)
      .sort((a, b) => a.start.localeCompare(b.start) || a.id.localeCompare(b.id));
    const result = results.get(resultKey(selected)), error = errors.get(resultKey(selected));
    const changes = changedInGroup(selected), pending = groupRemaining(selected);
    const attempted = hasAttempt(selected) || changes.length > 0;
    const complete = attempted && pending.length === 0;
    const canMove = base.some(leave => leave.movable !== false && leave.start >= model.firstDate && leave.end <= model.lastDate);
    ui('manager-title').textContent = group.manager;
    ui('manager-port').textContent = group.port;
    ui('manager-org').innerHTML = [group.level1, group.level2, group.level3].filter(Boolean).map(value => '<bdi>' + escape(value) + '</bdi>').join(' / ');
    ui('panel-note').classList.toggle('sp-success', complete && !isBefore);
    ui('panel-note').innerHTML = complete && !isBefore
      ? '<i class="sp-dot sp-green" aria-hidden="true"></i><span>تم تنسيق التقاطع ضمن الفلتر</span>'
      : `<i class="sp-dot" aria-hidden="true"></i><span>${number(new Set(base.map(leave => leave.employee_id)).size)} موظفين مرتبطين بالحالات الأصلية${isBefore ? ' · قبل التعديل' : ''}</span>`;
    const peak = baseCases.length ? Math.max(...baseCases.map(item => item.percent)) : 0;
    ui('case-summary').textContent = number(baseCases.length) + ' حالات مستقبلية أصلية · أعلى نسبة ' + number(peak) + '%';
    ui('case-list').innerHTML = baseCases.map(item => `<div class="sp-case-item"><strong>${escape(period(item.start, item.end))}</strong><div>${number(item.people)} من ${number(group.workforce)} موظفين · ${number(item.percent)}%</div><small>${escape(item.id)}</small></div>`).join('');
    ui('employee-heading').textContent = attempted && !isBefore ? 'المواعيد بعد التنسيق' : 'الموظفون في التقاطع';
    ui('employees').innerHTML = base.map(old => {
      const leave = shownById.get(old.id), changed = leave.start !== old.start || leave.end !== old.end;
      const locked = old.movable === false || old.start < model.firstDate || old.end > model.lastDate;
      const reason = old.locked_reason || (old.start < model.firstDate ? 'الإجازة بدأت بحلول تاريخ العرض؛ يبقى موعدها ثابتًا.' : 'هذه الإجازة ثابتة ضمن فترة التنسيق.');
      return `<div class="sp-employee" data-employee-leave="${escape(leave.id)}"><div class="sp-employee-head"><span class="sp-employee-name">${escape(leave.name)}<span class="sp-employee-id">${escape(leave.employee_id)}</span></span><span class="sp-duration">${durationText(toDay(leave.end) - toDay(leave.start) + 1)}</span></div>${changed
        ? `<div class="sp-old-period">${escape(leavePeriod(old))}</div><div class="sp-period sp-new-period"><span>${escape(leavePeriod(leave))}</span><span class="sp-tiny-label">مقترح</span></div>`
        : `<div class="sp-period">${escape(leavePeriod(leave))}</div>${complete && !isBefore ? '<div class="sp-unchanged">بقي في موعده · دون تقاطع ضمن الفلتر</div>' : ''}`}${locked ? '<div class="sp-locked">' + escape(reason) + '</div>' : ''}</div>`;
    }).join('') || '<p class="sp-rule">لا توجد سجلات إجازات مرتبطة بالحالات المختارة.</p>';
    ui('fixed').hidden = !fixed.length;
    ui('fixed-label').textContent = 'إجازات أخرى نراعيها (' + number(fixed.length) + ')';
    ui('fixed-list').innerHTML = fixed.map(leave => `<div class="sp-fixed-item"><b>${escape(leave.name)}</b><span>${escape(leavePeriod(leave))}${!isBefore && isChanged(leave) ? ' · مقترح سابق' : ''}</span></div>`).join('');
    ui('result').hidden = (!attempted && !error) || isBefore;
    ui('result').classList.toggle('sp-partial', !complete || Boolean(error));
    ui('result').innerHTML = error
      ? '<b>تعذر تكوين اقتراح</b>' + escape(error)
      : complete
        ? `<b>لا توجد تقاطعات متبقية مرتبطة بهذه الحالات</b>تغيّر موعد ${number(changes.length)} إجازات في المجموعة. مدة كل إجازة محفوظة.`
        : '<b>بقي تقاطع يحتاج مراجعة</b>' + number(pending.length) + ' تقاطعات في فترة التنسيق.' +
          (result?.search_limit_reached ? ' انتهت حدود البحث؛ لم يُعثر على بديل كامل في هذه المحاولة.' : ' عُرضت المواعيد البديلة التي اجتازت الفحص.') ;
    action('solve').hidden = false;
    action('solve').disabled = busy || isBefore || complete || !canMove || !pending.length;
    action('solve').textContent = busy ? 'جارٍ تنسيق الإجازات…' : complete ? 'تم تنسيق التقاطع' : 'تنسيق التقاطع';
    action('undo').hidden = !attempted && !error;
    action('undo').disabled = busy;
    action('close').disabled = busy;
    ui('rule').textContent = isBefore ? 'اختاري «بعد التعديل» لتنسيق المواعيد.' : !canMove && pending.length
      ? 'الإجازات المرتبطة بهذه الحالات ثابتة؛ تحتاج مراجعة المواعيد يدويًا.'
      : 'نفس مدة الإجازة، مع فحص إجازات المجموعة الأخرى حتى ' + dateText(model.lastDate) + '.';
  }

  function render() {
    if (selected && !scopedGroups().has(selected)) selected = null;
    renderCounts(); renderCalendar(); renderPanel();
  }
  function openGroup(id) {
    if (busy || !groups.has(id) || !scopedGroups().has(id)) return;
    selected = id;
    render();
    ui('live').textContent = 'تفاصيل إجازات ' + groups.get(id).manager + ' في ' + groups.get(id).port;
    if (window.matchMedia('(max-width: 840px)').matches) ui('panel').scrollIntoView({block: 'start', behavior: window.matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth'});
    action('close').focus({preventScroll: true});
  }
  function closePanel() {
    const previous = selected;
    selected = null; render();
    [...root.querySelectorAll('[data-group]')].find(item => item.dataset.group === previous)?.focus({preventScroll: true});
  }
  async function solveSelected() {
    if (!selected || busy || action('solve').disabled) return;
    const id = selected, chosenThreshold = threshold(), key = resultKey(id);
    errors.delete(key); busy = true; render();
    await new Promise(resolve => requestAnimationFrame(() => setTimeout(resolve, 0)));
    try {
      const result = model.solveGroup(id, current, {threshold: chosenThreshold});
      const checked = model.validate(original, result.leaves);
      if (!checked.ok) throw new Error(checked.errors.join(' '));
      current = result.leaves; results.set(key, result); view = 'after';
      ui('live').textContent = result.resolved ? 'تم تنسيق تقاطعات ' + groups.get(id).manager + '.' : 'انتهى البحث. بقيت تقاطعات تحتاج مراجعة.';
    } catch (error) {
      errors.set(key, error.message || 'يمكنك إعادة المحاولة.'); view = 'after';
      ui('live').textContent = 'تعذر تكوين اقتراح. ' + errors.get(key);
    } finally {
      busy = false; render();
      if (!action('undo').hidden) action('undo').focus({preventScroll: true});
    }
  }
  function undoSelected() {
    if (!selected || busy) return;
    current = current.map(leave => leave.group_id === selected ? {...originalById.get(leave.id)} : leave);
    for (const key of [...results.keys()]) if (key.startsWith(selected + '|')) results.delete(key);
    for (const key of [...errors.keys()]) if (key.startsWith(selected + '|')) errors.delete(key);
    view = 'after'; ui('live').textContent = 'استُعيدت المواعيد الأصلية للمجموعة.'; render();
    action('solve').focus({preventScroll: true});
  }
  root.addEventListener('click', event => {
    const group = event.target.closest('[data-group]');
    if (group && root.contains(group)) { openGroup(group.dataset.group); return; }
    const control = event.target.closest('[data-action]');
    if (!control || control.disabled || !root.contains(control)) return;
    switch (control.dataset.action) {
      case 'previous': monthIndex = Math.max(0, monthIndex - 1); render(); break;
      case 'next': monthIndex = Math.min(months.length - 1, monthIndex + 1); render(); break;
      case 'close': closePanel(); break;
      case 'before': view = 'before'; render(); break;
      case 'after': view = 'after'; render(); break;
      case 'solve': solveSelected(); break;
      case 'undo': undoSelected(); break;
      case 'reset': ui('manager-search').value = ''; ui('threshold').value = '0'; render(); break;
    }
  });
  root.addEventListener('keydown', event => {
    if (event.key === 'Escape' && selected && !busy) { event.preventDefault(); closePanel(); }
  });
  ui('manager-search').addEventListener('input', render);
  ui('threshold').addEventListener('change', render);
  ui('manager-names').innerHTML = [...new Set(model.originalCases().map(item => groups.get(item.group_id)?.manager).filter(Boolean))]
    .sort().map(name => '<option value="' + escape(name) + '">').join('');
  ui('footer-period').textContent = 'فترة التنسيق: ' + period(model.firstDate, model.lastDate) + ' ' + model.lastDate.slice(0, 4) + '.' +
    (input.planner.end_date_inclusive === false ? ' تواريخ بطاقات الموظفين حسب اصطلاح المصدر؛ يوم النهاية غير محسوب في مدة الإجازة أو التقويم.' : '');
  if (input.meta?.demo_label) { ui('demo-label').hidden = false; ui('demo-label').textContent = 'معاينة ببيانات تجريبية'; }
  root.__plannerUI = {model, render, openGroup, solveSelected, undoSelected,
    state: () => ({current: current.map(leave => ({...leave})), threshold: threshold(), selected, monthIndex, view, busy,
      original_case_count: scopedCases().length, groups: [...scopedGroups()]})};
  render();
})();


if (DATA.meta.demo_label) {
  const notice=document.getElementById('integration-demo-notice');
  notice.textContent=DATA.meta.demo_label; notice.hidden=false;
}

</script>
</body>
</html>
"""


if __name__ == '__main__':
    try:
        build()
    except Exception as exc:
        print('خطأ: ' + str(exc), file=sys.stderr)
        sys.exit(1)

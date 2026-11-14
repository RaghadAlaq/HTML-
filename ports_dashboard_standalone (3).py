# -*- coding: utf-8 -*-
"""
داشبورد تقاطع الإجازات — نفس منطق سكربت التقرير المرفق.

التثبيت مرة واحدة:
    python -m pip install pandas openpyxl
التشغيل:
    python ports_dashboard_standalone.py

عدّلي مسارات Employee Master وPorts Leaves وMapping أدناه، ثم شغّلي الملف.
الناتج: dashboard.html بجانب ملف الإجازات، ويُفتح بالمتصفح مباشرة.
ملف Python واحد، ويُنشئ HTML فقط. جميع مكونات العرض مضمنة وتعمل دون إنترنت.

الحالة = نفس مجموعة موظفين اثنين فأكثر في أيام تقويمية متصلة، ضمن
المدير نفسه + الموقع نفسه + المسار التنظيمي نفسه للمستويات الأول والثاني والثالث.
تغيّر الموظفين أو انقطاع الأيام يبدأ حالة جديدة. النسبة من موظفي المجموعة نفسها.
تترجم الأسماء والمسميات من Mapping عند وجود مطابقة صريحة وفريدة، وإلا يظهر نص المصدر.
ملف الربط للعرض فقط ولا يغيّر عدد الحالات أو المقام أو منطق التقرير. لا يلزم تقرير مرجعي.
تُعرض جميع الحالات بغض النظر عن نسبتها.
المنسّق الذكي يعرض حالات التقرير التي تبدأ بعد تاريخ العرض، ويقترح مواعيد
داخل الصفحة مع إبقاء مدة الإجازة وإجازات الموظفين الآخرين. لا يكتب تغييرات إلى المصادر.
"""
from pathlib import Path
from datetime import date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError
from collections import defaultdict, Counter
import json
import math
import os
import re
import sys
import tempfile
import unicodedata
import pandas as pd

SCRIPT_VERSION = '2026-11-14-standalone-v6'

# ===================== عدّلي المسارات هنا =====================
MANAGERS_FILE = r'C:\Users\rammutairi\Downloads\Employee Master.xlsx'
LEAVES_FILE = r'C:\Users\rammutairi\Downloads\Ports Leaves.xlsx'
MAPPING_FILE = r'C:\Users\rammutairi\Downloads\Mapping.xlsx'  # مرجع الترجمة والمواقع للعرض
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


# Mapping is display-only: no translated label is used as a calculation key.
MAPPING_FIELDS = {
    'port': ['المنفذ', 'اسم المنفذ', 'الموقع', 'موقع العمل', 'port', 'location', 'port name'],
    'title': ['المسمى الوظيفي', 'المسمى الوظيفي للمدير', 'المسمى', 'الدور', 'title', 'role', 'job title', 'manager role', 'manager title'],
    'manager': ['المدير', 'اسم المدير', 'manager', 'manager name'],
    'name': ['الاسم', 'اسم الموظف', 'اسم الموظف أو المدير', 'name', 'employee', 'employee name'],
    'level1': ['المستوى الأول', 'level1', 'level 1', 'lvl1'],
    'level2': ['المستوى الثاني', 'level2', 'level 2', 'lvl2'],
    'level3': ['المستوى الثالث', 'level3', 'level 3', 'lvl3'],
    'region': ['المنطقة', 'المنطقة الإدارية', 'المنطقه الإدارية', 'region', 'administrative region'],
    'type': ['نوع المنفذ', 'نوع الموقع', 'التصنيف', 'type', 'port type', 'category'],
}
MAPPING_WIDE = {
    'port': (['Location (Name)', '(Employee) Location (Name)', 'Name', 'Location', 'Port Name', 'اسم الموقع الأصلي', 'اسم المنفذ الأصلي', 'موقع العمل', 'المنفذ'],
             ['أسم الموقع عربي', 'اسم الموقع عربي', 'اسم الموقع بالعربي', 'اسم المنفذ عربي', 'اسم المنفذ بالعربي', 'Port Name Arabic', 'Location Arabic']),
    'title': (['Job Title', 'Manager Job Title', 'Manager Role', 'المسمى الوظيفي', 'المسمى الوظيفي الأصلي', 'المسمى الوظيفي للمدير'],
              ['المسمى الوظيفي للمدير عربي', 'المسمى الوظيفي عربي', 'المسمى الوظيفي بالعربي', 'Job Title Arabic', 'Manager Role Arabic']),
    'manager': (['Manager Name', '(Employee) Manager (Name)', 'اسم المدير الأصلي', 'اسم المدير'],
                ['اسم المدير عربي', 'اسم المدير بالعربي', 'Manager Name Arabic']),
    'name': (['Employee Name', '(Employee) Full Name', 'اسم الموظف الأصلي', 'اسم الموظف'],
             ['اسم الموظف عربي', 'اسم الموظف بالعربي', 'الاسم الرسمي بالعربي', 'Employee Name Arabic']),
    'level1': (['Level 1', 'Level 1 (Name)', 'Level 1 Name', 'المستوى الأول', 'المستوى الأول الأصلي'],
               ['المستوى الأول عربي', 'المستوى الأول بالعربي', 'Level 1 Arabic']),
    'level2': (['Level 2', 'Level 2 (Name)', 'Level 2 Name', 'المستوى الثاني', 'المستوى الثاني الأصلي'],
               ['المستوى الثاني عربي', 'المستوى الثاني بالعربي', 'Level 2 Arabic']),
    'level3': (['Level 3', 'Level 3 (Name)', 'Level 3 Name', 'المستوى الثالث', 'المستوى الثالث الأصلي'],
               ['المستوى الثالث عربي', 'المستوى الثالث بالعربي', 'Level 3 Arabic']),
    'region': (['Region', 'Administrative Region', 'المنطقة الإدارية الأصلية', 'المنطقة الإدارية', 'المنطقه الإدارية'],
               ['المنطقة الإدارية عربي', 'المنطقة الإدارية بالعربي', 'المنطقه الإدارية عربي', 'Region Arabic']),
    'type': (['Category (Picklist Label)', 'Category', 'Port Type', 'نوع المنفذ', 'نوع الموقع'],
             ['نوع المنفذ عربي', 'نوع المنفذ بالعربي', 'Port Type Arabic']),
}
PORT_METADATA = {
    'code': ['Code', 'Port Code', 'Location Code', 'رمز المنفذ', 'رمز الموقع', 'كود المنفذ'],
    'region': ['المنطقه الإدارية', 'المنطقة الإدارية', 'Region', 'Administrative Region'],
    'type': ['Category (Picklist Label)', 'Category', 'Port Type', 'نوع المنفذ', 'نوع الموقع'],
    'latitude': ['العرض2', 'Latitude', 'Lat', 'خط العرض', 'دائرة العرض'],
    'longitude': ['الطول2', 'Longitude', 'Lon', 'Lng', 'خط الطول'],
}


class MappingResolver:
    """Read explicit, exact mappings; keep conflicting or missing labels unchanged."""
    def __init__(self, path):
        self.path = str(path or '')
        self.file_name = self.path.replace('\\', '/').rsplit('/', 1)[-1]
        self.loaded = False
        self.sheets = []
        self.issues = []
        self._issues = set()
        self._missing = set()
        self._translated = set()
        self.translations = defaultdict(lambda: defaultdict(set))
        self.port_rows = []
        self.port_index = defaultdict(list)
        self.code_rows = defaultdict(list)
        self.unsafe_codes = set()
        self._port_cache = {}
        self._matched = set()
        self._unmatched = set()
        self._ambiguous = set()
        self._load()

    @staticmethod
    def text(value):
        return '' if value is None or pd.isna(value) else re.sub(r'\s+', ' ', str(value)).strip()

    def issue(self, field, source, message, missing=False):
        source = self.text(source)
        key = (field, source, message)
        if key not in self._issues:
            self._issues.add(key)
            self.issues.append({'field': field, 'source': source, 'issue': message})
        if missing:
            self._missing.add((field, norm(source)))

    def _add(self, field, source, target):
        source, target = self.text(source), self.text(target)
        if not source or not target:
            return
        if not re.search('[\u0621-\u064a\u066e-\u06d3]', target):
            self.issue(field, source, 'الترجمة في ملف الربط لا تحتوي نصًا عربيًا؛ حُفظ نص المصدر.')
            return
        self.translations[field][norm(source)].add(target)

    @staticmethod
    def _values(row, aliases):
        wanted = {norm(alias) for alias in aliases}
        return list(dict.fromkeys(MappingResolver.text(value) for key, value in row.items()
                                 if norm(key) in wanted and MappingResolver.text(value)))

    def _load(self):
        if not self.path or not Path(self.path).is_file():
            self.issue('mapping', self.file_name or 'Mapping.xlsx',
                       'ملف الربط غير موجود؛ ظهرت القيم الأصلية وبقي حساب التقاطعات كما هو.', missing=True)
            return
        type_keys = {norm(alias): field for field, aliases in MAPPING_FIELDS.items() for alias in aliases}
        type_keys.update({norm(alias): '*' for alias in ('عام', 'الكل', 'all', '*')})
        source_keys = ['النص الأصلي', 'القيمة الأصلية', 'Source', 'Source Text', 'Original Text']
        arabic_keys = ['الترجمة العربية', 'النص العربي', 'Arabic', 'Arabic Translation']
        kind_keys = ['نوع الحقل', 'الحقل', 'Field', 'Field Type']
        recognized = 0
        try:
            with pd.ExcelFile(self.path, engine='openpyxl') as book:
                self.loaded = True
                for sheet in book.sheet_names:
                    try:
                        # Read the header row explicitly so pandas cannot silently rename
                        # duplicate headings (which would otherwise conceal ambiguity).
                        raw_frame = pd.read_excel(book, sheet_name=sheet, dtype=object, header=None)
                        if raw_frame.empty:
                            self.sheets.append(sheet)
                            continue
                        header = [self.text(value) for value in raw_frame.iloc[0]]
                        frame = raw_frame.iloc[1:].copy().dropna(how='all')
                        frame.columns = header
                    except Exception as exc:
                        self.issue('mapping', sheet, 'تعذرت قراءة ورقة الربط: ' + type(exc).__name__)
                        continue
                    self.sheets.append(sheet)
                    headings = [value for value in frame.columns if value]
                    if len(set(norm(value) for value in headings)) != len(headings):
                        self.issue('mapping', sheet, 'رؤوس أعمدة مكررة بعد التطبيع؛ لم تُستخدم هذه الورقة.')
                        continue
                    for row_number, (_, series) in enumerate(frame.iterrows(), 2):
                        row = series.to_dict()
                        sources, targets = self._values(row, source_keys), self._values(row, arabic_keys)
                        if sources or targets:
                            kinds = self._values(row, kind_keys)
                            field = '*' if not kinds else type_keys.get(norm(kinds[0])) if len(kinds) == 1 else None
                            if field is None:
                                self.issue('mapping', f'{sheet}:{row_number}', 'نوع الحقل غير معروف؛ لم تُستخدم الترجمة العامة.')
                            elif len(sources) == 1 and len(targets) == 1:
                                self._add(field, sources[0], targets[0])
                                recognized += 1
                            else:
                                self.issue('mapping', f'{sheet}:{row_number}', 'صف ترجمة عام ناقص أو متعدد القيم؛ لم يُستخدم.')
                        for field, (original_columns, arabic_columns) in MAPPING_WIDE.items():
                            originals, arabics = self._values(row, original_columns), self._values(row, arabic_columns)
                            if field != 'port' and arabics and not originals:
                                self.issue(field, f'{sheet}:{row_number}',
                                           'توجد ترجمة دون نص أصلي مطابق؛ لم تُسند إلى الموظف أو المدير من اسم المنفذ وحده.')
                            for source in originals:
                                for target in arabics:
                                    self._add(field, source, target)
                                    recognized += 1
                        aliases = self._values(row, MAPPING_WIDE['port'][0])
                        arabic_ports = self._values(row, MAPPING_WIDE['port'][1])
                        metadata = {field: self._values(row, columns) for field, columns in PORT_METADATA.items()}
                        if not aliases or not (arabic_ports or any(metadata.values())):
                            continue
                        codes = list(dict.fromkeys(empid(value) for value in metadata['code']))
                        # All explicit source aliases and the approved Arabic label identify this row.
                        record = {'aliases': {norm(value) for value in aliases + arabic_ports},
                                  'arabic_ports': arabic_ports, 'code': codes,
                                  'region': metadata['region'] or self._values(row, MAPPING_WIDE['region'][1]),
                                  'type': metadata['type'] or self._values(row, MAPPING_WIDE['type'][1]),
                                  'latitude': metadata['latitude'], 'longitude': metadata['longitude'],
                                  'source': f'{sheet}:{row_number}'}
                        index = len(self.port_rows)
                        self.port_rows.append(record)
                        for alias in record['aliases']:
                            self.port_index[alias].append(index)
                        recognized += 1
            self._check_codes()
            if not recognized:
                self.issue('mapping', self.file_name, 'لم توجد أعمدة ربط مدعومة؛ ظهرت قيم المصادر الأصلية.', missing=True)
        except Exception as exc:
            self.issue('mapping', self.file_name,
                       'تعذر فتح ملف الربط (' + type(exc).__name__ + ')؛ ظهرت قيم المصادر الأصلية.', missing=True)

    def _check_codes(self):
        by_code = defaultdict(list)
        for row in self.port_rows:
            for code in row['code']:
                by_code[norm(code)].append(row)
        self.code_rows = by_code
        for code, rows in by_code.items():
            # Disjoint alias sets with the same code are not evidence of one physical port.
            components = []
            for row in rows:
                joined = set(row['aliases'])
                intersecting = [item for item in components if item & joined]
                for item in intersecting:
                    joined |= item
                    components.remove(item)
                components.append(joined)
            if len(components) > 1 or len({name for row in rows for name in row['arabic_ports']}) > 1:
                self.unsafe_codes.add(code)
                self.issue('port_code', rows[0]['code'][0],
                           'الكود مستخدم لسجلات مواقع غير محسومة؛ لم تُدمج المواقع على أساسه.')

    def translate(self, field, value, report_missing=True):
        original = self.text(value)
        if not original:
            return original
        scopes = [field] + (['name'] if field == 'manager' else []) + ['*']
        for scope in scopes:
            candidates = self.translations[scope].get(norm(original), set())
            if len(candidates) > 1:
                self.issue(field, original, 'توجد ترجمات عربية متعددة للنص نفسه؛ حُفظ نص المصدر.')
                return original
            if len(candidates) == 1:
                value = next(iter(candidates))
                if value != original:
                    self._translated.add((field, norm(original)))
                return value
        if report_missing and re.search('[A-Za-z]', original):
            self.issue(field, original, 'لا توجد ترجمة عربية صريحة وفريدة في ملف الربط؛ حُفظ نص المصدر.', missing=True)
        return original

    @staticmethod
    def _port_type(value):
        aliases = {
            'جوي': ['جوي', 'منفذ جوي', 'air', 'air port', 'airport', 'airports'],
            'بري': ['بري', 'منفذ بري', 'land', 'land port', 'land ports'],
            'بحري': ['بحري', 'منفذ بحري', 'sea', 'sea port', 'seaport', 'seaports', 'marine'],
        }
        return next((label for label, terms in aliases.items() if norm(value) in {norm(t) for t in terms}), None)

    def port(self, source):
        source = self.text(source)
        if norm(source) in self._port_cache:
            return dict(self._port_cache[norm(source)])
        fallback = {'port': self.translate('port', source), 'port_id': 'source:' + norm(source),
                    'port_code': '', 'type': 'غير محدد', 'region': 'غير محدد',
                    'type_raw': '', 'region_raw': '', 'latitude': None, 'longitude': None}
        rows = [self.port_rows[index] for index in self.port_index.get(norm(source), [])]
        if not rows:
            self._unmatched.add(norm(source))
            self.issue('port', source, 'لا توجد مطابقة موقع صريحة؛ لم تُفترض إحداثيات أو منطقة أو نوع منفذ.', missing=True)
            self._port_cache[norm(source)] = fallback
            return dict(fallback)
        codes = {value for row in rows for value in row['code']}
        names = {value for row in rows for value in row['arabic_ports']}
        if len(codes) > 1 or len(names) > 1:
            self._ambiguous.add(norm(source))
            self.issue('port', source, 'مطابقة الموقع متعددة ومختلفة في الاسم أو الكود؛ لم تُستخدم بياناتها.')
            fallback['port'] = source
            self._translated.discard(('port', norm(source)))
            self._port_cache[norm(source)] = fallback
            return dict(fallback)
        # Check every explicit alias of a trusted canonical code together. Inconsistent
        # coordinates cannot be hidden by looking up only one of the source aliases.
        if len(codes) == 1 and norm(next(iter(codes))) not in self.unsafe_codes:
            rows = self.code_rows[norm(next(iter(codes)))]
        self._matched.add(norm(source))
        if codes:
            code = next(iter(codes))
            fallback['port_code'] = code
            if norm(code) not in self.unsafe_codes:
                fallback['port_id'] = 'code:' + norm(code)
        for field in ('region', 'type'):
            raw = {value for row in rows for value in row[field]}
            fallback[field + '_raw'] = ' | '.join(sorted(raw))
            if field == 'region':
                translated = {self.translate(field, value) for value in raw}
            else:
                translated = {self._port_type(self.translate('type', value, report_missing=False)) for value in raw}
                if None in translated:
                    self.issue('type', source, 'تصنيف المنفذ غير معروف؛ ظهر «غير محدد».', missing=True)
            if len(translated) == 1 and None not in translated:
                fallback[field] = next(iter(translated))
            elif len(translated) > 1:
                self.issue(field, source, 'بيانات الموقع متعددة ومختلفة لهذا الحقل؛ ظهر «غير محدد».')
            elif not raw:
                self.issue(field, source, 'قيمة الحقل غير موجودة في سجل الموقع.', missing=True)
        points = set()
        invalid = False
        for row in rows:
            latitude, longitude = row['latitude'], row['longitude']
            if not latitude and not longitude:
                continue
            if len(latitude) != 1 or len(longitude) != 1:
                invalid = True
                continue
            try:
                lat, lon = float(latitude[0]), float(longitude[0])
                if not (math.isfinite(lat) and math.isfinite(lon) and -90 <= lat <= 90 and -180 <= lon <= 180
                        and 16.0 <= lat <= 33.0 and 34.0 <= lon <= 56.0):
                    invalid = True
                else:
                    points.add((lat, lon))
            except (TypeError, ValueError):
                invalid = True
        if invalid or len(points) > 1:
            self.issue('coordinates', source,
                       'الإحداثيات ناقصة أو متعددة أو خارج النطاق المتوقع للمملكة؛ لم تُرسم نقطة الموقع.')
        elif points:
            fallback['latitude'], fallback['longitude'] = next(iter(points))
        else:
            self.issue('coordinates', source, 'لا توجد إحداثيات للموقع؛ لم تُرسم نقطة افتراضية.', missing=True)
        self._port_cache[norm(source)] = fallback
        return dict(fallback)

    def summary(self, groups):
        return {
            'loaded': self.loaded, 'file_name': self.file_name, 'sheets': self.sheets,
            'issues': self.issues, 'issue_count': len(self.issues), 'missing_count': len(self._missing),
            'translated_count': len(self._translated), 'matched_source_ports': len(self._matched),
            'unmatched_source_ports': len(self._unmatched), 'ambiguous_source_ports': len(self._ambiguous),
            'total_ports': len({group['port_id'] for group in groups}),
            'located_ports': len({group['port_id'] for group in groups
                                  if group['latitude'] is not None and group['longitude'] is not None}),
            'policy': 'الربط للعرض فقط، بمطابقة نصية صريحة وفريدة، مع حفظ القيم الأصلية ومفاتيح الحساب.',
        }


def extract_data(result):
    """Build display data from the same cases; display labels do not change grouping."""
    mapping = MappingResolver(MAPPING_FILE)
    groups = []
    group_ids = {}
    for number, (key, members) in enumerate(sorted(result['group_staff'].items()), 1):
        group_id = f'G{number:04d}'
        group_ids[key] = group_id
        manager = result['managers_by_id'][key[0]]
        levels = result['display_levels'][key]
        raw = {'manager': manager['name'], 'title': manager['job'],
               'port': result['display_port'][key],
               'level1': levels[0], 'level2': levels[1], 'level3': levels[2]}
        port = mapping.port(raw['port'])
        raw.update({'region': port['region_raw'], 'type': port['type_raw']})
        group = {'id': group_id, 'manager_id': key[0], 'workforce': len(members),
                 **{field + '_raw': value for field, value in raw.items()}, 'raw': raw}
        group.update({field: mapping.translate(field, raw[field])
                      for field in ('manager', 'title', 'level1', 'level2', 'level3')})
        group.update(port)
        if not group['title']:
            mapping.issue('title', raw['manager'], 'المسمى الوظيفي للمدير غير موجود في ملف الموظفين.', missing=True)
        groups.append(group)

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
                'name_raw': result['employee_display_names'].get(employee_id) or 'اسم غير متوفر في المصدر',
                'name': mapping.translate('name', result['employee_display_names'].get(employee_id)
                                           or 'اسم غير متوفر في المصدر'),
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

    planner = extract_planner(result, group_ids, cases)
    for leave in planner['leaves']:
        leave['name_raw'] = leave['name']
        leave['name'] = mapping.translate('name', leave['name'])

    data = {
        'meta': {
            'script_version': SCRIPT_VERSION,
            'title': REPORT_TITLE, 'department': DEPARTMENT,
            'prepared_by': PREPARED_BY, 'supervised_by': SUPERVISED_BY,
            'as_of': result['asof'].isoformat(),
            'period_start': result['pstart'].isoformat(), 'period_end': result['pend'].isoformat(),
            'threshold': THRESHOLD, 'end_date_inclusive': END_DATE_INCLUSIVE,
        },
        'groups': groups,
        'cases': cases,
        'planner': planner,
        'mapping': mapping.summary(groups),
        'methodology': {
            'case_definition': 'الموظفون أنفسهم (اثنان فأكثر) في أيام تقويمية متصلة ضمن المدير والموقع والمسار التنظيمي نفسه؛ تغيّر الأسماء يبدأ حالة جديدة.',
            'percentage_definition': 'عدد الموظفين المجازين في الحالة ÷ إجمالي موظفي مجموعتها × 100. لا تتغير المجموعة أو المقام بسبب الترجمة أو توحيد اسم المنفذ للعرض.',
            'future_definition': 'الحالات المستقبلية في المنسّق يبدأ تقاطعها بعد تاريخ العرض، مع تثبيت الإجازات التي بدأت والإجازات غير المستهدفة.',
            'week_definition': 'الأسبوع من الأحد إلى السبت، مع قص الأسبوع الأول والأخير عند حدود فترة التقرير.',
            'weekly_count_definition': 'عدد حالات التقرير الفريدة التي تمتد إلى كل أسبوع ضمن التصفية؛ الحالة الممتدة إلى أسبوعين تظهر مرة في كل أسبوع، لذا لا يُجمع المنحنى لإجمالي الحالات.',
            'port_count_definition': 'عدد معرفات الحالات الفريدة المرتبطة بالمنفذ ضمن التصفية؛ يعرض الكود الموثوق اسم الموقع بصورة موحدة دون دمج مجموعات الحساب.',
            'mapping_policy': 'الترجمة وبيانات الموقع من Mapping فقط عند وجود مطابقة صريحة وفريدة، وإلا يبقى نص المصدر وتظهر ملاحظة ربط.',
            'classification_date': result['asof'].isoformat(),
            'presentation_date': PRESENTATION_DATE,
            'week_start': 'الأحد', 'week_end': 'السبت',
        },
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
    print('مصادر الحساب: ملف الموظفين وملف الإجازات. Mapping للترجمة وبيانات المواقع فقط؛ لا يلزم تقرير مرجعي.', flush=True)
    result = calculate()
    data = extract_data(result)
    mapping = data['mapping']
    if mapping['loaded']:
        print(f"تمت قراءة الربط: {mapping['file_name']} | أوراق: {', '.join(mapping['sheets'])} | مواقع بإحداثيات: {mapping['located_ports']} من {mapping['total_ports']}", flush=True)
    else:
        print('لم يُقرأ ملف Mapping؛ ستظهر قيم المصادر الأصلية. راجعي المسار في MAPPING_FILE.', flush=True)
    if mapping['issue_count']:
        print(f"ملاحظات الربط: {mapping['issue_count']}؛ التفاصيل في قسم ملاحظات الربط بالداشبورد.", flush=True)
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

/* Mapping labels and shared report filters. */
.filters{grid-template-columns:repeat(4,minmax(0,1fr));position:relative;z-index:5}
.filters .search-field{grid-column:1 / span 3}
.filters .reset-button{grid-column:4}
.filters .filter-scope-note{grid-column:1 / -1;margin:1px 0 0;font-size:11px;color:var(--muted)}
.filters .field{min-width:0}
.port-combobox{position:relative}
.port-input-row{display:flex;align-items:stretch;border:1px solid var(--line);border-radius:8px;background:white;overflow:hidden}
.port-input-row:focus-within{outline:3px solid #00828e55;outline-offset:2px}
.port-input-row input{min-width:0;border:0;border-radius:0;outline:none!important;padding-inline-end:5px}
.port-input-row button{border:0;border-radius:0;padding:0;min-width:33px;min-height:44px;flex:0 0 33px;font-size:18px}
.port-input-row .port-clear{font-size:20px;font-weight:400}
.port-options{position:absolute;inset-inline:0;top:calc(100% + 5px);max-height:300px;overflow:auto;overscroll-behavior:contain;background:#fff;border:1px solid #bdceca;border-radius:8px;box-shadow:0 14px 30px #12375226;z-index:30;padding:5px}
.port-options[hidden],.port-clear[hidden]{display:none!important}
.port-option{padding:10px 11px;min-height:45px;cursor:pointer;border-radius:5px;overflow-wrap:anywhere;font-size:12px}
.port-option strong{display:block;font-weight:700;line-height:1.6}
.port-option small{display:block;color:var(--muted);font-size:10px;line-height:1.6}
.port-option:hover,.port-option[aria-selected="true"]{background:#e4f1ef;color:#075d61}
.port-no-results{padding:13px 10px;color:var(--muted);font-size:12px}
.port-help{font-size:10px;color:var(--muted);min-height:16px;margin:4px 0 0}
.mapping-statrow{display:flex;gap:12px 20px;flex-wrap:wrap;margin:10px 0 15px}
.mapping-statrow span{font-size:11px;color:var(--muted)}
.mapping-statrow b{font-size:17px;color:var(--navy);margin-inline-end:5px}
.mapping-table-wrap{max-height:340px;overflow:auto;border:1px solid var(--line);border-radius:7px;background:var(--paper)}
.mapping-table-wrap table{font-size:11px;min-width:500px}
.mapping-table-wrap td{white-space:normal;overflow-wrap:anywhere;vertical-align:top}
.mapping-table-wrap th{position:sticky;top:0;background:#eef2ee;z-index:1}
.mapping-originals{margin:16px 0;padding:11px 14px;border:1px solid var(--line);border-radius:7px;font-size:12px;color:var(--muted)}
.mapping-originals summary{cursor:pointer;color:var(--navy);font-weight:700}
.mapping-originals dl{display:grid;grid-template-columns:140px minmax(0,1fr);gap:7px 12px;margin:12px 0 0}
.mapping-originals dd{margin:0;overflow-wrap:anywhere}
#smart-leave-calendar .sp-events{grid-auto-rows:62px}
#smart-leave-calendar .sp-event{min-height:60px}
#smart-leave-calendar .sp-event-person{display:flex;flex:1;min-width:0;flex-direction:column;gap:2px;align-items:flex-start}
#smart-leave-calendar .sp-event-name{display:block;max-width:100%;flex:0 1 auto;font-size:13px;line-height:1.6}
#smart-leave-calendar .sp-event-role{display:block;max-width:100%;font-size:10px;font-weight:400;line-height:1.6;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;opacity:.92}
#smart-leave-calendar .sp-manager-role{font-size:12px;color:var(--sp-muted);line-height:1.8;margin-top:3px}
@media(max-width:1080px){.filters{grid-template-columns:repeat(2,minmax(0,1fr))}.filters .search-field{grid-column:1}.filters .reset-button{grid-column:2}}
@media(max-width:700px){.filters{grid-template-columns:1fr}.filters .search-field,.filters .reset-button{grid-column:1}.mapping-originals dl{grid-template-columns:1fr}.mapping-originals dd{margin-bottom:7px}}
@media(max-width:540px){#smart-leave-calendar .sp-events{grid-auto-rows:64px}#smart-leave-calendar .sp-event-role{font-size:10px}#smart-leave-calendar .sp-event-person{max-width:calc(100% - 95px)}}
/* Port and weekly analytics: local SVG, one case-counting definition. */
.pa-panel,.wa-panel{margin-top:24px;scroll-margin-top:24px}
.pa-header{gap:18px;align-items:flex-start;margin-bottom:18px}
.pa-header .section-subtitle{max-width:830px;line-height:1.85;margin-top:6px}
.pa-scope-note{display:inline-block;white-space:nowrap;font-size:11px;color:var(--muted,#65757e);background:#f0f3ef;border:1px solid var(--line,#dce2e0);border-radius:5px;padding:6px 10px}
.pa-stats{display:flex;flex-wrap:wrap;gap:10px;margin:0 0 18px}
.pa-stat{display:flex;align-items:baseline;gap:8px;padding:8px 13px;border-radius:7px;background:#f3f6f2;color:#526773;font-size:12px}
.pa-stat b{font-size:20px;color:var(--navy,#123752);font-variant-numeric:tabular-nums}
.pa-stat.pa-stat-warning{background:#faf1e5;color:#82551f}
.pa-stat.pa-stat-warning b{color:#82551f}
.pa-grid{display:grid;grid-template-columns:minmax(0,1.7fr) minmax(280px,1fr);gap:22px;align-items:start}
.pa-map-column,.pa-list-column{min-width:0}
.pa-map-frame{position:relative;overflow:hidden;border:1px solid #dce5e1;background:#eff4f3;border-radius:9px}
.pa-map{display:block;width:100%;height:auto;aspect-ratio:1200/800;max-height:650px;touch-action:none;cursor:grab}
.pa-map:active{cursor:grabbing}
.pa-map-tools{position:absolute;top:10px;left:10px;z-index:1;display:flex;gap:4px;direction:ltr}
.pa-map-tools button{min-width:34px;min-height:34px;height:34px;padding:3px 9px;font-size:13px;border:1px solid #d3ded8;background:#fffefa;border-radius:5px;color:var(--navy,#123752);box-shadow:0 2px 6px #1237520b}
.pa-map-tools button:first-child,.pa-map-tools button:nth-child(2){font-size:20px}
.pa-region-label{font:600 18px Tahoma,Arial,sans-serif;fill:#726d60;paint-order:stroke;stroke:#f3f0e8;stroke-width:3px;pointer-events:none}
.pa-marker{cursor:pointer;outline:none}
.pa-marker .pa-pin-head{transition:stroke .12s,filter .12s}
.pa-marker:hover .pa-pin-head,.pa-marker:focus-visible .pa-pin-head{stroke:#ae641c;stroke-width:3px;filter:drop-shadow(0 2px 4px #12375240)}
.pa-marker:focus-visible .pa-marker-hit{fill:#d48b3033;stroke:#ae641c;stroke-width:2}
.pa-map-legend{display:flex;align-items:center;justify-content:center;flex-wrap:wrap;gap:10px 20px;margin:14px 0 6px;font-size:12px;color:#526873}
.pa-map-legend span{display:flex;align-items:center;gap:7px}
.pa-type-icon{display:inline-flex;align-items:center;justify-content:center;flex:0 0 28px;width:28px;height:28px;border-radius:50%;background:var(--pa-icon-color,#708183);color:#fff}
.pa-type-icon svg{width:17px;height:17px;overflow:visible}
.pa-map-hint,.pa-map-credit{font-size:11px;line-height:1.9;color:var(--muted,#65757e);text-align:center;margin:4px 0}
.pa-map-credit{font-size:10px;color:#819087}
.pa-list-heading{display:flex;align-items:center;justify-content:space-between;gap:10px;margin:0 0 10px}
.pa-list-heading h3{font-size:14px;margin:0;color:var(--navy,#123752)}
.pa-list-heading>span{color:var(--muted,#65757e);font-size:11px}
.pa-port-list{max-height:600px;overflow-y:auto;overscroll-behavior:contain;scrollbar-gutter:stable;padding-left:3px}
.pa-port-button{display:flex;align-items:center;gap:10px;width:100%;min-height:68px;padding:12px 10px;margin-bottom:7px;background:#fffefa;border:1px solid #dfe6df;border-radius:7px;text-align:right;color:var(--ink,#17334a)}
.pa-port-button:hover,.pa-port-button:focus-visible{background:#ecf5f0;border-color:#88b8af}
.pa-port-main{flex:1;min-width:0}
.pa-port-name{display:block;font-size:12px;line-height:1.75;overflow-wrap:anywhere}
.pa-port-description{display:block;font-size:10px;color:#77837b;line-height:1.8;margin-top:1px}
.pa-port-count{display:flex;flex-direction:column;align-items:center;flex:0 0 48px;color:var(--teal,#087f83)}
.pa-port-count b{font-size:23px;line-height:1.15;font-variant-numeric:tabular-nums}
.pa-port-count small{font-size:10px;color:#6f7d7a;margin-top:3px}
.pa-coordinate-tag{color:#946020;display:inline-block;margin-top:3px;font-size:10px}
.pa-empty{padding:30px 16px;text-align:center;line-height:1.9;font-size:13px;color:var(--muted,#65757e)}
.pa-location-notice{padding:12px 15px;background:#faf3e8;border:1px solid #eadac0;border-radius:7px;color:#76532d;font-size:12px;line-height:1.9;margin-top:16px}
.pa-location-notice[hidden]{display:none}
.pa-location-notice button{font-size:11px;min-height:30px;margin:5px 4px 0 0;padding:4px 9px;background:#fffefa;border:1px solid #dbc8aa;border-radius:5px;color:#76532d}
.pa-dialog-meta{display:flex;flex-wrap:wrap;gap:8px;margin:12px 0 20px}
.pa-dialog-meta span{background:#edf3ed;border:1px solid #dde5dd;border-radius:6px;padding:6px 10px;font-size:12px;color:#456356}
.pa-manager-card,.wa-case-card{padding:15px;border:1px solid #dce4dd;border-radius:8px;margin:10px 0;background:#fffefa}
.pa-manager-line,.wa-case-heading{display:flex;gap:16px;justify-content:space-between;align-items:flex-start}
.pa-manager-name{display:block;font-size:14px;color:var(--navy,#123752);line-height:1.75}
.pa-manager-role{display:block;margin-top:2px;font-size:12px;font-weight:400;color:#667971;line-height:1.8}
.pa-manager-meta{font-size:12px;margin-top:7px;color:#667971;line-height:1.8}
.pa-manager-case-count{flex-shrink:0;color:var(--teal,#087f83);font-size:13px;white-space:nowrap}
.pa-case-buttons{display:flex;flex-wrap:wrap;gap:7px;margin-top:12px}
.pa-case-buttons button,.wa-case-card button,.pa-filter-button{padding:7px 12px;min-height:38px;font-size:12px;border:1px solid #cad9cf;border-radius:5px;background:#f2f7f1;color:#194f43}
.pa-filter-button{margin-top:16px;background:var(--navy,#123752);border-color:var(--navy,#123752);color:#fff;min-height:42px}
.pa-dialog-empty{padding:20px;color:var(--muted,#65757e);line-height:1.8}
.wa-peak{display:flex;align-items:center;gap:24px;background:#eef5ee;border:1px solid #dce6d9;border-radius:8px;padding:15px 18px;margin-bottom:18px}
.wa-peak.wa-no-data{background:#f3f4ef;border-color:#e1e4dc;font-size:13px;color:#6e7a70}
.wa-peak-heading{display:flex;flex-direction:column;gap:2px;flex-shrink:0;min-width:116px}
.wa-peak-heading>span{font-size:12px;color:#567267}
.wa-peak-heading strong{font-size:30px;line-height:1.5;color:#0c6257;font-variant-numeric:tabular-nums}
.wa-peak-heading small{font-size:12px;font-weight:400;margin-right:8px}
.wa-peak-weeks{display:flex;flex-wrap:wrap;gap:8px;align-items:center}
.wa-peak-weeks button{display:flex;flex-direction:column;align-items:flex-start;gap:2px;min-height:48px;background:#fffefa;border:1px solid #cadbcb;border-radius:6px;padding:7px 12px;color:#174a40;font-size:12px;text-align:right}
.wa-peak-weeks button:hover{border-color:#489585;background:#e0eee3}
.wa-peak-weeks small{color:#668073;font-size:10px}
.wa-peak-ties{font-size:11px;color:#60796c;flex:1 1 100%;margin-bottom:2px}
.wa-chart-legend{display:flex;flex-wrap:wrap;gap:13px 22px;margin:0 0 12px;font-size:11px;color:#63766e}
.wa-chart-legend span{display:inline-flex;align-items:center;gap:7px}
.wa-line-key{width:24px;border-top:3px solid #087f83;position:relative}
.wa-line-key:after{content:'';position:absolute;left:9px;top:-5px;width:7px;height:7px;background:#fffefa;border:2px solid #087f83;border-radius:50%}
.wa-past-key,.wa-current-key{width:14px;height:14px;border:1px solid #d6d9d0;background:#edecea;border-radius:3px}
.wa-current-key{background:#dff0e8;border-color:#a6c7b5}
.wa-chart-scroll{overflow-x:auto;overscroll-behavior:contain;border:1px solid #e0e6dd;border-radius:8px;background:#fffefa;scrollbar-gutter:stable}
.wa-chart-scroll svg{display:block;width:100%;min-width:940px;height:auto;direction:ltr;font-family:Tahoma,Arial,sans-serif}
.wa-week-point{cursor:pointer;outline:none}
.wa-week-point .wa-dot{transition:stroke-width .15s}
.wa-week-point:hover .wa-dot,.wa-week-point:focus-visible .wa-dot{stroke:#b66a22;stroke-width:4px}
.wa-week-point:focus-visible .wa-focus-ring{fill:#b66a2220;stroke:#b66a22;stroke-width:2px}
.wa-count-note{font-size:11px;color:#708177;line-height:1.85;margin:12px 0 0}
.wa-case-dates{font-size:12px;line-height:1.9;color:#60766b;margin:10px 0}
.wa-case-count{font-size:12px;white-space:nowrap;color:#486355}
.wa-details-note{font-size:12px;line-height:1.9;color:#6b7d71;background:#f1f4ed;padding:12px;border-radius:6px}
@media(min-width:1450px){.pa-port-list{max-height:640px}}
@media(max-width:1050px){.pa-grid{grid-template-columns:minmax(0,1fr) 280px;gap:15px}.pa-region-label{font-size:18px}.pa-map-tools{top:7px;left:7px}.pa-map-tools button{font-size:11px;padding:3px 7px}.pa-port-button{padding:10px 8px}.pa-type-icon{flex-basis:25px;width:25px;height:25px}.pa-port-list{max-height:430px}.pa-map-legend{gap:8px 12px}.pa-port-name{font-size:11px}}
@media(max-width:820px){.pa-grid{grid-template-columns:1fr}.pa-port-list{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;max-height:330px}.pa-port-button{margin:0}.pa-list-heading{margin-top:8px}.pa-map-tools button{min-height:38px;height:38px;min-width:38px}.pa-header{display:block}.pa-scope-note{margin-top:8px}.pa-map{max-height:none}.wa-peak{gap:15px;align-items:flex-start}.wa-peak-heading{min-width:105px}.wa-peak-heading strong{font-size:26px}}
@media(max-width:520px){.pa-port-list{grid-template-columns:1fr;max-height:350px}.pa-map-frame{margin-inline:-4px}.pa-stat{padding:6px 9px;gap:6px;font-size:11px}.pa-stat b{font-size:18px}.wa-peak{display:block;padding:12px}.wa-peak-heading{flex-direction:row;align-items:center;justify-content:space-between;margin-bottom:10px}.wa-peak-weeks button{flex:1 1 135px}.pa-manager-line,.wa-case-heading{gap:10px;flex-wrap:wrap}.pa-map-legend{font-size:11px}.pa-map-hint{font-size:10px}}
@media print{.pa-map-tools{display:none!important}.pa-grid{grid-template-columns:1.4fr 1fr}.pa-port-list{max-height:none;overflow:visible}.pa-port-button{break-inside:avoid}.wa-chart-scroll{overflow:visible}.wa-chart-scroll svg{min-width:0}.pa-panel,.wa-panel{break-inside:avoid}.wa-peak-weeks button{display:flex!important}.pa-location-notice button{display:inline-block!important}}

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
      <div class="field"><label for="title-filter">المسمى الوظيفي للمدير</label><select id="title-filter"></select></div>
      <div class="field"><label for="type-filter">نوع المنفذ</label><select id="type-filter"><option value="">جميع الأنواع</option><option value="جوي">جوي</option><option value="بري">بري</option><option value="بحري">بحري</option><option value="غير محدد">غير محدد</option></select></div>
      <div class="field"><label for="port-search">اسم المنفذ</label><div class="port-combobox" id="port-combobox"><input type="hidden" id="port-filter"><div class="port-input-row"><input type="search" id="port-search" role="combobox" aria-expanded="false" aria-controls="port-options" aria-autocomplete="list" aria-describedby="port-help" placeholder="ابحثي أو اختاري المنفذ" autocomplete="off"><button class="port-clear" id="port-clear" type="button" aria-label="إزالة تصفية المنفذ" title="إزالة تصفية المنفذ" hidden>×</button><button id="port-toggle" type="button" aria-label="عرض قائمة المنافذ" aria-controls="port-options" aria-expanded="false">⌄</button></div><div id="port-options" class="port-options" role="listbox" aria-label="المنافذ" hidden></div></div><p class="port-help" id="port-help">البحث بالاسم العربي أو الأصلي أو رمز المنفذ.</p></div>
      <div class="field search-field"><label for="group-search">البحث في المجموعات</label><input type="search" id="group-search" placeholder="اسم المدير، المسمى، المنفذ، أو أحد المستويات التنظيمية"></div>
      <button class="reset-button" id="reset-filters" type="button">إعادة ضبط</button>
      <p class="filter-scope-note">هذه الفلاتر تطبّق على المؤشرات والسجل والخريطة والشارت والمنسّق الذكي.</p>
    </div>
    <div class="kpis" aria-label="مؤشرات حالات التقاطع" aria-live="polite">
      <article class="kpi"><div class="kpi-head"><span>جميع حالات التقاطع</span><span class="kpi-index" aria-hidden="true">01</span></div><div class="kpi-value" id="kpi-all">—</div><div class="kpi-note">بغض النظر عن نسبة الغياب</div></article>
      <article class="kpi"><div class="kpi-head"><span id="threshold-kpi-label">تقاطع 50% فأكثر</span><span class="kpi-index" aria-hidden="true">02</span></div><div class="kpi-value" id="kpi-threshold">—</div><div class="kpi-note">جزء من جميع حالات التقاطع</div></article>
      <article class="kpi"><div class="kpi-head"><span>الحالات الجارية</span><span class="kpi-index" aria-hidden="true">03</span></div><div class="kpi-value" id="kpi-ongoing">—</div><div class="kpi-note">من جميع الحالات، حسب تاريخ التصنيف</div></article>
      <article class="kpi"><div class="kpi-head"><span>الحالات القادمة</span><span class="kpi-index" aria-hidden="true">04</span></div><div class="kpi-value" id="kpi-upcoming">—</div><div class="kpi-note">من جميع الحالات، بعد تاريخ التصنيف</div></article>
    </div>
    <section class="highest" aria-labelledby="highest-title"><div class="highest-score"><h3 id="highest-title">أعلى نسبة غياب مجدول</h3><div class="highest-number" id="highest-number">—</div><p id="highest-ratio"></p></div><div id="highest-details" class="highest-main"></div></section>
  </section>
  <section class="panel pa-panel" id="ports-analytics" aria-labelledby="pa-title">
  <div class="section-header pa-header"><div><h2 class="section-title" id="pa-title">التقاطعات حسب المنفذ</h2><p class="section-subtitle">عدد حالات التقاطع لجميع النسب ضمن فلاتر المجموعات المحددة. اختر المنفذ لعرض المديرين ومسمياتهم الوظيفية.</p></div><span class="pa-scope-note">من واقع حالات التقرير</span></div>
  <div class="pa-stats" id="pa-stats" aria-live="polite"></div>
  <div class="pa-grid">
    <div class="pa-map-column">
      <div class="pa-map-frame">
        <div class="pa-map-tools"><button type="button" id="pa-zoom-in" aria-label="تكبير الخريطة">+</button><button type="button" id="pa-zoom-out" aria-label="تصغير الخريطة">−</button><button type="button" id="pa-map-reset">إعادة الخريطة</button></div>
        <svg id="pa-map" class="pa-map" viewBox="0 0 1200 800" role="group" aria-label="خريطة مناطق المملكة والمنافذ؛ افتح أي منفذ لعرض حالات التقاطع" style="direction:ltr">
          <rect width="1200" height="800" fill="#eff4f3"></rect>
          <g id="pa-map-layer">
<path d="M804.13,354.26L804.18,354.57L804.34,354.89L804.55,354.99L804.70,355.14L804.98,355.19L805.24,355.50L805.49,355.59L805.83,356.64L805.86,356.71L806.01,356.72L806.23,357.22L806.02,356.35L805.75,355.66L805.81,354.87L805.74,354.71L805.52,354.50L804.97,354.23L804.47,354.16L804.13,354.26ZM776.02,287.07L776.05,287.33L776.27,287.27L776.21,286.99L776.45,286.57L776.45,286.39L776.23,286.41L776.26,286.60L776.02,287.07ZM772.27,285.16L772.32,285.34L772.28,285.51L772.42,285.68L772.33,286.08L772.36,286.20L772.36,286.03L772.48,285.86L772.69,285.91L772.67,285.99L772.88,286.14L773.08,286.48L773.10,286.42L772.90,285.94L772.86,285.92L772.87,286.02L772.95,286.18L772.73,285.99L772.73,285.72L772.66,285.51L772.36,285.18L772.27,285.16ZM596.04,262.37L596.88,262.86L598.28,263.86L598.62,264.33L599.04,265.17L599.19,268.16L601.90,272.73L602.63,273.73L603.12,274.67L603.46,275.64L603.77,277.65L604.23,279.56L604.68,280.46L605.63,281.31L606.51,281.96L609.02,283.51L611.90,285.10L612.84,285.54L616.55,287.81L617.45,288.25L619.31,288.86L620.32,288.94L623.12,289.02L623.97,289.46L624.76,290.30L626.21,292.30L627.91,294.27L630.00,296.25L630.92,297.02L632.51,298.14L633.37,298.58L635.18,298.86L637.04,298.94L638.79,298.86L641.58,298.98L642.43,299.10L644.99,299.90L646.65,300.82L648.25,301.90L650.88,303.81L653.52,305.92L656.18,307.83L657.06,308.40L659.02,309.19L660.72,309.30L661.69,309.24L666.70,308.08L667.60,308.09L668.46,308.34L669.30,308.98L670.92,310.67L671.78,311.45L672.60,312.09L674.96,314.54L675.79,315.27L676.27,316.17L676.66,317.11L677.45,318.15L678.31,318.48L680.13,318.89L680.99,319.18L686.52,321.42L687.54,321.76L690.24,323.00L691.05,323.48L691.92,324.14L692.55,325.06L692.86,325.93L693.16,327.06L693.47,329.95L693.64,332.72L693.70,342.16L694.09,347.93L694.04,348.97L694.11,351.03L693.93,352.82L693.91,355.88L693.97,357.57L694.08,358.55L693.67,363.00L693.36,375.76L693.75,384.21L694.14,387.05L694.52,388.44L695.24,389.72L695.80,390.48L698.39,392.54L700.96,394.00L705.78,397.20L709.11,399.74L714.54,405.37L719.54,410.19L721.44,412.54L723.45,416.10L724.80,419.14L725.54,422.05L726.34,427.01L726.32,435.86L725.51,442.74L723.94,451.90L716.54,512.94L714.35,531.76L712.01,549.04L710.25,561.60L708.53,575.26L706.47,591.29L704.24,610.37L702.61,621.89L701.09,634.65L699.26,649.41L697.91,660.34L695.11,682.91L694.11,698.25L693.35,709.88L698.57,696.23L708.31,683.65L721.39,666.78L739.63,657.49L757.91,648.22L781.23,645.73L797.54,643.98L823.11,641.34L848.21,636.62L870.71,632.38L887.25,626.53L906.59,619.69L934.62,609.78L962.79,599.82L988.08,590.88L997.14,561.83L1004.61,537.89L1014.16,507.06L996.26,477.08L995.53,477.84L993.45,480.35L969.99,477.13L954.74,475.09L941.42,473.34L922.79,470.95L893.44,467.30L877.10,445.87L867.59,433.28L861.24,425.02L854.68,416.61L854.63,411.14L854.31,411.31L854.31,411.16L854.30,411.31L854.09,411.27L853.59,411.42L852.84,411.41L852.48,411.30L852.14,410.94L852.08,410.62L851.82,410.45L851.70,410.11L851.62,410.07L851.62,409.91L851.29,409.66L851.16,409.35L850.95,409.24L850.88,409.08L850.87,409.12L850.74,409.10L850.81,409.05L850.53,408.94L850.60,408.88L850.69,408.93L850.69,408.89L850.17,408.75L850.05,408.52L849.60,408.56L849.16,408.47L849.01,408.30L849.07,408.22L848.89,408.10L848.79,408.12L848.80,408.20L848.74,408.21L848.08,408.24L847.42,408.00L847.43,408.15L846.96,408.23L846.55,408.45L846.13,409.25L845.85,409.46L845.81,409.57L845.52,409.67L844.90,409.66L844.62,409.62L844.28,409.48L844.18,409.53L843.90,409.44L843.57,409.56L843.41,409.68L843.59,409.59L843.56,409.66L843.17,409.88L842.95,409.60L842.70,409.47L842.62,409.48L842.64,409.57L842.84,410.08L842.58,409.84L842.45,409.58L842.44,409.52L842.49,409.48L842.41,409.44L842.41,409.38L842.47,409.40L842.45,409.32L842.38,409.34L842.37,409.14L842.48,409.04L842.51,408.90L842.56,409.24L842.63,409.20L842.57,408.89L842.64,408.88L842.62,408.69L842.69,408.52L842.68,408.61L842.75,408.59L842.77,408.46L842.85,408.47L842.74,407.75L842.87,407.78L842.87,407.73L842.77,407.44L842.62,407.40L842.59,407.57L842.52,407.61L842.52,407.44L842.70,407.31L842.74,407.41L842.80,407.36L842.85,407.52L842.84,407.33L842.72,407.16L842.82,406.98L842.88,406.96L842.89,407.22L842.94,407.05L843.11,406.92L843.29,406.82L843.28,406.89L843.38,406.93L843.33,406.81L843.38,406.83L843.40,406.76L843.52,406.82L843.50,406.67L843.28,406.51L843.43,406.60L843.32,406.42L843.36,406.34L843.30,406.23L843.42,406.19L843.36,406.02L843.44,405.94L843.51,405.72L843.57,405.26L843.54,405.18L843.38,405.32L843.36,405.26L843.39,405.17L843.41,405.22L843.47,405.19L843.42,405.12L843.52,405.05L843.56,404.69L843.70,404.58L843.72,404.62L843.74,404.45L843.80,404.48L843.83,404.42L843.90,404.11L843.95,404.11L843.91,404.20L843.96,404.26L844.01,404.20L844.00,404.40L844.23,403.61L844.77,402.98L845.28,402.62L845.70,402.17L845.83,402.07L846.30,402.11L846.59,401.64L846.70,401.56L846.92,401.53L847.09,401.14L847.27,400.89L847.65,400.78L847.84,400.35L847.76,399.94L847.86,399.50L848.09,399.15L848.31,399.00L848.57,398.99L848.70,399.13L848.67,399.20L848.59,399.19L848.64,399.29L848.79,399.12L849.04,399.18L849.03,399.11L848.93,399.09L848.90,399.02L848.91,398.77L849.04,398.49L849.26,398.19L849.47,398.07L849.60,398.16L849.85,398.09L849.97,398.14L849.99,398.21L849.92,398.20L849.99,398.22L850.07,398.13L850.05,398.03L850.14,397.98L850.66,398.09L850.70,397.99L850.54,397.98L850.58,397.80L850.70,397.73L850.86,397.72L850.89,397.77L851.12,397.60L851.26,397.59L851.27,397.49L851.35,397.52L851.34,397.39L851.07,397.29L851.11,397.24L851.01,397.12L850.94,397.29L850.82,397.26L850.75,397.13L850.71,397.17L850.65,397.13L850.56,396.92L850.56,396.85L850.62,396.86L850.60,396.72L850.46,396.52L850.39,396.54L850.24,396.43L850.30,396.55L850.25,396.72L850.17,396.71L850.17,396.65L850.15,396.72L850.08,396.70L850.08,396.62L850.05,396.68L849.98,396.63L850.02,396.74L849.83,396.68L849.81,396.76L849.58,396.73L849.66,397.03L849.51,397.16L849.29,397.17L849.17,396.83L849.18,396.59L849.33,396.54L849.29,396.37L849.18,396.25L849.19,396.14L849.29,396.12L849.28,396.01L849.20,396.06L849.24,395.91L849.14,395.87L849.17,395.80L849.07,395.74L848.91,395.40L848.83,395.41L848.84,395.33L848.75,395.32L848.78,395.37L848.73,395.37L848.69,395.32L848.70,395.37L848.63,395.38L848.45,395.34L848.44,395.45L848.52,395.53L848.42,395.64L848.42,395.80L848.25,395.84L848.18,395.66L847.98,395.63L847.96,395.74L847.85,395.72L847.84,395.67L847.82,395.71L847.73,395.69L847.58,395.48L847.51,395.48L847.50,395.37L847.57,395.34L847.59,395.23L847.30,395.19L847.14,395.27L847.20,395.36L847.15,395.63L847.26,395.66L847.33,395.80L847.32,396.29L846.73,396.90L846.59,397.25L846.12,397.47L845.74,397.55L845.52,397.74L845.50,397.94L845.37,398.00L845.30,398.16L845.17,398.23L845.14,398.41L845.39,398.61L845.40,398.70L845.23,398.78L845.14,398.96L845.27,398.99L845.28,399.14L845.17,399.17L845.09,399.11L845.10,399.29L844.49,399.57L844.42,399.77L844.13,399.90L844.13,399.97L843.95,400.04L843.91,400.16L843.73,400.24L843.74,400.29L843.49,400.39L841.98,400.36L840.21,400.89L838.26,401.22L837.67,401.51L835.44,401.85L832.55,400.92L831.37,400.43L829.01,398.81L828.78,398.58L826.63,394.32L825.45,392.15L824.22,390.08L823.77,390.11L823.70,390.01L823.69,390.08L823.41,390.09L823.20,390.43L822.99,390.54L822.63,391.21L822.62,391.14L822.54,391.12L822.51,390.96L822.60,390.93L822.51,390.93L822.47,390.83L822.32,389.89L822.21,389.66L821.90,389.34L821.69,388.40L821.63,388.05L821.77,387.20L821.68,386.85L821.47,386.75L821.23,386.44L820.99,385.89L820.99,385.58L821.06,385.41L820.95,385.04L820.94,384.59L820.55,384.35L820.50,384.09L820.10,383.81L819.97,383.79L819.49,383.88L819.24,383.78L819.16,383.55L818.80,382.98L818.83,382.14L818.91,381.92L818.81,381.44L818.67,381.31L818.63,381.13L818.34,380.88L818.28,380.55L818.07,380.22L818.05,380.02L817.79,379.74L817.60,379.63L817.58,379.54L817.46,379.48L817.21,379.15L816.86,378.36L816.59,378.04L816.32,377.36L816.00,377.15L815.81,376.75L815.66,376.56L815.49,376.48L815.40,376.24L815.23,376.05L815.11,376.06L815.00,376.16L814.95,376.36L815.02,376.81L814.95,377.00L815.41,377.70L815.28,377.67L815.27,377.58L814.93,377.66L814.71,377.43L814.57,377.39L814.44,377.18L814.49,377.20L814.43,377.03L814.50,376.61L814.65,376.39L814.79,376.29L814.75,376.13L814.58,375.99L814.37,375.50L814.23,374.32L813.81,373.30L813.87,373.04L814.30,372.68L814.38,372.05L814.32,371.15L814.17,370.73L813.85,370.26L813.23,369.73L812.93,369.07L812.82,368.37L813.09,367.68L813.30,367.35L813.28,366.91L813.17,366.59L813.27,366.05L813.10,365.48L812.80,364.77L812.43,364.25L812.36,363.61L812.23,363.13L812.10,362.89L811.77,362.58L811.72,362.05L811.77,361.89L811.75,362.03L811.80,362.00L811.81,361.84L811.75,361.63L811.23,361.00L810.84,360.74L810.57,360.40L810.07,360.05L810.01,359.93L809.30,359.86L809.16,359.79L808.71,359.32L808.38,358.84L807.99,358.64L807.88,358.32L807.68,358.25L807.57,358.06L807.59,357.88L807.70,357.78L807.84,357.78L807.85,357.91L807.92,357.92L807.93,357.55L807.48,356.96L807.36,356.89L807.44,357.10L807.34,357.22L807.16,357.29L806.62,357.31L806.74,357.72L807.05,357.91L807.37,358.47L807.44,358.86L807.31,359.27L807.39,359.61L807.56,359.79L807.82,359.93L807.67,359.64L807.48,359.54L807.48,359.42L807.71,359.32L807.87,359.33L807.95,359.44L808.14,360.06L808.13,360.19L808.36,360.71L808.27,361.00L808.42,361.37L808.21,361.34L808.08,361.26L808.08,360.99L807.85,360.88L807.84,360.63L807.75,360.47L807.79,360.27L807.66,359.87L807.45,359.73L806.95,359.67L806.88,359.49L806.75,359.38L806.69,359.05L806.55,358.83L806.05,358.56L805.87,358.00L805.69,357.89L805.47,357.36L805.17,357.19L804.90,356.91L804.65,356.78L804.54,356.25L804.30,355.97L804.43,355.78L804.69,355.70L804.56,355.53L804.33,355.43L803.78,355.34L803.43,355.11L803.11,355.00L802.90,354.58L802.65,354.31L802.45,353.49L801.51,352.96L801.19,352.36L801.03,352.22L800.93,351.97L800.95,351.84L800.83,351.71L800.71,351.42L800.81,351.23L800.89,351.23L800.93,351.16L800.96,351.19L800.96,351.08L801.06,350.88L800.95,350.70L800.89,350.41L800.69,350.24L800.66,349.57L800.36,349.32L800.34,349.43L800.22,349.52L800.17,349.76L800.22,350.01L800.14,350.04L800.15,350.42L799.85,350.48L799.70,350.16L799.87,350.06L799.78,349.86L799.58,349.92L799.38,349.69L798.88,349.36L798.71,349.10L798.57,349.03L798.52,348.79L798.44,348.75L798.29,348.32L798.16,348.24L798.05,347.95L798.10,347.69L797.72,347.39L797.77,347.00L797.83,346.96L798.14,347.43L798.15,347.58L798.08,347.60L798.14,347.63L798.12,347.69L798.76,347.41L798.91,347.46L798.97,347.66L799.10,347.74L799.22,347.94L799.37,348.43L799.67,348.46L800.06,348.70L800.15,348.86L800.08,348.99L800.11,349.13L800.59,349.13L800.73,349.24L800.84,349.45L800.85,349.62L801.03,349.80L801.02,349.90L801.10,349.88L801.08,349.69L801.16,349.57L801.48,349.45L801.67,349.50L801.94,349.96L802.02,350.38L802.15,350.54L802.12,351.12L802.22,350.85L802.43,350.82L802.58,350.92L802.50,351.50L802.60,351.50L802.59,351.66L802.67,351.69L802.63,351.96L802.70,351.97L802.80,352.12L802.83,352.56L802.99,351.92L802.73,351.60L802.71,351.36L802.86,350.63L802.74,349.95L802.17,349.65L801.78,349.26L801.45,348.47L801.43,347.95L801.33,347.59L800.89,346.82L800.00,346.50L799.87,346.29L799.60,346.08L799.20,345.29L798.87,344.82L798.77,344.53L798.59,344.33L798.20,343.63L797.97,343.48L797.77,343.08L797.71,342.87L797.75,342.76L797.51,342.36L797.48,342.11L797.20,341.78L797.17,341.72L797.25,341.74L797.16,341.68L797.08,341.50L797.06,341.40L797.17,341.39L797.06,341.38L797.02,341.18L796.74,340.84L796.71,340.57L796.62,340.45L796.75,340.39L796.62,340.43L796.59,340.38L796.76,340.36L796.59,340.36L796.49,340.03L796.62,340.02L796.50,340.01L796.49,339.93L796.59,339.34L796.97,338.84L797.35,338.72L797.47,338.73L797.52,338.84L797.67,338.72L797.32,338.51L797.06,338.23L797.24,338.07L797.04,338.19L796.98,338.12L796.95,337.81L797.13,337.18L796.61,336.42L796.50,336.29L795.52,336.36L794.97,336.18L794.87,335.99L794.89,335.71L794.97,335.60L794.75,335.61L794.25,335.50L793.99,335.31L793.75,334.75L793.96,335.28L793.96,335.41L793.41,335.92L793.18,336.34L793.03,336.43L792.75,336.45L792.47,336.18L792.38,335.81L792.44,335.59L792.33,335.41L792.34,335.23L792.53,334.96L792.34,334.97L792.26,334.85L792.42,334.93L792.52,334.89L792.27,334.74L792.24,334.58L792.36,334.75L792.50,334.82L792.51,334.75L792.36,334.60L792.41,334.57L792.43,334.67L792.51,334.67L792.48,334.54L792.55,334.37L792.48,334.21L792.44,333.79L792.50,332.40L792.60,332.21L792.75,332.17L792.89,332.30L792.80,332.42L792.96,332.48L792.99,332.23L792.96,332.06L792.84,332.04L792.77,331.99L792.59,332.02L792.46,332.12L792.48,332.01L792.35,332.01L792.43,331.93L792.29,331.90L792.39,331.84L792.27,331.80L792.28,331.59L792.40,331.70L792.49,331.69L792.52,331.81L792.60,331.72L792.63,331.86L792.71,331.79L792.70,331.93L792.79,331.89L792.78,331.97L792.84,332.03L792.96,332.01L792.92,331.85L792.78,331.66L792.37,331.52L792.33,331.44L792.26,331.47L792.27,331.40L792.20,331.42L792.25,331.36L792.19,331.40L792.21,331.33L792.13,331.28L791.99,330.78L792.12,330.30L792.71,329.34L792.82,328.52L793.26,327.90L793.34,327.62L793.27,327.43L793.74,326.92L793.93,326.95L794.02,327.03L794.00,327.14L794.29,327.27L794.61,327.58L794.77,327.64L794.94,328.02L794.89,328.13L795.01,328.14L794.94,328.18L795.01,328.24L795.00,328.42L795.08,328.46L795.47,329.65L795.51,330.17L795.59,330.28L796.11,330.36L796.21,330.57L796.57,331.01L796.68,331.27L796.67,331.41L796.81,331.50L796.79,331.66L796.83,331.61L796.82,331.69L796.87,331.65L796.91,331.70L796.85,331.73L796.93,331.76L796.91,331.86L797.01,331.84L797.03,331.69L797.08,331.77L797.03,331.83L797.15,331.84L797.15,331.98L797.09,332.05L797.09,331.90L796.96,331.88L796.99,332.16L797.12,332.19L797.21,332.31L797.14,332.44L797.20,332.43L797.18,332.51L797.05,332.67L797.14,332.85L797.07,333.18L797.10,333.35L797.15,333.35L797.13,333.25L797.20,333.21L797.25,333.27L797.22,333.40L797.33,333.29L797.40,333.39L797.46,333.26L797.51,333.43L797.58,333.26L797.59,333.43L797.70,333.28L797.73,333.45L797.89,333.30L797.86,333.21L798.00,333.24L798.00,333.32L797.90,333.30L797.82,333.43L797.94,333.44L798.01,333.53L797.98,333.68L798.05,333.87L797.89,333.93L798.02,333.94L798.10,334.02L798.19,333.89L798.22,333.99L798.11,334.05L797.81,333.91L797.80,333.96L797.89,334.04L797.86,334.08L797.71,333.91L797.62,333.90L797.68,334.21L797.61,334.41L797.87,334.43L798.56,334.00L798.55,333.82L798.60,333.79L798.50,333.62L798.60,333.57L798.27,333.64L798.19,333.58L798.36,333.56L798.34,333.49L798.18,333.54L798.32,333.44L798.47,333.51L798.43,333.39L798.55,333.51L798.58,333.44L798.50,333.36L798.58,333.31L798.53,333.33L798.49,333.26L798.56,333.20L798.44,333.15L798.35,332.78L798.37,332.71L798.44,332.72L798.39,332.64L798.50,332.68L798.37,332.62L798.45,332.57L798.38,332.53L798.44,332.52L798.37,332.42L798.44,332.40L798.39,332.35L798.49,332.35L798.41,332.31L798.46,332.29L798.42,332.24L798.49,332.25L798.50,332.09L798.62,332.01L798.62,331.86L798.76,331.73L798.71,331.68L798.77,331.59L798.79,331.68L798.81,331.59L798.72,331.56L798.80,331.53L798.72,331.42L798.72,330.90L798.75,330.82L798.79,330.83L798.74,330.82L798.80,330.73L798.77,330.66L798.84,330.66L798.79,330.58L798.85,330.53L798.82,330.48L798.87,330.51L798.83,330.45L798.89,330.42L798.86,330.31L798.95,330.16L798.71,330.06L798.78,329.77L799.13,329.81L799.06,329.75L799.11,329.65L799.15,329.69L799.19,329.51L799.32,329.66L799.24,329.43L798.94,329.39L798.83,329.48L798.92,329.35L798.94,328.39L799.00,328.14L799.10,328.09L799.09,327.89L799.11,328.08L799.21,328.10L799.35,328.03L799.65,328.12L799.91,328.43L799.97,328.62L800.02,328.60L799.97,328.64L800.02,328.71L799.98,328.79L799.63,329.05L799.66,329.16L799.97,329.02L800.08,328.82L800.13,328.85L800.07,328.94L800.13,328.86L800.10,328.81L800.17,328.67L800.26,328.67L800.22,328.62L800.25,328.58L800.28,328.63L800.27,328.56L800.30,328.60L800.27,328.55L800.32,328.57L800.30,328.51L800.49,328.21L800.54,328.23L800.51,328.33L800.55,328.23L800.54,328.19L800.60,328.20L800.57,328.16L800.86,327.83L800.93,327.78L801.05,327.83L801.07,327.77L800.83,327.67L801.09,327.68L800.86,327.60L800.88,327.56L801.07,327.58L801.08,327.51L800.88,327.55L800.83,327.49L801.10,327.41L801.10,327.36L800.96,327.33L801.10,327.28L801.08,326.20L801.13,326.13L801.08,326.07L800.90,326.08L800.97,326.02L801.08,326.03L801.08,325.97L800.84,325.99L800.86,325.93L801.09,325.90L800.86,325.91L800.87,325.86L801.09,325.81L800.84,325.79L801.08,325.71L800.88,325.70L800.87,325.64L801.08,325.62L800.84,325.58L801.07,325.53L800.92,325.49L801.07,325.45L801.06,325.41L800.95,325.42L800.96,325.35L801.19,325.34L801.07,325.30L801.09,325.23L801.18,325.20L801.08,325.20L801.08,324.58L801.12,324.58L801.07,324.57L801.03,324.09L801.07,323.86L801.12,323.84L801.08,323.81L801.09,323.72L801.19,323.54L801.18,323.33L801.24,323.32L801.17,323.28L801.17,323.12L801.32,322.93L801.26,322.70L801.42,322.59L801.36,322.53L801.38,322.32L801.45,322.27L801.50,322.36L801.61,322.33L801.60,322.27L801.44,322.22L801.49,321.88L801.38,321.46L801.41,321.38L801.22,320.82L801.34,320.80L801.36,320.88L801.40,320.78L801.59,320.73L801.61,320.82L801.63,320.78L801.61,320.71L801.31,320.46L801.22,320.34L801.14,319.72L801.78,319.55L801.60,319.23L801.01,319.43L800.84,319.16L801.08,319.07L800.95,318.83L800.61,318.97L800.35,318.47L800.70,318.27L800.55,318.00L801.09,317.70L800.89,317.38L799.90,317.95L800.06,317.81L799.52,317.44L799.55,317.36L799.49,317.43L799.39,317.41L799.44,317.30L799.22,317.21L799.25,317.13L799.15,317.09L799.21,317.02L799.29,317.05L799.34,316.97L799.62,317.07L799.60,316.96L799.78,316.98L799.73,316.89L799.80,316.82L799.88,316.54L799.96,316.55L800.04,316.37L800.11,316.40L800.15,316.33L800.03,316.21L800.04,316.02L799.72,315.80L799.84,315.60L799.97,315.69L800.18,315.27L800.41,315.46L800.42,315.29L800.49,315.26L800.56,315.02L800.82,315.09L801.08,313.85L801.03,313.99L800.88,313.92L800.85,314.32L800.94,314.35L800.79,315.02L800.66,315.02L800.63,314.87L800.70,314.81L800.60,314.83L800.56,314.76L800.44,314.73L800.39,314.86L800.32,314.84L800.35,314.73L800.17,314.77L800.21,314.66L800.32,314.66L800.22,314.63L800.24,314.55L800.32,314.55L800.37,314.64L800.59,314.63L800.77,313.79L800.59,313.72L800.47,314.45L800.29,314.44L800.50,313.44L800.55,313.43L800.25,313.24L800.02,314.39L799.99,314.43L799.98,314.37L799.83,314.35L799.88,313.88L799.96,313.70L799.88,313.73L799.78,313.66L799.94,312.82L800.19,312.83L799.94,312.81L799.52,314.87L800.01,315.05L799.48,315.91L798.95,317.14L798.83,317.29L798.45,317.17L797.60,316.66L797.75,314.53L797.73,314.02L797.24,313.91L797.16,314.02L797.20,314.05L797.16,314.22L796.97,314.36L797.09,314.36L797.02,314.43L797.06,314.51L796.94,314.48L797.06,314.67L796.96,314.98L796.97,315.16L796.74,315.19L796.87,315.38L796.86,315.57L796.74,315.72L796.56,315.77L796.38,315.70L796.28,315.54L796.27,315.46L796.43,315.22L796.31,315.12L796.39,314.91L796.29,314.84L796.29,314.68L796.26,314.68L796.32,314.53L796.44,314.47L796.57,314.51L796.54,314.63L796.62,314.66L796.57,314.49L796.44,314.46L796.32,314.51L796.25,314.60L796.21,314.81L795.95,314.81L795.90,314.66L795.78,314.87L795.82,315.22L795.69,315.30L795.45,315.30L795.36,315.42L795.32,315.38L795.41,315.30L795.30,315.19L795.17,315.21L795.28,315.15L795.26,315.07L795.48,315.04L795.43,315.01L795.46,314.95L795.25,315.06L795.22,314.99L795.26,315.02L795.31,314.95L795.22,314.88L795.17,314.93L795.20,314.98L795.08,314.99L794.88,314.45L794.58,314.26L794.45,313.93L794.10,314.00L794.01,313.78L794.18,313.75L794.00,313.51L794.23,313.46L794.01,312.40L793.77,312.42L793.37,311.61L793.42,311.58L793.53,311.63L793.64,311.82L793.59,311.87L793.64,311.85L793.68,311.93L793.62,311.76L793.67,311.59L793.38,311.56L793.31,311.50L793.36,311.15L793.31,310.84L793.35,310.83L793.41,311.17L793.55,311.23L793.56,311.35L793.70,311.43L793.91,311.43L793.93,311.84L794.36,311.77L794.41,311.85L794.35,311.91L794.66,311.97L795.01,311.93L795.36,312.09L795.61,312.04L795.82,312.11L795.82,312.18L795.82,312.10L795.57,312.02L795.62,311.84L795.53,311.70L795.86,311.24L795.94,310.82L796.01,310.81L795.95,310.49L795.90,310.32L795.83,310.22L795.93,310.42L795.97,310.61L795.93,310.61L795.80,310.24L795.52,309.89L795.56,309.50L795.54,309.67L795.39,309.71L794.97,309.65L794.94,309.59L794.66,309.69L794.45,309.90L793.60,310.02L793.55,309.98L793.55,310.21L793.29,310.33L793.34,310.80L793.30,310.81L793.26,310.57L793.10,310.27L793.29,310.17L793.31,309.96L793.01,308.54L792.73,308.31L792.83,307.96L792.93,307.86L792.89,307.70L792.92,307.65L792.98,307.69L792.97,307.55L793.09,307.40L793.19,307.37L793.15,307.32L793.36,307.11L792.93,307.55L792.88,307.38L793.06,307.33L793.03,307.16L792.80,307.06L792.30,307.11L792.18,307.24L792.20,306.94L792.43,306.82L792.08,306.99L791.98,306.86L791.92,306.88L792.12,306.65L792.07,306.48L791.96,304.35L792.19,303.67L792.24,303.22L792.62,303.15L792.75,303.23L793.18,304.05L793.34,304.40L793.24,304.49L793.32,304.61L793.29,305.41L793.35,305.52L793.30,305.04L793.43,305.38L793.53,305.43L793.70,305.73L793.80,306.11L793.99,306.39L793.94,306.46L794.12,306.35L794.21,306.37L794.23,306.49L794.30,306.39L794.15,306.34L794.19,306.27L794.83,305.72L794.92,305.71L794.95,305.77L794.98,305.68L795.08,305.66L795.46,305.66L795.46,305.73L795.52,305.66L795.59,305.72L795.61,305.68L796.39,305.81L796.48,305.86L796.60,306.13L796.94,306.04L797.03,305.87L797.24,305.88L797.37,305.99L797.62,306.40L797.52,306.58L797.59,306.65L797.57,306.61L797.65,306.59L797.79,306.65L798.05,306.99L798.45,307.86L798.54,307.95L798.48,308.00L798.51,308.08L798.54,308.01L798.80,308.36L798.89,308.27L798.87,307.80L798.69,307.47L798.22,307.11L798.00,306.86L797.31,305.87L797.25,305.70L796.61,305.30L796.43,305.03L796.56,304.84L796.07,304.65L795.34,304.07L795.22,303.69L794.98,303.33L794.72,303.03L794.24,302.70L793.78,301.87L793.56,301.70L792.80,300.70L792.34,299.84L790.96,298.99L790.79,298.80L790.68,298.46L790.05,298.38L788.95,298.36L788.80,298.30L788.83,298.03L788.79,298.22L788.69,298.22L788.67,298.30L788.04,298.30L787.53,298.21L787.35,298.12L787.33,298.05L787.17,298.00L787.15,297.87L786.69,297.69L786.44,297.44L786.89,296.61L786.96,296.65L786.91,296.56L786.99,296.58L786.95,296.49L787.03,296.50L786.96,296.47L787.06,296.45L786.99,296.40L787.02,296.35L787.11,296.40L787.10,296.47L787.12,296.40L787.02,296.34L786.45,297.40L785.98,297.34L785.70,297.16L785.08,297.00L784.81,296.71L784.56,296.66L784.25,296.50L784.44,296.39L784.67,296.45L784.65,296.40L784.45,296.37L784.17,296.50L784.17,296.48L783.75,296.16L783.73,296.11L783.73,295.97L783.96,295.64L783.93,295.50L783.94,295.63L783.70,296.02L783.74,296.17L784.16,296.48L783.94,296.45L783.40,296.07L783.54,295.71L783.79,295.53L783.55,295.67L783.43,295.95L783.33,295.91L782.58,295.28L782.07,294.94L781.96,294.69L782.03,294.69L782.14,294.91L782.19,294.86L781.97,294.57L781.71,294.46L781.37,294.43L781.01,294.24L780.95,294.12L780.71,293.98L780.70,293.85L780.65,293.83L780.96,293.57L780.94,293.42L780.93,293.57L780.59,293.76L780.40,293.60L780.19,293.19L780.26,293.09L780.37,293.16L780.31,293.11L780.38,293.11L780.36,293.03L780.42,293.06L780.37,293.02L780.43,293.04L780.38,293.01L780.44,293.01L780.43,292.94L780.48,292.96L780.41,292.91L780.38,292.96L780.28,292.88L780.34,292.79L780.60,292.76L780.60,292.83L780.66,292.94L780.62,292.79L780.75,292.80L780.63,292.76L780.37,292.76L780.03,292.96L779.82,293.02L779.64,292.86L779.46,292.65L779.39,292.34L779.49,292.15L779.69,292.13L779.68,292.24L779.71,292.22L779.69,292.08L779.48,292.08L779.48,292.02L779.65,291.98L779.61,291.86L779.71,291.81L779.54,291.78L779.44,291.54L779.48,291.40L779.57,291.36L779.61,291.46L779.95,291.50L779.82,291.48L779.81,291.39L779.86,291.33L779.93,291.36L779.97,291.29L779.98,291.36L780.05,291.43L780.00,291.53L780.65,290.91L780.73,290.64L780.64,290.90L780.09,291.40L779.98,291.30L779.97,291.19L780.45,290.71L780.28,290.54L779.85,290.98L779.71,290.88L779.62,290.97L779.33,290.96L779.06,290.68L779.28,290.50L779.27,290.43L778.77,290.96L778.62,290.96L778.28,290.49L778.25,290.14L778.12,289.87L777.84,289.65L777.81,289.54L777.98,289.48L777.49,289.62L777.43,289.57L777.61,289.51L777.89,289.25L779.20,288.64L779.33,288.74L779.43,288.74L779.44,288.54L779.62,288.45L779.62,288.74L779.86,288.74L779.87,288.32L780.03,288.26L780.02,288.72L780.19,288.72L780.19,288.46L780.38,288.47L780.37,288.73L780.54,288.74L780.55,288.20L780.79,288.21L780.85,288.89L780.83,288.15L781.39,288.16L781.72,288.83L781.40,288.14L780.83,288.14L780.81,287.66L780.19,287.66L780.12,288.07L777.48,289.30L777.40,289.25L777.47,288.99L778.16,288.86L777.36,289.00L777.29,288.90L777.37,288.83L777.51,288.29L777.36,288.82L777.29,288.89L777.19,288.79L777.28,288.71L777.40,288.28L777.28,288.71L777.18,288.79L777.10,288.70L777.21,288.05L777.17,288.26L777.03,288.40L776.99,288.60L776.88,288.73L776.52,288.44L776.17,287.96L776.08,287.71L776.16,287.58L776.11,287.61L776.10,287.47L776.02,287.52L776.03,287.47L775.94,287.42L775.90,287.25L775.96,287.23L775.86,287.15L775.62,286.38L775.69,286.09L775.61,286.04L775.70,286.06L775.54,285.76L775.55,285.58L775.73,285.07L775.92,285.00L776.00,284.89L775.95,284.75L776.08,284.78L776.23,284.66L776.20,284.40L776.26,284.13L776.13,284.03L776.39,282.88L777.30,283.58L777.25,283.85L777.30,283.85L777.30,284.03L777.14,284.04L777.14,284.13L777.30,284.12L777.29,285.19L777.37,285.19L777.37,285.03L777.31,284.97L777.31,283.56L776.39,282.83L776.82,280.98L777.14,280.71L777.22,280.71L777.19,280.66L777.87,279.53L777.82,279.50L777.78,279.57L777.18,280.65L776.80,280.97L776.12,284.03L775.91,284.02L775.57,284.39L775.51,284.55L775.03,284.83L774.81,284.50L774.75,284.29L774.30,283.95L774.11,283.95L773.97,283.73L774.22,283.41L774.40,283.57L774.82,283.38L774.82,282.89L774.62,282.27L774.58,281.75L774.46,281.60L774.17,281.68L774.00,281.61L773.85,281.31L773.46,281.34L773.47,281.73L773.33,281.97L773.22,281.32L773.27,281.33L773.25,281.15L773.10,280.42L773.02,280.16L772.84,280.07L772.73,279.88L772.73,279.28L772.85,279.09L773.04,278.98L772.45,279.25L772.34,278.87L772.42,278.63L772.78,278.30L772.89,277.83L773.01,277.69L773.28,277.58L773.48,277.38L773.66,277.38L773.86,277.25L773.99,277.24L774.19,277.25L774.52,277.43L774.71,277.67L774.82,277.55L775.11,277.66L775.16,277.77L775.32,277.85L775.53,278.39L775.03,278.61L775.03,278.56L774.75,278.62L774.74,278.68L774.85,278.77L774.95,278.77L775.01,278.89L775.11,279.16L775.08,279.37L775.10,279.29L775.18,279.27L775.15,279.12L775.25,279.25L775.32,279.25L775.33,279.14L775.22,279.06L775.25,278.92L775.33,278.86L775.43,278.98L775.44,278.87L775.54,278.85L775.52,278.70L775.69,278.87L775.60,279.05L775.85,279.02L775.84,279.11L775.46,279.63L775.83,279.18L776.44,278.70L777.04,278.53L777.56,278.28L778.06,278.15L778.15,278.16L778.18,278.24L778.27,278.15L778.47,278.13L778.46,278.23L778.68,278.45L778.64,278.54L778.50,278.52L778.45,278.57L778.55,278.64L778.58,278.79L778.81,278.78L778.80,278.85L779.40,278.81L779.73,278.86L779.74,278.94L780.53,278.82L780.67,279.08L780.57,278.86L780.81,278.77L780.50,278.71L780.04,278.44L779.41,278.20L779.01,278.16L778.23,277.94L776.71,277.73L776.06,277.53L775.73,277.22L775.62,276.89L775.24,276.70L774.97,276.64L774.79,276.41L774.33,276.50L774.21,276.37L774.01,276.35L772.95,276.69L772.65,276.69L772.37,276.84L772.08,277.20L771.96,277.50L771.80,277.62L771.68,277.82L771.63,278.06L771.34,278.64L771.11,279.59L771.23,279.81L771.22,280.15L771.35,280.41L771.38,280.33L771.50,280.48L771.36,280.28L771.31,280.28L771.29,280.17L771.40,280.12L771.60,280.43L771.78,280.49L771.81,280.59L771.79,280.47L771.50,280.28L771.41,279.99L771.45,279.86L771.58,279.78L771.63,279.61L772.19,279.20L772.33,278.89L772.45,279.26L772.02,279.54L771.98,279.62L772.06,279.79L772.16,279.87L772.27,280.33L772.20,280.56L771.95,280.77L771.49,281.68L771.81,281.54L772.18,281.57L772.28,281.81L772.52,281.97L772.58,282.16L772.67,282.21L772.61,282.12L772.67,282.14L772.80,282.30L772.71,282.07L772.81,281.93L772.90,281.97L773.13,282.32L773.30,282.40L773.28,282.57L773.48,282.92L773.60,283.02L773.91,283.10L774.21,283.40L773.96,283.70L773.84,283.68L773.80,283.78L773.56,283.95L773.16,284.48L772.95,284.61L772.72,284.86L772.76,285.03L772.70,285.26L772.83,285.52L772.92,285.56L772.86,285.63L772.99,286.00L773.16,286.18L773.26,286.19L773.46,286.43L773.43,286.39L773.50,286.35L773.50,286.47L773.66,286.57L773.77,286.76L773.43,287.03L773.16,286.72L773.00,286.69L772.80,286.56L772.78,286.43L772.10,286.22L771.83,286.28L771.69,286.41L771.51,286.46L771.52,286.53L771.10,286.76L770.98,287.03L770.78,287.02L770.66,286.94L770.65,286.81L770.71,286.72L770.81,285.81L770.67,285.62L770.26,285.66L770.09,285.83L769.99,286.35L769.85,286.32L769.75,286.20L769.57,286.55L769.55,286.71L769.40,286.53L769.45,286.36L769.37,286.44L769.54,286.96L769.72,287.21L769.77,287.67L769.88,287.76L770.13,287.80L770.17,287.62L770.21,287.85L770.49,288.09L770.49,287.84L770.55,287.75L770.60,288.05L770.68,288.16L770.90,288.26L770.95,288.17L770.92,288.30L771.04,288.27L771.02,288.21L771.08,288.22L771.11,288.11L771.25,288.15L771.26,288.06L771.14,288.02L771.19,287.88L771.24,287.90L771.25,287.85L771.13,287.76L771.13,287.65L771.32,287.63L771.61,288.11L771.70,288.39L771.60,288.45L771.49,288.69L771.55,288.53L771.51,288.22L771.44,288.26L771.40,288.43L771.06,288.43L770.97,288.55L770.73,288.39L770.62,288.41L770.10,287.91L769.76,287.85L769.51,287.69L769.42,287.48L768.98,287.09L768.70,287.01L768.25,287.02L768.27,287.05L768.00,286.90L767.89,286.93L767.89,287.01L767.69,286.79L767.74,286.70L767.63,286.21L767.68,285.83L767.85,285.37L768.03,285.32L768.23,285.43L768.59,285.23L768.76,285.61L768.85,285.68L768.99,285.67L769.03,285.14L769.14,285.14L769.04,285.13L769.09,284.89L769.03,285.07L768.94,284.86L769.02,284.89L769.10,284.58L769.19,284.61L769.26,284.32L769.29,284.52L769.29,284.21L769.42,284.14L769.56,283.72L769.47,283.45L769.23,283.41L769.16,283.81L769.09,283.81L768.99,283.93L769.00,283.80L768.86,283.22L769.06,283.16L769.06,283.12L768.86,283.19L768.81,283.27L768.27,283.33L768.31,283.45L768.20,283.51L768.05,283.53L767.84,283.43L767.58,283.59L767.50,283.70L767.48,283.93L767.52,283.96L767.52,283.88L767.56,283.94L767.66,284.22L767.61,284.64L767.57,284.73L767.35,284.86L767.19,285.33L767.18,285.66L767.13,285.77L767.07,285.78L766.93,285.40L766.96,285.28L766.90,285.16L766.95,285.11L766.69,284.87L766.48,284.56L766.27,284.43L766.18,284.25L765.89,284.03L765.91,283.57L766.22,283.30L766.08,282.89L766.03,282.86L766.04,282.92L765.90,282.79L765.82,282.79L765.67,282.54L765.77,282.36L765.76,282.21L765.86,282.07L765.80,281.77L765.84,281.33L765.98,281.13L766.15,281.06L766.10,280.54L765.65,279.75L765.71,278.68L765.76,278.67L765.60,278.52L765.61,278.32L765.67,278.30L765.65,278.22L765.55,278.18L765.60,277.98L765.75,278.02L765.72,277.93L765.87,277.94L765.81,278.03L765.96,278.04L765.93,277.83L765.81,277.66L765.56,277.58L765.43,277.27L765.21,277.16L764.75,277.16L764.69,277.02L764.46,276.99L764.44,276.91L764.40,277.02L764.11,277.11L763.99,277.06L763.96,277.17L763.84,277.20L763.82,277.30L763.28,277.53L763.21,277.59L763.18,277.79L762.64,278.26L762.18,278.45L761.96,278.45L761.85,278.34L761.84,278.16L761.96,277.80L762.12,277.65L762.27,277.60L762.34,277.66L762.33,277.76L762.48,277.75L762.41,277.82L762.57,277.75L762.59,277.23L762.76,276.95L762.70,276.40L762.50,276.40L762.37,276.10L762.46,275.81L762.67,275.59L762.88,275.52L762.99,275.65L762.83,275.82L763.02,275.78L763.16,275.86L763.13,276.00L763.23,275.94L763.46,275.50L763.83,275.37L763.99,275.00L763.97,274.92L763.41,274.66L763.28,274.48L763.12,274.41L763.07,274.20L763.22,274.01L763.11,274.06L763.05,273.95L762.87,273.93L762.68,274.06L762.45,274.02L762.57,274.68L762.52,275.08L762.39,275.20L761.89,275.39L761.28,275.50L761.23,275.48L761.23,275.41L761.35,275.27L761.29,275.03L761.35,274.61L761.19,274.61L761.09,274.86L760.83,274.90L760.56,274.79L760.76,274.70L760.27,274.56L760.47,274.53L760.45,274.48L760.23,274.48L760.16,274.38L760.21,274.36L760.07,274.38L759.90,274.29L759.99,274.30L759.83,274.02L759.33,273.95L759.17,273.86L758.70,273.97L758.36,274.31L758.19,274.23L758.15,274.13L758.16,273.86L758.31,273.54L758.47,273.51L758.53,273.57L758.50,273.64L758.63,273.57L758.68,273.23L758.88,272.90L758.82,272.95L758.75,272.92L758.82,272.63L758.77,272.61L758.95,272.63L759.06,272.73L759.02,272.88L758.92,272.94L759.44,272.98L759.64,273.17L759.71,273.41L759.78,273.41L759.91,273.12L759.93,272.93L760.11,272.61L760.47,272.38L760.73,272.40L760.85,272.64L760.96,273.18L760.50,273.62L760.52,273.67L760.90,273.50L761.09,273.34L761.31,273.09L761.57,272.51L761.78,272.35L762.05,272.37L762.40,272.73L762.54,272.77L762.59,273.13L762.53,273.14L762.59,273.17L762.98,273.01L763.07,272.74L763.25,272.61L763.39,272.31L763.54,272.65L763.50,272.74L763.42,272.66L763.28,272.66L763.47,272.81L763.63,272.79L763.86,272.95L764.04,272.97L764.45,272.66L764.68,271.80L764.87,271.73L765.10,271.79L765.11,271.93L765.06,271.94L765.11,271.95L765.06,271.95L765.12,271.98L765.22,272.33L765.21,272.56L765.02,272.84L764.94,273.22L765.15,273.02L765.30,273.05L765.53,272.97L765.49,273.32L765.56,273.30L765.54,272.54L765.68,272.39L765.38,272.02L765.33,271.86L765.30,271.52L765.36,271.31L765.19,271.04L764.54,270.88L764.23,270.69L763.80,269.90L763.69,269.42L763.29,268.85L761.94,268.50L761.76,268.36L761.74,268.20L761.57,268.09L761.34,268.04L760.42,268.06L758.99,268.32L758.12,268.30L757.11,268.11L756.88,267.94L756.85,267.82L756.38,267.53L756.36,267.42L756.02,267.25L755.12,267.01L755.74,267.18L755.37,267.40L754.92,267.32L754.81,267.23L754.90,267.12L755.07,267.11L755.03,267.04L755.10,267.04L754.63,266.93L754.42,266.81L753.72,265.90L753.62,265.64L753.72,265.53L753.53,265.66L753.17,265.22L753.16,265.11L753.23,265.09L753.13,265.01L752.88,264.93L752.83,264.83L752.13,264.60L751.93,264.58L751.73,264.79L751.30,264.99L750.88,265.32L750.33,265.46L749.97,265.93L749.79,265.98L749.87,266.04L749.57,266.57L749.80,266.50L749.80,266.45L749.85,266.58L749.92,266.56L749.80,266.42L749.79,266.20L750.05,265.85L750.20,265.78L750.41,265.91L750.49,266.05L750.33,266.30L750.44,266.62L750.28,266.82L749.59,267.12L749.27,267.09L749.22,266.97L749.22,266.91L749.40,266.84L749.32,266.59L749.35,266.26L749.30,266.05L749.62,264.95L749.85,264.61L749.96,264.59L750.04,264.48L749.42,264.25L749.29,264.27L749.26,264.61L749.13,264.72L749.07,265.02L748.90,265.14L748.78,265.56L748.53,265.68L748.51,265.73L748.13,265.76L747.98,265.54L747.41,265.44L747.40,265.27L747.19,265.16L747.10,264.82L747.21,264.59L747.45,264.36L747.68,264.46L747.95,264.42L748.00,264.54L747.93,264.76L748.01,264.99L748.25,264.90L748.31,264.93L748.40,265.25L748.45,264.70L748.39,264.58L748.27,264.43L748.11,264.42L748.05,264.19L747.91,264.11L747.93,264.18L747.88,264.17L747.77,264.07L747.66,264.11L747.64,264.05L747.49,263.99L747.46,263.73L747.62,263.40L748.13,263.25L748.37,263.35L748.42,263.71L748.34,263.92L748.43,263.87L748.39,263.94L748.56,263.80L748.69,263.52L748.82,263.49L748.65,263.22L748.56,262.66L748.68,262.35L748.50,262.69L748.17,262.84L747.80,262.90L747.69,263.08L747.87,262.97L747.99,263.03L747.59,263.10L747.34,262.99L747.33,262.79L747.43,262.82L747.53,262.71L747.58,262.74L747.57,262.88L747.63,262.72L747.64,262.92L747.63,262.66L747.56,262.39L747.57,262.15L747.73,261.90L747.72,261.62L747.60,261.33L747.37,261.20L747.24,260.85L747.18,260.96L747.30,261.18L747.20,261.47L746.83,261.77L746.76,262.01L746.52,261.96L746.50,261.73L746.40,261.69L746.35,261.52L746.18,261.38L746.14,261.15L746.33,260.55L746.53,260.34L746.48,260.14L746.53,259.71L746.02,259.23L746.15,259.28L746.20,259.38L746.14,259.24L745.85,259.19L745.75,259.34L745.74,259.40L745.87,259.22L746.00,259.27L745.71,259.57L745.68,259.73L745.74,259.91L745.80,259.98L745.73,259.84L745.86,259.85L745.91,260.01L745.67,260.25L745.70,260.50L745.62,260.68L745.40,260.77L745.30,260.94L745.10,260.95L744.88,260.77L744.90,260.57L745.01,260.33L745.18,260.22L745.33,259.86L745.40,259.43L745.47,259.31L745.50,258.56L745.71,258.40L745.88,257.99L745.83,257.87L745.81,257.30L745.84,256.72L745.75,256.54L745.86,256.05L745.96,255.89L746.57,255.41L746.69,255.45L746.64,255.53L746.47,255.54L746.49,255.62L746.63,255.69L746.62,255.91L746.71,256.11L746.55,256.47L746.70,257.10L746.67,257.23L746.35,257.58L746.00,257.73L745.97,257.80L746.04,257.81L746.11,257.70L746.35,257.59L746.07,257.99L746.24,258.24L746.20,258.56L746.02,258.61L745.84,258.55L745.58,258.64L745.75,258.91L746.31,258.95L746.74,258.47L746.96,258.38L747.00,258.28L746.90,258.19L746.76,258.24L746.68,258.20L746.68,257.88L746.90,257.49L746.94,257.09L746.83,257.00L746.81,256.81L747.04,256.27L747.15,256.15L747.43,256.02L747.57,256.02L747.95,256.69L747.96,256.83L748.18,256.96L748.25,257.08L748.25,257.18L748.06,257.48L748.10,257.63L748.02,257.77L747.47,258.11L747.41,258.21L747.15,258.32L747.32,258.51L747.64,258.45L747.42,258.41L747.45,258.29L747.65,258.09L747.77,258.07L747.89,258.23L747.88,258.62L747.75,258.75L747.76,258.87L747.62,259.13L747.63,259.22L747.78,259.32L747.80,259.44L747.62,259.65L747.15,259.82L747.25,259.83L747.10,259.89L747.07,260.05L747.16,260.17L747.24,260.17L747.85,259.73L748.50,259.14L748.57,258.93L748.86,258.60L749.06,258.22L749.02,258.12L749.21,257.70L749.21,257.56L749.20,257.70L749.01,258.12L748.95,258.13L748.88,258.11L748.87,257.99L748.62,257.99L748.39,257.61L748.52,256.88L748.80,256.24L748.87,256.18L749.15,256.29L748.90,256.13L748.97,255.88L749.05,255.91L748.76,255.75L748.72,255.56L748.78,255.22L748.38,255.09L747.79,254.47L747.33,254.21L747.22,254.05L746.80,253.96L746.50,253.78L746.22,252.94L746.29,252.48L746.45,252.26L745.74,252.30L745.31,252.17L745.18,252.04L745.20,251.96L744.96,251.62L744.59,251.34L744.43,250.88L744.10,250.60L744.08,250.38L743.61,250.16L743.47,249.98L743.27,249.28L743.45,248.80L743.72,248.60L744.00,248.52L744.34,248.52L744.57,248.60L744.65,248.51L744.68,248.14L744.77,247.97L744.85,247.91L745.05,247.94L745.10,248.05L745.07,248.23L744.94,248.45L744.72,248.51L744.72,248.64L744.48,248.70L744.28,248.95L743.73,249.07L743.72,249.12L743.47,249.13L743.70,249.16L744.41,248.92L745.22,248.31L745.26,247.92L745.14,247.90L745.13,247.83L744.52,247.83L743.53,247.55L743.37,247.56L741.61,247.27L741.35,247.16L741.09,246.91L741.07,246.71L740.94,246.60L740.88,246.60L740.94,246.81L740.87,246.83L740.78,246.77L740.85,246.68L740.42,246.59L740.25,246.50L740.17,246.30L740.23,245.99L740.16,245.80L739.93,245.52L739.65,245.37L739.17,244.97L738.56,244.09L738.31,243.35L738.31,243.16L738.67,243.05L738.71,243.13L738.66,243.16L739.06,243.17L738.93,243.15L738.77,242.94L738.22,243.09L738.11,242.95L738.18,242.87L738.08,242.88L738.02,242.27L738.11,241.67L738.09,240.69L738.26,239.86L738.52,239.36L738.87,239.13L738.96,239.39L738.97,239.23L739.02,239.30L738.99,239.50L738.91,239.57L739.02,239.52L739.08,239.25L738.83,238.98L738.75,238.72L738.29,238.49L738.14,238.03L738.03,237.98L737.80,237.57L737.25,237.00L737.12,236.68L736.47,236.30L736.42,236.19L735.54,235.58L735.40,235.43L735.10,234.98L734.88,234.28L734.57,233.76L734.36,233.20L734.31,232.66L734.40,232.11L734.96,230.66L735.16,230.36L735.48,230.06L735.51,229.94L735.45,229.72L735.29,229.48L734.94,229.38L734.93,229.47L734.58,229.59L734.50,229.74L734.44,229.68L734.48,229.76L734.31,230.13L734.22,230.52L734.11,230.69L733.90,230.82L733.77,230.88L733.89,230.78L733.69,230.79L733.71,230.73L733.40,230.80L733.09,230.96L733.14,231.05L733.06,231.02L733.07,231.14L732.99,231.07L732.89,231.23L732.89,231.11L732.80,231.18L732.76,231.01L732.60,230.99L732.56,230.89L732.65,230.57L732.88,230.31L733.11,230.22L733.24,230.26L733.23,230.44L732.94,230.59L732.83,230.73L733.04,230.62L733.18,230.78L733.32,230.54L734.04,230.59L734.13,230.49L734.13,230.09L734.25,229.95L734.25,229.73L734.54,229.51L734.77,229.44L734.74,229.34L734.79,229.30L734.94,229.30L734.81,229.21L734.63,229.27L734.27,229.07L733.93,228.69L733.79,228.41L733.78,227.95L733.93,227.33L733.66,226.77L733.64,226.57L733.69,226.26L733.84,225.92L733.27,225.81L732.48,225.44L731.97,224.97L731.64,224.46L731.06,224.15L702.65,224.60L702.83,223.76L700.67,221.39L700.09,219.88L698.80,218.76L698.40,217.28L698.33,215.54L698.57,214.14L698.37,212.35L698.06,210.67L697.19,209.53L696.34,207.90L695.19,206.45L694.56,204.87L694.34,204.53L693.32,203.36L668.94,200.18L666.59,199.96L658.00,198.89L657.45,198.89L652.64,200.63L637.97,199.20L637.04,201.36L636.18,202.30L635.17,202.99L633.99,203.39L632.92,203.43L631.93,203.79L630.97,204.53L630.44,205.31L629.99,205.82L629.41,206.11L628.41,206.30L627.70,206.56L627.32,206.95L627.05,208.37L626.85,208.69L626.35,208.98L624.93,208.83L624.43,209.04L624.31,210.17L624.12,210.73L623.36,212.20L622.90,212.71L622.36,212.99L621.16,213.26L619.22,213.27L618.37,213.19L616.85,212.92L613.53,212.63L613.11,212.73L611.04,212.76L610.53,212.87L609.56,212.89L608.57,213.10L607.31,213.21L606.63,212.96L605.39,212.92L605.02,212.99L604.72,212.94L604.08,213.03L603.75,212.92L603.05,212.94L602.63,213.07L602.57,213.38L602.54,213.85L602.73,214.72L602.70,215.14L602.58,215.47L602.03,216.10L601.58,216.41L601.03,216.57L600.55,216.86L599.75,217.76L599.04,218.20L598.89,218.41L598.83,218.81L598.87,219.08L600.08,220.82L600.40,221.18L601.05,221.68L601.30,222.06L601.68,222.93L602.54,223.82L602.96,224.10L603.91,225.33L604.93,226.21L605.46,226.79L607.15,228.87L607.71,229.82L608.19,231.04L608.57,232.50L608.94,233.69L609.41,234.89L609.61,235.65L610.56,237.79L610.67,238.63L610.62,239.11L610.45,239.53L610.14,239.91L609.82,240.17L608.80,240.61L607.82,240.77L607.02,241.24L606.73,241.60L606.50,241.72L605.82,241.70L605.55,241.64L605.35,241.82L605.21,242.25L605.14,242.76L605.23,244.13L605.04,244.50L604.42,245.20L604.24,245.52L604.12,246.07L604.16,246.99L604.28,247.46L604.21,248.48L604.13,248.91L602.69,250.31L602.34,251.18L602.51,252.24L602.51,252.83L602.42,253.34L602.18,253.69L601.04,254.92L599.77,255.87L598.66,257.26L598.49,258.01L598.58,258.61L598.53,259.32L598.01,259.87L597.49,260.00L597.21,259.95L597.00,260.25L596.70,261.13L596.19,261.96L596.04,262.37ZM766.95,282.65L767.01,282.78L767.16,282.76L767.16,282.56L767.11,282.53L767.30,282.50L767.53,282.70L767.46,282.73L767.44,282.69L767.43,282.86L767.38,282.86L767.40,282.69L767.28,282.71L767.24,282.84L767.28,282.92L767.52,282.87L767.59,282.67L767.61,282.85L767.70,282.81L767.88,282.49L767.88,282.29L767.76,281.98L767.34,281.88L767.37,282.08L767.47,282.12L767.55,282.32L767.63,282.57L767.62,282.67L767.61,282.59L767.28,282.42L767.07,282.46L766.95,282.65ZM766.28,281.75L766.32,281.81L766.43,281.76L766.46,281.81L766.56,281.56L766.72,281.47L766.84,281.48L766.99,281.37L767.10,281.19L767.11,280.67L766.64,280.78L766.45,280.90L766.44,281.13L766.33,281.29L766.28,281.75ZM764.44,275.94L764.44,276.09L764.61,276.23L764.68,276.45L765.03,276.33L765.20,276.11L765.40,275.99L765.55,276.03L765.59,276.13L765.65,276.08L765.67,275.77L765.77,275.46L765.69,275.10L765.49,275.21L765.37,275.18L765.28,275.30L765.14,275.29L765.17,275.41L765.07,275.53L764.44,275.94ZM761.56,273.66L761.66,273.76L761.79,274.30L761.79,274.44L761.61,274.60L761.62,274.68L761.83,274.59L761.91,274.71L762.21,274.51L762.22,273.72L762.03,273.52L761.89,273.47L761.71,273.50L761.56,273.66Z" fill="#eee9df" stroke="#b4ad9f" stroke-width=".8" vector-effect="non-scaling-stroke" fill-rule="evenodd"><title>المنطقة الشرقية</title></path>
<path d="M704.24,610.37L698.84,611.94L692.68,613.39L684.65,614.43L674.07,616.12L660.85,617.27L648.04,618.46L638.44,619.46L627.91,620.21L618.13,621.12L611.30,621.83L607.26,621.88L603.75,621.78L597.21,621.09L593.53,620.26L589.58,618.56L587.60,617.53L584.52,615.33L581.70,613.43L577.94,611.14L577.99,612.00L577.78,613.69L577.83,614.58L578.15,616.40L578.08,617.19L577.90,618.09L577.45,618.93L576.62,619.68L573.11,621.20L572.26,621.89L571.73,622.81L571.47,623.71L571.56,624.58L571.76,625.40L572.31,627.13L573.03,629.74L573.22,630.70L573.46,633.50L573.41,635.13L573.28,635.92L572.60,637.69L572.17,638.46L571.53,639.29L570.79,640.07L568.08,641.69L566.27,642.45L565.42,642.70L562.80,643.11L561.94,643.38L560.91,643.99L559.83,645.20L559.50,646.13L559.04,647.77L558.63,648.65L557.40,650.33L556.57,651.23L553.85,653.37L551.13,655.79L548.09,658.97L544.59,662.89L543.59,664.82L543.45,665.73L543.49,666.71L543.92,668.45L544.62,670.38L544.83,671.19L545.00,671.99L545.07,673.71L545.04,674.49L544.70,677.19L544.32,678.86L544.22,679.65L544.28,682.44L544.18,683.26L543.90,684.14L543.42,684.94L543.52,685.77L543.73,686.58L544.36,688.28L544.84,689.94L544.96,690.88L544.90,691.71L544.66,692.48L542.77,697.21L545.34,699.64L548.64,699.33L549.47,699.30L550.13,700.06L551.15,700.78L554.03,700.67L556.42,701.12L558.17,698.18L558.64,698.63L559.66,698.08L560.34,698.04L561.66,699.67L562.06,699.65L562.84,698.95L563.13,697.91L564.75,697.71L565.88,697.71L566.61,697.54L572.05,696.91L572.93,696.84L575.98,696.84L577.93,697.40L578.46,697.57L578.53,697.64L579.06,697.72L579.90,697.97L583.14,696.84L599.74,696.91L605.35,696.83L607.42,697.97L612.50,701.01L626.12,702.71L639.88,704.42L650.32,705.11L665.31,703.06L675.10,716.66L682.27,716.70L693.35,709.88L694.11,698.25L695.11,682.91L697.91,660.34L699.26,649.41L701.09,634.65L702.61,621.89L704.24,610.37Z" fill="#e9e5db" stroke="#b4ad9f" stroke-width=".8" vector-effect="non-scaling-stroke" fill-rule="evenodd"><title>منطقة نجران</title></path>
<path d="M637.97,199.20L611.72,196.74L585.97,194.48L573.29,183.62L560.17,172.44L547.21,161.47L534.63,150.88L521.87,140.19L509.27,129.71L496.26,118.95L482.83,107.91L457.60,95.94L437.35,82.57L417.40,69.52L393.29,64.61L369.97,60.00L369.12,60.88L362.36,67.10L339.25,73.65L321.64,78.83L321.16,79.58L320.98,80.08L321.15,80.61L321.47,81.20L321.52,81.74L321.47,81.96L321.27,82.24L321.08,82.35L320.68,82.78L320.55,83.21L320.59,84.02L320.53,84.44L321.26,85.97L321.31,86.39L320.91,87.14L320.85,87.44L321.01,88.04L320.98,88.39L320.69,89.14L321.29,90.57L321.47,90.78L321.56,91.52L321.32,91.75L320.96,91.96L320.71,92.32L320.73,92.82L320.86,93.42L321.24,94.63L321.50,95.20L321.50,95.86L321.45,96.22L321.06,96.68L320.76,96.89L320.48,97.31L320.38,97.61L320.33,97.95L320.54,99.14L320.48,100.07L320.12,100.60L319.75,100.97L319.77,101.62L319.46,102.26L319.27,102.31L318.66,102.91L318.69,103.22L318.77,103.52L319.12,103.83L319.65,104.02L320.09,103.96L320.44,103.73L320.88,103.54L321.74,103.85L322.66,103.62L323.03,103.72L323.48,104.06L324.59,105.90L325.19,106.32L325.47,106.10L325.66,105.50L325.94,105.24L326.24,105.15L326.75,105.65L326.78,106.10L326.52,106.50L325.95,107.05L325.76,107.52L325.94,108.26L326.29,108.63L326.80,108.57L327.08,108.35L327.58,107.78L327.98,107.78L329.09,108.94L329.28,109.50L329.24,110.10L328.96,110.80L329.30,111.68L329.80,112.36L329.99,112.94L329.96,113.45L329.74,114.06L329.39,115.59L329.35,116.89L329.95,118.88L330.12,119.18L330.57,119.44L330.81,119.36L331.11,119.14L331.34,119.12L331.96,119.29L332.52,119.23L332.83,119.08L333.04,118.81L333.24,118.41L333.56,118.14L334.11,118.09L335.11,117.68L335.51,117.58L336.00,117.33L337.95,117.24L339.29,117.30L340.51,117.28L341.26,117.22L343.01,116.83L343.62,116.54L344.25,116.11L344.82,116.03L345.89,116.12L347.53,116.10L349.44,116.50L349.61,116.37L350.19,116.33L351.29,116.10L351.71,116.20L352.19,116.47L352.60,117.05L352.71,118.03L352.93,118.23L353.33,118.41L353.96,118.39L354.19,117.82L354.46,117.46L354.74,117.37L356.07,117.48L356.99,117.23L359.70,115.01L360.16,114.86L360.49,114.90L360.83,114.76L361.03,114.60L361.33,114.50L361.96,114.48L362.64,114.70L362.94,114.55L363.44,114.50L365.98,114.47L366.69,114.35L370.30,114.41L371.89,114.65L374.39,115.24L374.64,115.47L375.69,116.12L377.18,117.49L377.71,117.84L378.31,117.96L378.92,117.90L379.37,117.75L380.16,117.21L381.07,116.83L381.37,116.79L381.71,116.76L382.48,116.90L382.80,116.89L383.67,117.11L384.09,117.08L384.52,116.81L384.96,116.70L386.01,117.33L386.30,117.31L386.57,117.18L386.82,116.92L387.53,116.90L389.21,118.39L389.86,118.84L391.03,119.45L391.64,119.65L392.29,119.65L392.23,119.59L393.02,119.41L393.82,119.00L394.30,118.87L394.85,118.46L395.32,118.23L396.09,118.05L396.55,118.50L396.91,119.31L397.20,119.70L397.77,119.88L398.77,120.05L399.45,120.71L399.92,121.31L400.40,122.09L400.84,122.39L401.75,122.75L402.06,123.00L402.47,123.10L402.75,123.30L403.72,124.45L404.54,125.05L404.91,125.11L406.52,125.13L407.20,125.05L408.07,124.90L409.03,124.54L409.41,123.96L409.53,123.59L409.91,123.01L410.15,122.83L410.63,122.62L411.32,121.56L411.86,121.09L412.19,121.00L412.94,120.96L413.48,120.99L414.33,121.27L414.83,121.54L415.39,122.05L416.89,122.16L417.51,122.74L418.27,123.00L419.09,122.81L419.68,122.86L419.99,122.94L420.75,123.51L426.14,124.68L428.22,124.99L429.74,124.78L430.34,124.62L431.67,124.55L432.36,124.33L433.10,124.22L434.93,123.62L436.20,123.43L436.72,123.54L437.43,123.43L438.16,123.58L440.79,123.88L442.75,124.42L444.30,125.31L444.42,125.71L444.34,126.08L444.14,126.36L444.27,126.75L444.98,127.52L445.12,128.10L445.14,128.60L445.04,129.12L445.03,130.60L444.94,131.16L444.60,131.64L444.32,132.50L444.28,133.34L444.30,134.13L444.49,134.78L444.79,135.32L445.34,135.53L445.73,135.47L446.06,135.25L446.46,134.61L446.71,133.69L447.41,132.69L448.07,132.21L448.88,131.99L450.24,132.09L451.45,132.32L452.11,132.34L452.78,132.48L453.34,132.48L454.48,132.64L454.96,132.60L455.41,132.41L456.15,131.96L456.73,131.76L457.12,131.82L457.53,132.27L458.00,133.01L458.51,133.30L460.38,133.57L460.88,133.58L461.56,133.74L461.93,134.15L462.46,134.46L463.09,134.55L464.19,134.55L464.77,134.67L465.11,134.83L465.29,135.20L465.31,135.64L465.03,136.15L464.67,136.65L464.12,137.23L463.55,137.79L462.39,138.65L462.25,138.92L462.28,140.30L462.63,141.55L463.13,142.13L463.63,142.57L464.35,142.94L465.69,143.14L466.86,143.07L467.60,143.15L468.83,143.53L469.41,143.84L469.88,144.30L470.16,144.75L470.34,145.39L470.68,146.04L471.05,146.46L471.53,146.75L471.92,147.16L472.56,148.48L473.39,149.12L474.16,150.23L474.58,150.70L474.99,150.99L475.38,151.08L476.88,152.49L478.30,153.62L478.51,153.86L478.64,154.42L478.53,154.82L477.78,155.47L477.26,155.70L476.84,155.74L475.75,156.12L474.52,156.22L473.48,156.48L472.39,156.65L471.83,156.93L471.27,157.34L470.47,158.23L470.32,159.25L470.22,161.90L470.02,162.21L469.67,162.50L469.30,162.43L468.96,162.26L468.26,161.74L466.82,162.18L465.97,162.32L464.53,161.91L464.01,161.90L463.55,162.04L463.05,162.25L462.22,162.95L461.85,164.36L461.58,164.63L461.14,164.89L460.75,165.03L460.48,165.39L460.32,166.22L460.06,166.63L458.13,166.57L457.27,166.98L456.45,167.05L456.08,167.25L455.81,167.85L456.63,169.06L456.75,169.46L456.46,171.21L456.43,173.06L456.49,173.79L456.43,174.63L456.56,175.46L456.48,176.05L456.60,176.71L456.66,178.11L456.86,179.56L457.10,179.90L457.50,179.94L457.82,179.75L458.04,179.47L458.20,178.67L458.46,178.40L458.87,178.33L460.06,178.29L460.38,178.71L460.36,179.32L460.05,180.53L460.01,181.18L459.99,182.68L460.11,183.48L461.34,185.32L461.93,187.77L462.14,188.35L462.59,188.82L463.14,189.10L463.51,189.41L463.52,189.85L463.40,190.36L463.03,191.36L462.96,191.81L463.38,193.81L463.62,194.28L463.87,195.01L464.13,196.46L464.14,197.04L463.96,197.39L462.73,197.68L461.12,197.29L459.63,197.29L458.98,197.41L458.00,197.75L457.29,198.09L456.65,198.64L456.24,199.26L455.97,200.07L455.63,201.33L455.59,201.98L455.04,202.91L454.50,204.40L454.12,206.15L454.01,207.51L454.00,209.05L455.59,208.84L458.45,209.07L459.25,209.21L459.81,209.40L460.16,209.56L460.74,210.04L461.41,210.31L462.60,210.43L463.48,210.41L463.85,210.50L466.37,210.75L468.39,210.75L469.94,210.97L470.26,210.94L471.19,211.12L472.25,211.12L473.16,211.01L474.01,211.09L475.55,212.03L475.89,212.06L476.80,211.91L477.12,211.58L477.99,211.18L478.26,210.94L478.93,210.52L479.62,209.92L480.13,209.61L480.51,209.20L481.00,208.82L481.33,208.74L481.72,208.88L482.27,209.44L482.81,210.46L483.27,210.86L484.48,211.26L485.07,211.33L485.44,211.52L486.59,211.80L487.37,211.91L488.24,212.20L489.45,212.73L490.97,213.15L491.53,213.18L492.17,213.34L492.76,213.37L494.32,213.71L495.77,213.68L496.60,213.79L498.00,214.15L499.50,214.89L500.15,215.44L500.59,216.19L500.65,216.62L500.22,217.52L500.33,217.77L500.75,218.10L501.18,218.29L501.47,218.52L501.68,218.99L501.90,219.78L502.15,220.11L503.25,220.73L504.34,220.93L504.73,220.84L504.99,220.60L505.83,220.35L506.17,220.55L507.27,221.62L508.31,221.69L508.97,221.87L509.46,222.06L510.32,222.70L510.55,223.11L510.75,223.64L510.74,224.05L510.52,224.81L510.19,225.10L509.95,225.55L509.91,226.09L510.00,226.62L509.89,226.99L509.91,227.80L509.98,228.05L510.31,228.34L511.33,228.51L511.92,228.82L512.04,229.01L512.08,229.22L511.57,230.11L511.55,230.49L511.74,230.73L512.03,230.93L512.43,230.86L512.93,230.61L513.50,230.09L514.05,229.81L514.40,229.86L514.72,229.98L515.13,230.26L515.70,230.82L515.99,230.93L516.36,230.93L516.67,231.06L518.05,231.92L518.47,232.09L519.20,232.04L519.72,232.09L520.35,232.07L520.94,232.23L521.41,232.52L522.42,233.37L522.71,233.39L522.87,233.32L523.26,233.02L523.48,232.76L523.69,232.36L525.11,231.51L525.65,230.99L525.97,230.77L526.37,230.62L526.75,230.81L527.14,230.77L527.81,230.41L528.22,229.91L528.64,229.05L528.92,228.84L529.59,228.90L530.27,229.11L530.62,229.12L530.91,228.96L531.26,228.91L531.66,228.96L532.51,229.50L533.66,229.51L533.98,229.69L534.13,229.96L534.32,229.84L535.32,228.66L535.68,228.62L535.79,228.85L535.76,230.84L535.86,231.32L536.02,231.65L536.32,231.87L536.80,231.90L537.04,231.65L536.91,230.21L536.98,228.84L537.26,228.29L537.43,228.19L537.76,228.19L537.94,228.26L538.32,228.90L538.76,229.28L539.47,228.75L539.76,228.68L540.00,228.79L540.56,229.26L540.77,229.31L540.86,229.10L541.35,228.72L542.01,228.91L542.46,228.84L543.44,228.27L544.28,227.53L544.55,227.44L544.87,227.45L545.03,227.63L545.51,228.62L545.72,228.76L545.92,228.80L546.07,228.75L547.09,229.12L547.61,229.42L547.86,229.45L548.18,229.38L548.72,228.94L549.15,229.05L549.30,229.38L549.99,229.68L550.40,229.63L550.71,229.39L551.00,228.86L551.48,228.44L551.86,228.60L552.20,228.91L553.01,230.03L554.27,232.02L554.39,232.54L554.41,233.04L554.27,233.53L554.01,233.91L553.57,234.33L553.30,234.83L553.12,236.40L553.36,237.53L553.61,238.12L554.12,240.00L554.14,240.84L554.89,243.03L555.39,244.42L555.29,245.88L554.72,248.25L554.24,249.22L554.23,249.94L554.71,250.30L555.25,250.45L555.81,250.44L556.19,250.79L556.12,251.39L556.32,252.20L556.71,252.48L557.38,252.52L558.22,251.91L558.55,251.58L558.83,251.10L559.31,250.65L560.32,249.84L560.97,249.42L563.36,249.26L564.15,249.30L566.67,250.10L568.82,251.32L569.22,251.92L569.84,252.32L570.20,252.78L570.59,253.13L572.99,254.46L574.38,255.74L575.19,256.28L576.44,256.91L577.46,257.73L578.04,258.52L578.51,258.80L578.97,259.43L579.24,259.68L581.06,263.02L583.34,265.58L583.80,266.60L584.06,266.97L584.23,267.33L584.81,267.98L585.93,268.88L586.11,268.89L586.39,268.77L587.20,268.21L588.28,267.32L588.79,267.16L589.63,266.74L590.00,266.68L591.04,266.74L592.75,266.45L593.67,265.78L594.12,265.35L594.42,264.77L595.49,263.30L596.04,262.37L596.19,261.96L596.70,261.13L597.00,260.25L597.21,259.95L597.49,260.00L598.01,259.87L598.53,259.32L598.58,258.61L598.49,258.01L598.66,257.26L599.77,255.87L601.04,254.92L602.18,253.69L602.42,253.34L602.51,252.83L602.51,252.24L602.34,251.18L602.69,250.31L604.13,248.91L604.21,248.48L604.28,247.46L604.16,246.99L604.12,246.07L604.24,245.52L604.42,245.20L605.04,244.50L605.23,244.13L605.14,242.76L605.21,242.25L605.35,241.82L605.55,241.64L605.82,241.70L606.50,241.72L606.73,241.60L607.02,241.24L607.82,240.77L608.80,240.61L609.82,240.17L610.14,239.91L610.45,239.53L610.62,239.11L610.67,238.63L610.56,237.79L609.61,235.65L609.41,234.89L608.94,233.69L608.57,232.50L608.19,231.04L607.71,229.82L607.15,228.87L605.46,226.79L604.93,226.21L603.91,225.33L602.96,224.10L602.54,223.82L601.68,222.93L601.30,222.06L601.05,221.68L600.40,221.18L600.08,220.82L598.87,219.08L598.83,218.81L598.89,218.41L599.04,218.20L599.75,217.76L600.55,216.86L601.03,216.57L601.58,216.41L602.03,216.10L602.58,215.47L602.70,215.14L602.73,214.72L602.54,213.85L602.57,213.38L602.63,213.07L603.05,212.94L603.75,212.92L604.08,213.03L604.72,212.94L605.02,212.99L605.39,212.92L606.63,212.96L607.31,213.21L608.57,213.10L609.56,212.89L610.53,212.87L611.04,212.76L613.11,212.73L613.53,212.63L616.85,212.92L618.37,213.19L619.22,213.27L621.16,213.26L622.36,212.99L622.90,212.71L623.36,212.20L624.12,210.73L624.31,210.17L624.43,209.04L624.93,208.83L626.35,208.98L626.85,208.69L627.05,208.37L627.32,206.95L627.70,206.56L628.41,206.30L629.41,206.11L629.99,205.82L630.44,205.31L630.97,204.53L631.93,203.79L632.92,203.43L633.99,203.39L635.17,202.99L636.18,202.30L637.04,201.36L637.97,199.20Z" fill="#f3efe7" stroke="#b4ad9f" stroke-width=".8" vector-effect="non-scaling-stroke" fill-rule="evenodd"><title>منطقة الحدود الشمالية</title></path>
<path d="M454.00,209.05L453.03,209.14L451.99,209.00L450.96,208.74L450.30,208.69L448.30,209.15L446.01,209.15L445.03,208.98L444.23,209.07L443.48,208.96L442.06,209.07L441.69,209.02L440.76,208.74L439.73,208.82L438.86,208.73L437.24,208.70L436.00,208.57L435.41,208.35L433.75,208.23L433.15,208.46L432.47,208.60L432.15,208.75L431.58,208.84L430.28,208.91L429.09,208.88L428.66,208.57L428.10,208.29L427.47,208.17L426.63,208.21L425.17,208.45L424.19,208.47L423.89,208.42L423.51,208.47L423.25,208.39L422.28,208.43L422.01,208.50L421.92,208.65L421.03,209.23L419.95,210.09L419.64,210.41L418.45,211.28L417.84,211.62L413.79,213.04L412.39,214.01L411.57,214.45L411.24,214.51L410.95,214.70L410.16,214.92L409.39,215.34L409.07,215.58L408.14,215.87L407.54,216.36L407.28,216.72L406.91,216.91L406.17,217.13L405.63,217.44L402.65,218.83L399.30,219.73L396.98,220.87L394.85,222.49L394.50,222.64L394.31,222.93L393.66,223.39L393.07,223.69L392.99,223.98L392.41,224.76L391.61,226.09L391.26,227.12L390.83,227.77L389.16,229.87L388.22,230.31L385.90,232.16L385.26,233.45L383.59,235.05L382.90,236.22L381.76,237.31L381.01,238.56L379.94,239.59L379.54,240.37L378.87,241.25L378.41,242.37L377.87,243.04L376.56,245.13L376.40,245.77L376.49,246.89L376.91,247.87L377.87,248.73L378.34,249.01L380.80,250.07L382.36,250.64L383.75,251.05L384.56,251.46L385.66,251.83L386.76,252.46L387.12,252.86L387.98,253.37L389.02,253.56L389.66,253.83L390.37,254.38L391.37,254.86L391.98,255.44L392.53,255.82L394.68,258.30L395.36,259.49L396.61,262.06L396.86,262.86L397.18,263.61L397.40,264.65L397.65,265.32L397.79,266.03L398.07,268.71L398.13,270.09L398.04,272.02L398.30,273.15L398.54,273.65L399.43,274.99L399.80,276.08L399.92,276.68L400.14,277.31L400.58,278.22L400.78,278.62L401.67,279.74L401.75,280.32L401.75,280.94L401.64,281.56L401.69,282.00L401.51,282.44L401.41,283.38L401.66,284.56L401.23,286.31L400.71,286.58L400.38,286.48L400.15,286.05L400.02,285.50L399.80,285.34L399.57,285.24L398.75,285.31L398.20,285.09L397.63,285.01L397.26,284.74L396.97,284.39L396.80,283.94L396.82,282.80L396.31,282.55L396.07,282.52L395.28,281.88L394.41,281.51L393.65,281.54L392.84,281.88L392.17,282.58L391.89,283.01L391.53,284.05L391.46,284.50L391.19,284.86L390.81,286.73L389.86,287.94L387.09,291.05L386.70,291.56L385.61,293.75L384.77,294.39L384.35,295.58L383.67,296.18L383.01,296.64L382.28,297.48L381.76,298.21L381.52,299.28L381.41,300.27L381.45,301.39L381.53,301.83L381.70,302.28L381.93,302.64L382.22,302.85L382.73,302.99L383.74,303.12L385.69,303.22L387.29,303.42L387.59,303.51L387.91,303.80L387.99,304.38L387.94,305.35L388.03,306.23L388.25,306.61L388.62,306.96L390.34,308.11L390.77,308.55L391.16,308.84L391.89,309.62L392.35,310.34L392.68,311.98L393.05,312.84L393.37,313.24L393.81,313.58L394.13,313.98L394.50,314.65L394.74,315.25L396.30,317.40L396.78,318.37L396.70,318.76L397.01,319.70L397.02,320.07L396.65,322.91L396.65,323.30L396.86,324.19L397.31,325.08L397.82,325.77L398.03,326.21L399.14,327.90L399.64,328.79L399.75,329.20L399.69,330.11L399.43,332.04L399.50,332.67L399.35,335.35L399.42,335.77L399.38,336.83L399.59,337.90L399.62,338.54L399.77,338.91L399.79,339.43L399.68,339.99L399.73,340.74L399.49,341.72L399.41,342.39L399.34,344.21L399.54,345.08L399.78,345.69L400.35,346.63L400.41,347.13L400.30,347.55L400.10,349.78L399.68,351.57L399.67,354.42L399.77,355.84L399.60,356.33L399.82,358.35L400.16,359.81L400.37,360.43L401.28,361.61L401.79,362.46L402.71,363.04L403.24,363.20L404.17,363.26L405.25,364.05L405.55,364.20L407.13,364.53L408.29,364.93L410.46,364.99L411.15,365.16L412.25,365.61L412.79,365.71L413.34,365.69L414.88,365.40L416.18,365.51L416.68,365.78L418.08,366.28L418.59,366.45L419.01,366.48L419.27,366.25L419.30,365.85L419.15,365.27L418.92,363.43L419.10,362.51L419.60,361.44L420.67,360.41L421.09,360.15L421.53,360.03L422.60,360.05L423.23,360.93L423.70,361.21L424.42,361.52L426.18,361.15L428.56,361.14L429.17,361.02L429.84,360.63L430.54,359.72L430.75,359.27L431.08,358.76L431.18,358.34L431.51,357.83L431.98,356.92L432.92,355.47L433.34,355.02L433.75,354.73L435.06,354.14L435.76,353.90L436.66,353.83L437.35,353.95L438.35,354.42L439.70,354.77L441.69,354.36L442.85,354.48L444.00,355.25L444.98,356.42L445.46,356.89L446.29,357.54L447.18,357.88L447.97,358.10L449.65,358.31L451.23,356.65L451.83,355.68L452.77,355.05L453.68,354.73L454.16,354.23L455.03,352.78L455.74,351.74L456.26,350.52L456.54,349.58L457.01,348.56L457.49,347.75L458.28,347.22L458.91,347.14L459.74,347.50L460.57,347.60L461.31,347.42L464.49,347.69L465.59,347.90L466.17,347.89L466.89,347.67L467.45,346.94L467.88,346.24L468.10,345.29L468.42,344.51L468.82,343.84L469.42,343.22L471.68,341.46L472.37,340.83L472.40,340.29L471.79,339.24L471.97,338.18L472.67,337.13L473.09,336.84L473.98,337.13L474.71,337.73L475.48,338.16L476.03,338.15L476.47,337.72L476.35,336.96L476.07,336.45L475.99,335.72L476.76,334.79L477.22,334.36L477.23,333.50L476.99,332.78L476.93,332.22L477.45,331.55L478.10,331.09L478.84,330.77L479.72,330.47L480.23,330.55L480.72,330.19L481.11,329.59L481.43,328.85L481.85,328.28L482.24,327.87L482.94,327.83L483.64,328.08L484.97,329.02L485.75,328.96L486.90,327.49L487.22,326.57L487.37,325.32L487.25,324.58L487.03,323.98L486.01,322.20L485.84,320.69L486.23,320.17L486.71,319.75L487.42,319.31L488.09,318.73L488.84,317.92L489.21,317.39L489.83,315.98L490.37,315.29L491.10,314.99L491.88,314.83L492.69,314.96L493.27,315.47L493.48,315.82L492.98,316.37L492.86,316.73L493.17,317.22L493.56,317.33L494.20,316.80L494.55,315.89L495.07,315.03L495.22,314.32L495.15,313.91L494.98,313.56L495.33,313.23L495.91,313.05L496.58,313.39L497.37,313.64L498.32,313.12L500.75,311.07L501.25,310.46L501.92,309.81L502.64,309.70L503.18,309.39L503.67,308.96L505.50,305.81L506.09,305.11L507.39,303.09L507.54,302.17L507.41,301.48L507.14,300.77L507.15,300.08L507.29,299.77L507.80,299.88L508.08,300.12L508.46,300.18L508.80,299.94L509.09,299.55L509.53,299.52L509.92,299.61L511.39,298.79L511.89,298.45L512.28,297.80L512.65,297.35L513.16,297.38L513.85,297.87L514.29,297.72L514.87,296.56L515.32,296.01L515.72,295.67L516.20,295.52L516.67,295.51L517.37,295.74L518.37,297.37L518.71,297.49L519.05,297.36L519.44,297.02L519.90,296.18L520.32,296.01L520.74,296.07L523.16,293.74L524.14,292.58L524.45,292.03L525.47,290.56L525.63,289.96L525.64,289.29L525.73,288.67L526.04,288.05L526.48,287.75L526.84,287.27L527.06,286.77L527.16,286.24L527.11,285.71L526.57,284.85L526.77,284.39L527.00,284.19L526.99,283.67L526.77,283.12L526.96,282.79L527.59,282.44L530.28,281.32L531.82,281.00L531.98,280.52L531.91,280.12L531.67,279.45L531.07,279.13L530.41,279.03L530.01,279.06L529.38,278.86L529.16,278.64L528.99,278.29L528.90,277.79L528.93,276.98L529.06,276.07L529.56,274.27L529.81,273.79L530.29,273.52L530.79,273.37L533.12,273.27L533.58,273.05L533.72,272.40L533.54,271.83L533.72,271.16L534.28,270.46L534.85,270.19L535.76,270.31L536.81,271.17L537.17,271.26L537.58,271.04L538.07,270.03L538.51,269.90L539.24,269.96L540.69,269.90L541.81,269.76L542.86,269.48L543.85,269.32L545.76,269.35L546.61,269.24L546.92,268.34L546.83,267.13L546.79,264.02L546.48,263.21L546.19,260.82L546.12,259.49L546.37,257.31L546.35,256.30L546.56,254.60L546.96,252.96L547.13,251.90L547.11,251.02L547.23,250.30L547.58,249.81L548.15,249.37L549.29,248.00L553.29,244.28L554.89,243.03L554.14,240.84L554.12,240.00L553.61,238.12L553.36,237.53L553.12,236.40L553.30,234.83L553.57,234.33L554.01,233.91L554.27,233.53L554.41,233.04L554.39,232.54L554.27,232.02L553.01,230.03L552.20,228.91L551.86,228.60L551.48,228.44L551.00,228.86L550.71,229.39L550.40,229.63L549.99,229.68L549.30,229.38L549.15,229.05L548.72,228.94L548.18,229.38L547.86,229.45L547.61,229.42L547.09,229.12L546.07,228.75L545.92,228.80L545.72,228.76L545.51,228.62L545.03,227.63L544.87,227.45L544.55,227.44L544.28,227.53L543.44,228.27L542.46,228.84L542.01,228.91L541.35,228.72L540.86,229.10L540.77,229.31L540.56,229.26L540.00,228.79L539.76,228.68L539.47,228.75L538.76,229.28L538.32,228.90L537.94,228.26L537.76,228.19L537.43,228.19L537.26,228.29L536.98,228.84L536.91,230.21L537.04,231.65L536.80,231.90L536.32,231.87L536.02,231.65L535.86,231.32L535.76,230.84L535.79,228.85L535.68,228.62L535.32,228.66L534.32,229.84L534.13,229.96L533.98,229.69L533.66,229.51L532.51,229.50L531.66,228.96L531.26,228.91L530.91,228.96L530.62,229.12L530.27,229.11L529.59,228.90L528.92,228.84L528.64,229.05L528.22,229.91L527.81,230.41L527.14,230.77L526.75,230.81L526.37,230.62L525.97,230.77L525.65,230.99L525.11,231.51L523.69,232.36L523.48,232.76L523.26,233.02L522.87,233.32L522.71,233.39L522.42,233.37L521.41,232.52L520.94,232.23L520.35,232.07L519.72,232.09L519.20,232.04L518.47,232.09L518.05,231.92L516.67,231.06L516.36,230.93L515.99,230.93L515.70,230.82L515.13,230.26L514.72,229.98L514.40,229.86L514.05,229.81L513.50,230.09L512.93,230.61L512.43,230.86L512.03,230.93L511.74,230.73L511.55,230.49L511.57,230.11L512.08,229.22L512.04,229.01L511.92,228.82L511.33,228.51L510.31,228.34L509.98,228.05L509.91,227.80L509.89,226.99L510.00,226.62L509.91,226.09L509.95,225.55L510.19,225.10L510.52,224.81L510.74,224.05L510.75,223.64L510.55,223.11L510.32,222.70L509.46,222.06L508.97,221.87L508.31,221.69L507.27,221.62L506.17,220.55L505.83,220.35L504.99,220.60L504.73,220.84L504.34,220.93L503.25,220.73L502.15,220.11L501.90,219.78L501.68,218.99L501.47,218.52L501.18,218.29L500.75,218.10L500.33,217.77L500.22,217.52L500.65,216.62L500.59,216.19L500.15,215.44L499.50,214.89L498.00,214.15L496.60,213.79L495.77,213.68L494.32,213.71L492.76,213.37L492.17,213.34L491.53,213.18L490.97,213.15L489.45,212.73L488.24,212.20L487.37,211.91L486.59,211.80L485.44,211.52L485.07,211.33L484.48,211.26L483.27,210.86L482.81,210.46L482.27,209.44L481.72,208.88L481.33,208.74L481.00,208.82L480.51,209.20L480.13,209.61L479.62,209.92L478.93,210.52L478.26,210.94L477.99,211.18L477.12,211.58L476.80,211.91L475.89,212.06L475.55,212.03L474.01,211.09L473.16,211.01L472.25,211.12L471.19,211.12L470.26,210.94L469.94,210.97L468.39,210.75L466.37,210.75L463.85,210.50L463.48,210.41L462.60,210.43L461.41,210.31L460.74,210.04L460.16,209.56L459.81,209.40L459.25,209.21L458.45,209.07L455.59,208.84L454.00,209.05Z" fill="#eee9df" stroke="#b4ad9f" stroke-width=".8" vector-effect="non-scaling-stroke" fill-rule="evenodd"><title>منطقة حائل</title></path>
<path d="M596.04,262.37L595.49,263.30L594.42,264.77L594.12,265.35L593.67,265.78L592.75,266.45L591.04,266.74L590.00,266.68L589.63,266.74L588.79,267.16L588.28,267.32L587.20,268.21L586.39,268.77L586.11,268.89L585.93,268.88L586.82,269.90L587.18,270.72L588.00,272.05L588.96,273.32L589.62,274.35L589.91,275.11L590.39,275.90L590.74,276.82L591.97,278.88L592.61,282.16L592.22,285.78L592.07,288.24L592.04,290.63L593.03,294.61L592.77,295.04L592.24,294.95L591.61,294.33L590.94,293.32L589.39,292.51L586.54,290.62L585.41,290.03L584.49,290.05L582.70,290.60L581.18,290.63L579.69,291.23L578.56,291.23L577.30,291.49L575.80,291.93L574.45,292.76L573.96,293.65L573.79,294.77L573.73,296.67L573.61,298.34L573.15,301.37L572.91,304.20L573.04,305.57L573.04,308.92L573.30,310.44L573.63,311.78L574.50,318.03L575.26,321.23L575.39,322.24L575.84,323.93L576.46,325.75L576.94,328.05L577.36,331.93L577.53,332.73L578.11,333.91L578.64,334.77L579.10,335.66L579.52,336.69L579.64,337.66L580.04,338.79L581.03,341.11L581.47,341.83L582.28,343.58L582.91,344.57L583.49,345.15L585.32,346.20L586.20,346.60L587.78,347.77L588.80,348.74L588.60,349.95L587.77,352.21L587.41,353.75L587.03,354.87L585.77,357.18L584.88,357.08L579.99,354.45L578.43,353.49L576.37,352.65L575.59,352.63L574.93,352.97L574.49,353.84L573.97,354.53L573.27,355.61L572.37,356.04L569.37,357.22L567.32,357.95L566.07,358.20L564.55,358.65L562.31,358.93L561.05,359.22L560.58,359.34L559.86,359.69L559.40,360.50L558.24,361.98L556.57,364.51L554.47,364.46L551.25,364.67L549.08,365.11L548.13,365.24L547.31,365.61L546.93,366.06L546.92,366.74L547.18,367.72L547.29,368.57L547.18,369.57L546.59,370.45L545.55,371.10L544.25,371.66L542.33,371.91L541.16,371.79L539.62,370.68L539.28,370.19L538.33,369.76L537.41,369.83L536.36,370.18L535.77,370.91L534.85,371.84L534.86,372.17L533.17,374.75L532.68,375.21L529.42,377.86L528.94,378.43L528.12,378.59L526.49,378.34L525.26,378.34L524.18,378.55L523.50,379.61L522.56,381.54L522.46,382.64L522.64,383.39L523.26,383.50L523.97,382.73L524.95,381.97L525.67,381.79L526.65,381.93L527.19,382.51L527.15,383.14L526.72,383.96L525.99,384.72L525.20,385.16L524.52,385.27L523.45,385.59L523.11,385.90L523.18,386.50L523.83,387.24L524.79,387.87L524.82,388.15L524.29,388.49L522.82,388.86L522.29,389.27L522.26,389.87L522.68,390.41L523.24,390.47L524.92,389.97L526.22,389.72L526.74,390.20L526.78,390.71L526.53,391.75L525.37,393.59L524.28,395.95L523.83,396.80L523.35,398.18L521.23,399.90L520.48,399.91L519.71,399.46L518.48,398.94L517.26,398.71L514.80,398.57L513.06,399.15L512.24,399.64L511.74,400.14L511.67,401.38L511.74,401.93L511.37,402.40L510.24,403.01L508.13,403.07L503.09,402.93L501.53,402.95L500.91,402.22L500.33,401.29L499.26,399.12L498.92,398.36L498.09,395.79L497.54,394.58L496.95,393.08L496.43,391.99L495.17,391.65L491.39,391.61L489.39,391.81L487.51,391.48L485.80,391.44L484.56,391.67L483.32,396.52L483.35,397.20L483.02,398.73L482.78,399.39L482.33,400.36L481.68,401.35L479.88,403.36L479.44,404.24L479.45,405.47L480.32,407.46L480.50,408.35L479.87,409.62L479.58,410.48L478.90,413.92L478.85,415.55L478.95,417.14L478.75,420.41L478.93,422.05L479.02,423.91L478.95,424.23L478.69,425.63L478.45,428.39L478.46,430.17L478.56,431.07L478.86,431.96L479.64,432.76L482.09,434.14L482.93,434.81L483.58,435.73L483.95,436.74L484.63,440.44L484.86,441.28L485.20,442.21L486.61,444.89L487.25,446.67L487.41,448.53L487.84,451.31L488.28,453.11L489.45,455.90L489.93,456.90L490.68,458.84L493.02,464.13L493.48,465.07L494.13,465.96L494.94,466.68L496.71,467.55L497.54,467.83L499.30,468.05L502.01,467.96L502.96,467.86L504.77,467.85L506.53,467.97L510.81,468.41L513.55,468.55L514.36,468.68L516.18,469.19L517.06,469.52L517.89,470.01L518.67,470.68L519.15,471.66L519.18,472.49L518.18,475.01L518.08,475.91L518.22,476.81L518.18,478.53L517.91,480.24L517.85,481.10L517.95,481.96L518.44,482.82L519.29,483.74L522.56,485.09L523.49,485.20L524.40,485.07L527.20,484.23L528.12,483.90L529.91,483.51L530.72,483.57L531.58,483.71L533.58,484.24L536.28,484.63L537.23,485.13L537.64,485.97L537.58,486.91L536.93,489.52L536.60,491.30L536.49,492.27L536.49,493.95L536.59,494.89L537.06,497.70L537.65,502.50L537.86,503.45L537.95,504.32L538.54,508.02L538.52,508.91L538.41,509.88L538.16,510.73L537.75,511.58L537.31,512.38L536.71,513.15L535.99,513.91L535.44,514.70L535.10,515.49L535.12,516.38L535.30,518.23L535.47,519.17L535.62,520.82L535.72,523.42L535.44,525.24L534.15,529.69L533.76,531.37L533.72,533.11L533.89,533.94L534.38,534.79L537.31,538.39L538.46,540.19L539.79,542.05L542.57,546.58L543.08,547.60L543.71,548.49L544.72,550.11L545.67,550.72L546.48,551.39L548.83,553.65L550.51,555.01L551.26,555.83L552.07,556.86L554.75,561.46L556.41,563.43L557.31,564.31L558.83,565.57L559.62,566.49L560.04,567.56L560.24,568.45L560.26,569.26L560.19,570.14L559.94,570.93L559.62,571.72L559.17,572.50L557.27,575.25L556.45,576.88L556.15,577.89L555.56,580.62L555.47,581.46L555.26,582.24L554.76,584.86L554.65,585.74L555.01,586.69L555.82,587.59L557.36,589.12L558.13,590.07L558.68,590.99L560.56,594.58L561.14,595.50L562.57,597.34L566.70,601.09L568.33,602.83L570.61,604.96L573.96,607.74L576.28,609.86L577.15,610.59L577.94,611.14L581.70,613.43L584.52,615.33L587.60,617.53L589.58,618.56L593.53,620.26L597.21,621.09L603.75,621.78L607.26,621.88L611.30,621.83L618.13,621.12L627.91,620.21L638.44,619.46L648.04,618.46L660.85,617.27L674.07,616.12L684.65,614.43L692.68,613.39L698.84,611.94L704.24,610.37L706.47,591.29L708.53,575.26L710.25,561.60L712.01,549.04L714.35,531.76L716.54,512.94L723.94,451.90L725.51,442.74L726.32,435.86L726.34,427.01L725.54,422.05L724.80,419.14L723.45,416.10L721.44,412.54L719.54,410.19L714.54,405.37L709.11,399.74L705.78,397.20L700.96,394.00L698.39,392.54L695.80,390.48L695.24,389.72L694.52,388.44L694.14,387.05L693.75,384.21L693.36,375.76L693.67,363.00L694.08,358.55L693.97,357.57L693.91,355.88L693.93,352.82L694.11,351.03L694.04,348.97L694.09,347.93L693.70,342.16L693.64,332.72L693.47,329.95L693.16,327.06L692.86,325.93L692.55,325.06L691.92,324.14L691.05,323.48L690.24,323.00L687.54,321.76L686.52,321.42L680.99,319.18L680.13,318.89L678.31,318.48L677.45,318.15L676.66,317.11L676.27,316.17L675.79,315.27L674.96,314.54L672.60,312.09L671.78,311.45L670.92,310.67L669.30,308.98L668.46,308.34L667.60,308.09L666.70,308.08L661.69,309.24L660.72,309.30L659.02,309.19L657.06,308.40L656.18,307.83L653.52,305.92L650.88,303.81L648.25,301.90L646.65,300.82L644.99,299.90L642.43,299.10L641.58,298.98L638.79,298.86L637.04,298.94L635.18,298.86L633.37,298.58L632.51,298.14L630.92,297.02L630.00,296.25L627.91,294.27L626.21,292.30L624.76,290.30L623.97,289.46L623.12,289.02L620.32,288.94L619.31,288.86L617.45,288.25L616.55,287.81L612.84,285.54L611.90,285.10L609.02,283.51L606.51,281.96L605.63,281.31L604.68,280.46L604.23,279.56L603.77,277.65L603.46,275.64L603.12,274.67L602.63,273.73L601.90,272.73L599.19,268.16L599.04,265.17L598.62,264.33L598.28,263.86L596.88,262.86L596.04,262.37Z" fill="#e9e5db" stroke="#b4ad9f" stroke-width=".8" vector-effect="non-scaling-stroke" fill-rule="evenodd"><title>منطقة الرياض</title></path>
<path d="M577.94,611.14L577.15,610.59L576.28,609.86L573.96,607.74L570.61,604.96L568.33,602.83L566.70,601.09L562.57,597.34L561.14,595.50L560.56,594.58L558.68,590.99L558.13,590.07L557.36,589.12L555.82,587.59L555.01,586.69L554.65,585.74L554.76,584.86L555.26,582.24L555.47,581.46L555.56,580.62L556.15,577.89L556.45,576.88L557.27,575.25L559.17,572.50L559.62,571.72L559.94,570.93L560.19,570.14L560.26,569.26L560.24,568.45L560.04,567.56L559.62,566.49L558.83,565.57L557.31,564.31L556.41,563.43L554.75,561.46L552.07,556.86L551.26,555.83L550.51,555.01L548.83,553.65L546.48,551.39L545.67,550.72L544.72,550.11L542.02,550.96L541.08,551.32L540.15,551.79L539.23,552.31L537.40,553.57L535.47,555.21L534.62,555.69L533.67,556.09L532.77,556.14L531.89,556.00L531.11,555.72L530.25,555.31L528.55,554.73L527.74,554.60L526.00,554.71L525.18,554.88L524.23,555.24L522.46,556.29L521.59,556.61L520.77,556.75L518.23,556.62L517.42,556.68L516.43,556.93L515.47,557.40L514.58,558.15L513.79,558.95L512.50,560.52L511.59,561.27L510.71,561.76L509.02,562.30L507.88,562.76L506.08,563.69L504.31,564.77L503.35,565.52L501.55,567.10L496.97,571.40L494.71,573.75L492.84,575.33L491.91,575.97L491.08,576.34L490.22,576.63L489.40,576.70L488.53,576.60L487.62,576.35L486.84,575.88L486.03,575.18L485.39,574.33L484.88,573.38L483.48,570.41L482.37,567.77L481.44,565.94L480.91,565.03L480.73,564.82L480.31,566.22L479.86,567.20L478.55,570.47L478.35,571.26L478.30,572.14L478.56,573.04L479.35,573.79L480.16,574.40L480.88,575.29L481.10,576.20L481.00,577.11L480.69,577.94L479.83,578.77L476.94,580.79L476.02,581.34L475.11,581.99L474.16,582.55L473.22,583.30L472.54,584.56L472.26,585.35L472.17,586.26L472.46,588.47L472.53,589.40L472.18,590.22L470.21,593.63L469.82,594.57L468.99,596.16L468.72,596.99L468.32,597.88L468.05,598.67L467.68,599.33L469.40,598.48L470.29,598.41L471.12,598.47L471.99,598.63L472.78,598.98L473.57,599.87L474.09,600.78L474.34,601.72L474.33,602.51L473.84,605.09L473.35,606.78L472.89,608.74L472.78,609.58L472.84,610.42L473.01,611.27L473.50,612.13L474.77,613.85L475.22,614.68L475.33,615.49L475.26,616.59L474.56,618.18L474.05,619.88L473.88,620.74L473.57,621.56L473.03,622.36L472.19,622.93L471.30,623.22L469.66,623.28L468.83,623.11L465.92,622.72L464.27,622.86L463.37,623.15L461.65,624.19L460.76,624.99L458.97,626.37L458.13,626.44L457.36,625.86L456.68,625.02L455.89,624.30L455.65,625.13L455.56,625.99L455.69,627.86L455.68,628.67L455.49,630.29L455.32,633.85L455.33,635.71L455.55,637.40L456.03,639.10L456.43,639.99L457.19,640.79L457.96,641.18L458.82,641.34L462.18,641.15L463.05,641.16L464.75,641.30L466.44,641.58L468.04,642.20L468.83,642.96L469.74,645.68L470.38,648.25L470.68,649.14L472.07,652.59L472.21,653.40L472.20,654.27L471.89,655.65L471.66,656.43L471.49,657.26L470.94,659.00L470.50,660.71L471.35,660.43L473.06,660.34L474.77,660.13L475.57,660.18L476.34,660.53L477.00,661.35L477.11,662.20L477.13,663.04L476.33,667.14L476.07,669.66L476.06,670.48L476.38,671.37L477.22,671.91L478.01,671.89L480.57,671.40L481.44,671.37L482.23,671.45L483.01,671.66L483.76,672.30L484.24,673.20L484.71,674.81L484.81,675.61L485.21,676.50L486.03,677.24L487.67,677.49L488.58,677.50L490.30,677.65L493.83,677.44L494.65,677.50L495.43,677.69L496.19,678.23L496.89,679.11L497.46,680.88L498.07,682.53L498.69,683.37L499.44,684.02L500.20,684.52L500.95,685.20L501.61,686.12L502.38,686.98L503.29,687.46L504.11,687.70L504.90,687.75L505.70,687.54L506.34,686.79L507.43,681.70L508.24,679.07L508.63,678.31L509.60,676.68L510.29,675.93L511.17,675.27L512.87,674.21L513.72,673.74L514.57,673.37L515.38,673.20L516.18,673.64L516.51,674.48L516.10,678.72L516.35,679.53L517.10,680.15L519.61,681.07L520.44,681.61L521.29,682.53L521.79,683.40L522.46,685.21L522.85,686.04L523.44,686.90L524.22,687.55L525.07,687.80L526.75,687.95L527.57,688.24L528.19,689.07L529.28,690.33L531.16,692.83L531.74,692.54L532.48,691.86L533.30,692.37L533.45,691.90L534.22,691.86L535.02,691.48L535.48,691.55L536.03,692.12L536.78,692.03L537.56,692.33L538.00,693.07L538.41,693.19L538.96,693.80L539.23,693.59L540.77,694.86L542.77,697.21L544.66,692.48L544.90,691.71L544.96,690.88L544.84,689.94L544.36,688.28L543.73,686.58L543.52,685.77L543.42,684.94L543.90,684.14L544.18,683.26L544.28,682.44L544.22,679.65L544.32,678.86L544.70,677.19L545.04,674.49L545.07,673.71L545.00,671.99L544.83,671.19L544.62,670.38L543.92,668.45L543.49,666.71L543.45,665.73L543.59,664.82L544.59,662.89L548.09,658.97L551.13,655.79L553.85,653.37L556.57,651.23L557.40,650.33L558.63,648.65L559.04,647.77L559.50,646.13L559.83,645.20L560.91,643.99L561.94,643.38L562.80,643.11L565.42,642.70L566.27,642.45L568.08,641.69L570.79,640.07L571.53,639.29L572.17,638.46L572.60,637.69L573.28,635.92L573.41,635.13L573.46,633.50L573.22,630.70L573.03,629.74L572.31,627.13L571.76,625.40L571.56,624.58L571.47,623.71L571.73,622.81L572.26,621.89L573.11,621.20L576.62,619.68L577.45,618.93L577.90,618.09L578.08,617.19L578.15,616.40L577.83,614.58L577.78,613.69L577.99,612.00L577.94,611.14Z" fill="#f3efe7" stroke="#b4ad9f" stroke-width=".8" vector-effect="non-scaling-stroke" fill-rule="evenodd"><title>منطقة عسير</title></path>
<path d="M365.66,493.67L365.69,494.03L365.77,494.13L365.79,494.26L365.83,494.26L365.80,494.20L365.86,494.26L365.93,494.26L365.95,494.22L365.87,494.25L365.81,494.20L365.94,494.21L365.99,493.54L365.92,493.45L365.91,493.57L365.83,493.54L365.86,493.44L365.78,493.22L365.70,493.27L365.66,493.67ZM359.01,469.05L359.04,469.17L359.05,469.12L359.11,469.12L359.06,469.21L359.18,469.36L359.18,469.56L359.31,469.70L359.70,469.87L359.87,469.81L359.79,469.60L359.73,469.58L359.77,469.50L359.66,469.47L359.66,469.40L359.24,469.34L359.26,469.25L359.14,469.20L359.12,469.08L359.09,469.08L359.11,468.95L359.19,468.97L359.16,468.86L359.03,468.85L359.09,468.93L359.02,468.91L359.01,469.05ZM355.96,466.56L356.14,466.92L356.40,467.03L356.50,467.18L356.73,467.15L356.40,466.88L356.28,466.64L356.16,466.60L356.09,466.50L355.96,466.56ZM353.28,457.56L353.34,457.61L353.31,457.85L353.38,457.96L353.55,458.02L353.58,458.18L353.64,458.12L353.65,458.33L353.71,458.31L353.73,458.43L353.79,458.44L353.86,458.37L353.82,458.09L353.95,458.02L353.88,457.83L353.94,457.58L353.86,457.40L353.74,457.32L353.79,457.27L353.76,457.10L353.57,457.03L353.48,457.10L353.28,457.56ZM435.98,620.47L436.03,621.08L436.26,621.18L436.35,621.10L436.29,621.02L436.42,620.95L436.28,620.90L436.20,620.80L436.36,620.56L436.35,620.50L436.31,620.60L436.21,620.62L436.08,620.38L436.01,620.38L435.98,620.47ZM429.79,604.20L429.87,604.15L429.88,604.04L429.99,604.03L430.05,603.88L430.25,604.11L430.60,604.10L430.67,604.22L430.91,604.17L430.46,603.90L430.05,603.82L429.99,603.95L429.84,604.04L429.79,604.20ZM409.02,583.97L409.24,584.11L409.96,584.28L410.11,584.28L410.16,584.14L409.91,584.05L409.73,584.08L409.70,584.01L409.81,583.93L409.70,583.95L409.71,583.91L409.62,584.01L409.67,584.01L409.66,584.06L409.61,584.07L409.56,583.88L409.48,583.88L409.50,583.95L409.40,583.85L409.02,583.97ZM401.08,580.09L401.21,580.36L401.67,580.74L402.30,581.11L402.52,581.17L403.55,581.09L403.99,581.16L404.22,581.28L404.35,581.43L404.71,581.55L405.06,581.75L405.86,582.45L406.32,582.59L406.74,582.93L407.44,583.21L407.61,583.38L408.24,583.73L408.77,583.95L408.93,583.92L408.77,583.68L408.65,583.66L408.61,583.50L408.47,583.53L408.39,583.43L407.82,583.22L407.13,582.71L407.09,582.73L406.15,582.16L406.16,582.07L406.08,582.07L405.93,581.90L405.72,581.84L404.74,581.27L404.44,581.22L404.08,581.04L403.65,580.92L403.22,580.92L403.26,580.87L403.06,580.80L402.79,580.85L402.56,580.74L401.97,580.64L401.47,580.36L401.34,580.10L401.08,580.09ZM478.95,424.23L477.74,424.11L475.95,423.76L475.01,423.48L473.21,423.25L472.25,423.31L471.37,423.50L470.41,423.88L469.54,424.43L468.97,425.30L468.81,426.24L468.87,427.08L469.87,428.99L470.13,429.96L470.38,431.67L470.32,433.64L470.15,434.49L469.54,436.17L469.02,437.00L468.33,437.91L467.38,438.56L465.59,439.61L463.73,440.48L461.68,441.19L459.78,442.00L458.81,442.46L454.99,444.57L452.38,446.28L451.55,447.06L450.91,447.85L450.79,448.68L451.14,449.56L452.00,450.32L453.75,451.44L455.51,452.32L456.32,452.64L457.23,453.12L458.07,453.68L458.53,454.61L458.53,455.47L457.89,456.34L456.94,457.08L455.14,457.68L453.30,458.22L451.45,458.97L450.61,459.48L449.58,460.21L448.67,460.99L447.92,461.77L447.22,462.68L446.33,463.65L444.72,465.34L442.86,467.07L441.01,468.63L437.03,471.40L432.25,475.20L431.00,476.81L430.57,477.65L428.01,481.83L427.18,482.81L426.27,483.00L425.34,482.83L423.52,482.24L422.73,481.84L421.87,481.65L418.52,480.66L417.63,480.31L416.72,480.70L415.79,481.51L415.04,482.45L414.21,483.36L412.37,485.00L410.65,485.97L409.73,486.28L408.75,486.38L407.82,486.35L406.11,486.11L404.33,485.64L403.63,484.75L403.59,483.84L404.29,481.34L404.54,479.43L404.54,478.53L404.34,477.63L404.03,476.79L403.36,475.93L402.45,475.15L401.66,474.61L400.78,474.17L398.89,473.55L397.14,473.27L396.46,472.42L396.40,471.44L397.14,469.71L398.03,467.98L398.74,466.28L399.05,465.36L399.41,463.53L399.49,462.55L399.45,461.57L399.33,460.66L399.09,459.82L398.70,458.98L397.88,458.24L396.97,458.12L395.25,458.20L392.48,458.59L390.65,459.11L387.66,460.17L384.91,460.99L384.00,461.35L380.62,462.30L379.83,461.64L379.71,460.73L379.78,459.84L379.65,458.11L379.34,457.24L378.37,456.83L377.54,456.81L375.81,457.09L373.90,457.68L372.00,458.07L371.13,458.16L370.24,457.81L369.21,457.11L368.34,456.33L367.47,455.36L365.42,452.64L364.46,451.66L361.71,449.56L360.57,447.51L360.13,446.59L359.30,445.65L358.36,445.26L354.86,446.31L352.03,446.61L349.04,446.73L349.04,447.29L349.27,447.78L349.21,448.34L349.39,448.43L349.40,448.51L349.30,448.68L349.24,449.08L349.28,449.21L349.44,449.31L349.39,449.46L349.28,449.38L349.09,449.47L349.13,449.91L349.30,450.00L349.28,449.90L349.33,449.84L349.26,449.87L349.22,449.71L349.23,449.64L349.29,449.66L349.42,449.75L349.36,449.84L349.58,450.02L349.63,450.35L350.03,450.56L350.20,450.59L350.17,450.49L350.26,450.53L350.22,450.73L350.00,450.71L349.99,450.65L350.10,450.58L349.91,450.61L349.90,450.82L349.77,450.97L349.72,451.14L349.81,451.48L349.87,451.48L350.00,451.72L350.04,451.92L350.18,451.95L350.20,452.34L350.05,452.21L350.04,452.02L349.95,451.90L349.70,451.78L349.75,452.03L350.14,452.51L350.32,453.01L350.47,453.10L350.31,452.67L350.49,452.46L350.55,452.46L350.64,452.69L350.78,452.74L350.88,452.91L350.81,453.11L350.90,453.15L350.75,453.15L350.54,452.98L350.71,453.31L350.77,453.79L350.80,453.76L350.85,453.95L350.78,453.84L350.72,453.92L350.86,454.01L350.99,454.39L351.17,454.51L351.44,454.90L351.63,455.41L352.16,456.37L352.56,456.66L352.80,456.99L352.83,457.19L353.01,457.24L353.04,456.97L353.33,457.02L353.50,456.91L353.35,456.81L353.30,456.34L353.34,456.39L353.41,456.35L353.62,456.41L353.63,456.59L353.76,456.87L353.81,457.24L353.92,457.27L353.94,457.36L354.01,457.67L353.95,457.82L354.01,457.95L354.00,458.35L354.22,458.53L354.17,459.01L354.26,459.01L354.26,459.14L354.44,459.25L354.55,459.49L354.55,459.80L354.65,459.86L354.60,460.45L354.77,460.98L354.63,461.24L354.65,461.37L354.74,461.44L354.74,461.60L354.62,462.10L353.83,463.12L353.91,463.39L353.88,463.51L354.01,463.61L354.34,464.18L354.36,464.25L354.22,464.34L354.39,464.26L354.70,464.81L354.69,464.96L354.85,465.20L355.34,465.58L355.56,466.03L355.61,466.04L355.62,465.94L355.37,465.62L355.41,465.51L355.35,465.47L355.39,465.31L355.33,465.10L355.36,464.88L355.52,464.78L355.52,464.84L355.70,464.90L355.71,464.98L355.74,464.94L355.82,465.13L355.75,465.20L355.79,465.53L355.95,465.60L356.10,465.56L356.31,465.61L356.45,465.52L356.82,465.53L356.92,465.59L356.87,465.67L356.92,465.91L357.00,465.90L357.07,465.98L357.57,466.10L358.06,466.03L358.23,466.15L358.53,466.19L358.94,466.52L358.88,466.66L358.93,466.67L358.98,467.24L359.09,467.37L359.12,467.63L359.19,467.65L359.20,467.84L359.26,467.85L359.22,467.92L359.28,467.92L359.52,468.33L359.63,468.88L359.77,468.90L359.98,469.09L360.15,469.49L360.36,469.51L361.04,470.53L361.02,470.67L360.96,470.52L360.89,470.70L360.79,470.75L360.61,471.00L360.78,471.24L360.67,471.14L360.52,471.11L360.52,471.21L360.45,471.26L360.58,471.29L360.43,471.31L360.45,471.44L360.36,471.21L360.10,471.09L359.90,471.16L359.68,471.02L359.54,471.01L359.40,470.92L359.40,470.78L359.28,471.04L359.05,471.02L358.79,470.60L358.89,470.56L358.90,470.62L358.96,470.46L359.09,470.46L359.26,470.58L359.26,470.48L359.18,470.42L359.21,470.37L359.08,470.13L358.84,469.97L358.83,469.90L358.63,469.73L358.65,469.69L358.40,469.32L358.24,468.83L357.94,468.61L357.72,467.99L357.37,467.85L357.19,467.52L357.03,467.41L356.79,467.36L356.73,467.26L356.61,467.23L356.62,467.28L356.48,467.34L356.41,467.34L356.40,467.26L356.30,467.36L356.22,467.34L356.32,467.55L356.48,467.68L356.47,467.84L356.64,467.94L356.89,468.27L357.53,469.22L357.71,469.90L358.02,470.29L358.55,471.23L358.73,471.42L358.92,471.81L359.26,471.98L359.50,472.48L359.64,472.51L359.85,472.69L360.30,473.72L360.48,473.90L360.56,474.13L360.79,474.39L360.84,474.55L360.94,474.61L361.04,474.87L361.29,475.15L361.58,475.70L361.71,475.70L361.80,475.60L361.71,475.57L361.59,475.40L361.52,475.27L361.56,475.14L361.50,475.10L361.85,475.20L361.87,475.12L362.09,475.10L362.10,475.21L362.03,475.19L361.93,475.31L362.11,475.27L362.15,475.34L362.20,475.24L362.24,475.27L362.37,475.12L362.56,475.03L362.56,474.90L362.59,475.29L362.22,475.57L362.46,475.89L362.19,476.09L362.02,475.83L361.95,475.84L361.95,475.97L361.82,476.01L361.73,475.88L361.63,476.07L361.64,476.22L361.77,476.54L361.87,476.48L361.99,476.69L361.88,476.81L362.08,476.74L361.93,476.87L362.05,476.83L362.10,476.92L362.29,476.98L362.81,477.69L363.09,478.43L363.12,478.53L363.08,478.55L363.18,478.53L363.09,478.58L363.19,478.72L363.14,478.69L363.15,478.74L363.23,478.74L363.18,478.80L363.23,478.88L363.18,478.89L363.25,478.91L363.20,478.92L363.19,479.04L363.27,479.05L363.19,479.06L363.19,479.19L363.27,479.14L363.20,479.21L363.28,479.14L363.20,479.22L363.28,479.15L363.23,479.21L363.24,479.33L363.15,479.35L363.25,479.34L363.16,479.44L363.18,479.50L363.31,479.49L363.36,479.87L363.60,480.11L363.93,480.71L364.04,480.79L363.98,480.88L364.07,480.77L364.14,480.79L364.21,481.11L364.28,481.18L364.22,481.01L364.33,481.07L364.46,481.44L364.78,481.82L364.89,482.05L364.70,481.99L364.78,482.06L364.93,482.07L365.05,482.38L365.18,482.52L365.16,482.64L365.20,482.64L365.28,483.04L365.44,483.33L365.50,483.59L365.45,483.92L365.28,483.90L365.27,484.00L365.52,484.00L365.52,484.16L365.23,484.16L365.21,484.34L365.30,484.81L366.07,484.81L366.07,485.05L365.45,485.09L365.28,485.23L365.23,485.22L365.29,485.24L365.21,485.30L365.15,485.55L365.34,485.55L365.23,486.21L365.56,487.07L365.60,487.59L365.56,487.80L365.35,488.13L365.34,488.92L365.10,489.55L365.06,490.00L365.17,490.29L365.28,490.38L365.33,490.34L365.23,490.25L365.21,490.11L365.40,489.87L365.49,489.86L365.24,490.12L365.33,490.29L365.56,490.32L365.57,490.28L365.72,490.35L365.40,490.35L365.32,490.39L365.29,490.46L365.45,490.73L365.61,490.66L365.76,490.78L365.85,490.77L365.70,490.66L365.72,490.54L365.80,490.46L365.95,490.52L365.87,490.65L365.92,491.15L365.81,491.51L365.87,491.84L366.02,491.83L366.05,491.56L366.22,491.43L366.25,490.93L366.41,490.90L366.46,491.04L366.58,491.05L366.68,490.98L366.80,491.08L366.93,491.06L366.90,490.24L366.71,490.07L366.57,490.09L366.47,489.90L366.33,489.89L366.58,489.82L366.55,489.76L366.62,489.78L366.65,489.72L366.96,489.83L367.04,489.61L367.14,489.53L367.18,489.60L367.57,489.70L367.71,489.88L367.49,490.43L367.52,491.00L367.40,491.43L367.27,491.51L367.10,491.38L366.80,491.43L366.72,491.60L366.47,491.65L366.36,491.99L366.50,492.00L366.68,492.18L366.44,492.54L366.26,492.56L366.22,492.51L366.22,492.42L366.38,492.22L366.31,492.09L366.09,492.55L365.99,492.40L365.95,492.49L365.78,492.41L365.82,492.09L365.78,492.16L365.72,492.12L365.60,492.32L365.55,492.69L365.61,492.70L365.59,492.65L365.82,492.76L365.83,492.90L365.93,492.55L366.20,492.65L366.28,492.73L366.22,492.79L366.32,492.82L366.20,493.12L366.18,493.05L366.13,493.05L366.14,493.18L366.00,493.41L366.14,493.48L366.11,493.53L366.06,493.50L365.98,493.72L365.98,493.99L366.12,494.07L366.13,494.03L366.18,494.16L366.11,494.19L366.18,494.22L366.01,494.33L366.06,494.35L365.83,494.40L365.82,494.47L365.61,494.56L365.62,494.67L365.54,494.72L365.63,494.71L365.61,494.80L365.76,494.79L365.86,494.91L365.86,495.00L365.91,495.00L365.93,495.14L365.81,495.17L365.69,495.01L365.62,495.04L365.63,495.16L365.74,495.15L365.60,495.20L365.58,495.27L365.52,495.25L365.58,495.27L365.47,495.61L365.47,495.95L365.39,495.96L365.31,496.37L365.14,496.50L365.05,496.75L364.90,496.81L364.91,496.50L364.81,496.89L364.47,497.31L364.29,497.76L364.12,497.85L364.26,497.81L364.27,497.90L364.15,498.02L364.00,497.99L364.00,498.07L363.88,498.11L363.83,498.31L363.75,498.37L363.77,498.44L363.71,498.47L363.73,498.97L363.80,499.14L363.80,499.20L363.68,499.22L363.80,499.22L363.86,499.33L363.63,499.44L363.57,499.59L363.62,499.83L363.75,499.92L363.73,500.22L363.62,500.32L363.75,500.61L363.76,500.92L363.34,501.75L363.23,501.75L363.22,501.95L363.02,502.37L362.93,502.39L362.96,502.86L362.81,503.06L363.01,503.39L363.17,503.35L363.20,503.42L363.17,503.33L363.22,503.17L363.51,502.91L363.40,502.88L363.39,502.64L363.51,502.34L363.63,501.63L363.75,501.42L363.86,500.95L364.21,500.51L363.99,500.54L364.01,500.46L363.96,500.41L364.09,500.24L364.07,500.21L363.97,500.30L364.04,500.13L364.17,500.10L364.30,500.24L364.31,500.44L364.28,500.55L364.20,500.59L364.29,500.67L364.33,500.81L364.53,500.84L364.54,500.92L364.37,501.48L364.48,501.84L364.37,501.90L364.29,502.24L364.09,502.46L364.12,502.61L364.25,502.67L364.00,502.95L363.92,503.44L363.96,503.68L364.12,503.91L364.23,504.24L364.18,504.33L364.21,504.42L364.14,504.48L364.05,504.42L363.96,504.58L363.63,504.50L363.51,504.60L363.33,504.60L363.41,504.53L363.37,504.40L363.58,504.32L363.42,504.26L363.29,504.12L363.12,504.10L362.84,504.27L362.70,504.46L362.74,504.62L362.62,504.77L362.62,504.85L362.68,504.90L362.66,505.14L362.52,505.60L362.53,505.83L362.65,505.87L362.68,506.11L362.76,506.16L362.71,506.19L362.78,506.30L362.84,506.67L362.79,506.74L362.78,507.04L362.59,507.10L362.54,507.03L362.47,507.09L362.58,507.40L362.57,507.56L362.43,507.67L362.27,507.69L362.11,507.57L361.61,507.43L361.37,507.81L361.29,507.77L361.35,507.82L361.34,507.90L361.21,507.85L361.17,507.75L361.26,507.68L361.16,507.67L361.16,507.71L361.07,507.69L361.12,507.80L360.99,507.86L361.03,507.91L360.97,507.93L360.76,507.75L360.73,507.67L360.77,507.61L360.82,507.73L360.90,507.68L360.91,507.77L360.96,507.77L360.87,507.62L360.95,507.67L360.94,507.38L361.20,507.17L361.19,507.06L361.26,506.96L361.20,506.79L361.21,506.52L361.33,506.18L361.27,506.09L361.05,506.09L361.05,505.92L361.23,505.80L361.41,505.88L361.56,505.81L361.67,505.54L361.77,505.43L361.20,505.64L361.16,505.49L361.24,505.40L361.14,505.36L360.96,505.54L360.68,505.62L360.71,506.02L360.76,506.08L360.80,505.99L360.87,506.02L360.80,506.06L360.81,506.30L360.43,506.24L360.20,505.92L360.19,505.78L360.12,505.89L360.00,505.72L359.76,505.84L359.72,506.06L359.65,506.10L359.80,506.07L359.69,506.35L359.63,506.31L359.61,506.14L359.55,506.28L359.40,506.29L359.52,506.31L359.50,506.45L359.47,506.53L359.38,506.55L359.47,506.54L359.53,506.70L359.57,507.16L359.77,507.64L359.69,507.87L359.91,507.86L360.05,508.13L360.07,508.38L359.92,508.67L360.54,508.50L360.57,508.56L360.51,508.63L360.61,508.66L360.68,508.62L360.66,508.48L360.74,508.46L360.96,509.15L360.87,509.20L360.81,509.08L360.74,509.09L360.66,509.16L360.75,509.19L360.75,509.24L360.05,509.47L360.07,509.69L360.23,509.62L360.18,509.54L360.25,509.52L360.25,509.63L360.37,509.65L360.42,509.55L360.50,509.55L360.45,509.58L360.49,509.61L360.38,509.65L360.45,509.96L360.04,510.06L360.03,510.13L360.15,510.10L360.17,510.16L360.02,510.21L360.18,510.20L360.46,510.34L360.59,510.35L360.88,510.53L360.84,510.56L360.55,510.37L360.52,510.43L360.47,510.38L360.49,510.44L360.40,510.48L360.33,510.32L360.03,510.28L360.09,510.31L360.05,510.36L360.11,510.37L360.06,510.43L360.12,510.42L360.07,510.45L360.12,510.48L360.07,510.51L360.14,510.52L360.10,510.61L360.17,510.60L360.09,510.67L360.23,510.61L360.30,510.80L360.11,510.87L360.13,510.97L360.38,511.06L360.73,511.73L360.88,511.67L360.74,511.77L360.89,511.73L360.76,511.80L360.82,511.88L360.90,511.85L360.84,511.91L360.93,511.92L360.95,511.99L360.90,512.00L360.95,512.00L361.03,512.18L362.17,513.70L362.07,513.87L362.21,513.79L362.11,513.94L362.27,513.89L362.15,514.03L362.31,514.01L362.18,514.10L362.36,514.03L362.46,514.38L362.36,514.47L362.46,514.57L362.43,514.43L362.49,514.43L362.53,514.59L362.49,514.62L362.57,514.62L362.58,514.56L362.51,514.43L362.62,514.55L362.57,514.68L362.67,514.61L362.76,514.46L362.66,514.49L362.70,514.42L362.63,514.43L362.69,514.38L362.63,514.33L362.70,514.36L362.66,514.30L362.72,514.33L362.71,514.25L362.78,514.31L362.78,514.25L362.83,514.33L362.87,514.25L362.82,514.17L362.88,514.23L362.86,514.14L362.91,514.19L362.90,514.12L362.97,514.15L362.93,514.09L363.01,514.13L362.97,514.06L363.03,514.10L363.53,513.72L363.52,513.82L363.27,514.00L363.32,514.05L363.64,513.86L363.62,513.93L363.34,514.11L363.37,514.17L363.29,514.14L363.25,514.20L363.17,514.04L363.08,514.07L363.06,514.21L363.10,514.16L363.12,514.22L363.17,514.17L363.23,514.21L363.15,514.27L363.17,514.32L363.08,514.30L363.09,514.37L363.05,514.32L363.06,514.40L363.01,514.35L363.03,514.42L362.98,514.38L362.99,514.46L362.93,514.43L362.95,514.48L362.92,514.44L362.92,514.51L362.86,514.48L362.89,514.53L362.83,514.51L362.84,514.56L362.79,514.55L362.63,514.75L363.28,515.42L363.36,515.27L363.51,515.31L363.49,515.38L363.58,515.29L363.48,515.40L363.55,515.39L363.66,515.63L363.75,515.67L363.76,515.76L363.80,515.75L363.93,516.17L363.89,516.19L363.96,516.22L363.90,516.23L364.04,516.25L364.13,516.38L364.17,516.36L364.18,516.50L364.22,516.49L364.18,516.61L364.03,516.63L364.16,516.67L364.15,516.72L363.97,516.69L364.16,516.73L364.15,516.83L364.09,516.84L364.14,516.88L364.05,516.89L364.14,516.97L364.06,517.00L364.11,516.99L364.21,517.26L364.18,517.37L364.29,517.37L364.24,517.48L364.35,517.47L364.31,517.59L364.38,517.55L364.36,517.64L364.45,517.61L364.48,517.69L364.43,517.75L364.51,517.70L364.55,517.75L364.50,517.83L364.57,517.81L364.61,517.88L364.54,517.95L364.61,517.90L364.55,517.97L364.65,517.99L364.60,518.03L364.68,517.98L364.69,518.04L364.63,518.08L364.70,518.04L364.65,518.12L364.74,518.12L364.68,518.16L364.74,518.19L364.70,518.21L364.73,518.25L364.79,518.20L364.76,518.28L364.82,518.24L364.76,518.30L364.90,518.35L364.85,518.42L364.93,518.38L364.86,518.44L364.95,518.40L364.88,518.47L364.96,518.42L364.96,518.54L365.05,518.52L365.01,518.61L365.09,518.57L365.17,518.70L365.14,518.78L365.22,518.85L365.18,518.89L365.27,519.05L365.23,519.09L365.30,519.06L365.36,519.18L365.30,519.23L365.37,519.18L365.42,519.25L365.35,519.30L365.42,519.28L365.41,519.33L365.47,519.29L365.49,519.36L365.73,519.38L365.79,519.29L365.82,518.99L365.92,518.84L366.17,518.76L366.23,518.68L366.28,518.80L366.28,518.64L366.30,518.69L366.48,518.72L366.50,518.58L366.43,518.46L366.51,518.43L366.45,518.42L366.45,518.21L366.51,518.22L366.50,518.11L366.62,518.07L366.62,517.78L366.69,517.76L366.69,517.65L366.78,517.68L366.79,517.60L366.94,517.59L366.94,517.68L366.98,517.62L367.00,517.68L367.00,517.59L367.17,517.55L367.17,517.34L367.25,517.19L367.33,517.17L367.26,517.10L367.28,516.96L367.33,516.99L367.41,516.90L367.46,516.92L367.44,516.87L367.52,516.76L367.96,516.61L368.24,516.37L368.41,516.13L368.29,516.36L368.42,516.42L368.86,516.38L368.42,516.44L368.27,516.38L367.84,516.77L367.60,516.86L367.29,517.42L367.29,517.70L367.09,517.99L366.72,518.31L366.55,518.84L366.23,518.99L366.19,518.94L366.13,518.99L366.08,518.95L366.09,519.02L366.01,518.98L366.03,519.02L365.91,519.06L365.89,519.33L365.81,519.45L365.58,519.48L365.51,519.55L365.52,519.45L365.47,519.52L365.48,519.60L365.57,519.59L365.49,519.61L365.57,519.62L365.49,519.67L365.59,519.65L365.50,519.71L365.52,519.77L365.57,519.72L365.53,519.83L365.64,519.80L365.66,519.87L365.54,519.87L365.62,519.92L365.57,520.00L365.60,519.96L365.64,519.99L365.58,520.04L365.59,520.10L365.64,520.03L365.69,520.07L365.62,520.15L365.69,520.28L365.79,520.31L365.76,520.40L365.69,520.38L365.76,520.40L365.79,520.48L365.71,520.54L365.73,520.63L365.74,520.54L365.82,520.52L365.88,520.67L365.78,520.74L365.75,520.70L365.75,520.76L365.86,520.73L365.90,520.81L365.96,521.10L365.93,521.05L365.89,521.06L365.87,521.10L365.93,521.06L365.93,521.11L365.84,521.12L365.84,521.18L365.92,521.13L365.95,521.17L365.86,521.24L365.96,521.25L365.87,521.31L366.01,521.28L365.89,521.36L366.02,521.32L365.89,521.38L365.92,521.46L366.01,521.43L365.96,521.54L366.04,521.43L365.97,521.61L366.04,521.87L366.10,521.89L366.05,522.05L366.14,522.52L366.10,522.63L366.17,522.89L366.23,522.91L366.25,523.22L366.31,523.27L366.33,523.76L366.22,524.04L366.27,524.41L366.35,524.55L366.45,525.63L366.52,525.70L366.51,525.78L366.58,525.78L366.52,525.80L366.56,526.01L366.46,526.24L366.60,526.21L366.63,526.28L366.49,526.34L366.50,526.40L366.56,526.47L366.63,526.43L366.67,526.58L366.48,526.69L366.67,526.61L366.63,526.65L366.69,526.62L366.87,526.89L367.07,526.98L367.01,527.12L367.11,526.95L367.21,526.97L367.26,527.06L367.14,527.18L367.19,527.20L367.36,527.03L367.11,527.46L367.25,527.74L367.14,527.49L367.20,527.46L367.22,527.53L367.21,527.46L367.28,527.45L367.31,527.50L367.32,527.44L367.39,527.51L367.57,526.99L367.62,526.65L367.71,526.71L367.79,526.66L367.87,526.71L368.22,526.67L368.24,526.84L368.29,526.85L368.23,526.85L367.97,527.16L368.19,527.70L368.24,527.68L368.29,527.75L368.26,527.64L368.33,527.60L368.50,527.81L368.48,527.98L368.33,528.15L368.38,528.18L368.49,528.13L368.57,528.23L368.85,528.07L368.91,527.97L368.98,528.01L368.95,528.10L368.76,528.14L368.71,528.22L368.83,528.41L368.77,528.46L368.91,528.54L368.84,528.54L368.83,528.62L369.10,528.68L369.20,528.61L369.26,528.39L369.36,528.42L369.35,528.50L369.12,528.71L369.03,528.93L369.06,528.72L368.72,528.68L368.63,528.41L368.11,528.49L368.16,528.77L368.01,528.82L368.06,529.17L367.90,529.20L367.95,529.46L368.03,529.44L367.96,529.21L368.06,529.19L368.12,529.42L368.32,529.17L368.27,528.90L368.46,528.86L368.57,529.32L368.54,529.38L368.30,529.39L368.31,529.47L368.57,529.54L368.71,530.31L368.52,530.34L368.41,529.69L368.27,529.65L368.35,530.30L368.15,530.32L368.17,530.43L368.94,530.36L368.49,530.45L368.37,530.42L368.21,530.50L368.16,530.44L368.16,531.12L368.22,531.09L368.20,530.53L368.35,530.50L368.40,530.98L368.54,530.88L368.49,530.86L368.45,530.48L368.71,530.43L368.79,530.71L368.86,530.75L368.88,530.91L368.82,530.92L368.88,530.93L368.83,530.98L368.89,530.98L368.90,531.06L368.82,531.07L368.90,531.07L368.85,531.12L368.91,531.15L368.86,531.17L368.94,531.15L369.05,531.24L368.92,531.22L368.83,531.26L368.97,531.24L368.97,531.29L368.89,531.29L368.97,531.29L368.93,531.40L369.00,531.36L368.98,531.24L369.04,531.25L369.06,531.37L369.23,531.39L369.21,531.50L369.37,531.52L369.38,531.61L369.53,531.71L369.46,531.80L369.51,531.90L369.44,532.18L369.37,532.13L369.31,531.87L369.24,531.95L369.24,532.10L369.02,532.02L369.05,532.16L369.07,532.08L369.23,532.11L369.28,532.21L369.40,532.19L369.44,532.27L369.29,532.28L369.27,532.39L369.19,532.44L369.30,532.65L369.11,533.02L368.96,532.95L369.03,532.89L369.00,532.86L368.95,532.95L368.67,532.82L369.10,533.03L368.99,533.21L369.01,533.35L368.89,533.44L368.86,533.63L368.73,533.77L368.73,533.86L368.52,533.98L368.29,534.25L367.88,534.42L367.77,534.52L367.30,534.64L367.08,534.60L367.06,534.53L367.09,534.48L367.17,534.48L367.16,534.59L367.18,534.50L367.17,534.39L367.08,534.29L367.17,534.06L367.27,533.96L367.43,533.92L367.41,533.77L367.30,533.74L367.16,533.97L366.94,534.02L367.11,534.03L366.99,534.29L366.98,534.49L367.02,534.57L366.99,534.48L367.05,534.47L366.99,534.74L367.05,534.83L367.00,534.74L367.27,534.84L367.22,534.99L367.15,535.02L367.08,534.94L367.13,535.01L366.87,535.11L366.67,534.88L366.71,534.77L366.58,534.59L366.59,534.50L366.48,534.47L366.28,534.74L366.28,535.10L366.33,535.11L366.34,534.98L366.61,534.92L366.71,534.95L366.86,535.12L366.48,535.19L366.24,535.42L365.97,535.53L366.26,535.43L366.22,535.92L366.15,536.08L366.20,536.55L366.52,536.85L366.89,537.68L366.85,537.74L366.92,537.84L367.06,537.81L367.07,537.88L367.14,537.90L367.20,538.25L367.48,538.60L367.58,539.03L367.73,539.22L367.75,539.33L367.87,539.34L368.20,539.52L368.33,539.66L368.45,539.94L368.61,540.05L368.68,540.29L368.62,540.32L368.60,540.59L368.94,541.17L368.92,541.31L368.81,541.33L368.84,541.42L368.76,541.45L368.68,541.79L368.70,541.84L368.78,541.71L368.93,541.67L368.79,541.96L368.85,541.96L368.79,542.22L368.69,542.26L368.67,542.21L368.75,542.61L368.70,542.65L368.84,542.67L368.94,543.18L369.06,543.28L369.10,543.22L369.07,543.52L369.15,543.74L369.54,544.21L369.51,544.29L369.42,544.24L369.59,544.75L369.55,544.85L369.67,544.88L369.79,545.09L369.87,545.06L370.03,545.17L370.03,545.12L370.22,545.06L370.51,545.23L370.75,545.64L370.80,545.90L370.86,545.90L370.82,545.75L370.93,545.74L371.09,545.77L371.12,546.03L371.47,545.96L371.54,546.16L371.46,546.18L371.43,546.30L371.57,546.25L371.50,546.39L371.55,546.44L371.54,546.37L371.61,546.36L371.68,546.60L371.62,546.64L371.71,546.61L371.89,546.80L371.92,547.19L372.20,547.45L372.18,547.68L372.33,547.72L372.55,547.65L372.54,547.94L372.43,547.97L372.44,548.01L372.52,548.05L372.57,548.30L372.71,548.32L372.75,548.43L372.94,548.58L372.91,549.14L372.95,549.36L372.78,549.59L372.73,549.96L373.02,550.40L373.27,550.95L373.39,551.00L374.17,551.99L374.15,552.21L374.46,552.33L374.73,552.56L374.91,552.55L375.02,552.40L375.14,552.35L375.37,552.39L375.40,552.47L375.52,552.53L375.56,552.30L375.72,552.32L375.91,552.58L375.94,552.72L375.98,552.70L376.05,552.81L376.28,552.88L376.31,553.10L376.64,553.22L376.75,553.42L376.64,553.42L376.60,553.54L376.69,553.65L376.76,553.65L376.76,553.80L376.68,553.86L376.53,553.84L376.24,553.62L376.26,553.85L376.18,553.99L376.09,554.00L376.03,553.90L376.08,554.02L376.16,554.07L376.22,554.31L376.34,554.39L376.30,554.49L376.33,554.67L376.87,554.86L377.39,555.25L377.73,555.66L378.10,555.76L378.58,556.35L378.64,556.59L378.83,556.61L379.11,556.44L379.01,556.35L379.07,556.33L378.92,556.21L378.90,556.11L379.33,556.17L379.27,556.36L379.34,556.45L379.80,556.37L379.97,556.51L380.00,556.70L379.86,556.67L379.91,556.79L379.81,556.75L379.68,556.86L379.78,556.88L379.77,556.97L379.84,557.03L380.21,557.17L380.30,557.31L379.89,557.53L379.74,557.51L379.64,557.36L379.35,557.37L379.37,557.24L379.12,556.99L378.89,556.63L378.68,556.72L378.79,557.01L378.80,557.30L378.97,557.64L379.00,557.90L379.06,557.98L379.19,557.99L379.35,558.31L379.61,559.18L379.96,559.82L380.15,559.99L380.87,560.07L381.16,560.50L381.28,560.44L381.31,560.24L381.04,560.05L380.67,559.68L380.36,559.22L380.09,559.05L380.03,558.96L380.04,558.91L380.11,558.91L380.10,558.85L379.98,558.81L379.94,558.71L379.90,558.74L379.94,558.63L379.63,558.27L379.60,558.19L379.73,558.08L379.99,557.98L380.09,557.99L380.09,557.93L380.12,558.00L380.25,558.04L380.25,558.25L380.30,558.26L380.28,558.37L380.37,558.56L381.00,558.60L381.15,558.66L381.16,558.77L381.26,558.91L381.11,558.71L380.86,558.74L380.93,558.84L381.05,558.87L381.12,558.78L381.27,558.98L381.43,558.92L381.43,559.01L381.51,559.04L381.52,558.93L381.60,558.98L381.51,559.39L381.42,559.26L381.41,559.37L381.66,559.47L381.83,559.67L381.79,559.89L381.53,559.99L381.60,560.27L381.75,560.23L381.75,560.29L381.72,560.37L381.34,560.43L381.14,560.56L381.63,561.61L381.82,561.77L381.94,562.03L382.19,562.20L382.25,562.33L382.16,562.35L382.26,562.37L382.35,562.57L382.39,562.50L382.39,562.67L382.45,562.64L382.59,562.79L382.56,562.82L382.68,562.94L382.59,563.08L382.70,562.98L382.85,563.21L382.91,563.34L382.76,563.47L382.91,563.36L382.96,563.43L383.01,564.05L382.95,564.07L382.96,564.12L383.00,564.10L383.33,564.49L383.26,564.57L383.45,564.57L383.52,564.68L383.43,564.71L383.49,564.75L383.58,564.70L383.62,564.77L383.55,564.82L383.56,564.87L383.66,564.83L383.72,564.93L383.62,564.94L383.74,564.97L383.65,565.04L383.70,565.05L383.65,565.12L383.73,565.01L383.95,564.87L383.75,565.02L383.69,565.12L383.79,565.01L384.07,564.98L384.26,565.50L384.56,565.83L384.60,566.16L384.85,566.49L385.00,566.54L385.04,566.39L385.14,566.30L385.34,566.28L385.38,566.37L385.61,566.42L385.83,566.27L385.87,566.37L385.76,566.68L385.83,566.95L385.88,567.04L386.00,567.03L386.18,567.14L386.33,567.37L386.37,567.57L386.26,567.65L386.14,567.55L385.99,567.37L386.04,567.34L386.26,567.47L386.23,567.39L385.77,567.09L385.70,566.98L385.72,566.88L385.59,566.71L385.19,566.71L385.07,566.58L384.98,566.60L384.82,566.80L384.82,566.99L384.95,567.31L385.12,567.44L385.33,567.72L385.59,568.27L386.06,568.67L386.11,568.67L386.18,568.48L385.96,567.99L386.04,567.99L386.11,568.18L386.16,568.12L386.45,568.09L386.57,568.14L386.53,568.26L386.37,568.32L386.26,568.64L386.09,568.70L386.09,568.78L386.13,568.84L386.45,568.90L386.57,569.07L386.60,569.24L386.85,569.44L386.86,569.55L387.15,569.53L387.34,569.41L387.47,569.47L387.51,569.37L387.66,569.45L387.59,569.66L387.45,569.69L387.53,569.88L387.47,570.22L387.65,570.82L387.91,570.83L388.00,570.73L388.01,570.90L388.06,570.91L387.95,570.97L387.98,571.22L388.13,571.29L388.32,571.20L388.44,571.23L388.44,571.29L388.33,571.28L388.32,571.73L388.14,571.89L388.15,571.94L388.29,572.07L388.85,572.36L388.85,572.50L388.76,572.60L388.80,572.74L388.98,572.83L389.27,572.78L389.34,572.89L389.54,572.99L389.47,573.18L389.60,573.19L389.67,573.30L389.85,573.38L389.94,573.62L390.15,573.80L390.32,574.07L390.63,574.22L390.69,574.33L390.69,574.55L391.10,574.69L391.53,575.20L392.13,575.58L392.47,575.94L393.21,576.35L393.27,576.47L394.08,576.73L394.33,576.89L394.63,577.13L394.76,577.32L395.20,577.59L395.73,578.17L395.99,578.33L396.14,578.58L396.64,578.73L397.20,579.04L397.56,579.35L397.98,579.56L399.07,579.92L399.19,579.99L399.19,580.04L399.26,579.92L399.43,579.83L399.45,579.75L399.32,579.84L399.39,579.71L399.25,579.75L399.16,579.63L399.02,579.58L399.03,579.40L398.93,579.28L398.96,579.20L398.53,579.17L398.37,579.09L398.36,578.96L398.27,578.88L397.90,578.86L397.75,578.78L397.79,578.66L397.94,578.62L397.88,578.41L397.36,578.24L397.55,578.25L397.48,578.14L397.72,578.14L397.85,578.24L397.83,578.31L398.00,578.34L398.02,578.42L398.20,578.38L398.33,578.27L398.82,578.29L399.10,578.55L399.31,578.59L399.65,578.80L399.64,578.89L399.83,578.97L399.87,579.11L400.03,579.05L400.04,579.15L400.14,579.12L400.28,579.20L400.38,579.35L400.48,579.36L400.50,579.47L400.64,579.60L400.67,579.78L400.84,579.87L401.08,579.81L401.33,579.93L401.46,579.77L401.60,579.82L401.70,579.77L401.66,579.54L401.52,579.41L401.56,579.19L401.64,579.12L402.33,579.25L402.34,579.20L402.60,579.19L402.56,579.22L403.02,579.26L403.35,578.98L403.59,578.94L403.67,578.86L404.68,579.12L404.96,579.35L405.37,579.40L405.47,579.49L405.61,579.50L405.81,579.70L405.74,579.86L405.80,580.00L405.86,580.02L406.11,580.67L406.21,580.74L406.16,580.81L406.21,580.99L406.17,581.19L406.44,581.44L406.53,581.46L406.64,581.40L406.88,581.51L407.26,582.13L407.42,582.19L407.62,582.45L407.85,582.93L408.08,583.04L408.25,583.04L408.34,582.92L408.42,583.03L408.37,582.88L408.41,582.84L408.91,582.72L409.27,582.89L409.69,582.84L409.86,582.98L410.22,582.86L410.47,582.89L410.54,583.12L410.52,583.72L410.40,584.01L410.25,584.08L410.27,584.27L410.38,584.31L410.40,584.42L410.58,584.47L410.56,584.40L410.62,584.38L410.81,584.42L410.94,584.28L411.01,584.33L410.99,584.44L411.11,584.25L410.99,584.45L411.02,584.42L411.08,584.51L411.01,584.58L411.03,584.64L411.25,584.93L411.39,584.99L411.47,585.15L411.50,585.41L411.60,585.56L411.77,585.62L411.81,585.87L411.98,586.09L412.26,586.26L412.31,586.56L412.44,586.73L412.68,586.75L412.85,586.85L413.14,586.82L413.27,586.88L413.43,586.78L413.75,586.97L413.94,587.20L414.12,587.30L414.11,587.39L414.44,587.42L414.57,587.52L414.76,587.53L414.81,587.60L415.14,587.57L415.48,587.70L415.59,587.68L415.60,587.72L415.67,587.68L415.81,587.74L415.87,587.99L416.32,588.16L416.41,588.27L416.60,588.29L416.65,588.26L416.40,588.09L416.52,588.00L416.67,588.16L416.74,588.39L416.94,588.53L417.19,588.56L417.29,588.68L417.51,588.60L417.65,588.73L417.65,588.86L417.48,588.83L417.45,588.98L417.69,589.20L417.91,589.31L417.98,589.58L418.07,589.69L418.19,589.75L418.42,589.72L418.28,589.78L418.73,590.01L419.22,590.42L419.74,590.73L419.86,590.75L420.42,591.21L420.62,591.29L420.76,591.46L420.89,591.48L421.06,591.63L421.47,591.73L421.48,591.86L421.56,591.83L421.56,591.76L421.69,591.79L421.83,591.96L422.05,592.03L422.30,592.24L422.65,592.81L422.72,593.50L422.63,594.01L422.83,593.99L422.63,594.02L422.60,594.23L422.43,594.60L422.22,594.90L422.45,595.07L422.72,595.13L422.97,595.26L423.23,595.22L423.67,595.50L423.79,595.74L423.83,596.05L423.88,596.00L424.04,596.13L424.14,596.43L424.25,597.11L424.03,597.55L424.11,597.54L424.10,597.67L424.21,597.83L424.20,598.00L424.54,598.19L424.63,598.32L424.81,598.25L425.06,598.29L425.29,598.47L425.49,598.51L425.60,598.61L425.67,598.70L425.72,599.35L425.82,599.43L425.92,599.37L425.89,599.43L425.94,599.45L425.95,599.77L425.90,599.85L425.96,600.04L426.18,600.30L426.44,600.41L426.53,600.64L426.61,600.65L426.63,600.73L426.84,600.82L426.99,600.81L427.04,600.87L427.01,601.04L427.29,600.94L427.54,601.26L427.71,601.35L427.75,601.49L427.82,601.28L428.01,601.14L427.96,601.03L427.95,601.12L427.92,601.05L427.74,601.16L427.66,601.15L427.57,600.98L427.32,600.81L427.44,600.82L426.82,600.46L426.73,600.34L426.67,599.96L426.52,599.85L426.49,599.70L426.55,599.54L426.65,599.49L426.87,599.48L426.91,599.41L427.02,599.43L427.37,599.83L427.70,599.89L428.69,599.67L428.72,599.73L429.11,599.52L429.67,599.51L429.88,599.44L429.87,599.39L430.07,599.27L430.20,599.55L430.33,599.63L430.36,599.74L430.60,599.85L430.67,600.08L430.81,600.09L430.82,600.23L431.07,600.43L431.18,600.61L431.21,600.84L431.33,600.98L431.52,601.01L431.67,601.19L432.09,601.93L432.14,601.97L432.25,601.94L432.30,601.99L432.30,602.05L432.17,602.01L432.16,602.18L432.40,603.03L432.50,603.13L432.78,604.04L432.78,604.26L432.66,604.48L432.87,604.53L432.96,604.69L432.90,604.82L432.80,604.82L432.69,605.08L432.70,605.54L432.66,605.62L432.51,605.68L432.51,605.90L432.64,605.96L432.71,606.15L432.67,606.57L432.80,606.71L432.63,607.15L432.66,607.53L432.45,607.58L432.49,607.33L432.61,607.30L432.57,607.25L432.42,607.16L432.07,607.14L432.12,607.03L432.05,606.96L431.97,606.72L431.71,606.77L431.61,606.90L431.65,607.10L431.76,607.28L431.99,607.31L432.13,607.42L432.52,608.11L432.74,608.36L432.92,608.46L433.21,608.41L433.34,608.53L433.66,608.57L433.84,608.93L433.92,608.87L433.91,608.94L433.96,608.87L433.95,608.99L434.05,609.01L434.23,609.19L434.33,609.12L434.91,609.68L435.14,609.69L435.60,609.92L436.26,609.97L436.74,610.19L437.31,610.21L437.53,610.39L437.53,610.65L437.92,610.81L437.93,610.94L438.04,611.09L438.06,611.30L438.28,611.57L438.70,611.64L438.90,611.79L438.95,611.90L438.86,612.24L438.86,613.08L438.72,613.52L438.73,613.69L438.65,613.71L438.55,613.94L438.51,614.45L438.31,614.88L438.45,615.52L438.37,616.21L438.52,616.81L438.60,616.96L438.68,616.94L438.65,617.00L438.69,616.98L438.86,617.35L438.76,617.69L438.51,617.84L438.65,617.95L438.65,617.85L438.73,617.81L439.06,617.82L439.32,618.00L439.42,618.19L439.41,618.40L439.65,618.62L440.31,618.67L440.43,618.76L440.55,619.01L440.82,619.15L440.89,619.39L440.87,619.93L440.93,619.99L440.91,620.10L441.01,620.32L441.01,620.49L441.24,620.48L441.60,620.67L441.74,620.84L441.88,621.23L442.12,621.48L442.20,621.72L442.16,621.72L442.20,621.84L442.15,622.05L442.27,622.03L442.20,622.09L442.29,622.13L442.34,622.29L442.27,622.38L442.37,622.55L442.23,622.91L442.01,623.08L442.17,623.52L442.12,623.55L442.05,623.84L441.98,623.83L441.92,624.28L442.10,624.37L442.12,624.66L442.18,624.61L442.25,624.65L442.21,624.93L442.08,625.05L442.09,625.14L442.51,625.90L442.40,625.59L442.12,625.14L442.15,625.02L442.74,626.10L442.73,626.21L442.60,626.21L442.61,626.25L442.73,626.40L442.98,626.46L443.06,626.56L443.22,626.51L443.06,626.60L443.16,626.81L443.09,627.01L443.17,627.02L443.04,627.07L443.10,627.12L443.04,627.17L443.19,627.20L443.14,627.30L443.18,627.26L443.19,627.41L443.24,627.35L443.24,627.42L443.29,627.39L443.39,627.47L443.59,627.17L443.42,627.52L443.62,627.64L443.68,627.59L443.59,627.53L443.67,627.54L443.76,627.73L443.87,627.78L443.85,627.83L444.20,628.10L445.00,628.20L445.21,628.15L445.55,628.26L445.84,628.46L445.84,628.40L445.93,628.37L446.18,628.38L446.54,628.56L446.67,628.82L446.83,629.54L446.88,630.71L446.74,631.12L446.65,631.11L446.63,631.21L446.49,631.30L446.56,631.37L446.46,631.33L446.43,631.38L446.42,631.46L446.49,631.45L446.38,631.90L446.21,631.92L446.26,632.03L446.22,632.18L446.29,632.35L446.26,632.33L446.29,632.44L446.25,632.46L446.31,632.57L446.13,632.62L446.09,632.70L445.93,632.78L446.02,632.87L445.94,632.85L445.93,632.95L445.86,632.92L445.87,633.01L445.82,633.00L445.98,633.48L445.92,633.56L445.72,633.58L445.74,633.70L445.67,633.66L445.70,633.72L445.51,633.81L445.43,634.22L445.46,634.28L445.82,634.42L445.81,634.50L446.04,634.70L445.99,634.71L446.04,634.93L446.17,634.92L446.12,634.94L446.21,635.12L446.15,635.13L446.17,635.24L446.25,635.23L446.20,635.29L446.24,635.62L446.34,635.90L446.41,635.90L446.34,635.95L446.40,636.00L446.36,636.05L446.53,636.13L446.43,636.15L446.56,636.66L446.64,637.23L446.94,637.39L446.92,637.44L447.01,637.51L447.13,637.52L447.14,637.65L447.35,637.73L447.57,638.00L447.71,638.02L447.69,638.10L447.78,638.17L447.98,638.12L447.96,638.18L448.05,638.17L448.07,638.23L448.13,638.18L448.10,638.25L448.26,638.26L448.30,638.18L448.31,638.23L448.46,638.30L448.77,638.30L449.12,638.53L449.19,638.47L449.22,638.56L449.51,638.48L449.46,638.28L449.50,638.27L449.69,638.37L449.67,638.46L449.77,638.58L449.85,638.57L449.88,638.72L449.78,639.09L449.69,639.07L449.65,638.94L449.53,639.03L449.51,639.10L449.67,639.26L449.59,639.40L449.64,639.62L449.58,640.28L449.62,640.45L449.56,640.48L449.48,640.78L449.41,640.80L449.32,641.23L449.14,641.43L449.17,641.74L449.03,641.85L448.86,642.32L448.99,642.67L448.71,643.37L448.69,643.24L448.76,643.18L448.68,643.22L448.42,643.90L448.41,644.13L448.52,644.48L448.71,644.78L449.33,645.34L449.70,645.85L449.85,646.30L449.95,646.37L449.95,646.59L450.05,646.82L449.95,647.00L449.96,647.20L449.86,647.12L449.85,647.21L450.00,647.29L450.15,647.22L450.24,647.32L450.50,647.40L450.30,647.40L450.42,647.50L450.28,647.51L450.27,647.60L450.22,647.56L450.34,647.94L450.28,647.94L450.26,648.01L450.36,648.12L450.44,648.41L450.40,648.42L450.62,648.55L450.71,648.68L450.87,648.60L450.95,648.64L450.85,648.64L450.78,648.73L450.72,648.70L450.88,648.91L451.03,648.82L451.02,648.87L451.10,648.88L451.08,648.98L451.18,648.92L451.23,649.04L451.43,649.00L451.32,649.10L451.44,649.05L451.39,649.10L451.26,649.20L451.43,649.28L451.66,649.28L451.62,649.34L451.82,649.59L452.03,649.61L451.82,649.56L451.78,649.47L451.85,649.43L451.84,649.33L451.87,649.57L451.97,649.50L452.05,649.53L452.13,649.42L452.07,649.52L452.23,649.48L452.21,649.53L452.40,649.57L452.18,649.61L452.34,649.62L453.20,649.61L453.32,649.54L453.13,649.58L453.10,649.51L453.18,649.48L453.25,649.54L453.35,649.39L453.35,649.46L453.42,649.47L453.44,649.37L453.46,649.43L453.53,649.41L453.53,649.53L453.62,649.57L453.52,649.60L453.61,649.65L454.25,649.65L454.38,649.74L454.59,650.04L454.70,650.63L454.99,651.06L455.35,651.47L456.34,651.98L456.54,652.16L456.73,652.44L456.80,652.75L456.72,652.94L456.94,653.07L456.96,653.16L456.89,653.15L456.57,653.53L456.54,653.82L456.63,654.13L456.69,654.13L456.70,654.01L456.83,653.94L456.72,653.68L456.65,653.70L456.70,653.65L456.77,653.67L456.77,653.58L457.10,653.55L457.19,653.60L457.23,653.80L457.51,653.77L457.43,654.17L457.21,654.52L457.34,654.55L457.29,654.64L457.39,654.62L457.51,654.69L457.59,654.86L457.52,654.89L457.38,654.83L457.66,655.19L457.89,655.28L458.12,655.25L458.02,655.31L457.66,655.25L457.66,655.31L457.72,655.31L457.68,655.41L457.77,655.45L457.63,655.49L457.59,655.98L457.66,656.04L457.74,655.74L457.98,655.92L458.09,656.24L457.99,656.43L458.06,656.45L458.16,656.71L458.01,656.84L457.83,656.79L457.65,656.92L457.65,657.35L457.88,657.92L457.91,658.26L458.04,658.54L458.03,658.91L458.17,659.20L458.17,659.37L458.31,659.78L458.43,659.88L458.67,659.97L458.75,660.26L458.85,660.42L459.00,660.42L459.10,660.50L459.10,660.61L459.02,660.77L459.20,660.76L459.19,660.86L459.05,660.95L458.93,660.96L458.93,660.89L458.85,660.97L458.80,661.15L458.86,661.52L458.97,661.72L459.17,661.77L459.06,661.68L459.39,661.64L459.55,661.43L459.55,661.52L459.61,661.51L459.67,661.61L459.60,661.70L459.51,661.65L459.49,661.75L459.82,661.81L459.82,662.03L459.92,662.01L459.98,661.88L460.18,661.90L460.14,661.94L460.25,661.92L460.24,662.08L459.85,662.47L459.91,662.57L460.16,662.60L460.26,662.50L460.17,662.55L460.06,662.41L460.11,662.38L460.29,662.49L460.24,662.37L460.29,662.25L460.30,662.36L460.41,662.38L460.44,662.34L460.41,662.13L460.53,662.40L460.52,662.69L460.45,662.72L460.42,662.91L460.30,662.90L460.33,662.79L460.44,662.71L460.29,662.67L460.18,662.86L460.00,663.00L459.93,663.28L459.58,663.43L459.51,663.73L459.56,664.00L459.87,664.33L459.98,664.27L459.81,664.25L460.04,664.17L460.09,663.98L460.00,664.13L459.96,664.08L459.93,664.14L459.77,664.12L459.95,664.08L460.00,664.01L460.00,663.95L459.99,664.00L459.93,663.96L459.99,663.63L459.96,663.31L460.05,663.20L460.15,663.25L460.23,663.37L460.30,663.35L460.42,663.43L460.37,663.22L460.52,663.19L460.61,663.51L460.59,663.68L461.00,663.82L460.94,663.82L461.06,663.99L461.01,663.99L461.01,664.26L461.23,664.42L461.31,664.41L461.31,664.47L461.15,664.86L460.87,664.93L460.88,665.34L461.02,665.33L461.06,665.19L461.15,665.10L461.26,665.10L461.53,665.50L461.60,665.56L461.74,665.43L461.81,665.71L461.85,665.64L461.94,665.69L461.91,665.82L462.09,666.02L462.02,666.15L462.07,666.22L462.22,666.26L462.39,666.50L462.29,666.59L462.16,666.60L461.97,666.58L461.86,666.48L461.73,666.66L461.93,666.69L462.14,666.81L462.30,666.68L462.29,666.73L462.39,666.72L462.46,666.82L462.72,666.85L462.61,667.16L462.83,667.42L462.89,667.37L462.93,667.49L462.88,667.62L462.97,667.85L463.16,667.88L462.91,667.91L462.87,668.19L463.04,668.11L463.09,668.28L463.09,668.93L463.14,669.15L463.21,669.20L463.15,669.32L463.19,669.35L463.11,669.45L463.27,669.72L464.71,666.17L465.35,665.42L466.18,664.97L468.66,663.83L469.38,663.07L469.92,662.30L470.26,661.49L470.50,660.71L470.94,659.00L471.49,657.26L471.66,656.43L471.89,655.65L472.20,654.27L472.21,653.40L472.07,652.59L470.68,649.14L470.38,648.25L469.74,645.68L468.83,642.96L468.04,642.20L466.44,641.58L464.75,641.30L463.05,641.16L462.18,641.15L458.82,641.34L457.96,641.18L457.19,640.79L456.43,639.99L456.03,639.10L455.55,637.40L455.33,635.71L455.32,633.85L455.49,630.29L455.68,628.67L455.69,627.86L455.56,625.99L455.65,625.13L455.89,624.30L456.68,625.02L457.36,625.86L458.13,626.44L458.97,626.37L460.76,624.99L461.65,624.19L463.37,623.15L464.27,622.86L465.92,622.72L468.83,623.11L469.66,623.28L471.30,623.22L472.19,622.93L473.03,622.36L473.57,621.56L473.88,620.74L474.05,619.88L474.56,618.18L475.26,616.59L475.33,615.49L475.22,614.68L474.77,613.85L473.50,612.13L473.01,611.27L472.84,610.42L472.78,609.58L472.89,608.74L473.35,606.78L473.84,605.09L474.33,602.51L474.34,601.72L474.09,600.78L473.57,599.87L472.78,598.98L471.99,598.63L471.12,598.47L470.29,598.41L469.40,598.48L467.68,599.33L467.66,599.49L466.79,599.90L465.94,600.04L465.14,600.08L461.55,599.95L460.70,600.05L459.82,600.62L459.21,601.43L458.23,603.09L457.09,605.71L456.82,606.65L455.69,609.79L455.20,611.79L454.83,614.32L454.59,615.12L454.17,615.97L453.29,616.65L451.66,617.21L450.78,617.35L449.93,617.38L448.73,617.21L447.79,616.77L446.92,616.01L446.28,615.17L445.30,613.44L444.59,612.52L443.70,611.79L442.93,611.33L441.29,610.47L440.40,610.08L439.49,609.31L438.92,608.33L438.89,607.52L438.94,605.81L439.23,603.26L439.15,602.42L438.45,600.62L438.23,599.70L438.19,598.82L438.63,596.35L438.76,594.55L438.76,593.68L438.54,592.79L438.16,591.88L437.45,591.04L436.62,590.66L434.12,590.10L433.29,589.79L432.53,588.99L432.26,588.17L432.34,587.37L433.03,586.60L434.68,585.04L435.34,584.26L436.88,582.70L437.78,582.24L438.68,581.94L441.23,581.69L443.86,581.21L444.79,580.99L445.66,580.58L446.16,579.80L446.64,578.10L447.22,576.51L448.16,574.76L449.79,572.19L450.27,571.22L450.39,570.41L450.44,569.46L450.41,568.53L449.58,566.06L449.40,565.16L449.15,563.34L449.24,562.54L449.55,561.70L449.94,560.93L450.64,560.03L451.49,559.31L452.36,559.04L453.25,558.90L454.05,558.95L454.85,559.18L455.67,559.65L456.47,560.43L457.05,561.31L457.57,563.06L457.75,563.92L458.05,564.75L458.51,565.62L459.07,566.46L459.85,567.01L460.68,567.29L461.54,567.14L464.08,565.81L465.92,565.27L466.84,565.21L469.71,565.52L470.65,565.50L472.30,564.86L474.06,563.62L475.97,562.61L476.80,562.35L477.64,562.28L478.48,562.61L480.06,563.89L480.73,564.82L480.91,565.03L481.44,565.94L482.37,567.77L483.48,570.41L484.88,573.38L485.39,574.33L486.03,575.18L486.84,575.88L487.62,576.35L488.53,576.60L489.40,576.70L490.22,576.63L491.08,576.34L491.91,575.97L492.84,575.33L494.71,573.75L496.97,571.40L501.55,567.10L503.35,565.52L504.31,564.77L506.08,563.69L507.88,562.76L509.02,562.30L510.71,561.76L511.59,561.27L512.50,560.52L513.79,558.95L514.58,558.15L515.47,557.40L516.43,556.93L517.42,556.68L518.23,556.62L520.77,556.75L521.59,556.61L522.46,556.29L524.23,555.24L525.18,554.88L526.00,554.71L527.74,554.60L528.55,554.73L530.25,555.31L531.11,555.72L531.89,556.00L532.77,556.14L533.67,556.09L534.62,555.69L535.47,555.21L537.40,553.57L539.23,552.31L540.15,551.79L541.08,551.32L542.02,550.96L544.72,550.11L543.71,548.49L543.08,547.60L542.57,546.58L539.79,542.05L538.46,540.19L537.31,538.39L534.38,534.79L533.89,533.94L533.72,533.11L533.76,531.37L534.15,529.69L535.44,525.24L535.72,523.42L535.62,520.82L535.47,519.17L535.30,518.23L535.12,516.38L535.10,515.49L535.44,514.70L535.99,513.91L536.71,513.15L537.31,512.38L537.75,511.58L538.16,510.73L538.41,509.88L538.52,508.91L538.54,508.02L537.95,504.32L537.86,503.45L537.65,502.50L537.06,497.70L536.59,494.89L536.49,493.95L536.49,492.27L536.60,491.30L536.93,489.52L537.58,486.91L537.64,485.97L537.23,485.13L536.28,484.63L533.58,484.24L531.58,483.71L530.72,483.57L529.91,483.51L528.12,483.90L527.20,484.23L524.40,485.07L523.49,485.20L522.56,485.09L519.29,483.74L518.44,482.82L517.95,481.96L517.85,481.10L517.91,480.24L518.18,478.53L518.22,476.81L518.08,475.91L518.18,475.01L519.18,472.49L519.15,471.66L518.67,470.68L517.89,470.01L517.06,469.52L516.18,469.19L514.36,468.68L513.55,468.55L510.81,468.41L506.53,467.97L504.77,467.85L502.96,467.86L502.01,467.96L499.30,468.05L497.54,467.83L496.71,467.55L494.94,466.68L494.13,465.96L493.48,465.07L493.02,464.13L490.68,458.84L489.93,456.90L489.45,455.90L488.28,453.11L487.84,451.31L487.41,448.53L487.25,446.67L486.61,444.89L485.20,442.21L484.86,441.28L484.63,440.44L483.95,436.74L483.58,435.73L482.93,434.81L482.09,434.14L479.64,432.76L478.86,431.96L478.56,431.07L478.46,430.17L478.45,428.39L478.69,425.63L478.95,424.23Z" fill="#eee9df" stroke="#b4ad9f" stroke-width=".8" vector-effect="non-scaling-stroke" fill-rule="evenodd"><title>منطقة مكة المكرمة</title></path>
<path d="M210.51,257.02L210.55,257.20L210.97,257.59L211.25,258.03L211.36,258.01L211.25,257.91L211.36,257.94L211.35,257.85L211.10,257.30L210.82,257.26L210.51,257.02ZM212.79,249.21L212.90,249.82L213.23,250.26L213.45,250.34L213.52,250.48L213.79,250.66L214.43,250.62L214.37,250.56L214.09,250.57L213.96,250.50L213.73,250.00L213.58,249.94L213.44,249.67L213.39,249.52L213.46,249.24L213.15,249.14L213.04,248.94L212.80,249.05L212.79,249.21ZM192.00,251.22L192.19,251.37L192.68,251.36L192.64,251.25L192.80,251.18L192.84,251.09L192.73,250.68L192.73,250.46L192.89,250.40L193.16,250.44L193.20,250.38L193.31,250.41L193.33,250.57L193.75,250.96L193.78,251.24L193.55,251.37L193.73,251.62L193.85,251.59L193.98,251.69L194.03,252.07L194.16,252.18L194.15,252.09L194.37,251.97L194.53,251.97L194.58,252.03L194.56,251.96L194.68,251.85L194.99,251.70L195.17,251.71L195.45,251.36L195.28,251.04L195.27,250.94L195.32,250.89L195.28,250.77L195.16,250.72L194.99,250.50L194.84,250.53L194.84,250.22L194.77,250.34L194.63,250.32L194.71,250.12L194.55,249.92L194.11,249.73L193.54,249.70L193.38,249.77L193.29,249.71L193.26,249.83L192.87,250.08L192.28,250.23L192.34,250.29L192.22,250.48L192.62,250.62L192.63,250.70L192.36,251.05L192.00,251.22ZM185.84,247.85L186.01,248.30L186.06,249.31L186.27,249.74L186.16,250.31L186.65,250.64L187.11,250.65L187.32,251.18L187.71,251.40L188.13,251.78L188.74,251.78L189.42,251.61L189.75,251.76L189.99,251.68L190.11,251.55L190.76,251.46L190.84,251.36L190.87,251.11L190.72,250.86L190.40,250.74L190.11,250.48L189.56,250.24L189.54,250.12L189.43,250.02L189.45,249.88L189.36,249.78L189.45,249.70L189.46,249.48L189.58,249.50L189.55,249.26L189.32,249.37L188.92,249.29L188.78,249.20L188.41,249.27L188.34,249.42L188.00,249.55L187.55,249.51L186.77,249.29L186.54,248.96L186.19,248.73L186.12,248.63L186.11,248.43L186.18,248.21L186.26,248.24L186.27,248.33L186.43,248.38L186.47,248.50L186.63,248.62L187.59,248.15L187.92,248.30L188.08,248.19L188.09,247.86L187.76,247.60L187.74,247.42L187.58,247.33L186.68,247.26L186.67,247.34L186.56,247.42L186.39,247.44L186.28,247.52L186.22,247.86L185.84,247.85ZM206.54,245.21L206.65,245.60L206.90,245.71L207.06,245.73L207.54,245.58L207.35,245.52L207.30,245.30L207.12,245.11L207.00,244.81L206.81,244.72L206.65,244.73L206.54,245.21ZM194.06,244.44L194.57,244.36L194.87,244.49L195.05,244.69L195.13,244.87L195.09,245.12L195.21,245.19L195.14,244.63L195.01,244.38L194.86,244.27L194.73,244.28L194.71,244.21L194.70,244.26L194.34,244.26L194.06,244.44ZM258.81,327.60L258.84,327.68L259.07,327.74L259.15,327.86L259.28,327.91L259.18,327.92L259.29,327.96L259.33,327.92L259.65,328.12L259.79,328.09L259.81,328.03L259.54,327.79L259.03,327.59L258.81,327.60ZM244.67,295.73L244.69,295.86L244.84,295.97L245.06,296.41L245.27,296.66L245.51,296.72L245.64,296.62L245.75,296.29L245.83,296.26L245.77,296.11L245.48,296.05L245.33,295.86L245.09,295.73L244.95,295.87L244.86,295.88L244.67,295.73ZM234.94,286.41L235.04,286.45L235.17,286.83L235.25,287.38L235.71,287.81L235.88,288.08L235.94,288.39L236.10,288.37L236.23,288.52L236.31,288.33L236.28,288.21L236.04,288.23L235.98,288.18L236.05,287.83L235.90,287.51L235.91,287.38L235.87,287.43L235.77,287.39L235.86,287.20L235.83,287.04L235.30,286.54L235.08,286.12L234.95,286.28L234.94,286.41ZM290.97,387.93L291.25,388.48L291.44,388.68L291.50,388.65L291.63,388.75L291.57,388.43L292.25,388.11L292.34,388.01L292.31,387.93L292.19,388.00L292.12,387.96L292.16,387.92L292.01,387.91L291.89,387.79L291.70,387.80L291.69,387.66L291.58,387.70L291.42,387.67L291.20,387.85L290.98,387.86L290.97,387.93ZM286.42,380.89L286.67,380.64L286.99,380.67L287.16,380.86L287.42,380.99L287.56,380.98L287.91,381.26L288.11,381.33L288.23,381.25L288.33,380.87L288.15,380.57L287.88,380.43L287.57,379.99L287.23,379.77L286.79,379.62L286.64,379.68L286.52,380.05L286.54,380.22L286.45,380.56L286.53,380.59L286.55,380.68L286.42,380.89ZM279.36,362.51L279.46,362.76L279.60,363.43L280.27,363.56L280.35,363.43L280.36,363.21L280.23,363.15L280.25,363.20L280.05,363.19L280.00,363.26L280.02,363.09L279.89,362.86L279.88,362.78L279.97,362.79L279.93,362.60L279.62,362.40L279.73,362.42L279.69,362.35L279.42,362.40L279.36,362.51ZM277.91,360.87L277.97,360.97L278.07,360.97L278.09,361.36L278.23,361.42L278.51,361.71L278.61,361.70L279.16,362.09L279.54,361.99L279.66,362.04L279.72,361.91L279.73,361.96L279.85,361.95L279.81,361.78L279.69,361.90L279.56,361.82L279.56,361.65L279.46,361.54L279.45,361.43L279.16,361.21L278.90,361.15L278.66,360.99L278.40,360.99L278.17,360.89L278.22,360.84L278.13,360.91L278.02,360.90L278.09,360.80L278.01,360.82L278.00,360.91L277.99,360.82L277.91,360.87ZM282.57,358.66L282.60,358.72L282.89,358.70L282.94,358.78L282.90,358.79L283.01,358.90L283.00,358.81L283.07,358.81L283.10,358.89L283.20,358.91L283.32,359.04L283.34,359.13L283.26,359.14L283.29,359.18L282.95,359.24L282.95,359.30L282.86,359.32L282.92,359.40L282.93,359.74L282.87,359.75L282.87,359.85L282.96,359.83L283.00,360.02L283.19,360.29L283.38,359.98L283.43,360.18L284.19,360.28L284.29,360.50L284.36,360.44L284.29,360.67L284.23,360.69L284.40,360.68L284.35,360.64L284.42,360.65L284.42,360.61L284.50,360.66L284.60,360.61L284.62,360.52L284.79,360.45L284.78,359.78L284.52,359.73L284.10,359.82L283.71,359.23L283.71,359.09L283.83,358.91L283.68,358.89L283.71,358.99L283.65,358.98L283.43,358.58L283.16,358.54L283.16,358.60L283.16,358.51L283.11,358.49L282.78,358.45L282.63,358.52L282.57,358.66ZM281.53,357.53L281.60,357.61L281.56,357.53L281.84,357.58L282.22,357.53L282.88,357.77L282.95,357.68L283.09,357.68L283.23,357.77L283.33,357.73L283.20,357.57L283.05,357.60L282.98,357.54L282.90,357.33L282.83,357.38L282.85,357.51L282.75,357.47L282.71,357.32L282.66,357.33L282.63,357.02L282.07,356.86L282.02,356.78L281.67,357.05L281.53,357.53ZM281.84,355.59L282.05,356.15L282.13,356.15L282.31,356.31L282.52,356.31L282.41,356.21L282.37,356.03L282.29,355.99L282.22,355.81L282.02,355.59L281.84,355.59ZM277.11,355.44L277.63,355.94L277.72,355.83L277.93,355.80L278.00,355.47L277.93,355.44L278.12,355.50L278.07,355.63L278.50,355.55L278.59,355.42L278.52,355.29L278.44,355.37L278.43,355.47L278.39,355.46L278.42,355.24L278.50,355.25L278.45,355.17L278.48,355.13L278.52,355.20L278.67,355.02L278.78,354.99L278.56,354.97L278.40,354.78L278.31,354.77L278.21,354.46L278.04,354.31L277.92,354.32L277.87,354.39L277.62,354.37L277.36,354.71L277.18,355.19L277.22,355.37L277.11,355.44ZM280.25,353.21L280.52,353.73L280.71,353.84L280.84,354.05L280.91,353.97L280.85,353.82L281.11,353.34L280.88,353.21L280.87,353.26L280.74,353.24L280.64,352.89L280.43,352.95L280.36,353.13L280.33,353.02L280.25,353.07L280.25,353.21ZM263.61,351.17L263.77,351.31L263.82,351.57L264.09,351.78L264.17,352.12L264.60,352.51L265.06,352.50L265.34,352.40L265.57,352.41L265.88,352.23L265.99,352.06L266.06,352.13L266.14,352.02L266.04,352.05L266.07,351.97L265.96,351.81L265.99,351.58L265.93,351.74L265.53,351.88L265.41,351.66L265.11,351.55L265.05,351.20L265.12,351.05L264.99,350.95L264.88,350.94L264.80,351.08L264.85,351.29L264.91,351.32L264.81,351.28L264.79,351.35L264.85,351.35L264.78,351.46L264.72,351.53L264.62,351.52L264.59,351.46L264.70,351.36L264.70,351.28L264.51,351.26L264.38,351.09L264.73,351.07L264.63,350.96L264.63,350.88L264.71,350.96L264.78,350.94L264.79,350.61L264.67,350.39L264.22,350.38L263.78,350.76L263.61,351.17ZM266.48,347.42L266.55,347.54L266.50,347.61L266.53,347.71L266.80,347.92L266.98,348.25L267.22,348.42L267.33,348.69L267.58,348.82L267.88,348.68L267.69,348.53L267.53,348.56L267.23,348.32L267.20,348.17L267.07,348.12L267.10,348.07L267.00,347.97L266.97,347.83L267.09,347.67L267.01,347.70L266.77,347.51L266.90,347.51L266.81,347.41L266.91,347.45L267.02,347.51L267.03,347.64L267.10,347.65L267.10,347.54L267.19,347.52L267.23,347.22L267.54,347.04L267.46,346.84L267.52,346.65L267.35,346.59L267.15,346.66L266.98,346.89L266.98,347.00L266.82,347.09L266.73,347.36L266.48,347.42ZM280.12,355.50L280.16,355.55L280.24,355.50L280.58,355.53L280.69,355.59L280.80,355.71L280.86,356.43L280.89,356.47L281.29,356.43L281.08,356.04L280.93,355.42L280.74,355.22L280.61,354.85L280.37,354.49L280.34,354.54L280.34,354.44L280.23,354.54L280.22,354.62L280.27,354.58L280.32,354.66L280.23,354.93L280.30,354.96L280.30,355.12L280.21,355.21L280.27,355.31L280.12,355.50ZM265.75,345.63L265.92,346.32L265.92,347.21L266.15,347.48L266.31,347.57L266.36,347.28L266.23,346.97L266.30,346.78L266.14,346.23L266.08,345.36L265.90,345.29L265.75,345.63ZM393.07,223.69L392.64,222.39L392.17,221.85L391.49,221.26L390.21,220.52L388.08,220.05L386.92,219.36L385.79,219.13L384.95,218.70L382.91,217.95L382.47,217.91L380.97,217.29L377.69,216.29L376.82,215.71L376.13,215.37L375.40,214.91L374.36,214.44L374.02,214.16L373.62,213.94L371.46,212.04L371.02,211.89L370.30,211.21L369.10,210.55L368.27,209.91L366.85,208.54L365.83,207.67L364.70,206.60L364.30,206.09L363.66,205.57L363.46,205.18L362.95,204.73L362.04,203.70L360.81,202.70L358.91,201.46L358.63,201.18L358.12,201.02L357.10,200.18L354.11,198.82L352.24,197.44L351.52,196.79L349.15,195.25L348.10,194.38L346.88,193.52L346.67,193.30L345.70,192.82L344.10,192.11L343.34,192.01L342.08,192.01L341.29,192.32L339.56,192.68L338.76,192.96L337.21,193.69L336.23,194.24L334.22,195.04L333.09,195.12L332.30,195.39L331.73,195.09L331.48,194.83L331.20,194.37L330.98,193.29L331.27,192.44L331.63,191.85L331.73,191.54L332.23,190.58L332.81,188.98L333.51,187.73L335.53,184.97L336.95,183.40L337.65,182.97L338.32,182.21L338.55,181.77L338.66,181.30L338.64,180.88L338.54,180.71L337.51,180.07L336.10,179.35L335.30,178.74L332.51,177.14L332.11,176.70L331.02,176.18L327.61,175.44L326.97,175.35L325.36,174.60L322.82,173.58L320.68,172.39L319.92,172.07L319.57,171.57L317.58,170.29L317.22,169.76L315.63,168.59L312.08,166.40L310.93,165.54L309.97,165.05L308.69,163.96L307.51,163.09L306.49,162.15L305.64,160.95L305.22,160.56L303.91,158.82L303.49,158.43L289.44,161.28L274.22,164.35L264.48,180.95L251.18,191.94L247.59,195.16L244.84,194.63L243.42,194.47L238.71,193.48L214.69,189.20L212.67,188.72L210.91,188.51L206.68,187.72L204.08,187.35L203.95,187.41L203.97,187.59L203.81,187.72L203.58,188.14L203.75,188.39L203.64,188.44L203.44,188.88L203.59,189.11L203.56,189.23L203.72,189.83L203.68,189.92L203.52,189.96L203.45,189.83L203.14,189.90L203.07,189.85L203.00,190.25L202.83,190.35L202.94,191.30L203.14,191.43L203.18,192.00L203.27,192.09L203.04,192.32L203.32,192.44L203.32,192.59L202.84,193.01L202.62,193.07L202.26,193.04L202.32,193.16L202.17,193.29L202.08,193.49L202.11,193.58L202.18,193.57L202.18,193.87L202.34,194.20L202.14,194.35L202.00,194.32L201.89,194.48L201.97,194.51L202.00,194.70L201.73,194.78L201.55,195.05L201.56,195.35L201.47,195.42L201.51,195.60L201.61,195.61L201.65,195.70L201.54,195.76L201.51,195.99L201.41,196.14L201.41,196.68L200.98,197.00L201.00,197.41L200.88,197.71L200.91,197.95L200.84,198.02L200.74,198.40L200.86,198.70L200.69,199.09L200.67,199.47L200.59,199.62L200.63,200.24L200.36,200.68L199.96,201.13L199.89,201.49L199.73,201.78L199.75,202.02L199.64,202.11L199.66,202.25L200.00,202.49L200.16,202.81L200.19,203.60L200.01,203.90L200.00,204.05L199.91,204.12L199.94,204.20L199.77,204.31L199.77,204.38L199.55,204.63L199.49,206.19L199.10,206.91L199.03,207.17L199.00,207.97L198.64,208.56L198.66,208.64L198.98,208.77L199.32,209.05L199.37,209.28L199.14,209.99L199.16,210.12L199.29,210.14L199.33,210.29L199.23,210.51L198.91,210.89L199.10,211.16L199.03,211.57L199.14,211.79L199.04,211.80L198.93,212.08L198.79,212.54L198.81,212.73L198.72,212.76L198.55,213.08L198.53,213.28L198.39,213.36L198.56,213.70L198.36,213.71L198.32,213.80L198.37,213.93L198.25,214.23L198.29,214.41L198.17,214.81L197.96,215.19L198.02,215.69L197.87,215.79L197.70,216.05L197.56,216.72L197.40,216.93L197.39,217.12L197.16,217.40L197.16,217.53L197.02,217.69L196.82,218.46L197.06,218.67L197.20,218.93L197.32,219.35L197.30,219.57L197.70,220.64L197.67,220.95L197.49,221.45L197.49,222.23L197.59,222.35L197.58,222.44L197.84,223.00L197.83,223.20L198.00,223.89L198.01,224.65L197.69,225.03L197.54,225.39L197.52,225.84L197.20,226.08L197.12,226.06L196.80,226.48L196.57,226.99L196.53,227.41L196.35,227.45L196.38,227.52L196.29,227.64L196.40,227.78L196.35,227.86L196.23,227.89L196.23,228.01L196.11,228.12L196.05,228.29L196.15,228.50L196.02,228.75L196.03,229.05L195.92,229.22L195.76,229.30L195.61,229.67L195.42,229.86L195.49,230.09L195.38,230.18L195.12,231.16L194.94,231.45L194.87,231.99L194.75,232.20L194.78,232.40L194.60,232.52L194.48,233.06L194.33,233.22L194.32,233.43L194.20,233.55L194.00,233.98L194.02,234.11L193.88,234.27L193.73,234.88L193.58,235.16L193.41,235.32L193.31,235.60L192.96,236.04L193.09,236.04L193.10,235.98L193.27,236.06L193.33,235.95L193.28,236.14L193.18,236.22L192.99,236.19L192.86,236.42L192.84,236.64L192.94,236.67L192.81,236.73L192.82,236.90L192.87,236.96L192.95,236.83L193.05,237.01L192.77,237.06L192.70,237.18L192.68,237.54L192.62,237.63L192.69,237.70L192.61,237.81L192.61,237.99L192.39,238.32L192.38,238.44L192.48,238.64L192.60,238.69L192.41,238.78L192.48,238.97L192.15,239.00L192.09,239.09L191.98,239.03L191.99,239.09L191.89,239.17L191.89,239.34L191.76,239.53L191.81,239.51L191.90,239.71L192.01,239.78L192.12,239.69L192.10,239.93L191.76,240.02L191.89,240.15L192.05,240.09L192.25,240.24L192.27,240.60L191.94,240.65L191.72,240.60L191.80,240.37L191.51,240.26L191.60,240.06L191.39,240.17L191.30,240.12L191.20,240.18L190.97,240.59L191.01,240.60L190.96,240.70L191.08,240.72L191.09,240.81L190.60,241.03L190.57,241.26L190.32,241.52L190.41,241.71L190.16,241.74L190.01,241.65L189.62,242.03L189.65,242.08L189.56,242.12L189.54,242.27L189.64,242.35L189.66,242.52L189.34,242.50L189.27,242.55L189.18,242.77L189.21,242.86L188.97,243.07L189.07,243.31L188.96,243.55L188.84,243.62L188.98,243.79L188.96,244.02L189.26,244.09L189.23,243.99L189.30,243.87L189.64,243.89L189.71,243.64L189.66,243.55L189.82,243.44L190.10,243.48L190.35,243.65L190.38,243.81L190.24,243.85L190.21,243.96L190.34,244.11L190.28,244.32L190.42,244.48L190.30,244.56L190.35,244.77L190.19,244.88L190.13,245.15L190.42,245.37L190.38,245.69L190.52,245.79L190.54,245.91L190.61,245.89L190.55,245.82L190.61,245.69L190.74,245.69L190.84,245.79L190.87,245.98L190.63,246.02L190.74,246.29L190.55,246.21L190.60,246.12L190.51,246.15L190.58,246.55L190.92,246.82L191.24,246.72L191.24,246.66L191.30,246.66L191.24,246.65L191.26,246.55L191.50,246.59L191.45,246.19L191.64,246.23L191.71,246.43L191.80,245.94L191.75,245.49L191.71,245.49L191.85,245.15L192.11,245.03L192.13,244.88L192.26,244.82L192.33,244.81L192.53,245.05L192.67,244.88L192.54,244.60L192.58,244.45L192.32,244.37L192.34,244.23L192.28,244.16L191.97,244.01L191.91,243.93L191.98,243.86L191.86,243.86L191.90,243.77L192.34,243.42L192.47,243.13L192.75,243.11L192.96,242.84L193.04,242.61L193.05,242.28L193.09,242.27L193.08,242.05L193.26,241.96L193.48,242.03L193.64,242.40L193.83,242.20L193.99,242.27L194.06,242.37L194.07,241.88L194.22,241.73L194.35,241.74L194.44,242.25L194.76,242.61L194.73,242.68L194.57,242.72L194.43,242.90L194.47,243.07L194.64,243.11L194.76,242.97L194.91,242.96L195.00,243.02L195.07,243.21L195.49,243.31L195.62,243.48L195.69,243.21L196.03,243.38L196.03,243.44L195.83,243.49L195.84,243.80L195.98,243.86L196.04,243.97L195.83,244.01L195.69,243.88L195.65,243.67L195.44,243.78L195.56,244.17L195.45,244.31L195.58,244.31L195.62,244.17L195.73,244.14L195.79,244.30L195.92,244.34L195.94,244.26L195.86,244.22L195.90,244.10L196.04,244.09L196.10,244.20L195.95,244.45L195.83,244.48L195.88,244.73L195.72,244.72L195.60,244.55L195.61,244.65L195.50,244.71L195.41,245.24L195.87,245.12L196.27,245.17L196.63,245.12L196.50,244.92L196.20,244.89L196.00,244.64L196.04,244.57L196.17,244.58L196.17,244.49L196.34,244.56L196.24,244.40L196.24,244.27L196.42,244.28L196.33,244.23L196.24,244.05L196.12,244.06L196.06,243.99L196.04,243.81L196.11,243.82L196.20,243.99L196.26,243.93L196.37,244.06L196.56,244.10L196.30,243.93L196.11,243.70L196.09,243.59L196.58,243.91L196.60,243.83L196.81,243.93L197.21,243.95L197.32,243.89L197.16,243.61L197.44,243.53L197.55,243.42L197.53,243.28L197.61,243.24L197.79,243.32L197.94,243.10L198.26,243.08L198.34,242.99L198.58,243.13L198.34,243.44L198.38,243.60L198.52,243.67L198.51,243.73L198.28,243.80L198.28,243.87L198.20,243.93L197.89,244.01L197.88,244.14L197.75,244.13L197.75,244.38L197.66,244.42L197.72,244.58L197.63,244.72L198.13,244.69L198.22,244.41L198.51,244.11L198.62,243.86L198.74,243.77L198.87,243.77L198.84,243.63L198.93,243.37L198.71,243.30L198.71,243.25L198.88,243.16L199.05,243.17L199.18,243.34L199.15,242.99L199.40,243.07L199.51,243.19L199.55,242.93L199.70,242.87L200.09,242.89L200.28,243.07L200.02,243.30L200.18,243.24L200.11,243.33L199.63,243.49L199.62,243.58L199.59,243.53L199.46,243.66L199.29,243.65L199.45,243.88L199.54,244.20L199.53,244.28L199.36,244.42L199.37,244.51L199.46,244.65L199.65,244.56L199.59,244.63L199.71,244.77L200.19,244.68L200.21,244.60L200.11,244.56L200.09,244.42L200.26,244.33L200.39,244.34L200.54,244.51L200.72,244.38L200.99,244.38L201.09,244.28L201.31,244.34L201.39,244.47L201.64,244.42L201.79,244.49L202.21,244.52L202.40,244.44L202.55,244.52L202.78,244.53L203.34,243.97L203.87,243.84L204.16,243.67L203.60,243.86L203.60,243.78L203.79,243.75L203.77,243.64L203.87,243.73L204.09,243.68L204.11,243.59L204.24,243.54L204.64,243.54L205.08,243.42L205.28,243.09L205.54,242.99L205.72,243.01L205.98,243.24L206.43,243.95L206.51,244.00L206.58,243.95L206.49,243.86L206.54,243.89L206.68,243.77L206.82,243.82L206.93,243.96L206.84,243.71L206.70,243.58L206.78,243.37L206.56,243.39L206.53,243.20L206.30,243.07L206.29,242.90L206.22,242.82L206.11,242.81L206.13,243.04L206.06,243.08L205.88,243.01L205.86,242.83L206.00,242.75L206.04,242.58L206.14,242.51L206.43,242.62L206.54,242.90L206.50,243.00L206.54,242.94L206.71,243.03L206.97,243.08L207.03,243.04L207.16,243.23L207.45,243.09L207.60,242.93L207.94,242.98L208.73,243.77L208.98,243.88L209.10,244.09L209.34,244.12L209.75,244.36L209.83,244.85L209.99,244.76L210.46,244.74L211.34,244.79L211.45,244.86L211.48,245.04L211.69,245.29L211.89,245.31L212.11,245.45L212.26,245.32L212.58,245.45L213.02,245.51L213.48,245.44L213.68,245.55L213.86,245.55L213.89,245.62L214.32,245.83L214.38,246.12L214.51,246.25L214.43,246.53L213.95,246.47L214.03,246.65L213.85,246.82L213.86,246.89L213.78,246.83L213.77,246.92L213.77,246.84L213.60,246.87L213.52,246.71L213.39,246.64L212.92,246.97L212.48,247.10L212.17,247.32L211.99,247.27L211.97,247.36L212.11,247.60L212.09,247.83L212.32,248.05L212.34,247.97L212.55,248.00L212.58,248.13L212.76,248.05L213.01,248.11L213.29,248.36L213.75,248.51L213.84,248.64L213.86,248.89L214.12,248.93L214.21,249.12L214.80,249.51L215.63,249.73L216.09,250.10L216.17,250.31L216.14,250.48L216.06,250.51L216.19,250.56L216.64,251.60L216.88,251.87L217.23,252.10L217.45,252.40L217.57,252.44L217.74,252.92L217.89,253.05L218.11,253.10L218.33,253.57L218.56,253.68L218.63,254.05L218.77,254.07L218.89,254.23L219.01,254.25L219.27,254.45L219.43,255.27L219.40,256.04L219.52,256.38L219.50,256.50L219.61,256.52L219.72,256.89L219.82,256.88L219.83,257.00L219.92,256.96L219.99,257.05L220.49,257.22L220.60,257.24L220.63,257.13L220.72,257.13L221.23,257.39L221.60,257.43L221.82,257.55L222.01,257.76L222.15,257.73L222.42,257.90L222.71,258.37L222.81,258.70L222.85,259.17L222.98,259.17L223.23,259.41L223.69,260.18L223.72,260.39L224.02,261.00L224.15,261.42L224.11,262.17L224.28,262.42L224.41,262.91L224.51,263.09L224.57,263.10L224.55,263.18L224.65,263.28L224.64,263.34L224.89,263.56L224.97,263.55L224.97,263.63L225.22,263.92L225.34,263.93L225.32,264.06L225.70,264.47L225.80,264.69L225.89,264.67L225.94,264.55L226.04,264.52L226.12,264.24L226.18,264.32L226.13,264.42L226.19,264.59L226.31,264.62L226.29,264.71L226.18,264.68L226.15,264.60L226.07,264.61L225.90,264.81L226.17,265.33L226.40,265.28L226.39,265.37L226.20,265.44L226.20,266.07L226.32,266.02L226.32,265.97L226.48,265.95L226.47,266.11L226.42,266.24L226.32,266.13L226.19,266.26L226.45,266.82L226.40,267.17L226.86,267.22L226.92,267.18L226.89,267.09L226.98,267.04L227.08,267.21L227.07,267.46L226.99,267.60L226.90,267.64L226.84,267.64L226.84,267.59L226.94,267.53L226.75,267.32L226.33,267.25L226.29,267.35L226.43,267.75L226.42,268.02L226.55,268.60L226.86,269.37L227.04,269.65L227.63,270.34L227.90,270.42L227.89,270.66L228.00,270.77L228.07,270.96L228.17,271.05L228.29,271.06L228.52,271.33L228.65,271.72L228.56,271.99L228.79,272.91L229.02,273.11L229.19,273.09L229.51,273.40L229.50,273.63L229.62,273.85L229.97,274.16L229.98,274.27L230.74,275.30L230.87,275.39L231.34,275.95L231.59,275.83L231.61,275.90L231.45,276.01L231.50,276.14L232.13,276.48L232.28,276.66L232.67,276.82L232.90,276.69L232.95,276.53L232.96,276.64L232.82,276.86L233.26,277.07L233.05,276.99L233.02,277.03L233.11,277.02L233.10,277.07L233.13,277.03L233.14,277.09L233.16,277.04L233.16,277.11L233.20,277.06L233.17,277.12L233.25,277.09L233.23,277.16L233.27,277.12L233.25,277.17L233.39,277.22L233.37,277.26L233.43,277.21L233.44,277.30L233.47,277.23L233.46,277.31L233.54,277.27L233.49,277.33L233.56,277.28L233.57,277.37L233.59,277.29L233.60,277.38L233.63,277.32L234.00,277.45L234.24,277.66L234.22,277.79L234.26,277.66L234.44,277.80L234.44,277.87L234.46,277.81L234.49,277.87L234.51,277.80L234.60,277.80L234.79,278.14L235.16,279.33L235.47,279.67L236.08,280.63L237.20,282.10L237.27,282.47L237.20,282.66L237.20,282.91L237.58,283.22L237.60,283.34L237.30,283.83L237.03,283.99L236.84,284.01L236.76,284.11L236.62,284.41L236.51,284.50L236.52,284.83L236.43,284.98L236.50,285.04L236.52,285.46L236.46,285.81L236.64,285.92L237.05,286.45L236.92,287.01L237.38,287.21L237.62,287.62L238.11,288.16L238.80,288.76L239.55,289.94L240.06,290.27L240.15,290.47L240.45,290.76L240.51,290.95L240.68,291.16L240.73,291.44L240.95,291.48L241.25,291.84L241.38,291.90L241.45,292.16L241.77,292.38L242.78,292.74L243.25,293.24L243.36,293.49L243.38,293.29L243.54,293.34L243.58,293.52L243.43,293.55L243.44,293.66L243.60,293.77L244.08,294.41L244.39,294.59L244.81,294.63L244.90,294.77L244.79,294.79L244.80,294.84L245.04,294.86L245.08,295.04L245.36,295.32L245.30,295.60L245.49,295.93L245.72,296.07L246.03,296.09L246.30,296.44L246.34,296.64L246.50,296.83L246.56,297.20L246.43,297.64L246.56,297.84L246.50,298.30L246.59,298.61L246.81,298.93L246.88,299.29L246.95,299.36L247.12,299.44L247.16,299.34L247.27,299.32L247.18,299.47L247.19,300.04L247.29,300.07L247.31,300.34L247.39,300.40L247.39,300.31L247.46,300.37L247.43,300.47L247.32,300.45L247.35,300.58L247.65,300.97L247.82,301.35L247.81,301.50L248.32,301.88L248.38,301.99L248.31,302.33L248.34,302.85L248.41,302.85L248.47,302.98L248.44,303.09L248.56,303.21L248.58,303.36L248.86,303.60L248.97,303.79L249.44,304.09L249.53,303.94L249.68,303.91L249.53,304.12L249.92,303.98L250.02,304.04L250.05,304.19L250.22,304.15L250.40,304.27L250.77,304.70L250.93,305.27L250.97,305.68L251.04,305.68L251.28,305.99L251.55,306.49L252.69,307.00L252.75,307.14L252.96,307.26L253.00,307.51L253.17,307.68L253.24,307.71L253.20,307.61L253.29,307.53L253.57,307.62L253.60,307.73L253.55,307.84L253.71,307.95L253.66,308.00L253.62,307.93L253.47,307.90L253.62,308.28L253.60,308.39L253.80,308.32L253.66,308.49L253.73,308.68L253.92,308.87L254.03,308.68L254.13,308.65L254.39,308.97L254.29,309.08L254.00,308.92L253.95,308.98L254.36,309.82L254.52,309.75L254.63,309.78L254.65,309.91L254.41,309.98L254.43,310.05L254.60,310.25L254.67,310.26L254.69,310.40L254.88,310.57L255.09,310.94L255.42,311.23L255.50,311.64L255.83,311.84L255.96,312.20L256.06,312.30L256.20,312.88L256.39,313.26L256.77,313.76L256.95,314.17L257.28,315.30L258.10,316.67L258.29,317.41L258.20,317.79L258.22,318.03L258.36,318.21L258.37,318.39L258.76,319.16L258.98,319.38L259.30,319.44L259.09,319.48L259.06,319.59L259.35,320.44L259.56,320.81L259.96,321.29L260.36,322.01L260.80,323.19L261.06,323.56L261.12,323.62L261.13,323.55L261.22,323.58L261.32,323.87L261.71,324.24L262.04,324.79L262.03,324.73L262.09,324.75L262.06,324.81L262.25,325.07L262.34,325.37L262.59,325.88L262.70,325.90L262.79,325.76L262.88,325.81L262.86,325.97L262.72,326.07L262.92,326.39L263.23,327.25L263.31,327.66L263.45,327.78L263.63,328.10L263.65,328.35L263.72,328.40L263.75,328.97L263.69,329.50L264.02,330.07L264.03,330.36L264.13,330.48L264.15,330.64L264.46,330.64L264.63,330.95L264.83,331.04L264.89,331.29L265.03,331.37L265.12,331.55L265.29,331.59L265.30,331.66L265.89,332.03L266.15,332.10L266.46,332.33L266.58,332.52L266.85,332.75L267.31,332.56L267.36,332.63L267.34,332.76L267.11,332.92L267.20,333.00L267.86,333.19L268.04,333.18L268.13,333.05L268.07,333.04L268.17,333.00L268.18,333.19L268.29,333.25L268.36,333.17L268.50,333.19L268.51,333.05L268.69,332.99L268.78,333.00L268.74,333.27L268.92,333.25L268.93,333.12L269.09,333.00L269.49,333.05L269.51,333.20L270.20,333.17L271.03,333.51L271.29,333.54L271.50,333.73L271.68,333.76L271.68,333.90L272.04,334.14L272.08,334.39L272.36,334.37L272.37,334.52L272.46,334.58L272.46,334.77L272.33,335.13L272.34,335.47L272.48,335.67L272.46,335.79L272.58,335.92L272.58,336.59L272.53,336.72L272.14,337.09L272.09,337.20L272.10,337.31L272.13,337.28L272.23,337.40L272.26,337.61L272.06,338.30L271.92,339.11L271.61,339.94L271.40,340.30L271.17,340.62L270.99,340.68L270.73,340.90L270.56,341.21L270.26,341.50L270.14,341.80L270.01,341.95L269.96,342.17L270.05,342.25L270.64,342.36L271.04,342.51L271.21,342.83L271.27,343.30L271.51,343.54L271.60,343.78L272.02,344.44L271.89,344.52L271.86,344.65L271.90,344.70L271.86,344.74L271.95,344.83L271.99,345.12L272.35,345.36L272.29,345.23L272.37,345.16L272.40,345.20L272.41,345.02L272.62,344.89L272.83,344.93L272.81,345.22L272.96,345.28L272.93,345.82L273.00,345.85L273.12,346.10L273.25,346.19L273.14,346.17L273.19,346.27L273.34,346.35L273.26,346.25L273.30,346.21L273.52,346.35L273.78,346.70L273.81,346.91L273.92,347.07L274.16,347.17L274.22,347.08L274.12,347.14L274.11,347.08L274.25,347.03L274.40,347.09L274.38,347.12L274.67,347.04L274.96,347.12L274.97,347.20L275.07,347.22L275.15,347.56L275.38,347.76L275.43,348.02L275.43,347.83L275.64,348.02L275.92,348.09L275.99,348.35L276.07,348.40L276.06,348.29L276.12,348.32L276.26,348.25L276.12,348.28L276.24,348.15L276.23,348.05L276.26,348.15L276.32,348.13L276.36,348.21L276.61,348.00L276.50,347.94L276.43,347.74L276.34,347.07L276.21,346.97L275.99,346.54L276.19,346.35L276.19,346.22L276.04,346.22L276.16,346.20L276.44,346.33L276.47,346.50L276.51,346.44L276.87,346.52L276.89,346.78L276.83,346.90L277.03,347.06L277.10,347.30L277.32,347.37L277.32,347.45L277.44,347.54L277.56,347.54L277.56,347.64L277.80,347.74L277.85,348.13L278.08,348.42L278.32,348.55L278.37,348.65L278.35,348.82L278.50,348.77L278.62,348.87L278.96,349.39L278.98,349.54L278.87,349.80L278.98,350.06L279.06,349.86L279.20,349.78L279.43,349.92L279.45,350.01L279.50,349.98L279.62,350.05L279.69,350.25L279.73,350.23L279.84,350.40L279.86,350.61L279.95,350.35L280.15,350.34L280.31,350.42L280.39,350.60L280.48,350.65L280.78,350.68L280.79,350.60L280.92,350.59L281.47,351.08L281.78,352.14L281.96,352.19L282.19,352.52L282.45,352.54L282.48,352.73L282.59,352.77L282.58,352.90L282.37,353.17L282.35,353.35L282.44,353.52L282.42,353.77L282.78,354.02L283.10,354.62L283.03,354.60L283.00,354.47L282.95,354.50L282.85,354.33L282.89,354.27L282.83,354.28L282.80,354.43L282.85,354.53L282.92,354.55L282.80,354.79L282.88,355.12L283.16,355.34L283.21,355.13L283.09,355.19L283.11,355.03L283.33,354.97L283.25,355.09L283.60,354.90L283.65,355.10L283.57,355.16L283.60,355.23L283.77,355.28L283.88,355.42L283.77,355.64L283.90,355.73L283.88,355.84L283.79,355.88L283.87,355.93L283.77,356.09L283.84,356.17L283.72,356.29L283.89,356.32L283.97,356.42L283.87,356.56L283.89,356.70L283.79,356.79L283.92,357.04L283.82,357.19L283.85,357.45L284.06,357.50L284.25,357.69L284.49,357.75L284.54,357.82L284.76,357.85L284.86,358.06L285.11,358.15L285.29,358.33L285.34,358.59L285.40,358.61L285.39,358.78L285.56,358.74L285.75,358.82L285.86,359.19L285.93,359.19L285.99,359.27L286.11,359.24L286.38,359.31L286.47,359.51L286.74,359.66L286.73,359.79L286.88,359.94L286.87,360.32L286.76,360.54L286.90,360.69L286.88,360.92L286.75,361.12L286.81,361.15L286.79,361.53L286.81,361.43L286.84,361.48L286.80,361.72L286.88,361.85L287.05,361.86L287.02,361.75L287.12,361.83L287.20,362.05L287.19,362.24L287.43,362.58L287.31,363.53L287.34,363.81L287.72,364.26L288.12,364.51L288.14,364.62L288.06,364.57L288.04,364.65L288.18,364.86L288.07,365.02L288.21,365.51L288.38,365.77L288.44,365.75L288.43,365.85L288.51,365.82L288.59,366.02L289.01,366.26L289.06,366.43L289.45,366.86L289.49,367.01L289.58,367.04L289.80,367.37L289.97,367.49L290.09,367.78L290.29,367.94L290.52,368.27L290.64,368.67L290.83,368.83L291.26,368.90L291.65,369.30L291.74,369.64L291.83,369.69L291.98,370.02L292.26,370.21L292.24,370.26L292.36,370.23L292.47,370.31L292.81,370.23L292.81,370.48L292.93,370.53L293.22,370.84L293.23,371.34L293.43,371.36L293.66,371.70L293.75,372.16L294.00,372.59L294.04,372.99L294.15,373.14L294.20,373.41L294.04,373.52L294.04,373.58L294.18,373.94L294.16,374.10L294.23,374.15L294.31,374.58L294.16,375.20L293.41,375.96L293.42,376.30L293.68,376.96L293.80,377.51L294.10,377.88L294.15,378.04L294.30,378.10L294.26,378.16L294.44,378.52L294.40,378.70L294.60,378.99L294.79,379.06L295.06,379.46L295.01,379.76L294.92,379.94L294.99,380.61L294.95,380.80L294.38,381.66L293.94,382.01L293.94,382.25L293.80,382.50L293.77,382.81L293.61,383.09L293.62,383.57L293.70,383.62L293.99,384.13L294.04,384.34L294.24,384.49L294.27,384.78L294.14,384.95L293.91,385.08L293.84,385.32L293.08,385.89L292.87,386.15L292.75,386.45L292.57,386.63L292.54,386.87L292.65,386.95L292.58,387.08L292.52,387.09L292.50,387.02L292.51,386.98L292.59,386.99L292.55,386.93L292.04,386.92L292.07,387.04L291.97,387.16L291.66,387.23L291.59,387.07L291.39,386.93L291.32,386.75L291.58,386.68L291.71,386.73L291.83,386.58L292.14,386.42L292.30,386.44L292.33,386.31L291.55,386.27L290.86,385.97L290.29,386.00L290.27,386.09L290.12,386.20L290.09,386.39L289.93,386.57L289.94,386.66L290.16,387.04L290.37,387.18L290.58,387.61L290.63,387.94L290.74,388.05L290.87,388.09L290.95,387.92L290.95,387.77L290.83,387.78L290.75,387.63L290.71,387.64L290.71,387.51L290.80,387.39L290.86,387.40L290.96,387.63L290.93,387.72L291.04,387.84L291.27,387.77L291.42,387.63L291.59,387.66L292.21,387.52L292.26,387.75L292.38,387.84L292.87,387.98L292.88,388.19L292.75,388.68L292.48,388.87L292.69,388.89L292.70,388.97L292.98,389.11L292.92,389.44L292.70,389.57L292.62,389.41L292.43,389.28L292.15,389.18L292.05,389.18L292.01,389.23L292.07,389.50L292.39,390.04L292.51,390.66L292.77,390.98L292.77,391.37L292.72,391.50L292.52,391.56L292.52,391.70L292.78,391.88L293.22,391.97L293.32,392.14L293.64,392.14L294.12,392.35L294.40,392.59L294.50,392.62L294.68,392.87L295.17,393.13L295.27,392.99L295.35,392.99L295.34,393.15L295.41,393.15L295.41,393.28L295.62,393.48L295.82,393.91L296.00,394.08L296.05,394.30L296.30,394.76L296.37,395.08L296.34,395.28L296.69,395.36L296.71,395.22L296.98,395.11L296.90,395.05L296.85,394.76L297.06,394.76L297.24,394.99L297.22,395.22L297.26,395.22L297.09,395.31L297.05,395.50L297.12,395.74L297.46,396.00L297.53,396.42L297.63,396.51L297.74,396.94L298.10,397.33L298.20,397.92L298.28,398.01L298.33,398.29L298.44,398.40L298.45,398.52L298.81,398.74L298.85,398.98L298.96,399.10L299.59,398.94L299.99,398.70L300.34,398.43L300.98,397.67L301.27,397.21L301.93,395.58L302.21,395.18L302.65,394.22L303.00,392.78L303.25,392.27L303.49,392.11L303.80,392.03L304.17,392.21L305.37,393.46L305.66,393.61L305.85,393.61L306.77,393.27L307.22,393.22L307.66,393.24L308.03,393.33L308.24,393.41L308.68,393.72L308.96,393.75L309.58,393.73L310.74,393.54L311.97,393.48L312.35,393.52L313.39,394.31L313.76,394.46L314.06,394.47L314.60,394.39L315.38,393.94L315.96,393.03L316.31,392.30L317.10,391.26L317.06,390.90L316.68,390.52L315.93,390.17L315.22,389.98L315.00,389.81L314.99,389.65L315.12,389.45L315.36,389.26L316.48,388.92L316.81,388.66L317.60,388.22L318.65,388.20L319.24,388.03L319.81,387.98L321.24,388.08L322.23,388.27L322.88,388.29L323.08,388.58L323.89,389.31L324.18,389.36L324.42,389.24L324.66,389.02L324.97,388.20L324.92,385.13L324.79,384.69L324.62,384.43L324.07,384.09L323.78,384.19L323.43,384.43L322.77,385.13L322.49,385.30L321.85,384.98L321.52,384.45L321.35,384.04L321.22,383.41L320.91,382.59L320.43,381.72L319.20,380.75L318.83,380.39L318.21,379.40L317.88,378.38L317.66,377.46L317.42,375.79L317.49,375.19L317.49,373.71L317.66,372.47L317.69,371.92L317.62,371.60L317.49,371.38L316.30,370.25L316.00,369.79L315.56,369.35L314.73,368.31L314.55,367.70L314.48,366.92L314.51,366.42L314.68,366.17L314.92,365.47L315.03,365.08L315.01,364.72L314.44,362.70L313.94,361.54L313.17,360.15L313.04,359.64L312.77,359.04L312.66,358.21L312.23,356.39L312.25,355.54L312.53,354.49L312.82,354.13L313.06,353.91L314.13,353.33L315.58,352.18L316.30,351.55L316.98,350.82L317.14,350.46L317.14,350.12L316.90,349.26L316.61,348.78L316.04,348.27L315.34,347.80L314.29,347.37L312.95,346.22L311.98,346.10L310.06,345.22L308.98,345.25L307.84,344.98L307.40,344.74L305.93,343.59L305.17,342.54L304.88,341.50L304.92,340.64L305.14,340.30L305.19,339.79L305.32,339.28L305.30,338.11L304.99,336.84L304.75,336.18L304.27,335.23L303.89,334.74L303.61,334.21L301.42,331.23L300.58,330.32L300.19,329.79L299.22,329.05L298.52,328.15L298.19,327.91L297.27,326.93L296.55,325.83L296.48,325.31L296.64,324.90L297.27,324.34L297.75,324.07L298.78,323.06L298.98,322.28L298.94,321.82L298.74,321.35L298.43,321.11L297.34,320.99L296.53,321.18L295.06,321.84L294.14,322.14L293.65,322.14L293.24,322.06L292.75,321.84L291.96,321.30L291.70,320.86L291.61,320.25L291.64,319.57L292.07,318.42L292.80,317.19L292.91,316.73L292.84,316.31L292.62,315.84L292.29,315.53L290.50,314.08L289.51,313.10L288.83,312.07L288.35,311.02L288.28,309.37L288.35,307.68L288.76,306.77L289.33,306.01L289.49,305.57L289.64,304.91L289.62,304.52L289.42,303.78L288.59,302.82L287.34,302.11L286.59,301.91L286.11,301.67L285.87,301.45L285.41,300.59L284.60,299.65L284.14,299.38L283.54,299.11L283.13,299.01L282.23,299.24L281.61,299.19L281.18,299.01L280.94,298.57L280.89,298.08L281.02,297.39L281.26,296.73L281.51,296.41L281.83,295.30L281.88,294.61L281.75,293.48L281.86,293.01L282.16,292.25L282.34,291.17L282.10,289.96L281.81,289.42L281.51,289.00L280.96,288.63L280.30,288.56L279.64,288.76L278.81,289.22L278.30,289.40L277.87,289.54L277.38,289.54L276.29,289.10L275.67,288.66L274.88,287.75L274.49,287.43L274.16,286.66L274.05,285.11L273.76,283.46L273.63,282.00L273.94,280.52L274.14,279.88L274.42,279.29L274.75,278.87L275.34,278.43L275.69,278.08L278.35,276.55L278.70,275.81L279.05,275.59L280.01,275.32L280.54,275.07L281.02,274.67L281.51,274.11L282.54,272.57L282.97,272.15L283.52,272.01L284.01,272.15L285.34,273.07L285.85,273.27L286.86,273.44L288.28,273.50L288.61,273.56L289.62,273.52L290.21,273.56L291.64,273.45L293.63,272.74L294.17,272.72L294.48,272.79L294.70,272.86L295.18,273.32L295.41,273.63L295.57,274.91L295.73,275.49L295.95,275.77L296.20,275.87L296.79,275.75L297.82,275.43L299.52,275.41L300.25,275.29L300.92,274.82L301.58,273.89L302.25,273.45L303.34,273.45L304.18,273.30L304.93,272.89L305.72,272.23L306.36,271.36L306.79,270.90L308.26,269.54L309.10,268.98L309.67,268.82L310.21,268.92L311.11,269.41L312.91,270.95L313.39,271.47L313.83,272.13L313.99,272.54L314.39,273.07L315.09,273.50L315.41,273.61L316.04,273.50L316.95,272.94L317.43,272.86L317.90,272.96L318.55,273.22L319.39,273.78L319.97,274.36L320.13,274.70L320.34,275.71L320.37,276.92L320.31,277.43L320.16,277.75L319.68,278.37L318.95,278.91L318.76,279.21L318.63,279.69L318.67,280.64L319.04,282.48L319.23,283.02L319.42,284.08L319.45,284.67L319.61,285.21L319.99,285.56L320.50,285.75L320.82,285.95L321.13,286.20L321.35,286.51L321.78,286.80L322.19,287.18L322.30,287.49L322.32,289.10L322.47,290.60L322.78,291.68L323.14,292.50L323.73,293.22L324.77,293.96L325.10,294.32L325.44,294.86L325.57,295.19L326.14,297.42L326.46,298.04L326.78,298.42L326.94,298.71L327.40,299.17L327.73,299.39L328.65,299.55L331.59,299.52L332.47,299.68L334.50,300.20L336.50,301.01L337.45,301.09L338.92,300.84L340.96,300.34L341.81,299.97L343.68,299.54L344.21,299.34L345.36,299.06L346.83,298.75L348.51,298.68L349.50,298.73L349.92,298.81L350.60,299.06L351.24,299.42L351.62,299.50L354.47,299.55L355.10,299.47L356.42,299.52L357.34,299.43L358.71,299.02L361.43,298.99L363.77,299.50L364.92,299.55L367.04,299.93L369.22,300.17L372.13,300.65L375.55,300.74L376.02,300.73L376.91,300.53L378.57,299.50L379.36,298.71L379.90,298.53L380.34,298.60L381.11,298.44L381.76,298.21L382.28,297.48L383.01,296.64L383.67,296.18L384.35,295.58L384.77,294.39L385.61,293.75L386.70,291.56L387.09,291.05L389.86,287.94L390.81,286.73L391.19,284.86L391.46,284.50L391.53,284.05L391.89,283.01L392.17,282.58L392.84,281.88L393.65,281.54L394.41,281.51L395.28,281.88L396.07,282.52L396.31,282.55L396.82,282.80L396.80,283.94L396.97,284.39L397.26,284.74L397.63,285.01L398.20,285.09L398.75,285.31L399.57,285.24L399.80,285.34L400.02,285.50L400.15,286.05L400.38,286.48L400.71,286.58L401.23,286.31L401.66,284.56L401.41,283.38L401.51,282.44L401.69,282.00L401.64,281.56L401.75,280.94L401.75,280.32L401.67,279.74L400.78,278.62L400.58,278.22L400.14,277.31L399.92,276.68L399.80,276.08L399.43,274.99L398.54,273.65L398.30,273.15L398.04,272.02L398.13,270.09L398.07,268.71L397.79,266.03L397.65,265.32L397.40,264.65L397.18,263.61L396.86,262.86L396.61,262.06L395.36,259.49L394.68,258.30L392.53,255.82L391.98,255.44L391.37,254.86L390.37,254.38L389.66,253.83L389.02,253.56L387.98,253.37L387.12,252.86L386.76,252.46L385.66,251.83L384.56,251.46L383.75,251.05L382.36,250.64L380.80,250.07L378.34,249.01L377.87,248.73L376.91,247.87L376.49,246.89L376.40,245.77L376.56,245.13L377.87,243.04L378.41,242.37L378.87,241.25L379.54,240.37L379.94,239.59L381.01,238.56L381.76,237.31L382.90,236.22L383.59,235.05L385.26,233.45L385.90,232.16L388.22,230.31L389.16,229.87L390.83,227.77L391.26,227.12L391.61,226.09L392.41,224.76L392.99,223.98L393.07,223.69Z" fill="#e9e5db" stroke="#b4ad9f" stroke-width=".8" vector-effect="non-scaling-stroke" fill-rule="evenodd"><title>منطقة تبوك</title></path>
<path d="M381.76,298.21L381.11,298.44L380.34,298.60L379.90,298.53L379.36,298.71L378.57,299.50L376.91,300.53L376.02,300.73L375.55,300.74L372.13,300.65L369.22,300.17L367.04,299.93L364.92,299.55L363.77,299.50L361.43,298.99L358.71,299.02L357.34,299.43L356.42,299.52L355.10,299.47L354.47,299.55L351.62,299.50L351.24,299.42L350.60,299.06L349.92,298.81L349.50,298.73L348.51,298.68L346.83,298.75L345.36,299.06L344.21,299.34L343.68,299.54L341.81,299.97L340.96,300.34L338.92,300.84L337.45,301.09L336.50,301.01L334.50,300.20L332.47,299.68L331.59,299.52L328.65,299.55L327.73,299.39L327.40,299.17L326.94,298.71L326.78,298.42L326.46,298.04L326.14,297.42L325.57,295.19L325.44,294.86L325.10,294.32L324.77,293.96L323.73,293.22L323.14,292.50L322.78,291.68L322.47,290.60L322.32,289.10L322.30,287.49L322.19,287.18L321.78,286.80L321.35,286.51L321.13,286.20L320.82,285.95L320.50,285.75L319.99,285.56L319.61,285.21L319.45,284.67L319.42,284.08L319.23,283.02L319.04,282.48L318.67,280.64L318.63,279.69L318.76,279.21L318.95,278.91L319.68,278.37L320.16,277.75L320.31,277.43L320.37,276.92L320.34,275.71L320.13,274.70L319.97,274.36L319.39,273.78L318.55,273.22L317.90,272.96L317.43,272.86L316.95,272.94L316.04,273.50L315.41,273.61L315.09,273.50L314.39,273.07L313.99,272.54L313.83,272.13L313.39,271.47L312.91,270.95L311.11,269.41L310.21,268.92L309.67,268.82L309.10,268.98L308.26,269.54L306.79,270.90L306.36,271.36L305.72,272.23L304.93,272.89L304.18,273.30L303.34,273.45L302.25,273.45L301.58,273.89L300.92,274.82L300.25,275.29L299.52,275.41L297.82,275.43L296.79,275.75L296.20,275.87L295.95,275.77L295.73,275.49L295.57,274.91L295.41,273.63L295.18,273.32L294.70,272.86L294.48,272.79L294.17,272.72L293.63,272.74L291.64,273.45L290.21,273.56L289.62,273.52L288.61,273.56L288.28,273.50L286.86,273.44L285.85,273.27L285.34,273.07L284.01,272.15L283.52,272.01L282.97,272.15L282.54,272.57L281.51,274.11L281.02,274.67L280.54,275.07L280.01,275.32L279.05,275.59L278.70,275.81L278.35,276.55L275.69,278.08L275.34,278.43L274.75,278.87L274.42,279.29L274.14,279.88L273.94,280.52L273.63,282.00L273.76,283.46L274.05,285.11L274.16,286.66L274.49,287.43L274.88,287.75L275.67,288.66L276.29,289.10L277.38,289.54L277.87,289.54L278.30,289.40L278.81,289.22L279.64,288.76L280.30,288.56L280.96,288.63L281.51,289.00L281.81,289.42L282.10,289.96L282.34,291.17L282.16,292.25L281.86,293.01L281.75,293.48L281.88,294.61L281.83,295.30L281.51,296.41L281.26,296.73L281.02,297.39L280.89,298.08L280.94,298.57L281.18,299.01L281.61,299.19L282.23,299.24L283.13,299.01L283.54,299.11L284.14,299.38L284.60,299.65L285.41,300.59L285.87,301.45L286.11,301.67L286.59,301.91L287.34,302.11L288.59,302.82L289.42,303.78L289.62,304.52L289.64,304.91L289.49,305.57L289.33,306.01L288.76,306.77L288.35,307.68L288.28,309.37L288.35,311.02L288.83,312.07L289.51,313.10L290.50,314.08L292.29,315.53L292.62,315.84L292.84,316.31L292.91,316.73L292.80,317.19L292.07,318.42L291.64,319.57L291.61,320.25L291.70,320.86L291.96,321.30L292.75,321.84L293.24,322.06L293.65,322.14L294.14,322.14L295.06,321.84L296.53,321.18L297.34,320.99L298.43,321.11L298.74,321.35L298.94,321.82L298.98,322.28L298.78,323.06L297.75,324.07L297.27,324.34L296.64,324.90L296.48,325.31L296.55,325.83L297.27,326.93L298.19,327.91L298.52,328.15L299.22,329.05L300.19,329.79L300.58,330.32L301.42,331.23L303.61,334.21L303.89,334.74L304.27,335.23L304.75,336.18L304.99,336.84L305.30,338.11L305.32,339.28L305.19,339.79L305.14,340.30L304.92,340.64L304.88,341.50L305.17,342.54L305.93,343.59L307.40,344.74L307.84,344.98L308.98,345.25L310.06,345.22L311.98,346.10L312.95,346.22L314.29,347.37L315.34,347.80L316.04,348.27L316.61,348.78L316.90,349.26L317.14,350.12L317.14,350.46L316.98,350.82L316.30,351.55L315.58,352.18L314.13,353.33L313.06,353.91L312.82,354.13L312.53,354.49L312.25,355.54L312.23,356.39L312.66,358.21L312.77,359.04L313.04,359.64L313.17,360.15L313.94,361.54L314.44,362.70L315.01,364.72L315.03,365.08L314.92,365.47L314.68,366.17L314.51,366.42L314.48,366.92L314.55,367.70L314.73,368.31L315.56,369.35L316.00,369.79L316.30,370.25L317.49,371.38L317.62,371.60L317.69,371.92L317.66,372.47L317.49,373.71L317.49,375.19L317.42,375.79L317.66,377.46L317.88,378.38L318.21,379.40L318.83,380.39L319.20,380.75L320.43,381.72L320.91,382.59L321.22,383.41L321.35,384.04L321.52,384.45L321.85,384.98L322.49,385.30L322.77,385.13L323.43,384.43L323.78,384.19L324.07,384.09L324.62,384.43L324.79,384.69L324.92,385.13L324.97,388.20L324.66,389.02L324.42,389.24L324.18,389.36L323.89,389.31L323.08,388.58L322.88,388.29L322.23,388.27L321.24,388.08L319.81,387.98L319.24,388.03L318.65,388.20L317.60,388.22L316.81,388.66L316.48,388.92L315.36,389.26L315.12,389.45L314.99,389.65L315.00,389.81L315.22,389.98L315.93,390.17L316.68,390.52L317.06,390.90L317.10,391.26L316.31,392.30L315.96,393.03L315.38,393.94L314.60,394.39L314.06,394.47L313.76,394.46L313.39,394.31L312.35,393.52L311.97,393.48L310.74,393.54L309.58,393.73L308.96,393.75L308.68,393.72L308.24,393.41L308.03,393.33L307.66,393.24L307.22,393.22L306.77,393.27L305.85,393.61L305.66,393.61L305.37,393.46L304.17,392.21L303.80,392.03L303.49,392.11L303.25,392.27L303.00,392.78L302.65,394.22L302.21,395.18L301.93,395.58L301.27,397.21L300.98,397.67L300.34,398.43L299.99,398.70L299.59,398.94L298.96,399.10L299.16,399.28L299.29,399.26L299.36,399.32L299.63,399.95L299.85,399.97L300.09,400.12L300.11,400.30L300.39,400.68L300.47,401.16L300.59,401.37L300.61,401.65L300.83,401.74L300.70,401.52L300.87,401.48L300.95,401.68L300.91,402.00L301.07,402.19L301.32,402.32L301.30,402.61L301.92,403.32L301.82,403.66L301.83,404.00L301.92,404.30L301.78,404.49L301.53,404.60L301.03,404.15L300.89,404.13L300.91,404.36L300.76,404.56L300.85,404.91L300.71,404.97L300.70,405.09L300.61,405.15L300.49,405.15L300.48,405.30L300.60,405.66L300.50,405.83L300.57,406.21L300.75,406.30L301.06,406.73L302.23,407.42L302.30,407.54L302.30,407.73L302.53,407.89L302.57,408.08L302.52,408.18L303.24,408.99L303.46,409.76L304.04,410.16L304.30,410.48L304.53,410.47L304.62,410.40L304.70,410.43L304.69,410.37L304.65,410.40L304.68,410.34L304.75,410.40L304.71,410.33L304.91,410.28L305.46,410.33L305.53,410.37L305.49,410.45L305.61,410.41L305.67,410.45L305.65,410.52L305.76,410.48L305.83,410.56L305.81,410.61L305.88,410.58L306.09,410.70L306.19,410.89L306.32,411.41L306.39,411.49L306.61,411.51L306.76,411.46L306.82,411.36L307.18,411.27L307.68,411.31L308.08,411.47L308.53,411.45L308.71,411.13L308.52,411.21L308.33,411.10L308.32,411.03L308.38,411.03L308.39,410.80L308.45,410.73L308.66,410.76L308.68,410.53L308.61,410.55L308.50,410.41L308.55,410.41L308.53,410.26L308.29,410.24L308.00,409.85L307.88,409.63L307.90,409.51L308.02,409.51L307.98,409.43L308.19,409.47L308.16,409.43L308.29,409.35L308.36,409.50L308.32,409.57L308.47,409.63L308.47,409.68L308.72,409.76L308.73,409.67L308.93,409.68L309.14,409.78L309.52,409.52L309.90,409.53L309.96,409.62L310.44,409.64L310.43,409.56L310.24,409.53L310.02,409.30L310.18,409.27L310.34,409.33L310.30,409.26L310.38,409.20L310.94,409.15L311.17,409.43L311.05,409.91L311.14,410.13L311.58,410.18L311.75,410.31L311.74,410.26L311.78,410.27L311.91,410.42L311.80,410.49L311.92,410.43L311.95,410.51L311.92,410.72L311.83,410.72L311.88,410.85L311.98,410.89L312.02,410.99L312.29,411.01L312.53,411.12L312.72,411.11L312.87,410.85L313.16,410.84L313.23,411.04L313.46,411.09L313.53,411.27L313.73,411.42L313.76,411.65L313.64,411.68L313.62,411.76L313.70,411.89L313.92,411.98L314.08,412.23L314.18,412.80L314.35,412.92L314.48,413.15L314.88,413.28L315.01,413.50L315.65,413.81L315.88,414.25L316.33,414.51L316.47,414.80L316.82,415.05L317.02,415.31L317.67,415.57L318.83,415.40L319.61,415.70L320.21,415.78L320.40,415.52L320.36,415.29L320.41,415.24L320.17,415.01L320.18,414.66L320.01,414.43L320.27,414.32L320.14,414.11L319.81,414.01L319.67,414.07L319.51,413.86L319.62,413.69L319.54,413.59L319.86,413.57L319.78,413.48L319.82,413.42L320.19,413.47L320.62,413.40L320.65,413.22L320.72,413.20L320.46,413.06L320.49,413.02L320.75,412.99L321.11,413.04L321.17,413.16L320.99,413.35L321.09,413.40L321.22,413.37L321.22,413.44L321.13,413.48L321.11,413.60L320.75,414.09L320.45,414.15L320.38,414.34L320.31,414.69L320.43,414.94L320.62,415.21L321.02,415.29L321.19,415.20L321.26,414.71L321.55,414.58L321.94,414.60L322.13,414.53L322.22,414.62L322.24,414.83L322.36,414.86L322.56,415.07L322.64,415.33L322.35,414.97L322.25,414.93L322.23,415.05L322.18,415.06L322.04,415.04L321.77,414.84L321.72,414.89L321.51,414.85L321.32,414.99L321.24,415.45L321.09,415.48L321.11,415.55L321.03,415.60L320.76,415.56L320.54,415.61L320.40,415.81L320.37,415.78L320.40,415.84L321.07,416.14L321.58,416.57L321.83,416.91L322.17,417.12L322.44,417.22L322.79,417.50L322.96,417.77L323.36,418.02L323.71,418.50L324.43,418.95L324.51,419.06L324.64,419.05L324.89,419.14L325.04,418.90L325.22,418.89L325.44,418.67L325.60,418.79L325.77,418.53L326.01,418.66L325.71,419.08L325.49,418.94L325.11,419.20L325.53,419.52L325.68,419.56L325.80,419.52L325.75,419.31L326.05,418.96L326.11,418.85L326.05,418.82L326.18,418.76L326.17,418.92L326.05,419.01L326.03,419.16L326.27,419.23L326.36,419.41L326.67,419.59L326.78,419.77L326.87,419.70L326.88,419.62L326.82,419.67L326.87,419.58L327.17,419.76L327.26,420.01L327.45,420.17L327.50,420.35L327.65,420.51L327.90,420.64L327.88,420.68L328.13,420.60L328.28,420.68L328.27,420.94L328.37,420.94L328.28,421.04L328.40,421.08L328.61,421.33L328.73,421.27L328.70,421.19L328.76,421.24L328.80,421.19L328.89,421.20L328.99,421.40L328.93,421.40L328.81,421.30L328.79,421.34L328.88,421.37L328.83,421.49L329.03,421.39L329.13,421.27L329.27,421.23L329.22,421.43L329.15,421.41L329.15,421.59L329.07,421.62L329.10,421.82L329.02,421.77L328.89,421.82L328.82,421.93L328.84,422.00L328.92,422.00L328.95,421.91L329.06,421.92L329.10,421.84L329.42,422.14L329.44,422.23L329.51,422.24L329.54,422.39L329.68,422.43L329.64,422.16L329.76,422.16L329.67,422.26L329.74,422.37L329.84,422.24L329.77,422.16L330.02,422.29L330.00,422.12L330.12,422.17L330.19,422.10L330.22,422.15L330.22,422.10L330.49,422.02L330.51,422.09L330.42,422.17L330.50,422.75L330.20,422.92L330.43,423.09L330.58,423.33L331.00,423.69L331.07,423.67L331.40,424.00L331.48,423.88L331.57,424.01L331.51,424.00L331.53,424.05L331.47,423.98L331.44,424.07L331.72,424.17L331.77,424.07L331.82,424.09L331.88,424.02L332.01,424.11L332.23,424.11L332.20,424.23L331.93,424.25L332.11,424.35L332.15,424.45L332.24,424.43L332.25,424.48L332.71,424.69L332.74,424.97L332.92,425.05L332.74,424.96L332.73,424.75L332.88,424.67L332.82,424.63L332.77,424.68L332.76,424.63L332.97,424.16L333.23,424.32L333.21,424.40L333.31,424.46L333.32,424.40L333.42,424.42L333.58,424.54L333.92,424.62L333.94,424.55L334.00,424.55L334.01,424.68L334.16,424.68L334.23,424.79L333.96,425.59L334.27,425.69L334.01,425.56L334.05,425.46L334.25,425.49L334.24,425.41L334.07,425.35L334.27,424.81L334.49,424.99L334.49,425.11L334.82,425.51L334.82,425.69L334.91,425.73L334.97,425.91L334.73,425.92L334.77,426.08L334.78,425.96L334.96,425.92L335.00,425.98L334.92,426.03L335.12,426.15L334.87,426.37L335.14,426.20L335.18,426.30L336.16,426.94L336.90,427.24L337.20,427.54L337.20,427.63L337.29,427.62L337.34,427.68L337.32,427.77L337.37,427.83L337.52,427.82L337.80,428.11L338.20,428.82L338.46,429.48L338.81,429.95L338.65,430.20L338.86,429.99L339.10,430.08L339.08,430.28L339.13,430.07L339.53,430.15L339.52,430.07L339.73,430.05L339.91,430.11L340.20,430.35L340.96,431.38L341.12,431.46L341.18,432.00L341.14,432.70L341.07,432.95L341.11,433.34L341.33,434.09L341.52,434.26L341.54,434.52L341.65,434.77L341.89,435.71L342.22,436.53L342.36,436.43L342.27,436.27L342.25,436.00L342.25,435.83L342.34,435.83L342.42,435.89L342.39,436.02L342.50,436.17L342.48,436.28L342.63,436.32L342.56,436.44L342.74,436.50L342.64,436.56L342.42,436.48L342.36,436.58L342.43,436.66L342.62,436.65L342.65,436.71L342.79,436.73L342.65,436.96L342.72,437.12L342.65,437.13L342.65,437.23L342.71,437.26L342.67,437.29L342.74,437.43L342.83,437.41L343.03,437.48L343.43,437.82L343.46,437.73L343.71,437.61L343.86,437.79L344.07,437.76L344.08,437.84L343.94,437.88L343.63,437.79L343.61,437.87L343.56,437.84L343.59,437.92L343.54,437.97L343.46,437.96L343.59,438.20L343.77,438.28L343.70,438.40L343.54,438.38L343.46,438.25L343.21,438.13L343.26,438.28L343.55,438.60L343.72,438.75L343.86,438.76L343.98,438.90L344.05,439.23L344.21,439.19L344.26,439.23L344.22,439.34L344.14,439.28L344.07,439.31L344.18,439.65L344.15,439.83L344.29,439.86L344.40,440.13L344.39,441.10L344.23,441.57L344.24,441.72L344.63,441.83L345.20,442.23L345.29,442.21L345.31,442.07L345.60,441.81L345.70,441.81L345.76,441.95L346.36,442.08L345.92,441.90L345.84,441.76L345.65,441.78L345.32,441.66L345.23,441.67L345.20,441.77L345.21,441.68L344.96,441.62L344.91,441.47L344.75,441.30L344.76,441.12L344.83,441.08L344.76,440.73L344.87,440.67L344.87,440.52L345.07,440.46L345.25,440.54L345.43,440.93L345.52,440.89L345.68,441.00L345.76,440.83L346.01,440.81L345.99,440.68L346.21,440.48L346.24,440.36L346.66,440.38L346.65,440.65L346.58,440.78L346.48,440.76L346.53,440.94L346.45,441.63L346.64,441.80L347.12,442.00L347.17,442.44L347.33,442.71L347.63,443.00L347.68,443.02L347.66,442.94L347.70,442.94L347.93,443.11L348.00,444.09L348.04,444.30L348.14,444.38L348.18,444.65L348.11,445.04L348.28,445.62L348.34,445.60L348.49,445.84L348.53,445.79L348.47,445.90L348.48,446.41L348.60,446.74L348.65,446.75L348.65,446.70L348.71,446.79L348.83,446.81L348.86,446.69L348.79,446.61L348.80,446.36L348.87,446.29L348.82,446.32L348.86,446.26L348.77,446.11L348.81,446.10L348.95,446.27L349.04,446.73L352.03,446.61L354.86,446.31L358.36,445.26L359.30,445.65L360.13,446.59L360.57,447.51L361.71,449.56L364.46,451.66L365.42,452.64L367.47,455.36L368.34,456.33L369.21,457.11L370.24,457.81L371.13,458.16L372.00,458.07L373.90,457.68L375.81,457.09L377.54,456.81L378.37,456.83L379.34,457.24L379.65,458.11L379.78,459.84L379.71,460.73L379.83,461.64L380.62,462.30L384.00,461.35L384.91,460.99L387.66,460.17L390.65,459.11L392.48,458.59L395.25,458.20L396.97,458.12L397.88,458.24L398.70,458.98L399.09,459.82L399.33,460.66L399.45,461.57L399.49,462.55L399.41,463.53L399.05,465.36L398.74,466.28L398.03,467.98L397.14,469.71L396.40,471.44L396.46,472.42L397.14,473.27L398.89,473.55L400.78,474.17L401.66,474.61L402.45,475.15L403.36,475.93L404.03,476.79L404.34,477.63L404.54,478.53L404.54,479.43L404.29,481.34L403.59,483.84L403.63,484.75L404.33,485.64L406.11,486.11L407.82,486.35L408.75,486.38L409.73,486.28L410.65,485.97L412.37,485.00L414.21,483.36L415.04,482.45L415.79,481.51L416.72,480.70L417.63,480.31L418.52,480.66L421.87,481.65L422.73,481.84L423.52,482.24L425.34,482.83L426.27,483.00L427.18,482.81L428.01,481.83L430.57,477.65L431.00,476.81L432.25,475.20L437.03,471.40L441.01,468.63L442.86,467.07L444.72,465.34L446.33,463.65L447.22,462.68L447.92,461.77L448.67,460.99L449.58,460.21L450.61,459.48L451.45,458.97L453.30,458.22L455.14,457.68L456.94,457.08L457.89,456.34L458.53,455.47L458.53,454.61L458.07,453.68L457.23,453.12L456.32,452.64L455.51,452.32L453.75,451.44L452.00,450.32L451.14,449.56L450.79,448.68L450.91,447.85L451.55,447.06L452.38,446.28L454.99,444.57L458.81,442.46L459.78,442.00L461.68,441.19L463.73,440.48L465.59,439.61L467.38,438.56L468.33,437.91L469.02,437.00L469.54,436.17L470.15,434.49L470.32,433.64L470.38,431.67L470.13,429.96L469.87,428.99L468.87,427.08L468.81,426.24L468.97,425.30L469.54,424.43L470.41,423.88L471.37,423.50L472.25,423.31L473.21,423.25L475.01,423.48L475.95,423.76L477.74,424.11L478.95,424.23L479.02,423.91L478.93,422.05L478.75,420.41L478.95,417.14L478.85,415.55L478.90,413.92L479.58,410.48L479.87,409.62L480.50,408.35L480.32,407.46L479.45,405.47L479.44,404.24L479.88,403.36L481.68,401.35L482.33,400.36L482.78,399.39L483.02,398.73L483.35,397.20L483.32,396.52L484.56,391.67L484.61,389.49L484.58,388.66L484.92,386.15L484.85,385.47L484.29,384.91L483.59,384.79L482.89,384.41L481.46,383.17L480.57,382.79L479.54,381.57L479.25,380.64L478.90,379.84L478.35,379.38L476.23,377.23L474.72,375.97L473.97,375.61L472.36,373.84L471.63,373.39L470.94,372.76L470.33,372.43L469.03,372.03L468.52,371.81L467.81,371.29L466.91,369.83L466.47,368.41L465.63,367.30L465.09,366.81L463.65,365.77L462.01,364.79L461.01,364.48L460.31,364.47L459.65,364.35L458.74,364.29L458.04,364.03L457.35,363.55L456.94,362.99L456.63,362.20L456.24,360.00L455.78,359.31L455.12,358.92L452.97,358.17L451.83,358.24L450.89,358.14L449.97,358.32L449.65,358.31L447.97,358.10L447.18,357.88L446.29,357.54L445.46,356.89L444.98,356.42L444.00,355.25L442.85,354.48L441.69,354.36L439.70,354.77L438.35,354.42L437.35,353.95L436.66,353.83L435.76,353.90L435.06,354.14L433.75,354.73L433.34,355.02L432.92,355.47L431.98,356.92L431.51,357.83L431.18,358.34L431.08,358.76L430.75,359.27L430.54,359.72L429.84,360.63L429.17,361.02L428.56,361.14L426.18,361.15L424.42,361.52L423.70,361.21L423.23,360.93L422.60,360.05L421.53,360.03L421.09,360.15L420.67,360.41L419.60,361.44L419.10,362.51L418.92,363.43L419.15,365.27L419.30,365.85L419.27,366.25L419.01,366.48L418.59,366.45L418.08,366.28L416.68,365.78L416.18,365.51L414.88,365.40L413.34,365.69L412.79,365.71L412.25,365.61L411.15,365.16L410.46,364.99L408.29,364.93L407.13,364.53L405.55,364.20L405.25,364.05L404.17,363.26L403.24,363.20L402.71,363.04L401.79,362.46L401.28,361.61L400.37,360.43L400.16,359.81L399.82,358.35L399.60,356.33L399.77,355.84L399.67,354.42L399.68,351.57L400.10,349.78L400.30,347.55L400.41,347.13L400.35,346.63L399.78,345.69L399.54,345.08L399.34,344.21L399.41,342.39L399.49,341.72L399.73,340.74L399.68,339.99L399.79,339.43L399.77,338.91L399.62,338.54L399.59,337.90L399.38,336.83L399.42,335.77L399.35,335.35L399.50,332.67L399.43,332.04L399.69,330.11L399.75,329.20L399.64,328.79L399.14,327.90L398.03,326.21L397.82,325.77L397.31,325.08L396.86,324.19L396.65,323.30L396.65,322.91L397.02,320.07L397.01,319.70L396.70,318.76L396.78,318.37L396.30,317.40L394.74,315.25L394.50,314.65L394.13,313.98L393.81,313.58L393.37,313.24L393.05,312.84L392.68,311.98L392.35,310.34L391.89,309.62L391.16,308.84L390.77,308.55L390.34,308.11L388.62,306.96L388.25,306.61L388.03,306.23L387.94,305.35L387.99,304.38L387.91,303.80L387.59,303.51L387.29,303.42L385.69,303.22L383.74,303.12L382.73,302.99L382.22,302.85L381.93,302.64L381.70,302.28L381.53,301.83L381.45,301.39L381.41,300.27L381.52,299.28L381.76,298.21Z" fill="#f3efe7" stroke="#b4ad9f" stroke-width=".8" vector-effect="non-scaling-stroke" fill-rule="evenodd"><title>منطقة المدينة المنورة</title></path>
<path d="M554.89,243.03L553.29,244.28L549.29,248.00L548.15,249.37L547.58,249.81L547.23,250.30L547.11,251.02L547.13,251.90L546.96,252.96L546.56,254.60L546.35,256.30L546.37,257.31L546.12,259.49L546.19,260.82L546.48,263.21L546.79,264.02L546.83,267.13L546.92,268.34L546.61,269.24L545.76,269.35L543.85,269.32L542.86,269.48L541.81,269.76L540.69,269.90L539.24,269.96L538.51,269.90L538.07,270.03L537.58,271.04L537.17,271.26L536.81,271.17L535.76,270.31L534.85,270.19L534.28,270.46L533.72,271.16L533.54,271.83L533.72,272.40L533.58,273.05L533.12,273.27L530.79,273.37L530.29,273.52L529.81,273.79L529.56,274.27L529.06,276.07L528.93,276.98L528.90,277.79L528.99,278.29L529.16,278.64L529.38,278.86L530.01,279.06L530.41,279.03L531.07,279.13L531.67,279.45L531.91,280.12L531.98,280.52L531.82,281.00L530.28,281.32L527.59,282.44L526.96,282.79L526.77,283.12L526.99,283.67L527.00,284.19L526.77,284.39L526.57,284.85L527.11,285.71L527.16,286.24L527.06,286.77L526.84,287.27L526.48,287.75L526.04,288.05L525.73,288.67L525.64,289.29L525.63,289.96L525.47,290.56L524.45,292.03L524.14,292.58L523.16,293.74L520.74,296.07L520.32,296.01L519.90,296.18L519.44,297.02L519.05,297.36L518.71,297.49L518.37,297.37L517.37,295.74L516.67,295.51L516.20,295.52L515.72,295.67L515.32,296.01L514.87,296.56L514.29,297.72L513.85,297.87L513.16,297.38L512.65,297.35L512.28,297.80L511.89,298.45L511.39,298.79L509.92,299.61L509.53,299.52L509.09,299.55L508.80,299.94L508.46,300.18L508.08,300.12L507.80,299.88L507.29,299.77L507.15,300.08L507.14,300.77L507.41,301.48L507.54,302.17L507.39,303.09L506.09,305.11L505.50,305.81L503.67,308.96L503.18,309.39L502.64,309.70L501.92,309.81L501.25,310.46L500.75,311.07L498.32,313.12L497.37,313.64L496.58,313.39L495.91,313.05L495.33,313.23L494.98,313.56L495.15,313.91L495.22,314.32L495.07,315.03L494.55,315.89L494.20,316.80L493.56,317.33L493.17,317.22L492.86,316.73L492.98,316.37L493.48,315.82L493.27,315.47L492.69,314.96L491.88,314.83L491.10,314.99L490.37,315.29L489.83,315.98L489.21,317.39L488.84,317.92L488.09,318.73L487.42,319.31L486.71,319.75L486.23,320.17L485.84,320.69L486.01,322.20L487.03,323.98L487.25,324.58L487.37,325.32L487.22,326.57L486.90,327.49L485.75,328.96L484.97,329.02L483.64,328.08L482.94,327.83L482.24,327.87L481.85,328.28L481.43,328.85L481.11,329.59L480.72,330.19L480.23,330.55L479.72,330.47L478.84,330.77L478.10,331.09L477.45,331.55L476.93,332.22L476.99,332.78L477.23,333.50L477.22,334.36L476.76,334.79L475.99,335.72L476.07,336.45L476.35,336.96L476.47,337.72L476.03,338.15L475.48,338.16L474.71,337.73L473.98,337.13L473.09,336.84L472.67,337.13L471.97,338.18L471.79,339.24L472.40,340.29L472.37,340.83L471.68,341.46L469.42,343.22L468.82,343.84L468.42,344.51L468.10,345.29L467.88,346.24L467.45,346.94L466.89,347.67L466.17,347.89L465.59,347.90L464.49,347.69L461.31,347.42L460.57,347.60L459.74,347.50L458.91,347.14L458.28,347.22L457.49,347.75L457.01,348.56L456.54,349.58L456.26,350.52L455.74,351.74L455.03,352.78L454.16,354.23L453.68,354.73L452.77,355.05L451.83,355.68L451.23,356.65L449.65,358.31L449.97,358.32L450.89,358.14L451.83,358.24L452.97,358.17L455.12,358.92L455.78,359.31L456.24,360.00L456.63,362.20L456.94,362.99L457.35,363.55L458.04,364.03L458.74,364.29L459.65,364.35L460.31,364.47L461.01,364.48L462.01,364.79L463.65,365.77L465.09,366.81L465.63,367.30L466.47,368.41L466.91,369.83L467.81,371.29L468.52,371.81L469.03,372.03L470.33,372.43L470.94,372.76L471.63,373.39L472.36,373.84L473.97,375.61L474.72,375.97L476.23,377.23L478.35,379.38L478.90,379.84L479.25,380.64L479.54,381.57L480.57,382.79L481.46,383.17L482.89,384.41L483.59,384.79L484.29,384.91L484.85,385.47L484.92,386.15L484.58,388.66L484.61,389.49L484.56,391.67L485.80,391.44L487.51,391.48L489.39,391.81L491.39,391.61L495.17,391.65L496.43,391.99L496.95,393.08L497.54,394.58L498.09,395.79L498.92,398.36L499.26,399.12L500.33,401.29L500.91,402.22L501.53,402.95L503.09,402.93L508.13,403.07L510.24,403.01L511.37,402.40L511.74,401.93L511.67,401.38L511.74,400.14L512.24,399.64L513.06,399.15L514.80,398.57L517.26,398.71L518.48,398.94L519.71,399.46L520.48,399.91L521.23,399.90L523.35,398.18L523.83,396.80L524.28,395.95L525.37,393.59L526.53,391.75L526.78,390.71L526.74,390.20L526.22,389.72L524.92,389.97L523.24,390.47L522.68,390.41L522.26,389.87L522.29,389.27L522.82,388.86L524.29,388.49L524.82,388.15L524.79,387.87L523.83,387.24L523.18,386.50L523.11,385.90L523.45,385.59L524.52,385.27L525.20,385.16L525.99,384.72L526.72,383.96L527.15,383.14L527.19,382.51L526.65,381.93L525.67,381.79L524.95,381.97L523.97,382.73L523.26,383.50L522.64,383.39L522.46,382.64L522.56,381.54L523.50,379.61L524.18,378.55L525.26,378.34L526.49,378.34L528.12,378.59L528.94,378.43L529.42,377.86L532.68,375.21L533.17,374.75L534.86,372.17L534.85,371.84L535.77,370.91L536.36,370.18L537.41,369.83L538.33,369.76L539.28,370.19L539.62,370.68L541.16,371.79L542.33,371.91L544.25,371.66L545.55,371.10L546.59,370.45L547.18,369.57L547.29,368.57L547.18,367.72L546.92,366.74L546.93,366.06L547.31,365.61L548.13,365.24L549.08,365.11L551.25,364.67L554.47,364.46L556.57,364.51L558.24,361.98L559.40,360.50L559.86,359.69L560.58,359.34L561.05,359.22L562.31,358.93L564.55,358.65L566.07,358.20L567.32,357.95L569.37,357.22L572.37,356.04L573.27,355.61L573.97,354.53L574.49,353.84L574.93,352.97L575.59,352.63L576.37,352.65L578.43,353.49L579.99,354.45L584.88,357.08L585.77,357.18L587.03,354.87L587.41,353.75L587.77,352.21L588.60,349.95L588.80,348.74L587.78,347.77L586.20,346.60L585.32,346.20L583.49,345.15L582.91,344.57L582.28,343.58L581.47,341.83L581.03,341.11L580.04,338.79L579.64,337.66L579.52,336.69L579.10,335.66L578.64,334.77L578.11,333.91L577.53,332.73L577.36,331.93L576.94,328.05L576.46,325.75L575.84,323.93L575.39,322.24L575.26,321.23L574.50,318.03L573.63,311.78L573.30,310.44L573.04,308.92L573.04,305.57L572.91,304.20L573.15,301.37L573.61,298.34L573.73,296.67L573.79,294.77L573.96,293.65L574.45,292.76L575.80,291.93L577.30,291.49L578.56,291.23L579.69,291.23L581.18,290.63L582.70,290.60L584.49,290.05L585.41,290.03L586.54,290.62L589.39,292.51L590.94,293.32L591.61,294.33L592.24,294.95L592.77,295.04L593.03,294.61L592.04,290.63L592.07,288.24L592.22,285.78L592.61,282.16L591.97,278.88L590.74,276.82L590.39,275.90L589.91,275.11L589.62,274.35L588.96,273.32L588.00,272.05L587.18,270.72L586.82,269.90L585.93,268.88L584.81,267.98L584.23,267.33L584.06,266.97L583.80,266.60L583.34,265.58L581.06,263.02L579.24,259.68L578.97,259.43L578.51,258.80L578.04,258.52L577.46,257.73L576.44,256.91L575.19,256.28L574.38,255.74L572.99,254.46L570.59,253.13L570.20,252.78L569.84,252.32L569.22,251.92L568.82,251.32L566.67,250.10L564.15,249.30L563.36,249.26L560.97,249.42L560.32,249.84L559.31,250.65L558.83,251.10L558.55,251.58L558.22,251.91L557.38,252.52L556.71,252.48L556.32,252.20L556.12,251.39L556.19,250.79L555.81,250.44L555.25,250.45L554.71,250.30L554.23,249.94L554.24,249.22L554.72,248.25L555.29,245.88L555.39,244.42L554.89,243.03Z" fill="#eee9df" stroke="#b4ad9f" stroke-width=".8" vector-effect="non-scaling-stroke" fill-rule="evenodd"><title>منطقة القصيم</title></path>
<path d="M480.73,564.82L480.06,563.89L478.48,562.61L477.64,562.28L476.80,562.35L475.97,562.61L474.06,563.62L472.30,564.86L470.65,565.50L469.71,565.52L466.84,565.21L465.92,565.27L464.08,565.81L461.54,567.14L460.68,567.29L459.85,567.01L459.07,566.46L458.51,565.62L458.05,564.75L457.75,563.92L457.57,563.06L457.05,561.31L456.47,560.43L455.67,559.65L454.85,559.18L454.05,558.95L453.25,558.90L452.36,559.04L451.49,559.31L450.64,560.03L449.94,560.93L449.55,561.70L449.24,562.54L449.15,563.34L449.40,565.16L449.58,566.06L450.41,568.53L450.44,569.46L450.39,570.41L450.27,571.22L449.79,572.19L448.16,574.76L447.22,576.51L446.64,578.10L446.16,579.80L445.66,580.58L444.79,580.99L443.86,581.21L441.23,581.69L438.68,581.94L437.78,582.24L436.88,582.70L435.34,584.26L434.68,585.04L433.03,586.60L432.34,587.37L432.26,588.17L432.53,588.99L433.29,589.79L434.12,590.10L436.62,590.66L437.45,591.04L438.16,591.88L438.54,592.79L438.76,593.68L438.76,594.55L438.63,596.35L438.19,598.82L438.23,599.70L438.45,600.62L439.15,602.42L439.23,603.26L438.94,605.81L438.89,607.52L438.92,608.33L439.49,609.31L440.40,610.08L441.29,610.47L442.93,611.33L443.70,611.79L444.59,612.52L445.30,613.44L446.28,615.17L446.92,616.01L447.79,616.77L448.73,617.21L449.93,617.38L450.78,617.35L451.66,617.21L453.29,616.65L454.17,615.97L454.59,615.12L454.83,614.32L455.20,611.79L455.69,609.79L456.82,606.65L457.09,605.71L458.23,603.09L459.21,601.43L459.82,600.62L460.70,600.05L461.55,599.95L465.14,600.08L465.94,600.04L466.79,599.90L467.66,599.49L467.68,599.33L468.05,598.67L468.32,597.88L468.72,596.99L468.99,596.16L469.82,594.57L470.21,593.63L472.18,590.22L472.53,589.40L472.46,588.47L472.17,586.26L472.26,585.35L472.54,584.56L473.22,583.30L474.16,582.55L475.11,581.99L476.02,581.34L476.94,580.79L479.83,578.77L480.69,577.94L481.00,577.11L481.10,576.20L480.88,575.29L480.16,574.40L479.35,573.79L478.56,573.04L478.30,572.14L478.35,571.26L478.55,570.47L479.86,567.20L480.31,566.22L480.73,564.82Z" fill="#e9e5db" stroke="#b4ad9f" stroke-width=".8" vector-effect="non-scaling-stroke" fill-rule="evenodd"><title>منطقة الباحة</title></path>
<path d="M475.93,710.84L476.15,711.25L476.37,711.45L476.97,711.72L476.98,711.64L476.73,711.32L476.39,711.10L476.20,710.87L475.98,710.80L475.93,710.84ZM469.21,712.31L469.27,712.56L469.36,712.63L470.30,712.64L470.57,712.52L470.56,712.39L470.46,712.27L470.05,712.12L469.94,712.14L469.69,711.98L469.45,712.03L469.34,712.19L469.26,712.19L469.21,712.31ZM462.54,719.60L462.60,720.59L462.66,720.67L462.77,720.66L462.93,720.92L463.12,721.02L463.28,721.23L463.29,721.39L463.14,721.47L463.20,721.65L464.24,722.31L464.84,722.40L465.21,722.29L465.16,722.20L464.85,722.09L464.83,722.01L464.42,721.73L464.18,721.32L464.20,721.26L463.98,721.02L463.91,720.79L462.93,719.47L462.81,718.96L462.70,718.83L462.60,718.80L462.62,719.19L462.54,719.60ZM463.21,719.07L463.33,719.67L463.50,719.83L463.90,720.55L464.11,720.75L464.09,720.85L464.35,721.15L464.48,721.17L464.58,721.26L464.61,721.56L464.94,721.67L464.75,721.43L464.61,721.39L464.66,721.36L464.61,721.22L464.02,720.40L463.74,719.83L463.38,719.36L463.29,719.08L463.21,719.07ZM468.15,718.31L468.28,718.58L468.31,718.40L468.60,718.30L468.62,718.11L468.67,718.14L468.65,718.07L468.76,717.93L468.87,717.97L468.79,718.02L468.75,718.12L468.79,718.14L468.94,718.01L469.00,718.02L468.99,718.10L469.14,718.09L469.12,717.98L469.06,717.85L468.86,717.83L468.46,717.98L468.19,718.16L468.15,718.31ZM471.18,715.25L471.19,715.38L471.35,715.40L471.35,715.88L471.24,715.97L471.35,716.04L471.48,716.35L471.50,716.54L471.65,716.64L471.94,716.58L472.02,716.50L472.10,716.21L472.51,715.97L472.10,715.79L472.06,715.54L471.96,715.42L471.32,715.28L471.31,715.23L471.18,715.25ZM473.49,714.17L473.64,714.18L473.83,714.33L474.02,714.27L474.41,714.31L475.09,714.63L475.58,715.04L475.61,715.12L475.57,715.13L475.50,715.01L475.47,715.03L475.50,715.16L475.81,715.27L476.12,715.50L476.22,716.04L476.39,716.25L476.80,716.36L477.11,716.65L477.22,716.61L477.20,716.42L477.29,716.28L477.19,716.04L477.22,715.79L477.01,715.77L477.03,715.49L476.91,715.51L476.86,715.60L476.74,715.59L476.41,715.39L476.16,715.38L475.93,715.09L475.65,715.07L475.60,714.86L475.21,714.67L475.18,714.62L475.30,714.60L475.37,714.52L475.43,714.58L475.42,714.54L475.32,714.48L475.18,714.49L475.04,714.33L475.03,714.44L474.91,714.44L474.66,714.28L474.29,714.16L474.24,714.09L474.03,714.16L473.90,714.08L473.87,713.97L473.60,713.91L473.54,713.93L473.55,714.09L473.49,714.17ZM488.73,737.75L489.01,737.94L489.31,737.90L489.73,738.07L489.88,738.37L489.79,738.64L489.69,738.68L489.71,738.74L489.74,738.70L489.90,738.73L490.07,738.82L489.97,738.56L490.02,738.29L490.14,738.07L489.57,737.90L489.24,737.72L488.85,737.69L488.73,737.75ZM475.55,737.96L475.63,738.20L475.67,738.18L475.64,738.43L476.06,738.38L476.14,738.28L476.01,738.22L476.00,738.16L475.88,738.15L475.81,737.96L475.64,737.89L475.55,737.96ZM480.87,733.06L480.97,733.47L481.24,734.07L481.60,734.11L481.82,734.03L481.92,733.90L482.07,733.40L482.20,733.26L482.17,733.14L482.36,732.80L482.26,732.77L482.22,732.47L482.03,732.53L481.96,732.62L481.90,733.00L481.67,733.49L481.58,733.48L481.64,733.18L481.77,732.92L481.80,732.62L481.75,732.50L481.62,732.44L481.22,732.45L481.00,732.68L480.87,733.06ZM487.71,724.85L487.78,724.96L487.74,725.13L487.79,725.15L487.88,725.01L487.89,724.95L487.81,724.93L487.85,724.87L488.10,724.72L487.97,724.33L487.83,724.36L487.71,724.85ZM496.80,722.99L496.99,723.42L497.32,723.72L497.34,723.59L497.24,723.55L497.07,723.28L497.13,722.98L496.91,722.88L496.80,722.99ZM468.81,723.09L468.94,723.26L469.17,723.34L469.40,723.40L469.83,723.37L469.93,723.28L469.96,722.96L470.11,722.77L469.74,722.74L469.10,722.55L469.32,722.73L469.35,722.82L469.33,722.87L469.07,722.84L469.03,722.68L468.95,722.69L468.81,723.09ZM479.49,723.27L479.70,723.52L479.73,723.75L479.81,723.79L479.88,724.07L480.10,724.13L480.16,723.97L480.27,723.90L480.26,723.82L480.14,723.88L480.07,723.84L480.06,723.63L480.32,723.54L480.28,723.41L479.90,723.36L479.85,723.31L479.85,723.05L479.69,722.85L479.67,722.93L479.75,723.03L479.66,723.28L479.49,723.27ZM475.81,729.93L476.20,730.32L476.52,730.47L476.65,730.43L476.48,730.25L476.44,729.99L476.23,729.89L476.25,729.77L476.43,729.80L476.43,729.72L476.15,729.68L476.20,729.79L476.15,729.84L475.93,729.84L475.81,729.93ZM473.33,728.27L473.35,728.34L473.56,728.44L473.58,728.55L473.87,728.67L474.02,728.81L474.01,729.03L474.10,729.15L474.67,729.17L474.86,729.37L475.00,729.40L474.99,729.32L474.81,729.17L474.73,728.95L474.56,728.71L474.20,728.59L473.74,728.29L473.39,728.18L473.33,728.27ZM468.59,725.75L468.68,725.98L468.96,726.17L469.21,726.49L469.47,726.67L469.93,726.80L470.10,727.04L471.49,727.21L471.87,727.20L472.10,727.33L472.19,727.52L472.42,727.52L472.83,727.78L472.94,727.77L472.70,727.18L472.14,726.89L471.81,726.53L471.68,726.48L471.47,726.24L471.23,726.12L471.44,726.10L471.61,726.28L471.84,726.28L471.86,726.36L472.04,726.32L471.95,726.15L471.40,725.79L469.41,725.36L469.91,725.81L469.69,725.75L469.77,725.87L469.67,725.79L469.47,725.77L469.41,725.71L469.51,725.52L469.31,725.59L469.01,725.58L468.61,725.70L468.59,725.75ZM478.44,731.34L478.44,731.41L478.53,731.49L478.87,731.45L479.29,731.00L479.01,731.09L478.70,731.05L478.52,731.17L478.44,731.34ZM479.37,730.09L479.55,729.93L479.77,730.03L479.88,730.18L479.79,730.32L479.82,730.52L480.13,730.84L480.19,730.84L480.19,730.48L480.06,730.34L480.10,730.07L480.17,730.07L480.26,730.22L480.48,729.81L480.63,729.76L480.66,729.65L480.85,729.67L480.77,729.91L480.71,729.96L480.63,729.94L480.51,730.15L480.56,730.67L480.65,730.80L480.63,730.92L480.67,730.91L480.73,731.05L480.84,729.92L480.94,729.67L481.17,729.41L481.11,729.08L480.91,728.94L480.88,728.81L480.78,728.75L479.90,728.90L479.78,729.01L479.60,729.39L479.48,729.45L479.40,729.68L479.37,730.09ZM469.89,719.17L469.94,719.61L470.08,719.79L470.12,720.12L470.09,720.20L470.37,720.16L470.60,720.23L470.60,720.43L470.72,720.62L471.01,720.81L471.13,720.98L471.21,721.19L471.17,721.45L471.06,721.67L470.98,721.80L470.75,721.95L470.65,722.17L470.59,722.54L470.70,722.62L471.15,722.54L471.29,722.58L471.74,722.54L471.62,722.45L471.24,722.45L471.50,722.33L472.02,722.46L472.27,722.66L472.29,722.77L472.40,722.83L472.53,723.03L472.78,723.21L472.91,723.48L472.98,723.47L473.03,723.69L472.96,723.72L473.01,724.20L473.20,724.69L473.69,725.19L474.32,725.62L474.46,725.83L475.13,726.26L475.24,726.22L475.30,726.30L475.37,726.26L475.45,726.32L475.89,726.30L476.20,726.54L476.26,726.79L476.43,727.07L476.78,727.29L476.87,727.51L477.42,727.70L477.61,727.97L478.55,728.25L478.70,728.22L478.83,728.10L478.84,727.77L478.95,727.60L478.88,727.63L478.90,727.52L478.93,727.58L480.28,727.49L481.32,727.17L481.77,727.11L482.82,727.28L483.32,727.47L483.33,727.55L483.34,727.47L483.48,727.52L483.74,727.67L483.91,727.97L484.00,728.50L483.90,728.82L483.69,729.18L483.60,729.25L483.28,729.40L482.94,729.42L482.51,729.52L482.34,729.64L482.34,729.75L482.92,729.90L484.18,730.07L484.42,730.17L484.98,730.60L485.37,731.21L485.29,731.20L485.43,731.29L485.53,731.57L485.71,731.76L486.44,731.82L486.79,731.67L486.90,731.56L486.93,731.50L486.92,731.44L486.86,731.43L486.87,731.29L487.07,730.75L486.89,730.72L486.55,730.18L486.51,729.29L486.60,728.91L486.53,728.88L486.53,728.79L486.36,728.78L486.28,728.70L486.30,728.27L486.58,727.93L486.83,727.80L486.87,727.65L487.00,727.55L486.94,727.52L486.99,727.37L486.92,726.99L487.02,726.72L486.98,726.70L487.00,726.52L486.88,726.42L486.60,726.60L486.63,726.65L486.50,726.79L485.99,726.95L485.87,727.08L485.79,727.32L485.73,727.07L485.83,726.96L486.16,726.82L486.16,726.76L486.01,726.84L486.02,726.77L485.96,726.79L485.97,726.74L485.90,726.74L486.02,726.72L485.97,726.77L486.02,726.75L486.08,726.78L486.17,726.73L486.25,726.76L486.46,726.59L486.51,726.40L486.57,726.36L486.59,726.11L486.54,726.01L486.31,726.00L486.03,726.15L485.75,726.21L485.54,726.19L485.50,726.13L485.55,726.12L485.41,725.92L485.33,725.89L485.28,725.75L485.14,725.64L484.88,725.06L484.65,725.06L484.63,725.02L484.65,725.08L484.55,725.12L484.51,725.06L484.47,725.12L484.33,725.08L484.25,724.79L483.92,724.56L483.91,724.37L483.82,724.28L483.79,724.11L483.59,723.94L483.53,723.80L483.53,723.73L483.57,723.74L483.56,723.68L483.51,723.70L483.52,723.60L483.74,723.38L483.68,723.41L483.68,723.36L483.63,723.37L483.63,723.27L483.56,723.22L483.54,723.41L483.46,723.40L483.54,723.20L483.51,723.14L483.50,723.20L483.42,723.21L483.46,723.03L483.50,723.04L483.58,722.87L483.52,722.83L483.52,722.94L483.36,723.10L483.34,723.30L483.33,723.25L483.23,723.27L483.27,723.00L483.38,722.94L483.37,722.75L483.25,722.77L483.15,722.33L483.12,722.48L483.13,722.43L482.96,722.41L483.03,722.28L482.88,722.28L482.90,722.05L482.80,722.04L482.78,722.10L482.83,722.18L482.78,722.25L482.43,722.30L482.29,722.24L482.19,722.34L482.20,722.56L482.02,722.62L482.05,722.69L482.17,722.67L482.18,722.91L482.07,723.01L481.91,722.96L481.75,722.73L481.69,722.71L481.58,722.84L481.45,722.81L481.59,722.91L481.59,722.97L481.35,723.17L481.28,723.77L481.35,724.22L481.04,724.61L480.76,724.84L480.79,724.99L480.89,724.79L481.01,724.78L480.91,725.02L481.05,725.12L480.96,725.12L481.02,725.40L481.15,725.59L481.77,725.54L482.17,725.58L482.24,725.40L481.99,725.19L481.82,724.87L481.83,724.75L482.01,724.69L482.09,724.89L482.04,724.91L482.33,724.94L482.50,725.08L482.70,725.65L482.68,726.01L482.50,726.25L482.27,726.35L482.07,726.24L481.99,726.12L481.98,726.03L482.03,726.00L481.92,725.92L481.71,725.93L481.62,726.07L481.06,726.55L480.80,726.55L480.55,726.35L480.51,726.19L480.60,725.99L480.43,725.81L480.28,725.43L479.66,725.23L479.49,724.98L479.42,724.96L479.18,725.08L479.08,725.21L478.73,725.35L478.44,725.58L478.36,725.45L477.89,725.42L477.65,725.31L477.01,724.87L475.99,724.97L475.73,724.92L475.35,724.72L474.81,724.21L474.40,723.59L474.28,723.33L474.21,723.06L474.20,722.52L474.30,722.07L473.95,721.92L473.80,721.78L473.73,721.26L473.79,721.08L473.58,721.09L473.39,720.98L473.33,720.86L472.99,720.73L472.63,720.30L472.58,720.34L472.34,720.29L472.33,720.20L472.08,720.20L472.09,720.14L471.78,720.04L471.83,720.08L471.78,720.14L471.33,719.97L471.25,719.86L471.29,719.81L471.42,719.84L471.43,719.80L470.93,719.60L470.84,719.52L470.73,719.21L470.49,719.17L470.35,719.06L470.15,719.07L470.02,719.00L469.89,719.17ZM479.13,721.55L479.17,721.69L479.74,721.50L480.22,721.47L480.22,721.52L480.59,721.56L480.66,721.69L480.71,721.60L480.97,721.46L481.03,721.27L480.98,721.15L480.82,721.09L480.71,721.20L480.54,721.23L480.43,721.06L480.32,721.15L480.32,721.22L479.98,721.34L479.73,721.25L479.69,721.18L479.24,721.39L479.13,721.55ZM473.10,718.79L473.38,719.51L474.22,720.56L474.40,721.12L474.74,721.25L474.86,721.38L474.92,721.87L474.81,722.03L475.07,722.07L475.99,722.81L476.49,723.11L477.13,723.31L478.02,723.43L478.10,723.54L478.22,723.47L478.37,723.51L478.36,723.45L478.48,723.37L478.89,723.53L479.02,723.77L479.05,723.95L479.02,724.01L479.00,723.97L478.99,724.10L479.14,724.28L479.25,724.58L479.42,724.82L479.54,724.62L479.91,724.40L479.94,724.33L479.84,724.30L479.47,723.81L479.42,723.61L479.23,723.33L479.27,723.15L479.11,723.11L478.84,722.84L478.76,722.67L478.59,722.63L478.42,722.45L478.46,722.39L478.42,722.33L478.09,722.42L478.00,722.35L477.96,722.38L477.77,722.30L477.58,722.06L477.16,722.00L477.01,721.80L476.91,721.75L476.86,721.62L476.91,721.57L476.80,721.54L476.48,721.16L476.51,720.90L476.59,720.93L476.61,720.87L476.69,720.88L476.90,720.68L477.46,720.53L477.81,720.55L477.95,720.68L478.28,720.54L478.47,720.61L478.71,720.81L478.80,720.81L478.83,720.63L478.76,720.51L478.89,720.54L478.80,720.39L478.82,720.16L478.65,720.09L478.60,720.00L478.81,719.91L478.90,719.99L479.11,719.99L479.09,719.91L479.04,719.98L478.91,719.94L478.73,719.74L478.69,719.83L478.60,719.85L478.55,719.78L478.31,719.77L478.15,719.60L477.64,719.54L477.51,719.44L477.44,719.27L476.94,718.62L476.86,718.29L477.08,717.91L477.24,717.85L477.22,717.80L477.33,717.66L477.41,717.41L477.21,717.09L477.03,717.21L477.04,716.82L476.89,716.63L476.75,716.54L476.53,716.53L476.15,716.39L476.23,716.62L476.04,716.98L475.32,717.79L475.08,717.96L474.52,718.13L474.16,718.16L473.36,718.12L473.36,718.27L473.10,718.79ZM495.26,719.34L495.36,719.52L495.57,719.60L495.76,719.57L495.88,719.41L495.68,719.23L495.51,719.20L495.26,719.34ZM465.83,717.96L465.94,718.10L466.24,718.26L466.29,718.40L466.55,718.61L466.63,718.78L466.79,718.79L466.85,719.05L466.94,719.13L467.10,719.10L467.39,718.86L467.61,718.79L468.04,718.79L468.05,718.57L467.98,718.36L468.17,718.08L468.31,717.94L468.47,717.94L468.53,717.82L468.51,717.76L468.34,717.68L468.73,717.61L468.81,717.46L468.61,717.35L468.57,717.23L468.17,717.09L468.15,716.99L468.02,716.98L468.01,716.90L467.95,716.93L467.93,716.67L467.81,716.56L467.34,716.29L467.05,716.23L466.84,716.31L466.52,716.79L466.34,716.90L466.06,717.24L465.85,717.73L465.83,717.96ZM470.50,660.71L470.26,661.49L469.92,662.30L469.38,663.07L468.66,663.83L466.18,664.97L465.35,665.42L464.71,666.17L463.27,669.72L463.34,669.73L463.23,669.76L463.37,669.95L463.49,669.95L463.57,669.87L463.46,670.00L463.59,670.06L463.62,670.15L463.70,670.11L463.68,670.23L463.90,670.32L464.10,670.31L463.88,670.36L463.99,670.68L463.93,671.11L464.38,671.02L464.59,671.09L464.81,671.28L464.93,671.22L464.90,671.33L464.98,671.39L465.19,671.38L465.24,671.46L465.47,671.41L465.52,671.45L465.49,671.62L465.65,671.70L465.53,671.93L465.64,671.98L465.65,672.10L465.42,672.36L465.51,672.53L465.76,672.65L465.85,672.85L465.93,672.86L465.85,673.00L465.48,672.96L465.28,673.03L465.16,672.99L465.10,673.20L465.34,673.40L465.30,673.70L465.49,673.88L465.42,673.80L465.72,673.85L465.73,673.99L465.82,674.00L465.79,673.92L465.85,673.91L465.89,673.99L465.68,674.03L465.84,674.06L465.90,674.00L465.95,673.82L466.12,673.75L466.16,673.62L466.27,673.54L466.18,673.58L466.18,673.40L466.23,673.27L466.31,673.24L466.16,673.07L466.30,673.03L466.27,672.99L466.40,673.02L466.48,672.91L466.44,673.00L466.53,672.95L466.45,673.06L466.54,673.21L466.54,673.35L466.74,673.40L466.75,673.46L466.57,673.66L466.72,673.65L466.76,673.77L466.68,673.79L466.69,673.93L466.63,673.87L466.68,674.02L466.58,674.14L466.43,674.19L466.35,674.12L466.19,674.20L466.17,674.30L466.28,674.35L466.20,674.24L466.32,674.22L466.48,674.55L466.70,674.60L466.78,674.58L466.78,674.52L466.91,674.58L466.92,674.92L466.99,675.03L466.95,675.07L466.85,675.03L466.89,675.11L467.10,675.11L467.15,675.15L467.14,675.28L467.38,675.29L467.43,675.35L467.41,675.57L467.02,675.75L467.12,675.91L467.18,675.93L467.09,675.77L467.22,675.82L467.20,675.87L467.30,675.92L467.26,675.98L467.41,675.96L467.41,676.04L467.61,676.10L467.60,676.37L467.73,676.50L467.95,676.53L468.07,676.85L468.27,676.89L468.49,676.63L468.46,676.77L468.27,677.00L468.45,677.28L470.14,678.65L470.28,679.00L470.54,679.13L470.58,679.25L470.80,679.32L470.90,679.55L470.88,679.84L470.63,680.10L470.54,680.08L470.64,680.30L470.87,680.44L470.89,680.35L471.18,680.23L471.60,680.25L471.67,680.33L471.81,680.33L471.89,680.38L472.01,680.66L472.40,680.87L472.39,680.93L472.71,681.04L473.04,681.43L473.29,681.49L473.39,681.43L473.99,681.62L474.00,681.56L474.07,681.55L474.11,681.61L474.17,681.48L474.36,681.63L474.42,681.61L474.72,681.77L474.85,681.71L474.97,681.48L475.10,681.61L475.23,681.59L475.27,681.70L475.21,681.82L475.03,681.85L475.02,681.74L474.89,681.91L474.74,681.82L474.87,681.94L474.61,682.18L474.88,681.97L475.07,682.09L475.14,682.25L475.43,682.39L475.50,682.88L475.68,683.06L475.90,683.09L475.94,683.22L476.11,683.29L476.43,683.25L476.59,683.35L476.80,683.68L476.89,683.95L477.25,684.20L477.59,684.26L477.94,684.81L478.19,685.07L478.60,685.15L478.75,685.27L478.87,685.49L478.98,685.55L479.20,685.58L479.32,685.45L479.31,685.57L479.20,685.69L479.36,685.77L479.38,685.89L479.44,685.86L479.37,685.90L479.45,685.87L479.54,685.99L479.80,686.07L479.91,686.19L480.06,686.15L479.99,686.09L480.26,686.09L480.58,686.22L480.64,686.32L480.97,686.49L481.31,686.88L481.59,687.63L481.92,687.76L481.91,687.91L481.92,687.66L481.98,687.66L481.97,687.94L481.88,687.95L481.96,687.95L481.98,687.93L481.98,687.70L482.36,687.70L482.55,687.79L482.93,688.25L483.16,688.37L483.96,689.01L484.57,689.78L484.80,689.95L484.87,690.08L484.92,690.59L484.81,691.20L484.91,691.36L484.96,691.63L485.11,691.79L485.19,691.83L485.03,691.65L485.00,691.53L485.33,691.31L485.56,691.28L485.74,691.38L485.82,691.51L486.33,691.66L486.50,691.77L486.53,691.74L486.50,691.85L486.58,691.87L486.53,691.94L486.68,691.89L486.57,692.03L486.66,692.00L486.64,692.06L486.73,692.07L486.81,692.19L487.00,692.25L487.01,692.35L487.01,692.27L487.18,692.36L487.32,692.60L487.30,692.51L487.34,692.50L487.32,692.65L487.61,692.96L487.99,693.11L488.03,693.08L488.11,693.29L488.35,693.35L488.46,693.44L488.67,693.94L488.73,693.88L488.86,693.97L488.79,694.11L488.84,694.19L489.31,694.64L489.59,694.56L489.53,694.66L489.64,694.68L489.44,694.77L489.58,694.99L489.66,695.30L489.85,695.38L489.99,695.53L490.25,696.32L490.43,696.40L490.78,696.31L490.36,696.36L490.50,696.31L490.40,696.32L490.45,696.31L490.39,696.18L490.45,696.09L490.41,696.10L490.38,695.99L490.29,695.99L490.30,695.87L490.44,695.91L490.52,696.13L490.61,696.11L490.47,695.65L490.62,695.86L490.62,696.01L490.70,696.09L490.66,696.16L490.74,696.13L490.57,695.65L490.73,695.72L490.79,695.89L490.85,695.88L490.83,696.00L490.93,696.17L491.65,696.41L491.54,696.39L491.84,696.65L492.05,697.05L492.15,697.46L492.19,698.45L492.10,698.96L491.77,699.72L491.78,700.50L491.60,700.81L491.69,700.73L491.65,700.89L491.73,700.83L491.71,700.95L491.81,700.84L491.86,700.87L491.82,701.07L491.86,701.03L491.90,701.14L491.91,700.91L491.94,700.93L491.98,700.80L492.04,700.81L492.11,700.70L492.07,700.39L492.13,700.41L492.32,700.84L492.28,700.87L492.36,701.17L492.28,701.16L492.38,701.38L492.34,701.54L492.40,701.60L492.36,701.60L492.36,701.67L492.44,701.99L492.39,702.09L492.53,702.23L492.68,703.03L492.68,703.76L492.56,704.73L492.72,705.38L492.70,706.13L492.77,706.16L492.81,706.51L492.98,706.71L493.22,707.59L493.35,707.76L493.60,708.63L493.76,709.50L493.90,713.37L494.03,714.06L494.13,714.08L494.15,714.05L494.14,713.96L494.10,714.07L494.05,714.04L494.07,713.89L494.35,713.65L494.49,713.10L494.39,713.45L494.35,713.49L494.29,713.45L494.32,713.50L494.28,713.57L494.28,713.47L494.24,713.50L494.20,713.40L494.29,713.41L494.26,713.30L494.19,713.19L494.24,713.18L494.21,713.13L494.25,713.08L494.27,713.13L494.31,712.96L494.24,712.83L494.21,712.85L494.23,712.79L494.28,712.80L494.29,712.70L494.23,712.68L494.24,712.49L494.16,712.29L494.16,712.07L494.22,711.94L494.15,711.84L494.28,711.55L494.19,711.39L494.07,711.34L494.10,711.28L493.96,711.00L493.92,710.50L494.10,710.18L494.36,709.97L494.65,709.87L494.91,709.93L494.99,709.91L494.92,709.91L494.91,709.80L494.74,709.66L494.73,709.58L494.70,709.63L494.63,709.60L494.71,709.56L494.68,709.47L494.67,709.52L494.59,709.49L494.43,709.20L494.36,709.18L494.38,709.09L494.53,709.01L494.52,708.95L494.31,709.11L494.45,708.96L494.16,708.58L494.18,708.52L494.28,708.55L494.20,708.33L494.07,708.25L494.00,708.08L493.95,708.12L493.91,708.04L493.81,708.02L493.67,707.82L493.65,707.67L493.53,707.48L493.67,707.40L493.53,706.88L493.82,707.17L494.01,707.21L494.18,707.37L494.36,707.37L494.49,707.44L494.86,707.74L494.96,707.71L494.95,707.78L495.08,707.88L495.42,707.95L495.81,708.16L495.88,708.12L495.82,708.06L496.01,708.08L496.07,708.18L496.10,708.15L496.11,708.44L496.03,708.42L495.93,708.48L495.95,708.87L495.86,708.96L495.87,709.10L495.77,709.29L495.71,709.30L495.77,709.90L495.70,709.99L495.72,710.36L495.86,710.49L495.87,710.58L496.09,710.75L496.11,710.85L496.24,710.89L496.25,711.00L496.30,710.98L496.28,711.07L496.32,711.05L496.36,711.14L496.41,711.08L496.47,711.26L496.54,711.21L496.54,711.29L496.58,711.27L496.64,711.36L496.71,711.31L496.87,711.48L496.84,711.54L496.97,711.66L496.94,711.91L497.11,712.04L497.05,712.14L497.19,712.25L497.17,712.38L497.31,712.36L497.28,712.41L497.39,712.55L497.58,712.62L497.65,712.75L497.70,712.78L497.72,712.73L497.72,712.80L497.88,712.73L498.00,712.87L498.28,712.84L498.32,712.88L498.62,712.84L498.69,712.78L498.67,712.85L498.76,712.88L498.79,712.82L498.85,712.94L499.01,712.97L499.05,712.89L499.04,712.97L499.11,712.91L499.13,713.01L499.24,713.08L499.34,713.03L499.32,713.12L499.64,713.16L499.64,713.26L499.76,713.23L499.79,713.30L499.87,713.26L500.06,713.33L500.12,713.60L500.29,713.76L500.41,713.67L500.43,713.75L500.49,713.68L500.64,713.73L500.68,713.80L500.59,713.82L500.60,713.90L500.69,714.01L500.74,713.99L500.68,714.12L500.60,714.15L500.97,714.28L501.23,714.55L501.17,714.98L501.08,714.92L501.04,715.21L500.72,715.17L500.70,715.37L500.50,715.86L500.54,716.64L500.75,716.91L500.80,717.39L500.87,717.46L500.88,717.64L500.73,717.84L500.84,718.02L500.70,718.20L500.75,718.66L500.64,718.66L500.61,718.71L500.59,718.67L500.58,718.72L500.48,718.67L500.52,718.73L500.39,718.89L500.48,718.79L500.55,718.85L500.51,718.95L500.44,718.97L500.42,718.91L500.21,719.06L500.19,719.37L500.27,719.30L500.27,719.04L500.38,719.03L500.37,719.39L500.46,719.39L500.47,719.31L500.57,719.29L500.51,719.46L500.27,719.53L500.02,719.45L500.27,719.54L500.57,719.48L500.70,719.57L500.70,719.66L500.88,719.85L500.95,719.82L501.15,719.88L501.06,719.92L501.09,719.95L501.23,719.85L501.31,719.88L501.26,719.96L501.32,719.87L501.41,719.89L501.75,720.30L501.81,720.45L501.80,720.38L501.88,720.48L501.97,720.42L502.03,720.47L502.05,720.51L501.97,720.56L502.03,720.75L501.99,720.90L501.87,720.83L501.86,720.60L501.80,720.46L501.88,721.10L501.82,721.14L501.81,721.19L501.84,721.13L501.86,721.32L502.04,721.51L502.24,721.55L502.31,721.48L502.15,721.51L502.02,721.44L501.92,721.30L501.92,721.12L502.11,721.25L502.31,721.24L502.31,721.33L502.34,721.22L502.39,721.23L502.42,721.34L502.46,721.34L502.40,721.25L502.48,721.21L502.76,721.27L502.87,721.32L503.04,721.56L502.89,721.33L503.11,721.55L503.10,721.63L503.16,721.62L503.30,721.87L503.30,721.72L503.26,721.76L503.23,721.69L503.35,721.62L503.45,721.64L503.49,721.74L503.46,721.85L503.23,721.94L503.35,721.95L503.68,721.81L503.87,721.97L503.98,721.82L504.23,721.91L504.49,722.30L504.98,722.62L505.04,722.82L504.90,722.83L504.97,722.99L505.28,723.23L505.37,723.24L505.40,723.14L505.54,723.17L505.89,723.44L505.78,723.70L506.00,724.01L505.92,724.04L506.00,724.09L506.05,724.28L506.16,724.40L506.12,724.42L506.18,724.58L506.30,724.67L506.24,724.68L506.30,724.89L506.44,724.96L506.35,725.09L506.39,725.14L506.60,725.17L506.89,725.46L507.19,725.62L506.74,725.45L506.60,725.21L506.33,725.31L506.49,725.60L506.73,725.72L507.04,726.12L507.21,726.68L507.15,727.16L507.20,727.29L507.31,727.41L507.45,727.46L507.42,727.35L507.49,727.02L507.33,726.65L507.65,727.04L507.61,727.50L507.49,727.53L507.56,727.96L507.74,727.95L507.94,727.74L508.24,727.89L508.28,728.00L508.23,728.05L508.31,728.20L508.10,728.16L508.04,728.28L507.72,728.17L507.39,728.16L507.28,728.01L507.13,728.05L507.17,728.17L507.32,728.28L507.43,728.19L507.77,728.20L508.02,728.29L508.09,728.46L508.04,728.51L508.04,728.71L507.90,728.66L507.89,728.69L507.68,728.68L507.90,728.72L508.04,728.98L507.98,729.55L508.03,729.71L507.92,729.88L507.78,730.57L507.67,730.62L507.69,730.75L507.52,730.82L507.57,731.13L507.67,731.12L507.68,731.19L507.56,731.18L507.44,731.66L507.38,731.69L507.45,732.31L507.57,732.58L507.71,732.68L507.89,733.02L507.92,733.38L507.87,733.40L507.87,733.54L508.04,733.69L507.99,733.78L508.17,733.91L508.12,733.98L508.13,734.31L508.19,734.34L508.14,734.38L508.23,734.62L508.14,734.65L508.16,734.72L508.25,734.72L508.19,734.74L508.26,734.78L508.23,734.83L508.33,734.88L508.32,735.01L508.80,735.26L508.90,735.24L508.90,735.14L509.04,735.23L509.41,735.62L509.62,735.98L509.56,736.00L509.65,736.12L509.74,736.11L509.74,736.26L509.93,736.37L509.72,736.29L509.87,737.02L509.79,737.18L509.82,737.29L509.67,737.44L509.62,737.67L509.41,737.79L509.40,737.94L509.24,738.03L509.22,738.28L509.35,738.30L509.39,738.43L509.00,738.67L509.40,738.45L509.48,738.64L509.47,738.90L509.72,739.04L510.19,738.93L510.89,739.08L511.22,739.34L511.86,739.65L511.97,740.00L515.60,739.37L515.92,739.21L516.57,739.18L516.70,738.98L516.64,738.87L516.18,738.82L516.50,737.77L516.16,737.09L516.44,735.50L516.66,735.21L516.89,735.02L517.34,734.78L518.07,734.27L518.22,734.20L519.38,734.14L519.50,734.18L519.55,734.10L520.80,733.95L521.00,733.86L521.68,733.94L522.58,733.79L523.04,733.83L523.13,733.79L523.39,732.96L523.62,731.36L523.68,731.26L523.86,731.15L524.06,730.11L524.14,729.03L524.12,728.80L523.89,728.45L524.19,728.08L524.27,727.81L524.56,727.81L524.88,728.07L525.04,728.29L525.54,728.47L525.82,728.64L527.57,729.01L527.69,729.00L527.75,728.94L527.67,728.46L527.74,728.17L527.68,727.56L527.51,727.25L527.53,727.16L527.39,726.95L527.16,726.93L526.84,726.49L527.06,726.07L527.52,725.80L527.42,725.63L527.27,725.63L527.37,725.40L527.22,725.13L527.28,725.12L527.26,724.99L527.56,724.95L527.68,725.12L527.91,725.04L528.06,725.17L528.82,724.70L528.87,724.59L528.60,724.46L528.32,723.85L528.51,723.58L528.48,723.39L528.64,723.22L529.03,723.05L528.75,722.82L528.60,722.78L528.56,722.70L528.24,722.54L528.07,722.62L527.69,722.56L527.39,721.88L527.26,721.28L526.52,720.81L526.40,720.66L526.19,720.73L526.04,720.86L525.65,720.83L525.45,720.47L525.49,720.10L524.53,719.56L524.46,719.15L524.29,718.75L524.06,718.67L523.98,718.32L524.01,718.03L524.23,717.74L524.64,717.57L524.87,717.59L525.03,717.50L525.34,717.47L525.47,717.24L525.76,717.16L526.22,716.93L526.18,716.68L526.22,716.50L525.57,716.16L525.56,715.56L525.77,715.52L525.69,715.46L525.72,715.32L525.55,715.08L525.53,714.93L525.65,714.29L525.81,714.07L525.61,713.97L525.63,713.42L525.68,713.35L526.05,713.64L526.12,713.59L526.08,714.00L526.26,714.11L526.35,714.09L526.53,714.25L526.71,713.80L526.88,713.73L527.00,713.76L527.42,714.19L527.50,713.81L527.69,713.75L527.64,713.45L527.89,713.46L527.98,713.68L528.33,714.03L527.98,712.72L527.55,712.25L527.49,711.57L526.74,711.46L525.85,710.62L524.87,711.10L524.78,709.99L525.01,709.53L525.56,709.31L525.29,707.59L525.71,707.19L526.05,707.07L526.15,706.94L526.01,706.93L527.07,706.05L526.42,704.72L526.58,704.52L526.51,704.19L526.59,704.01L527.70,703.65L528.18,703.37L528.61,703.26L529.06,702.70L529.25,702.74L530.43,702.65L530.74,702.39L531.77,702.21L531.36,701.21L529.43,699.99L528.73,699.82L527.52,698.87L527.79,698.72L528.27,698.22L527.43,697.20L528.64,695.91L528.02,694.98L528.79,694.59L529.77,693.47L530.45,693.05L530.52,692.87L530.74,693.04L531.16,692.83L529.28,690.33L528.19,689.07L527.57,688.24L526.75,687.95L525.07,687.80L524.22,687.55L523.44,686.90L522.85,686.04L522.46,685.21L521.79,683.40L521.29,682.53L520.44,681.61L519.61,681.07L517.10,680.15L516.35,679.53L516.10,678.72L516.51,674.48L516.18,673.64L515.38,673.20L514.57,673.37L513.72,673.74L512.87,674.21L511.17,675.27L510.29,675.93L509.60,676.68L508.63,678.31L508.24,679.07L507.43,681.70L506.34,686.79L505.70,687.54L504.90,687.75L504.11,687.70L503.29,687.46L502.38,686.98L501.61,686.12L500.95,685.20L500.20,684.52L499.44,684.02L498.69,683.37L498.07,682.53L497.46,680.88L496.89,679.11L496.19,678.23L495.43,677.69L494.65,677.50L493.83,677.44L490.30,677.65L488.58,677.50L487.67,677.49L486.03,677.24L485.21,676.50L484.81,675.61L484.71,674.81L484.24,673.20L483.76,672.30L483.01,671.66L482.23,671.45L481.44,671.37L480.57,671.40L478.01,671.89L477.22,671.91L476.38,671.37L476.06,670.48L476.07,669.66L476.33,667.14L477.13,663.04L477.11,662.20L477.00,661.35L476.34,660.53L475.57,660.18L474.77,660.13L473.06,660.34L471.35,660.43L470.50,660.71Z" fill="#f3efe7" stroke="#b4ad9f" stroke-width=".8" vector-effect="non-scaling-stroke" fill-rule="evenodd"><title>منطقة جازان</title></path>
<path d="M321.64,78.83L310.42,82.23L308.91,82.62L299.60,85.48L283.90,90.12L294.21,101.99L302.02,111.10L304.69,114.16L306.26,115.82L313.01,123.62L313.32,124.19L314.77,126.08L322.88,135.74L309.91,143.37L308.76,145.90L304.24,156.50L303.64,158.38L303.49,158.43L303.91,158.82L305.22,160.56L305.64,160.95L306.49,162.15L307.51,163.09L308.69,163.96L309.97,165.05L310.93,165.54L312.08,166.40L315.63,168.59L317.22,169.76L317.58,170.29L319.57,171.57L319.92,172.07L320.68,172.39L322.82,173.58L325.36,174.60L326.97,175.35L327.61,175.44L331.02,176.18L332.11,176.70L332.51,177.14L335.30,178.74L336.10,179.35L337.51,180.07L338.54,180.71L338.64,180.88L338.66,181.30L338.55,181.77L338.32,182.21L337.65,182.97L336.95,183.40L335.53,184.97L333.51,187.73L332.81,188.98L332.23,190.58L331.73,191.54L331.63,191.85L331.27,192.44L330.98,193.29L331.20,194.37L331.48,194.83L331.73,195.09L332.30,195.39L333.09,195.12L334.22,195.04L336.23,194.24L337.21,193.69L338.76,192.96L339.56,192.68L341.29,192.32L342.08,192.01L343.34,192.01L344.10,192.11L345.70,192.82L346.67,193.30L346.88,193.52L348.10,194.38L349.15,195.25L351.52,196.79L352.24,197.44L354.11,198.82L357.10,200.18L358.12,201.02L358.63,201.18L358.91,201.46L360.81,202.70L362.04,203.70L362.95,204.73L363.46,205.18L363.66,205.57L364.30,206.09L364.70,206.60L365.83,207.67L366.85,208.54L368.27,209.91L369.10,210.55L370.30,211.21L371.02,211.89L371.46,212.04L373.62,213.94L374.02,214.16L374.36,214.44L375.40,214.91L376.13,215.37L376.82,215.71L377.69,216.29L380.97,217.29L382.47,217.91L382.91,217.95L384.95,218.70L385.79,219.13L386.92,219.36L388.08,220.05L390.21,220.52L391.49,221.26L392.17,221.85L392.64,222.39L393.07,223.69L393.66,223.39L394.31,222.93L394.50,222.64L394.85,222.49L396.98,220.87L399.30,219.73L402.65,218.83L405.63,217.44L406.17,217.13L406.91,216.91L407.28,216.72L407.54,216.36L408.14,215.87L409.07,215.58L409.39,215.34L410.16,214.92L410.95,214.70L411.24,214.51L411.57,214.45L412.39,214.01L413.79,213.04L417.84,211.62L418.45,211.28L419.64,210.41L419.95,210.09L421.03,209.23L421.92,208.65L422.01,208.50L422.28,208.43L423.25,208.39L423.51,208.47L423.89,208.42L424.19,208.47L425.17,208.45L426.63,208.21L427.47,208.17L428.10,208.29L428.66,208.57L429.09,208.88L430.28,208.91L431.58,208.84L432.15,208.75L432.47,208.60L433.15,208.46L433.75,208.23L435.41,208.35L436.00,208.57L437.24,208.70L438.86,208.73L439.73,208.82L440.76,208.74L441.69,209.02L442.06,209.07L443.48,208.96L444.23,209.07L445.03,208.98L446.01,209.15L448.30,209.15L450.30,208.69L450.96,208.74L451.99,209.00L453.03,209.14L454.00,209.05L454.01,207.51L454.12,206.15L454.50,204.40L455.04,202.91L455.59,201.98L455.63,201.33L455.97,200.07L456.24,199.26L456.65,198.64L457.29,198.09L458.00,197.75L458.98,197.41L459.63,197.29L461.12,197.29L462.73,197.68L463.96,197.39L464.14,197.04L464.13,196.46L463.87,195.01L463.62,194.28L463.38,193.81L462.96,191.81L463.03,191.36L463.40,190.36L463.52,189.85L463.51,189.41L463.14,189.10L462.59,188.82L462.14,188.35L461.93,187.77L461.34,185.32L460.11,183.48L459.99,182.68L460.01,181.18L460.05,180.53L460.36,179.32L460.38,178.71L460.06,178.29L458.87,178.33L458.46,178.40L458.20,178.67L458.04,179.47L457.82,179.75L457.50,179.94L457.10,179.90L456.86,179.56L456.66,178.11L456.60,176.71L456.48,176.05L456.56,175.46L456.43,174.63L456.49,173.79L456.43,173.06L456.46,171.21L456.75,169.46L456.63,169.06L455.81,167.85L456.08,167.25L456.45,167.05L457.27,166.98L458.13,166.57L460.06,166.63L460.32,166.22L460.48,165.39L460.75,165.03L461.14,164.89L461.58,164.63L461.85,164.36L462.22,162.95L463.05,162.25L463.55,162.04L464.01,161.90L464.53,161.91L465.97,162.32L466.82,162.18L468.26,161.74L468.96,162.26L469.30,162.43L469.67,162.50L470.02,162.21L470.22,161.90L470.32,159.25L470.47,158.23L471.27,157.34L471.83,156.93L472.39,156.65L473.48,156.48L474.52,156.22L475.75,156.12L476.84,155.74L477.26,155.70L477.78,155.47L478.53,154.82L478.64,154.42L478.51,153.86L478.30,153.62L476.88,152.49L475.38,151.08L474.99,150.99L474.58,150.70L474.16,150.23L473.39,149.12L472.56,148.48L471.92,147.16L471.53,146.75L471.05,146.46L470.68,146.04L470.34,145.39L470.16,144.75L469.88,144.30L469.41,143.84L468.83,143.53L467.60,143.15L466.86,143.07L465.69,143.14L464.35,142.94L463.63,142.57L463.13,142.13L462.63,141.55L462.28,140.30L462.25,138.92L462.39,138.65L463.55,137.79L464.12,137.23L464.67,136.65L465.03,136.15L465.31,135.64L465.29,135.20L465.11,134.83L464.77,134.67L464.19,134.55L463.09,134.55L462.46,134.46L461.93,134.15L461.56,133.74L460.88,133.58L460.38,133.57L458.51,133.30L458.00,133.01L457.53,132.27L457.12,131.82L456.73,131.76L456.15,131.96L455.41,132.41L454.96,132.60L454.48,132.64L453.34,132.48L452.78,132.48L452.11,132.34L451.45,132.32L450.24,132.09L448.88,131.99L448.07,132.21L447.41,132.69L446.71,133.69L446.46,134.61L446.06,135.25L445.73,135.47L445.34,135.53L444.79,135.32L444.49,134.78L444.30,134.13L444.28,133.34L444.32,132.50L444.60,131.64L444.94,131.16L445.03,130.60L445.04,129.12L445.14,128.60L445.12,128.10L444.98,127.52L444.27,126.75L444.14,126.36L444.34,126.08L444.42,125.71L444.30,125.31L442.75,124.42L440.79,123.88L438.16,123.58L437.43,123.43L436.72,123.54L436.20,123.43L434.93,123.62L433.10,124.22L432.36,124.33L431.67,124.55L430.34,124.62L429.74,124.78L428.22,124.99L426.14,124.68L420.75,123.51L419.99,122.94L419.68,122.86L419.09,122.81L418.27,123.00L417.51,122.74L416.89,122.16L415.39,122.05L414.83,121.54L414.33,121.27L413.48,120.99L412.94,120.96L412.19,121.00L411.86,121.09L411.32,121.56L410.63,122.62L410.15,122.83L409.91,123.01L409.53,123.59L409.41,123.96L409.03,124.54L408.07,124.90L407.20,125.05L406.52,125.13L404.91,125.11L404.54,125.05L403.72,124.45L402.75,123.30L402.47,123.10L402.06,123.00L401.75,122.75L400.84,122.39L400.40,122.09L399.92,121.31L399.45,120.71L398.77,120.05L397.77,119.88L397.20,119.70L396.91,119.31L396.55,118.50L396.09,118.05L395.32,118.23L394.85,118.46L394.30,118.87L393.82,119.00L393.02,119.41L392.23,119.59L392.29,119.65L391.64,119.65L391.03,119.45L389.86,118.84L389.21,118.39L387.53,116.90L386.82,116.92L386.57,117.18L386.30,117.31L386.01,117.33L384.96,116.70L384.52,116.81L384.09,117.08L383.67,117.11L382.80,116.89L382.48,116.90L381.71,116.76L381.37,116.79L381.07,116.83L380.16,117.21L379.37,117.75L378.92,117.90L378.31,117.96L377.71,117.84L377.18,117.49L375.69,116.12L374.64,115.47L374.39,115.24L371.89,114.65L370.30,114.41L366.69,114.35L365.98,114.47L363.44,114.50L362.94,114.55L362.64,114.70L361.96,114.48L361.33,114.50L361.03,114.60L360.83,114.76L360.49,114.90L360.16,114.86L359.70,115.01L356.99,117.23L356.07,117.48L354.74,117.37L354.46,117.46L354.19,117.82L353.96,118.39L353.33,118.41L352.93,118.23L352.71,118.03L352.60,117.05L352.19,116.47L351.71,116.20L351.29,116.10L350.19,116.33L349.61,116.37L349.44,116.50L347.53,116.10L345.89,116.12L344.82,116.03L344.25,116.11L343.62,116.54L343.01,116.83L341.26,117.22L340.51,117.28L339.29,117.30L337.95,117.24L336.00,117.33L335.51,117.58L335.11,117.68L334.11,118.09L333.56,118.14L333.24,118.41L333.04,118.81L332.83,119.08L332.52,119.23L331.96,119.29L331.34,119.12L331.11,119.14L330.81,119.36L330.57,119.44L330.12,119.18L329.95,118.88L329.35,116.89L329.39,115.59L329.74,114.06L329.96,113.45L329.99,112.94L329.80,112.36L329.30,111.68L328.96,110.80L329.24,110.10L329.28,109.50L329.09,108.94L327.98,107.78L327.58,107.78L327.08,108.35L326.80,108.57L326.29,108.63L325.94,108.26L325.76,107.52L325.95,107.05L326.52,106.50L326.78,106.10L326.75,105.65L326.24,105.15L325.94,105.24L325.66,105.50L325.47,106.10L325.19,106.32L324.59,105.90L323.48,104.06L323.03,103.72L322.66,103.62L321.74,103.85L320.88,103.54L320.44,103.73L320.09,103.96L319.65,104.02L319.12,103.83L318.77,103.52L318.69,103.22L318.66,102.91L319.27,102.31L319.46,102.26L319.77,101.62L319.75,100.97L320.12,100.60L320.48,100.07L320.54,99.14L320.33,97.95L320.38,97.61L320.48,97.31L320.76,96.89L321.06,96.68L321.45,96.22L321.50,95.86L321.50,95.20L321.24,94.63L320.86,93.42L320.73,92.82L320.71,92.32L320.96,91.96L321.32,91.75L321.56,91.52L321.47,90.78L321.29,90.57L320.69,89.14L320.98,88.39L321.01,88.04L320.85,87.44L320.91,87.14L321.31,86.39L321.26,85.97L320.53,84.44L320.59,84.02L320.55,83.21L320.68,82.78L321.08,82.35L321.27,82.24L321.47,81.96L321.52,81.74L321.47,81.20L321.15,80.61L320.98,80.08L321.16,79.58L321.64,78.83Z" fill="#eee9df" stroke="#b4ad9f" stroke-width=".8" vector-effect="non-scaling-stroke" fill-rule="evenodd"><title>منطقة الجوف</title></path>
<text class="pa-region-label" x="792.46" y="485.93" text-anchor="middle">المنطقة الشرقية</text>
<text class="pa-region-label" x="659.44" y="669.52" text-anchor="middle">منطقة نجران</text>
<text class="pa-region-label" x="518.60" y="158.42" text-anchor="middle">منطقة الحدود الشمالية</text>
<text class="pa-region-label" x="459.92" y="305.15" text-anchor="middle">منطقة حائل</text>
<text class="pa-region-label" x="620.32" y="464.71" text-anchor="middle">منطقة الرياض</text>
<text class="pa-region-label" x="502.95" y="628.24" text-anchor="middle">منطقة عسير</text>
<text class="pa-region-label" x="452.09" y="549.11" text-anchor="middle">منطقة مكة المكرمة</text>
<text class="pa-region-label" x="276.04" y="283.20" text-anchor="middle">منطقة تبوك</text>
<text class="pa-region-label" x="397.32" y="357.43" text-anchor="middle">منطقة المدينة المنورة</text>
<text class="pa-region-label" x="526.43" y="348.75" text-anchor="middle">منطقة القصيم</text>
<text class="pa-region-label" x="467.74" y="586.71" text-anchor="middle">منطقة الباحة</text>
<text class="pa-region-label" x="514.69" y="731.01" text-anchor="middle">منطقة جازان</text>
<text class="pa-region-label" x="338.64" y="189.93" text-anchor="middle">منطقة الجوف</text>
            <g id="pa-markers"></g>
          </g>
        </svg>
      </div>
      <div class="pa-map-legend" id="pa-map-legend"></div>
      <p class="pa-map-hint">الرقم بجوار الرمز هو عدد الحالات في المنفذ. الرموز المتقاربة متصلة بمواقعها بخط رفيع لتسهيل اختيارها.</p>
      <p class="pa-map-credit">حدود المناطق: geoBoundaries · بيانات 2017 · الإحداثيات من ملف الربط.</p>
    </div>
    <div class="pa-list-column"><div class="pa-list-heading"><h3>عدد التقاطعات لكل منفذ</h3><span id="pa-list-count"></span></div><div id="pa-port-list" class="pa-port-list"></div></div>
  </div>
  <div id="pa-location-notice" class="pa-location-notice" hidden></div>
</section>
<section class="panel wa-panel" id="weekly-analytics" aria-labelledby="wa-title">
  <div class="section-header pa-header"><div><h2 class="section-title" id="wa-title">حالات التقاطع أسبوعيًا</h2><p class="section-subtitle">حالات جميع نسب الغياب ضمن فلاتر المجموعات المحددة · من الأحد إلى السبت، مع اقتصار أول وآخر أسبوع على فترة التقرير.</p></div><span class="pa-scope-note" id="wa-asof"></span></div>
  <div id="wa-peak" class="wa-peak" aria-live="polite"></div>
  <div class="wa-chart-legend"><span><i class="wa-line-key"></i>عدد حالات التقاطع</span><span><i class="wa-past-key"></i>أسبوع منقضٍ</span><span><i class="wa-current-key"></i>الأسبوع الحالي بتاريخ التصنيف</span></div>
  <div id="wa-chart" class="wa-chart-scroll"></div>
  <p class="wa-count-note">تُحسب الحالة مرة واحدة في كل أسبوع تمتد إليه. لذلك قد يتجاوز مجموع الأعداد الأسبوعية عدد الحالات الفريدة في التقرير. اضغط على نقطة لعرض الحالات.</p>
</section>
  <section class="panel register" aria-labelledby="register-title">
    <div class="section-header"><div class="table-heading"><h2 class="section-title" id="register-title">سجل حالات التقاطع</h2><span id="table-count" class="count-label"></span></div></div>
    <div class="table-controls">
      <div class="table-filters"><div class="field"><label for="table-threshold">نسبة التقاطع</label><select id="table-threshold"><option value="all">جميع النسب</option><option value="threshold" id="threshold-option">50% فأكثر فقط</option></select></div><div class="field"><label for="table-status">حالة الفترة</label><select id="table-status"><option value="all">جميع الحالات</option><option value="جارية">جارية</option><option value="قادمة">قادمة</option><option value="منتهية">منتهية</option></select></div></div>
      <p class="table-scope-note">فلتر النسبة والحالة يغيّر السجل فقط. المؤشرات والخريطة والشارت تشمل جميع النسب ضمن الفلاتر العامة.</p>
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
        <div class="sp-panel-top"><div class="sp-panel-label-row"><span>إجازات مجموعة المدير والموقع</span><button class="sp-close" type="button" data-action="close" aria-label="إغلاق تفاصيل المدير">×</button></div><h3 id="sp-manager-title" data-ui="manager-title"></h3><p class="sp-manager-role" data-ui="manager-role"></p><p class="sp-port" data-ui="manager-port"></p><p class="sp-org" data-ui="manager-org"></p><div class="sp-panel-note" data-ui="panel-note"></div><details class="sp-original-cases"><summary data-ui="case-summary"></summary><div data-ui="case-list"></div></details></div>
        <div class="sp-panel-content"><div class="sp-result" data-ui="result" role="status" hidden></div><h3 data-ui="employee-heading">الموظفون في التقاطع</h3><div data-ui="employees"></div><details class="sp-fixed" data-ui="fixed"><summary data-ui="fixed-label"></summary><div class="sp-fixed-list" data-ui="fixed-list"></div></details></div>
        <div class="sp-panel-footer"><button type="button" class="sp-solve" data-action="solve">تنسيق التقاطع</button><button type="button" class="sp-undo" data-action="undo" hidden>استعادة المواعيد الأصلية</button><p class="sp-rule" data-ui="rule">نفس مدة الإجازة، مع مراعاة إجازات باقي الموظفين.</p></div>
      </aside>
    </div>
    <footer class="sp-footer"><span data-ui="footer-period"></span><span>الجمعة والسبت ضمن مدة الإجازة. المقترحات داخل الصفحة ولا تعدّل ملفات Excel.</span></footer>
    <div class="sp-live" data-ui="live" aria-live="polite" aria-atomic="true"></div>
    <noscript>يلزم تفعيل JavaScript لعرض المنسّق الذكي.</noscript>
  </div>
</section>

  <details class="methodology" id="methodology"><summary>كيف تُحتسب الحالة والنسبة؟</summary><div class="method-body"><p>الحالة هي إجازة <strong>الموظفين أنفسهم، وعددهم اثنان فأكثر، في أيام تقويمية متصلة</strong>. عند تغيّر أي موظف في المجموعة المتقاطعة، أو انقطاع الأيام، تبدأ حالة جديدة.</p><p>تُحسب كل مجموعة بصورة مستقلة حسب <strong>المدير المباشر + موقع العمل + المستويات التنظيمية الأول والثاني والثالث</strong>. نسبة الغياب = عدد المجازين معاً ÷ إجمالي موظفي المجموعة نفسها × 100.</p><p id="method-dates"></p><p>تُعرض الترجمات المتوفرة في ملف الربط، وتبقى القيمة الأصلية ظاهرة عند غياب ترجمة معتمدة. الترجمة لا تغيّر مجموعات الحساب أو أعداد الموظفين. النسبة المعروضة مقربة إلى منزلة عشرية، ويُطبق حد النسبة على القيمة قبل التقريب.</p><div class="qa-tiles" id="qa-tiles"></div><details class="qa-issues" id="qa-issues" hidden><summary id="qa-summary">ملاحظات معالجة البيانات</summary><ul id="qa-list"></ul></details></div></details>
  <details class="methodology" id="mapping-details"><summary id="mapping-summary">الترجمات وربط المنافذ</summary><div class="method-body"><p id="mapping-status"></p><div class="mapping-statrow" id="mapping-stats"></div><p>عند غياب ترجمة أو وجود أكثر من مطابقة، يُحتفظ بالقيمة الأصلية. الإحداثيات غير المتاحة لا تُستبدل بمواقع تقديرية.</p><div class="mapping-table-wrap" id="mapping-issues-wrap" hidden><table><thead><tr><th scope="col">البيان</th><th scope="col">القيمة الأصلية</th><th scope="col">ملاحظة الربط</th></tr></thead><tbody id="mapping-issues-list"></tbody></table></div></div></details>
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
const roleOf = g => {const value=String(g?.title ?? '').trim();return !value || ['غير محدد','غير متاح'].includes(value) ? 'المسمى الوظيفي غير متوفر' : value;};
const portIdOf = g => String(g?.port_id || norm(g?.port_raw || g?.raw?.port || g?.port || ''));
const originalValue = (g,field) => g?.[field+'_raw'] ?? g?.raw?.[field] ?? g?.[field] ?? '';
const originalValues = g => [...Object.values(g?.raw || {}),...Object.keys(g||{}).filter(key=>key.endsWith('_raw')).map(key=>g[key])];
const portCatalog=new Map();
for(const g of DATA.groups){
  const id=portIdOf(g);if(!id)continue;
  if(!portCatalog.has(id))portCatalog.set(id,{id,name:g.port || originalValue(g,'port') || 'منفذ غير محدد',code:g.port_code || '',types:new Set(),terms:new Set()});
  const entry=portCatalog.get(id);entry.types.add(g.type || 'غير محدد');
  [g.port,originalValue(g,'port'),g.port_code].filter(Boolean).forEach(value=>entry.terms.add(String(value)));
}
let portQuery='',portHighlight=-1;
function scopeActive(){return Boolean($('manager-filter').value || $('title-filter').value || $('type-filter').value || $('port-filter').value || portQuery.trim() || $('group-search').value.trim());}
function groupMatches(g){
  if(!g)return false;
  if($('manager-filter').value && String(g.manager_id) !== $('manager-filter').value)return false;
  if($('title-filter').value && roleOf(g) !== $('title-filter').value)return false;
  if($('type-filter').value && (g.type || 'غير محدد') !== $('type-filter').value)return false;
  if($('port-filter').value && portIdOf(g) !== $('port-filter').value)return false;
  if(!$('port-filter').value && norm(portQuery) && !norm([g.port,originalValue(g,'port'),g.port_code].join(' ')).includes(norm(portQuery)))return false;
  const query=norm($('group-search').value);
  return !query || norm([g.manager,g.manager_id,roleOf(g),g.port,g.port_code,g.type,g.region,g.level1,g.level2,g.level3,...originalValues(g)].join(' ')).includes(query);
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
  $('highest-details').innerHTML=`<div class="highest-top"><div><div class="highest-name">${raw(g.manager)}</div><div class="highest-job">${raw(roleOf(g))}</div></div>${badge(c.status)}</div><div class="highest-meta"><div><span class="info-label">موقع العمل / المجموعة التنظيمية</span><div class="info-value">${raw(g.port)}</div>${pathHtml(g)}</div><div><span class="info-label">فترة التقاطع</span><div class="info-value">${esc(dateRange(c))}</div><div class="highest-job">${number(c.days)} أيام متصلة بنفس الموظفين</div></div></div><div class="highest-bottom"><p class="subtle">${ties.length>1?`تكررت أعلى نسبة في ${number(ties.length)} حالات؛ المعروضة هي الأسبق تاريخاً.`:'هذه أعلى نسبة بين جميع حالات التقاطع ضمن النطاق المعروض.'}</p><button type="button" class="link-button" ${ties.length>1?'id="show-highest-ties"':`data-case-detail="${esc(c.id)}"`}>${ties.length>1?`عرض جميع الحالات الأعلى (${number(ties.length)})`:'عرض الموظفين والتفاصيل'}</button></div>`;
  if(ties.length>1)$('show-highest-ties').onclick=()=>showHighest(ties);
}
function renderTable(){
  const rows=getTableCases(),scopeRows=getScopeCases();
  page=Math.min(page,Math.max(0,Math.ceil(rows.length/PAGE_SIZE)-1));
  const shown=printing?rows:rows.slice(page*PAGE_SIZE,(page+1)*PAGE_SIZE);
  $('table-count').textContent=number(rows.length)+' حالة';
  $('case-rows').innerHTML=shown.map(c=>{
    const g=groupOf(c);
    return `<tr class="${c.status==='منتهية'?'ended-row':''}"><td class="manager-cell"><div class="cell-title">${raw(g.manager)}</div><div class="cell-note">${raw(roleOf(g))}</div><span class="case-id">${esc(c.id)}</span></td><td class="port-cell"><div class="cell-title">${raw(g.port)}</div><div class="cell-note">${pathHtml(g)}</div></td><td class="people-cell"><span class="people-number" dir="ltr">${number(c.people)} / ${number(c.workforce)}</span><div class="cell-note">المتبقون: ${number(c.remaining)}</div></td><td class="ratio-cell"><span class="ratio-value ${c.flag?'over':''}" dir="ltr">${pct(c)}</span><div class="ratio-track" aria-hidden="true"><span class="${c.flag?'over':''}" style="width:${Math.max(0,Math.min(100,Number(c.percent)))}%"></span></div></td><td class="date-cell"><div>${esc(dateRange(c))}</div><div class="cell-note">${number(c.days)} أيام متصلة</div></td><td>${badge(c.status)}</td><td class="screen-detail-col"><button type="button" class="detail-button" data-case-detail="${esc(c.id)}" aria-label="تفاصيل الحالة ${esc(c.id)} للمدير ${esc(g.manager)}، ${esc(roleOf(g))}">التفاصيل</button></td></tr>`;
  }).join('') || '<tr><td colspan="7" class="empty-cell">لا توجد حالات مطابقة للتصفية الحالية.</td></tr>';
  $('page-info').textContent=rows.length ? `${number(page*PAGE_SIZE+1)}–${number(Math.min(rows.length,(page+1)*PAGE_SIZE))} من ${number(rows.length)} حالة` : 'لا توجد حالات لعرضها';
  $('previous-page').disabled=page===0;
  $('next-page').disabled=(page+1)*PAGE_SIZE>=rows.length;
  const parts=[];
  if($('manager-filter').value)parts.push('المدير: '+$('manager-filter').selectedOptions[0].textContent);
  if($('title-filter').value)parts.push('المسمى: '+$('title-filter').value);
  if($('type-filter').value)parts.push('النوع: '+$('type-filter').value);
  if($('port-filter').value)parts.push('المنفذ: '+(portCatalog.get($('port-filter').value)?.name || $('port-search').value));
  else if(portQuery.trim())parts.push('بحث المنفذ: '+portQuery.trim());
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
  if(typeof window.renderPortAnalytics==='function')window.renderPortAnalytics(rows,DATA.groups.filter(groupMatches));
  document.dispatchEvent(new CustomEvent('dashboard:scopechange'));
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
  const summary=`<div class="detail-summary"><div><span class="info-label">اسم المدير</span><strong>${raw(g.manager)}</strong><div class="cell-note">${raw(roleOf(g))}</div></div><div><span class="info-label">موقع العمل</span><strong>${raw(g.port)}</strong></div><div class="full"><span class="info-label">المستوى الأول / الثاني / الثالث</span>${pathHtml(g)}</div><div><span class="info-label">فترة التقاطع</span><strong>${esc(dateRange(c))}</strong><div class="cell-note">${number(c.days)} أيام متصلة بنفس الموظفين</div></div><div><span class="info-label">حالة الفترة بتاريخ ${esc(fmt(DATA.meta.as_of))}</span>${badge(c.status)}<div class="cell-note">${esc(c.action||'')}</div></div></div><div class="detail-metrics"><span>المجازون: <b>${number(c.people)}</b></span><span>إجمالي المجموعة: <b>${number(c.workforce)}</b></span><span>نسبة الغياب: <b dir="ltr">${pct(c)}</b></span><span>المتبقون: <b>${number(c.remaining)}</b></span></div>`;
  const list=employees.map((e,i)=>`<li class="employee-item"><span class="employee-index">${i+1}</span><div><div class="employee-name">${raw(e.name||'اسم غير متوفر في المصدر')}<span class="employee-id">الرقم الوظيفي: <bdi>${esc(e.id)}</bdi></span></div>${(e.leave_periods||[]).length?(e.leave_periods||[]).map(p=>`<div class="leave-period">من ${esc(fmt(p.start))} إلى ${esc(fmt(p.end))}</div>`).join(''):'<div class="leave-period">تواريخ الإجازة الأصلية غير متاحة.</div>'}</div></li>`).join('');
  openDialog('تفاصيل حالة التقاطع · '+c.id,summary+originalLabels(g)+`<h3 class="employee-title">الموظفون المشاركون في الحالة (${number(c.people)})</h3><p class="employee-note">الفترات أدناه هي تواريخ الإجازات الأصلية في المصدر، وقد تكون أطول من فترة التقاطع.</p><ol class="employee-list">${list||'<li class="subtle">تفاصيل الموظفين غير متاحة.</li>'}</ol>`);
}
function showHighest(ties){
  openDialog('الحالات ذات أعلى نسبة غياب',`<p class="tie-intro">${number(ties.length)} حالات تشترك في أعلى نسبة (${pct(ties[0])}) ضمن نطاق المجموعة المعروض. المقارنة على النسبة قبل التقريب.</p>`+ties.map(c=>{const g=groupOf(c);return `<article class="tie-card"><div><div class="cell-title">${raw(g.manager)}</div><div class="cell-note">${raw(roleOf(g))}</div><div class="cell-note">${raw(g.port)}</div><div>${pathHtml(g)}</div><div class="cell-note">${esc(dateRange(c))} · ${number(c.people)} من ${number(c.workforce)} · ${number(c.days)} أيام</div><div style="margin-top:7px">${badge(c.status)}</div></div><button type="button" class="detail-button" data-case-detail="${esc(c.id)}">عرض الموظفين</button></article>`}).join(''));
}
function portChoices(){
  const query=norm(portQuery);
  const choices=[...portCatalog.values()].filter(entry=>!query || norm([...entry.terms].join(' ')).includes(query))
    .sort((a,b)=>String(a.name).localeCompare(String(b.name),'ar') || a.id.localeCompare(b.id));
  if(!query)choices.unshift({id:'',name:'جميع المنافذ',code:'',types:new Set(),terms:new Set()});
  return choices;
}
function closePortOptions(){
  $('port-options').hidden=true;$('port-search').setAttribute('aria-expanded','false');$('port-toggle').setAttribute('aria-expanded','false');$('port-search').removeAttribute('aria-activedescendant');portHighlight=-1;
}
function updatePortHelp(){
  const id=$('port-filter').value;
  $('port-clear').hidden=!id && !portQuery;
  $('port-help').textContent=id ? 'منفذ محدد. اضغطي × لعرض جميع المنافذ.' : portQuery.trim() ? (portChoices().length ? number(portChoices().length)+' منافذ مطابقة للبحث؛ اختاري واحدًا لتحديده.' : 'لا يوجد منفذ يطابق البحث؛ امسحي النص أو غيّريه.') : 'البحث بالاسم العربي أو الأصلي أو رمز المنفذ.';
}
function renderPortOptions(open=true){
  const options=portChoices();
  if(portHighlight>=options.length)portHighlight=options.length-1;
  $('port-options').innerHTML=options.map((entry,index)=>{
    const original=[...entry.terms].filter(value=>value!==entry.name && value!==entry.code).join(' / ');
    const detail=[...entry.types].join(' / ')+(entry.code?' · '+entry.code:'')+(original?' · '+original:'');
    return `<div role="option" class="port-option" id="port-choice-${index}" aria-selected="${index===portHighlight}" data-port-choice="${esc(entry.id)}"><strong>${raw(entry.name)}</strong>${detail?`<small>${raw(detail)}</small>`:''}</div>`;
  }).join('') || '<div class="port-no-results" role="status">لا توجد منافذ مطابقة.</div>';
  $('port-options').hidden=!open;$('port-search').setAttribute('aria-expanded',String(open));$('port-toggle').setAttribute('aria-expanded',String(open));
  if(open && portHighlight>=0 && options[portHighlight]){
    $('port-search').setAttribute('aria-activedescendant','port-choice-'+portHighlight);
    $('port-choice-'+portHighlight)?.scrollIntoView({block:'nearest'});
  }else $('port-search').removeAttribute('aria-activedescendant');
  updatePortHelp();
}
function selectPort(id,focus=false){
  id=String(id ?? '');if(id && !portCatalog.has(id))return false;
  $('port-filter').value=id;$('port-search').value=id?portCatalog.get(id).name:'';portQuery='';
  if(focus)$('port-search').focus({preventScroll:true});
  closePortOptions();updatePortHelp();page=0;render();return true;
}
window.selectPortAnalytics=portId=>selectPort(portId,false);
function initializePortSearch(){
  $('port-search').addEventListener('focus',()=>{portHighlight=-1;renderPortOptions();});
  $('port-search').addEventListener('input',()=>{
    $('port-filter').value='';portQuery=$('port-search').value;portHighlight=-1;renderPortOptions();page=0;render();
  });
  $('port-search').addEventListener('keydown',event=>{
    if(event.key==='Escape'){if(!$('port-options').hidden){event.preventDefault();event.stopPropagation();closePortOptions();}return;}
    if(event.key==='Tab'){closePortOptions();return;}
    if(event.key==='ArrowDown' || event.key==='ArrowUp'){
      event.preventDefault();const options=portChoices();if(!options.length){renderPortOptions();return;}
      portHighlight=event.key==='ArrowDown' ? Math.min(options.length-1,portHighlight+1) : (portHighlight<0?options.length-1:Math.max(0,portHighlight-1));renderPortOptions();return;
    }
    if(event.key==='Enter' && !$('port-options').hidden){
      const options=portChoices();const exact=options.filter(entry=>entry.name===$('port-search').value || (entry.code && entry.code===$('port-search').value));
      const chosen=portHighlight>=0?options[portHighlight]:exact.length===1?exact[0]:options.length===1?options[0]:null;
      if(chosen){event.preventDefault();selectPort(chosen.id,true);}else closePortOptions();
    }
  });
  $('port-options').addEventListener('pointerdown',event=>{if(event.target.closest('[data-port-choice]'))event.preventDefault();});
  $('port-options').addEventListener('click',event=>{const option=event.target.closest('[data-port-choice]');if(option)selectPort(option.dataset.portChoice,true);});
  $('port-toggle').addEventListener('pointerdown',event=>event.preventDefault());
  $('port-toggle').addEventListener('click',()=>{const wasOpen=!$('port-options').hidden;if(wasOpen)closePortOptions();else{$('port-search').focus({preventScroll:true});portHighlight=-1;renderPortOptions();}});
  $('port-clear').addEventListener('click',()=>selectPort('',true));
  $('port-combobox').addEventListener('focusout',()=>setTimeout(()=>{if(!$('port-combobox').contains(document.activeElement))closePortOptions();},0));
  $('port-filter').addEventListener('change',()=>selectPort($('port-filter').value,false));
  updatePortHelp();
}
function renderMappingDiagnostics(){
  const mapping=DATA.mapping || {},issues=Array.isArray(mapping.issues)?mapping.issues:[];
  const labels={manager:'اسم المدير',title:'المسمى الوظيفي',port:'اسم المنفذ',level1:'المستوى الأول',level2:'المستوى الثاني',level3:'المستوى الثالث',type:'نوع المنفذ',region:'المنطقة الإدارية',coordinates:'الإحداثيات',mapping:'ملف الربط',port_code:'رمز المنفذ',name:'اسم الموظف',employee:'اسم الموظف'};
  $('mapping-summary').textContent='الترجمات وربط المنافذ'+(issues.length?' · '+number(issues.length)+' ملاحظات':'');
  $('mapping-status').textContent=mapping.loaded ? 'مصدر الترجمة: '+String(mapping.file_name || 'ملف الربط')+(Array.isArray(mapping.sheets)&&mapping.sheets.length?' · الأوراق: '+mapping.sheets.join('، '):'') : 'لم يُحمّل ملف الربط؛ تظهر الأسماء والمسميات المتاحة من المصادر.';
  $('mapping-stats').innerHTML=[['قيم مترجمة',mapping.translated_count],['ملاحظات ربط',mapping.issue_count],['عدد المنافذ',mapping.total_ports],['منافذ بإحداثيات',mapping.located_ports]].filter(([,value])=>value!==undefined && value!==null && Number.isFinite(Number(value))).map(([label,value])=>`<span><b>${number(value)}</b>${esc(label)}</span>`).join('');
  $('mapping-issues-wrap').hidden=!issues.length;
  $('mapping-issues-list').innerHTML=issues.map(item=>`<tr><td>${esc(labels[item.field] || item.field || 'بيان الربط')}</td><td>${raw(item.source || 'غير متاح')}</td><td>${esc(item.issue || 'تحتاج مراجعة')}</td></tr>`).join('');
}
function originalLabels(g){
  const fields=[['manager','اسم المدير'],['title','المسمى الوظيفي'],['port','اسم المنفذ'],['level1','المستوى الأول'],['level2','المستوى الثاني'],['level3','المستوى الثالث'],['region','المنطقة الإدارية'],['type','نوع المنفذ']];
  const changed=fields.filter(([field])=>String(originalValue(g,field)||'') && String(originalValue(g,field))!==String(g[field]??''));
  if(!changed.length)return '';
  return `<details class="mapping-originals"><summary>القيم الأصلية قبل الترجمة</summary><dl>${changed.map(([field,label])=>`<dt>${esc(label)}</dt><dd>${raw(originalValue(g,field))}</dd>`).join('')}</dl></details>`;
}
function resetFilters(){
  $('manager-filter').value='';$('title-filter').value='';$('type-filter').value='';$('port-filter').value='';$('port-search').value='';portQuery='';portHighlight=-1;closePortOptions();updatePortHelp();$('group-search').value='';$('table-threshold').value='all';$('table-status').value='all';page=0;render();
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
  const managers=new Map();for(const g of DATA.groups){const id=String(g.manager_id);if(!managers.has(id))managers.set(id,{name:g.manager,roles:new Set()});managers.get(id).roles.add(roleOf(g));}
  $('manager-filter').innerHTML='<option value="">جميع المدراء</option>'+[...managers].sort((a,b)=>String(a[1].name).localeCompare(String(b[1].name),'ar')||a[0].localeCompare(b[0])).map(([id,entry])=>`<option value="${esc(id)}">${esc(entry.name)} — ${esc([...entry.roles].join(' / '))} · ${esc(id)}</option>`).join('');
  $('title-filter').innerHTML='<option value="">جميع المسميات الوظيفية</option>'+[...new Set(DATA.groups.map(roleOf))].sort((a,b)=>a.localeCompare(b,'ar')).map(title=>`<option value="${esc(title)}">${esc(title)}</option>`).join('');
  initializePortSearch();renderMappingDiagnostics();
  const qa=DATA.qa||{};
  $('qa-tiles').innerHTML=[['موظفون ضمن المجموعات',qa.employees],['مجموعات تنظيمية',qa.groups],['سجلات إجازات مطابقة',qa.matched_leaves]].filter(([,v])=>v!==undefined).map(([label,value])=>`<span><b>${number(value)}</b>${esc(label)}</span>`).join('');
  const problems=(qa.problems||[]).filter(p=>Number(p.count)>0);
  $('qa-issues').hidden=!problems.length;
  $('qa-summary').textContent='ملاحظات معالجة البيانات ('+number(problems.length)+' أنواع)';
  $('qa-list').innerHTML=problems.map(p=>`<li>${esc(p.label)}: <b>${number(p.count)}</b></li>`).join('');
  for(const id of ['manager-filter','title-filter','type-filter'])$(id).addEventListener('change',()=>{page=0;render();});
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
window.Dashboard={render,renderTable,getScopeCases,getTableCases,getHighestCases,showCase,resetFilters,groupMatches,roleOf,portIdOf,getPage:()=>page,getData:()=>DATA};
/* Offline analytics. Count original case IDs; never count employee pairs. */
(function(root){
  'use strict';
  const DAY=86400000;
  const hasDOM=typeof document!=='undefined';
  const ui=id=>hasDOM?document.getElementById(id):null;
  const escapeHTML=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const num=value=>new Intl.NumberFormat('en-US',{maximumFractionDigits:1}).format(Number(value)||0);
  const text=value=>String(value??'').trim();
  const dateNumber=value=>{
    const s=text(value);if(!/^\d{4}-\d{2}-\d{2}$/.test(s))return NaN;
    const stamp=Date.parse(s+'T00:00:00Z');
    return Number.isFinite(stamp)&&new Date(stamp).toISOString().slice(0,10)===s?stamp/DAY:NaN;
  };
  const iso=value=>new Date(value*DAY).toISOString().slice(0,10);
  const dateLabel=value=>Number.isFinite(dateNumber(value))?new Intl.DateTimeFormat('ar-SA-u-ca-gregory-nu-latn',{day:'numeric',month:'short',year:'numeric',timeZone:'UTC'}).format(new Date(dateNumber(value)*DAY)):'تاريخ غير متاح';
  const rangeLabel=(start,end)=>dateLabel(start)+' — '+dateLabel(end);
  const roleOf=g=>text(g?.title)||'المسمى الوظيفي غير متوفر في المصدر';
  const managerOf=g=>text(g?.manager)||'اسم المدير غير متوفر في المصدر';
  const portKey=g=>text(g?.port_id)||text(g?.port_code)||text(g?.port_source)||text(g?.port)||'group:'+String(g?.id??'');
  const managerKey=g=>text(g?.manager_id)||managerOf(g);
  const types={'جوي':{color:'#123752',label:'منفذ جوي'},'بري':{color:'#156342',label:'منفذ بري'},'بحري':{color:'#087f83',label:'منفذ بحري'},'غير محدد':{color:'#74827d',label:'نوع غير محدد'}};
  const iconPaths={
    'جوي':'<path d="M12 2l2.5 7L22 12v2l-7.5-1L14 19l3 2v1l-5-1-5 1v-1l3-2-.5-6L2 14v-2l7.5-3L12 2Z" fill="currentColor" stroke="none"/>',
    'بري':'<path d="M5 8l2-4h10l2 4m-15 0h16v10H4V8Zm3 10v2m10-2v2M5 12h14M7 15h1m8 0h1" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round" stroke-linecap="round"/>',
    'بحري':'<circle cx="12" cy="4" r="2.4" fill="none" stroke="currentColor" stroke-width="1.8"/><path d="M12 6.5V21M7 10h10M3 14v3c2 6 16 6 18 0v-3m-18 0 3 2m15-2-3 2" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round" stroke-linecap="round"/>',
    'غير محدد':'<circle cx="12" cy="12" r="5" fill="none" stroke="currentColor" stroke-width="2"/>'
  };
  const kind=value=>Object.prototype.hasOwnProperty.call(types,value)?value:'غير محدد';
  const typeIcon=(value,extra='')=>{const k=kind(value);return `<span class="pa-type-icon ${extra}" style="--pa-icon-color:${types[k].color}" aria-hidden="true"><svg viewBox="0 0 24 24">${iconPaths[k]}</svg></span>`;};
  function mapPoint(longitude,latitude){
    const my=Math.log(Math.tan(Math.PI/4+latitude*Math.PI/360))*180/Math.PI;
    return [185.8411075183024+(longitude-34.4944055)*39.12272444762573,60.0+(33.988539793063666-my)*39.12272444762573];
  }
  function finiteCoordinate(value,limit){
    if(value===null||value===undefined||typeof value==='boolean'||text(value)==='')return null;
    const n=Number(value);return Number.isFinite(n)&&Math.abs(n)<=limit?n:null;
  }
  function readCoordinates(g){
    const latitude=finiteCoordinate(g?.latitude,90),longitude=finiteCoordinate(g?.longitude,180);
    if(latitude===null||longitude===null||Math.abs(latitude)>=90)return null;
    const [x,y]=mapPoint(longitude,latitude);
    if(!Number.isFinite(x)||!Number.isFinite(y))return null;
    return {latitude,longitude,x,y,inView:x>=0&&x<=1200&&y>=0&&y<=800};
  }
  function uniqueCases(rows){
    const seen=new Set();return (Array.isArray(rows)?rows:[]).filter(c=>{
      const id=text(c?.id);if(!id||seen.has(id))return false;seen.add(id);return true;
    });
  }
  function weeklyRows(rows,start,end,asof){
    const first=dateNumber(start),last=dateNumber(end),ref=dateNumber(asof);
    if(!Number.isFinite(first)||!Number.isFinite(last)||first>last)return [];
    const cases=uniqueCases(rows).map(c=>({c,s:dateNumber(c.start),e:dateNumber(c.end)})).filter(o=>Number.isFinite(o.s)&&Number.isFinite(o.e)&&o.s<=o.e);
    const out=[];let cursor=first;
    while(cursor<=last){
      const stop=Math.min(last,cursor+(6-new Date(cursor*DAY).getUTCDay()));
      const members=cases.filter(o=>o.s<=stop&&o.e>=cursor).map(o=>o.c);
      out.push({week:out.length+1,start:iso(cursor),end:iso(stop),count:members.length,case_ids:members.map(c=>String(c.id)),cases:members,phase:Number.isFinite(ref)?stop<ref?'past':cursor>ref?'future':'current':'unknown'});
      cursor=stop+1;
    }
    return out;
  }
  function peakWeeks(weeks){
    const maximum=Math.max(0,...(weeks||[]).map(w=>Number(w.count)||0));
    return {maximum,weeks:maximum>0?weeks.filter(w=>w.count===maximum):[]};
  }
  function aggregatePorts(rows,groups){
    const allGroups=Array.isArray(groups)?groups:groups instanceof Map?[...groups.values()]:groups instanceof Set?[...groups]:[];
    const byGroup=new Map(allGroups.filter(g=>g&&typeof g==='object').map(g=>[String(g.id),g]));
    const ports=new Map();
    byGroup.forEach(g=>{
      const id=portKey(g);
      if(!ports.has(id))ports.set(id,{id,name:text(g.port)||text(g.port_source)||'اسم المنفذ غير متوفر في المصدر',code:text(g.port_code),type:kind(g.type),region:text(g.region)||'المنطقة غير متوفرة',groups:[],cases:[],case_ids:new Set(),manager_ids:new Set(),coordinates:null,coordinateIssue:''});
      const p=ports.get(id);p.groups.push(g);p.manager_ids.add(managerKey(g));
    });
    uniqueCases(rows).forEach(c=>{
      const g=byGroup.get(String(c.group_id));if(!g)return;
      const p=ports.get(portKey(g));p.cases.push(c);p.case_ids.add(String(c.id));
    });
    ports.forEach(p=>{
      p.count=p.case_ids.size;p.managerCount=p.manager_ids.size;
      const coordinates=p.groups.map(readCoordinates).filter(Boolean);
      const unique=new Map(coordinates.map(c=>[c.latitude.toFixed(5)+','+c.longitude.toFixed(5),c]));
      if(unique.size>1)p.coordinateIssue='إحداثيات متعددة تحتاج مراجعة الربط';
      else if(unique.size===0)p.coordinateIssue='إحداثيات غير متاحة أو غير صالحة';
      else{
        const c=[...unique.values()][0];
        if(c.inView)p.coordinates=c;else p.coordinateIssue='الإحداثيات خارج نطاق الخريطة';
      }
      const distinctTypes=new Set(p.groups.map(g=>kind(g.type)));
      if(distinctTypes.size>1){p.type='غير محدد';p.typeIssue='نوع المنفذ غير متفق في الربط';}
    });
    return [...ports.values()].sort((a,b)=>b.count-a.count||a.name.localeCompare(b.name,'ar')||a.id.localeCompare(b.id));
  }
  function layoutMarkers(ports){
    const points=ports.filter(p=>p.coordinates).map(p=>({port:p,anchorX:p.coordinates.x,anchorY:p.coordinates.y,x:p.coordinates.x,y:p.coordinates.y}));
    const visited=new Set(),clusters=[];
    for(let i=0;i<points.length;i++){
      if(visited.has(i))continue;const cluster=[],queue=[i];visited.add(i);
      while(queue.length){const j=queue.shift();cluster.push(j);for(let k=0;k<points.length;k++)if(!visited.has(k)&&Math.hypot(points[j].anchorX-points[k].anchorX,points[j].anchorY-points[k].anchorY)<58){visited.add(k);queue.push(k);}}
      clusters.push(cluster);
    }
    clusters.forEach(indices=>{
      if(indices.length<2)return;
      const cx=indices.reduce((n,i)=>n+points[i].anchorX,0)/indices.length,cy=indices.reduce((n,i)=>n+points[i].anchorY,0)/indices.length,radius=Math.max(40,indices.length*60/(Math.PI*2));
      indices.sort((a,b)=>points[a].port.id.localeCompare(points[b].port.id));
      indices.forEach((j,i)=>{const angle=-Math.PI/2+i*2*Math.PI/indices.length;points[j].x=cx+radius*Math.cos(angle);points[j].y=cy+radius*Math.sin(angle);});
    });
    const placed=[];
    points.forEach(p=>{
      const cx=p.x,cy=p.y;
      const clamp=(v,lo,hi)=>Math.max(lo,Math.min(hi,v));
      p.x=clamp(p.x,32,1168);p.y=clamp(p.y,32,748);
      for(let attempt=0;attempt<600&&placed.some(q=>Math.hypot(q.x-p.x,q.y-p.y)<58);attempt++){
        const angle=attempt*2.399963229728653,radius=12*Math.sqrt(attempt+1);
        p.x=clamp(cx+radius*Math.cos(angle),32,1168);p.y=clamp(cy+radius*Math.sin(angle),32,748);
      }
      p.shifted=Math.hypot(p.x-p.anchorX,p.y-p.anchorY)>2;placed.push(p);
    });
    return points;
  }
  const model={dateNumber,weeklyRows,peakWeeks,aggregatePorts,layoutMarkers,readCoordinates,mapPoint,uniqueCases};
  root.PortAnalyticsModel=model;
  if(typeof module==='object'&&module.exports)module.exports=model;
  if(!hasDOM)return;
  let state={ports:[],weeks:[],groups:new Map(),rows:[],meta:{}};
  let bound=false,mapZoom=1,mapX=0,mapY=0,drag=null;
  function dashboardData(){return typeof DATA!=='undefined'?DATA:root.DATA||{groups:[],cases:[],meta:{}};}
  function dialog(title,html){
    if(typeof root.openDialog==='function'){root.openDialog(title,html);return;}
    let d=ui('pa-details-dialog');
    if(!d){d=document.createElement('dialog');d.id='pa-details-dialog';d.style.cssText='max-width:760px;width:94vw;max-height:88vh;overflow:auto;background:#fffefa;border:1px solid #dce2e0;border-radius:10px;padding:24px';document.body.appendChild(d);}
    d.innerHTML=`<div class="pa-manager-line"><h2>${escapeHTML(title)}</h2><button type="button" data-pa-close aria-label="إغلاق">×</button></div>${html}`;
    if(!d.open){if(typeof d.showModal==='function')d.showModal();else d.setAttribute('open','');}
  }
  function closeAnalyticsDialog(){
    if(typeof root.closeDialog==='function')root.closeDialog();
    const d=ui('pa-details-dialog');if(d?.open){if(typeof d.close==='function')d.close();else d.removeAttribute('open');}
  }
  function showCaseDetails(id){
    if(typeof root.showCase==='function'){root.showCase(String(id));return;}
    const c=state.rows.find(item=>String(item.id)===String(id));if(!c)return;
    const g=state.groups.get(String(c.group_id));
    dialog('تفاصيل حالة التقاطع',`<strong class="pa-manager-name">${escapeHTML(managerOf(g))}</strong><span class="pa-manager-role">${escapeHTML(roleOf(g))}</span><p>${escapeHTML(g?.port||'')}</p><p>${escapeHTML(rangeLabel(c.start,c.end))}</p><p>${num(c.people)} موظفين · ${num(c.percent)}%</p>`);
  }
  function portDetails(index){
    const p=state.ports[index];if(!p)return;
    const caseGroups=new Map();p.cases.forEach(c=>{const id=String(c.group_id);if(!caseGroups.has(id))caseGroups.set(id,[]);caseGroups.get(id).push(c);});
    const groups=p.groups.filter(g=>caseGroups.has(String(g.id))||!p.count);
    const cards=groups.map(g=>{
      const cases=caseGroups.get(String(g.id))||[];
      return `<article class="pa-manager-card"><div class="pa-manager-line"><div><strong class="pa-manager-name"><bdi>${escapeHTML(managerOf(g))}</bdi></strong><span class="pa-manager-role"><bdi>${escapeHTML(roleOf(g))}</bdi></span></div><span class="pa-manager-case-count">${num(cases.length)} حالات تقاطع</span></div>${cases.length?`<div class="pa-case-buttons">${cases.map(c=>`<button type="button" data-pa-case="${escapeHTML(c.id)}">${escapeHTML(rangeLabel(c.start,c.end))}</button>`).join('')}</div>`:'<p class="pa-manager-meta">لا توجد حالات تقاطع ضمن النطاق المحدد.</p>'}</article>`;
    }).join('');
    dialog('التقاطعات في '+p.name,`<div class="pa-dialog-meta"><span>${num(p.count)} حالات فريدة</span><span>عدد المدراء: ${num(p.managerCount)}</span><span>${escapeHTML(types[p.type].label)}</span><span>${escapeHTML(p.region)}</span>${p.code?`<span>رمز المنفذ: <bdi>${escapeHTML(p.code)}</bdi></span>`:''}</div>${p.coordinateIssue?`<p class="pa-location-notice">${escapeHTML(p.coordinateIssue)}. يظهر المنفذ في الأعداد والقائمة، ولا يوضع له موقع افتراضي على الخريطة.</p>`:''}${p.typeIssue?`<p class="pa-location-notice">${escapeHTML(p.typeIssue)}</p>`:''}${cards||'<p class="pa-dialog-empty">لا توجد حالات تقاطع ضمن الفلاتر المحددة.</p>'}<button type="button" class="pa-filter-button" data-pa-filter="${escapeHTML(p.id)}">تصفية الداشبورد على هذا المنفذ</button>`);
  }
  function weekDetails(weekNumber){
    const w=state.weeks.find(v=>v.week===Number(weekNumber));if(!w)return;
    const ordered=w.cases.slice().sort((a,b)=>String(a.start).localeCompare(String(b.start))||String(a.id).localeCompare(String(b.id)));
    const cards=ordered.map(c=>{
      const g=state.groups.get(String(c.group_id));
      const within=Math.max(0,Math.min(dateNumber(c.end),dateNumber(w.end))-Math.max(dateNumber(c.start),dateNumber(w.start))+1);
      return `<article class="wa-case-card"><div class="wa-case-heading"><div><strong class="pa-manager-name"><bdi>${escapeHTML(managerOf(g))}</bdi></strong><span class="pa-manager-role"><bdi>${escapeHTML(roleOf(g))}</bdi></span><div class="pa-manager-meta"><bdi>${escapeHTML(g?.port||'اسم المنفذ غير متوفر في المصدر')}</bdi></div></div><span class="wa-case-count">الحالة ${escapeHTML(c.id)}</span></div><p class="wa-case-dates">الفترة الكاملة: ${escapeHTML(rangeLabel(c.start,c.end))}<br>${num(within)} أيام من الحالة تقع ضمن هذا الأسبوع · ${num(c.people)} موظفين · نسبة الغياب ${num(c.percent)}%</p><button type="button" data-pa-case="${escapeHTML(c.id)}">عرض موظفي الحالة</button></article>`;
    }).join('');
    dialog('حالات التقاطع · الأسبوع '+num(w.week),`<div class="pa-dialog-meta"><span>${escapeHTML(rangeLabel(w.start,w.end))}</span><span>${num(w.count)} حالات فريدة خلال الأسبوع</span></div>${cards||'<p class="pa-dialog-empty">لا توجد حالات تقاطع في هذا الأسبوع ضمن الفلاتر المحددة.</p>'}<p class="wa-details-note">كل بطاقة تمثل حالة التقرير نفسها، وليست عدد موظفين أو عدد أزواج متقاطعة. الحالة الممتدة إلى أسبوع آخر تظهر مرة واحدة في ذلك الأسبوع أيضًا.</p>`);
  }
  function renderMap(){
    const ports=state.ports,mapped=ports.filter(p=>p.coordinates),missing=ports.filter(p=>!p.coordinates),withCases=ports.filter(p=>p.count>0);
    ui('pa-stats').innerHTML=`<span class="pa-stat"><b>${num(state.rows.length)}</b>حالات تقاطع</span><span class="pa-stat"><b>${num(withCases.length)}</b>منافذ بها تقاطع</span><span class="pa-stat"><b>${num(mapped.length)}</b>منافذ على الخريطة</span>${missing.length?`<span class="pa-stat pa-stat-warning"><b>${num(missing.length)}</b>منافذ تحتاج مراجعة الإحداثيات</span>`:''}`;
    ui('pa-list-count').textContent=num(ports.length)+' منافذ ضمن النطاق';
    ui('pa-port-list').innerHTML=ports.map((p,index)=>`<button type="button" class="pa-port-button" data-pa-port="${index}" aria-label="${escapeHTML(p.name)}، ${num(p.count)} حالات تقاطع">${typeIcon(p.type)}<span class="pa-port-main"><b class="pa-port-name"><bdi>${escapeHTML(p.name)}</bdi></b><span class="pa-port-description">${escapeHTML(p.region)} · عدد المدراء: ${num(p.managerCount)}</span>${p.coordinateIssue?`<span class="pa-coordinate-tag">${escapeHTML(p.coordinateIssue)}</span>`:''}</span><span class="pa-port-count"><b>${num(p.count)}</b><small>حالات</small></span></button>`).join('')||'<div class="pa-empty">لا توجد منافذ مطابقة للفلاتر المحددة.</div>';
    const markers=layoutMarkers(ports);
    ui('pa-markers').innerHTML=markers.map(p=>{
      const index=ports.indexOf(p.port),t=types[p.port.type],x=p.x.toFixed(2),y=p.y.toFixed(2),ax=p.anchorX.toFixed(2),ay=p.anchorY.toFixed(2),count=num(p.port.count),badgeWidth=Math.max(25,12+count.length*8);
      return `<g class="pa-marker" tabindex="0" role="button" data-pa-port="${index}" aria-label="${escapeHTML(p.port.name)}، ${count} حالات تقاطع">${p.shifted?`<path d="M${ax} ${ay}L${x} ${y}" fill="none" stroke="${t.color}" stroke-width="1.1" stroke-opacity=".6" vector-effect="non-scaling-stroke"/><circle cx="${ax}" cy="${ay}" r="2.8" fill="${t.color}" stroke="#fffefa" stroke-width="1"/>`:''}<g transform="translate(${x} ${y})"><circle class="pa-marker-hit" r="27" fill="transparent"/><circle class="pa-pin-head" r="17.5" fill="${t.color}" stroke="#fffefa" stroke-width="2.5"/><svg x="-10" y="-10" width="20" height="20" viewBox="0 0 24 24" style="color:#fffefa">${iconPaths[p.port.type]}</svg><rect x="${-badgeWidth/2}" y="19" width="${badgeWidth}" height="18" rx="5" fill="#fffefa" stroke="${t.color}" stroke-width="1"/><text x="0" y="32.5" text-anchor="middle" font-size="12" font-weight="700" font-family="Tahoma,Arial,sans-serif" fill="${t.color}">${count}</text></g><title>${escapeHTML(p.port.name)} · ${count} حالات تقاطع</title></g>`;
    }).join('');
    ui('pa-map-legend').innerHTML=['جوي','بري','بحري'].map(t=>`<span>${typeIcon(t)}${types[t].label}</span>`).join('')+(ports.some(p=>p.type==='غير محدد')?`<span>${typeIcon('غير محدد')}نوع غير محدد</span>`:'');
    const notice=ui('pa-location-notice');notice.hidden=!missing.length;
    notice.innerHTML=missing.length?`<strong>${num(missing.length)} منافذ لا تظهر كنقاط على الخريطة.</strong> أعداد حالاتها محسوبة وتفاصيلها متاحة في القائمة.${missing.map(p=>`<button type="button" data-pa-port="${ports.indexOf(p)}"><bdi>${escapeHTML(p.name)}</bdi> · ${num(p.count)}</button>`).join('')}`:'';
    updateMap();
  }
  function renderWeekly(){
    const weeks=state.weeks,{maximum,weeks:peaks}=peakWeeks(weeks),peak=ui('wa-peak');
    ui('wa-asof').textContent='تاريخ التصنيف: '+dateLabel(state.meta.as_of);
    peak.classList.toggle('wa-no-data',!maximum);
    peak.innerHTML=maximum?`<div class="wa-peak-heading"><span>أعلى عدد أسبوعي</span><strong>${num(maximum)}<small>حالات تقاطع</small></strong></div><div class="wa-peak-weeks">${peaks.length>1?`<span class="wa-peak-ties">${num(peaks.length)} أسابيع متساوية في أعلى عدد</span>`:''}${peaks.map(w=>`<button type="button" data-wa-week="${w.week}"><b>الأسبوع ${num(w.week)}</b><small>${escapeHTML(rangeLabel(w.start,w.end))}</small></button>`).join('')}</div>`:'لا توجد حالات تقاطع في أسابيع فترة التقرير ضمن الفلاتر المحددة.';
    if(!weeks.length){ui('wa-chart').innerHTML='<div class="pa-empty">تعذر قراءة فترة التقرير لعرض الأسابيع.</div>';return;}
    const width=1180,height=365,left=57,right=1098,top=65,bottom=265,scale=Math.max(maximum,1),step=weeks.length>1?(right-left)/(weeks.length-1):150;
    const x=i=>weeks.length>1?right-i*step:(left+right)/2,y=count=>bottom-count/scale*(bottom-top);
    const ticks=[...new Set([0,Math.ceil(scale/4),Math.ceil(scale/2),Math.ceil(3*scale/4),scale])].sort((a,b)=>a-b);
    let h=`<svg viewBox="0 0 ${width} ${height}" role="group" aria-label="عدد حالات التقاطع حسب الأسبوع؛ يبدأ الأسبوع الأول من اليمين. ${escapeHTML(maximum?'أعلى عدد '+num(maximum)+' حالات':'لا توجد حالات')}"><rect width="${width}" height="${height}" fill="#fffefa"/>`;
    weeks.forEach((w,i)=>{
      const begin=Math.max(12,x(i)-step/2),finish=Math.min(1140,x(i)+step/2);
      h+=`<rect x="${begin}" y="16" width="${finish-begin}" height="316" fill="${w.phase==='past'?'#edecea':w.phase==='current'?'#e3f1e9':'#fffefa'}"/>`;
      if(w.phase==='current')h+=`<line x1="${x(i)}" x2="${x(i)}" y1="45" y2="${bottom+13}" stroke="#27604c" stroke-width="1" stroke-dasharray="4 4"/><rect x="${x(i)-34}" y="17" width="68" height="22" rx="11" fill="#123752"/><text x="${x(i)}" y="31.5" text-anchor="middle" font-size="10" fill="#fffefa">الأسبوع الحالي</text>`;
    });
    ticks.forEach(tick=>{h+=`<line x1="${left-28}" x2="${right+28}" y1="${y(tick)}" y2="${y(tick)}" stroke="#d9e0d7" stroke-width="1"/><text x="1154" y="${y(tick)+4}" font-size="12" fill="#65756c" text-anchor="middle">${num(tick)}</text>`;});
    weeks.forEach((w,i)=>{if(i){const previous=weeks[i-1];h+=`<path d="M${x(i-1)} ${y(previous.count)}L${x(i)} ${y(w.count)}" stroke="${previous.phase==='past'&&w.phase==='past'?'#a3aaa4':'#087f83'}" stroke-width="3" fill="none" stroke-linejoin="round" stroke-linecap="round"/>`;}});
    weeks.forEach((w,i)=>{
      const xx=x(i),yy=y(w.count),past=w.phase==='past',isPeak=maximum>0&&w.count===maximum,col=past?'#929b95':'#087f83',period=w.phase==='past'?'منقضٍ':w.phase==='current'?'حالي':w.phase==='future'?'قادم':'';
      h+=`<g class="wa-week-point" tabindex="0" role="button" data-wa-week="${w.week}" aria-label="الأسبوع ${w.week}، ${num(w.count)} حالات، ${escapeHTML(rangeLabel(w.start,w.end))}، ${period}"><circle class="wa-focus-ring" cx="${xx}" cy="${yy}" r="15" fill="transparent"/><circle class="wa-dot" cx="${xx}" cy="${yy}" r="${isPeak?7:5.5}" fill="${col}" stroke="${isPeak?'#b9752d':'#fffefa'}" stroke-width="2.5"/><text x="${xx}" y="${yy-17}" text-anchor="middle" font-size="13" font-weight="${isPeak?'700':'500'}" fill="${isPeak?'#805119':'#345446'}">${num(w.count)}</text><text x="${xx}" y="292" text-anchor="middle" font-size="11" fill="${past?'#7f8881':'#173c30'}">الأسبوع ${w.week}</text><text x="${xx}" y="311" text-anchor="middle" font-size="10" fill="#7b897e">${w.start.slice(8)}/${w.start.slice(5,7)} — ${w.end.slice(8)}/${w.end.slice(5,7)}</text><text x="${xx}" y="348" text-anchor="middle" font-size="10" fill="${past?'#91968f':'#657e70'}">${period}</text><title>${escapeHTML(rangeLabel(w.start,w.end))} · ${num(w.count)} حالات</title></g>`;
    });
    ui('wa-chart').innerHTML=h+'</svg>';
  }
  function updateMap(){const layer=ui('pa-map-layer');if(layer)layer.setAttribute('transform',`translate(${mapX} ${mapY}) scale(${mapZoom})`);}
  function zoomMap(factor){
    const next=Math.max(1,Math.min(4,mapZoom*factor)),ratio=next/mapZoom;
    mapX=600-(600-mapX)*ratio;mapY=400-(400-mapY)*ratio;mapZoom=next;
    if(next===1)mapX=mapY=0;updateMap();
  }
  function bind(){
    if(bound)return;bound=true;
    document.addEventListener('click',event=>{
      const node=event.target.closest?.('[data-pa-port],[data-wa-week],[data-pa-case],[data-pa-filter],[data-pa-close]');if(!node)return;
      if(node.hasAttribute('data-pa-port'))portDetails(Number(node.dataset.paPort));
      else if(node.hasAttribute('data-wa-week'))weekDetails(Number(node.dataset.waWeek));
      else if(node.hasAttribute('data-pa-case'))showCaseDetails(node.dataset.paCase);
      else if(node.hasAttribute('data-pa-filter')){const id=node.dataset.paFilter;closeAnalyticsDialog();if(typeof root.selectPortAnalytics==='function')root.selectPortAnalytics(id);}
      else if(node.hasAttribute('data-pa-close'))closeAnalyticsDialog();
    });
    document.addEventListener('keydown',event=>{
      const node=event.target.closest?.('g[data-pa-port],g[data-wa-week]');
      if(node&&(event.key==='Enter'||event.key===' ')){event.preventDefault();node.dispatchEvent(new MouseEvent('click',{bubbles:true}));}
    });
    ui('pa-zoom-in').addEventListener('click',()=>zoomMap(1.3));ui('pa-zoom-out').addEventListener('click',()=>zoomMap(1/1.3));
    ui('pa-map-reset').addEventListener('click',()=>{mapZoom=1;mapX=mapY=0;updateMap();});
    const svg=ui('pa-map');
    const svgPoint=event=>{const pt=svg.createSVGPoint();pt.x=event.clientX;pt.y=event.clientY;const matrix=svg.getScreenCTM();return matrix?pt.matrixTransform(matrix.inverse()):null;};
    svg.addEventListener('pointerdown',event=>{if(event.target.closest?.('.pa-marker'))return;const p=svgPoint(event);if(!p)return;drag={x:p.x,y:p.y,mx:mapX,my:mapY};svg.setPointerCapture(event.pointerId);});
    svg.addEventListener('pointermove',event=>{if(!drag)return;const p=svgPoint(event);if(!p)return;mapX=drag.mx+p.x-drag.x;mapY=drag.my+p.y-drag.y;updateMap();});
    svg.addEventListener('pointerup',()=>drag=null);svg.addEventListener('pointercancel',()=>drag=null);svg.addEventListener('lostpointercapture',()=>drag=null);
  }
  root.renderPortAnalytics=function(rows,matchedGroups){
    if(!ui('pa-stats')||!ui('wa-chart'))return;
    const data=dashboardData();
    const supplied=Array.isArray(matchedGroups)?matchedGroups:matchedGroups instanceof Map?[...matchedGroups.values()]:matchedGroups instanceof Set?[...matchedGroups]:data.groups||[];
    const groups=supplied.map(g=>typeof g==='object'?g:(data.groups||[]).find(item=>String(item.id)===String(g))).filter(Boolean);
    const ids=new Set(groups.map(g=>String(g.id))),scope=uniqueCases(rows).filter(c=>ids.has(String(c.group_id)));
    const meta=data.meta||{};
    state={ports:aggregatePorts(scope,groups),weeks:weeklyRows(scope,meta.period_start,meta.period_end,meta.as_of),groups:new Map(groups.map(g=>[String(g.id),g])),rows:scope,meta};
    bind();renderMap();renderWeekly();
  };
})(typeof window!=='undefined'?window:globalThis);

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
  const groupJob = group => typeof window.Dashboard?.roleOf === 'function' ? window.Dashboard.roleOf(group) : (String(group?.title || '').trim() || 'المسمى الوظيفي غير متوفر');
  const plannerGroupMatches = id => {
    const group=groups.get(id);if(!group)return false;
    if(typeof window.Dashboard?.groupMatches==='function' && !window.Dashboard.groupMatches(group))return false;
    const query=norm(ui('manager-search').value);
    return !query || norm([group.manager,group.manager_raw,group.raw?.manager,groupJob(group)].filter(Boolean).join(' ')).includes(query);
  };
  const scopedCases = () => model.originalCases(threshold()).filter(item => plannerGroupMatches(item.group_id));
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
      '، حسب الحالات الأصلية في التقرير والفلاتر العامة' + (threshold() ? ' وبنسبة 50% فأكثر.' : '، بجميع النسب.');
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
        const label = group.manager + '، ' + groupJob(group) + '، ' + group.port + (org ? '، ' + org : '') + '، ' + (good ? 'موعد بعد التنسيق' : 'تقاطع ' + number(ratio) + '%') + '، ' + period(event.start, event.end) + '، عرض الموظفين';
        bars += `<button type="button" class="sp-event${good ? ' sp-good' : ''}${selected === event.group_id ? ' sp-selected' : ''}" data-group="${escape(event.group_id)}" data-event="${escape(event.id)}" aria-label="${escape(label)}" title="${escape(label)}" aria-expanded="${selected === event.group_id}" ${busy ? 'disabled' : ''} style="grid-column:${start - week + 1} / span ${end - start + 1};grid-row:${lane + 1}"><span class="sp-event-person"><span class="sp-event-name">${escape(group.manager)}</span><span class="sp-event-role">${escape(groupJob(group))}</span></span><span class="sp-event-count">${tail}</span><span class="sp-mobile-period">${escape(period(toIso(start), toIso(end)))}</span></button>`;
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
    ui('manager-role').textContent = groupJob(group);
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
    ui('live').textContent = 'تفاصيل إجازات ' + groups.get(id).manager + '، ' + groupJob(groups.get(id)) + ' في ' + groups.get(id).port;
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
  document.addEventListener('dashboard:scopechange',()=>{if(!busy)render();});
  const plannerManagerNames=new Map();
  for(const item of model.originalCases()){const group=groups.get(item.group_id);if(!group?.manager)continue;if(!plannerManagerNames.has(group.manager))plannerManagerNames.set(group.manager,new Set());plannerManagerNames.get(group.manager).add(groupJob(group));}
  ui('manager-names').innerHTML = [...plannerManagerNames].sort((a,b)=>a[0].localeCompare(b[0],'ar')).map(([name,roles]) => '<option value="' + escape(name) + '" label="' + escape([...roles].join(' / ')) + '"></option>').join('');
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

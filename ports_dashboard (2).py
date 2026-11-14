# -*- coding: utf-8 -*-
"""
داشبورد تقاطع الإجازات — نفس منطق سكربت التقرير المرفق.

التثبيت مرة واحدة:
    python -m pip install pandas openpyxl
التشغيل:
    python ports_dashboard.py

عدّلي مساري Employee Master وPorts Leaves أدناه، ثم شغّلي الملف.
الناتج: dashboard.html بجانب ملف الإجازات، ويُفتح بالمتصفح مباشرة.
ملف Python واحد، ويُنشئ HTML فقط. جميع مكونات العرض مضمنة وتعمل دون إنترنت.

الحالة = نفس مجموعة موظفين اثنين فأكثر في أيام تقويمية متصلة، ضمن
المدير نفسه + الموقع نفسه + المسار التنظيمي نفسه للمستويات الأول والثاني والثالث.
تغيّر الموظفين أو انقطاع الأيام يبدأ حالة جديدة. النسبة من موظفي المجموعة نفسها.
تعرض الأسماء والمسميات كما وردت في المصادر، وتُعرض جميع الحالات بغض النظر عن نسبتها.
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
        'problems': problems, 'ambiguous': ambiguous,
        'source_rows': {'managers': len(managers), 'staff': len(staff), 'leaves': len(leaves)},
    }


def short_action(case):
    if case['status'] == 'منتهية':
        return 'منتهية — للعلم'
    if case['status'] == 'جارية':
        return 'مراجعة التغطية الحالية'
    return 'مراجعة الجدولة قبل تاريخ البداية'


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
</style>
</head>
<body>
<header class="hero">
  <div class="wrap hero-inner">
    <div><p class="department" id="department"></p><h1 id="report-title">لوحة تقاطع الإجازات</h1><div class="period" id="report-period"></div></div>
    <aside class="hero-aside"><div><span>تاريخ تصنيف الحالات</span><strong id="report-asof"></strong></div><button type="button" class="print-button" id="print-report">طباعة العرض الحالي</button></aside>
  </div>
</header>
<main class="wrap">
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
</body>
</html>
"""


if __name__ == '__main__':
    try:
        build()
    except Exception as exc:
        print('خطأ: ' + str(exc), file=sys.stderr)
        sys.exit(1)

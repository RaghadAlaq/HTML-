# -*- coding: utf-8 -*-
"""
تقرير تنفيذي لتقاطعات إجازات موظفي المنافذ — الحالات التي تبلغ 50% فأكثر.
Input: Excel workbook with 'مدراء المنافذ' and 'الموظفين تحت مدراء المنافذ',
       plus approved leaves workbook with 'البيانات الخام'.
Output: Excel with 50%+ priority cases, all intersections, manager/location summary, QA.
Install: pip install pandas openpyxl
Run offline: python ports_ceo_by_org_group.py

Core definition: an intersection CASE is a CONSECUTIVE run of calendar days with
THE EXACT SAME SET of 2+ employees on approved leave under the same manager,
Location (Name), and organization levels 1, 2 and 3. If the cohort changes, a new case begins. Therefore a 6-person
2-day case really means ALL SIX were away together for BOTH days.
Period-end inclusive by default; confirm HRIS semantics before final sign-off.
"""
from pathlib import Path
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo
from collections import defaultdict, Counter
import re
import unicodedata
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

# ===================== EDIT THESE PATHS =====================
MANAGERS_FILE = r'C:\Users\rammutairi\Downloads\Employee Master.xlsx'
LEAVES_FILE = r'C:\Users\rammutairi\Downloads\Ports Leaves.xlsx'
MANAGERS_SHEET = 'مدراء المنافذ'
STAFF_SHEET = 'الموظفين تحت مدراء المنافذ'
LEAVES_SHEET = 'البيانات الخام'
OUTPUT_FILE = r''  # blank => next to LEAVES_FILE
# ============================================================
AS_OF = datetime.now(ZoneInfo('Asia/Riyadh')).date().isoformat()  # تاريخ الرياض الفعلي عند التشغيل
PERIOD_START = '2026-10-01'
PERIOD_END = '2026-12-31'
THRESHOLD = 50.0
END_DATE_INCLUSIVE = True  # DO NOT change unless HR confirms its meaning
REPORT_TITLE = 'تقرير تقاطع الإجازات حسب المدير وموقع العمل والوحدة التنظيمية'
DEPARTMENT = 'إدارة أداء وكالة رأس المال البشري'
PREPARED_BY = 'تنفيذ رغد المطيري'
SUPERVISED_BY = 'إشراف أمل السويلم'

# ضعي الترجمات المعتمدة من الهيئة هنا: المصطلح الإنجليزي كما ورد في المصدر -> الاسم العربي الرسمي.
# إن لم تتوفر ترجمة معتمدة، يبقى المسمى الأصلي كما هو لتجنب اختلاق مسمى تنظيمي.
OFFICIAL_ORG_TRANSLATIONS = {
    # 'Security Affairs': 'المسمى العربي المعتمد هنا',
}
# ترجمات وصفية مقترحة فقط (ليست مسميات تنظيمية معتمدة):
SUGGESTED_TRANSLATIONS = {
    'operations':'العمليات',
    'security affairs':'الشؤون الأمنية',
    'legal enforcement':'الإنفاذ القانوني',
    'inspection':'التفتيش',
    'land':'البري',
    'air':'الجوي',
    'sea':'البحري',
    'marine':'البحري',
}


# APPROVAL field absent => don't assume data is approved; set flag below
ALLOW_NO_APPROVAL_COLUMN = True  # source must be an approved-leaves-only export

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

def translate_org(value):
    name=str(value or '').strip()
    if not name:return 'غير محدد'
    # Mapping is case/Arabic-normalization insensitive.
    official={norm(k):v for k,v in OFFICIAL_ORG_TRANSLATIONS.items()}.get(norm(name))
    if official:return official
    suggested=SUGGESTED_TRANSLATIONS.get(norm(name))
    return f'{suggested} ({name})' if suggested else name

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

def build():
    asof=date.fromisoformat(AS_OF)
    print(f'تاريخ إعداد التقرير والتصنيف: {asof:%Y-%m-%d} (تاريخ الجهاز المحلي)', flush=True)
    pstart=date.fromisoformat(PERIOD_START);pend=date.fromisoformat(PERIOD_END)
    print('1/5 قراءة الشيتات...',flush=True)
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
        if sn and cell(r,sn):staff_by_name[norm(cell(r,sn))].add(eid)
    if not group_staff:raise ValueError('لم تتم مطابقة موظفين مع المدراء حسب رقم المدير المباشر والموقع.')
    print(f'2/5 مطابقة الإجازات مع {len(staff_by_id)} موظفًا مستهدفًا...',flush=True)
    daily=defaultdict(set);day_leaves=defaultdict(list); record_signatures=set()
    matched=0;out_of_port=0
    for rowno,r in leaves.iterrows():
        st=dt(r[ls]);en=dt(r[le])
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
                         employees=','.join(sorted(cohort)),flag=ratio>=THRESHOLD))
    print('4/5 إعداد جداول الحالات والملخص...',flush=True)
    def fmt_date(d):return d.strftime('%d/%m/%Y') if d else ''
    def common(g):
        m=managers_by_id[g[0]]
        levels=display_levels[g]
        return [m['name'], m['job'], display_port[g],
                *(translate_org(v) for v in levels),
                'مجموعة تنظيمية مستقلة']
    def short_action(c):
        if c['status']=='منتهية':return 'منتهية — للعلم'
        if c['status']=='جارية':return 'مراجعة التغطية الحالية'
        return 'مراجعة الجدولة قبل تاريخ البداية'
    priority=[c for c in cases if c['flag']]
    priority.sort(key=lambda c:({'جارية':0,'قادمة':1,'منتهية':2}[c['status']],
                                -c['percent'],-c['people'],c['start']))
    # All qualifying overlaps (2+ employees), regardless of percentage.
    # Each run is a fixed cohort on consecutive calendar days; no misleading
    # long duration assembled from different employees.
    all_cases = sorted(cases, key=lambda c: (
        {'جارية': 0, 'قادمة': 1, 'منتهية': 2}[c['status']],
        -c['percent'], -c['people'], c['start'], c['group']))
    all_rows = []
    for c in all_cases:
        g=c['group']
        all_rows.append([c['status'], *common(g), c['workforce'], c['people'],
            round(c['percent'], 1), fmt_date(c['start']), fmt_date(c['end']),
            c['days'], short_action(c)])

    priority_rows=[]
    for c in priority:
        g=c['group']
        priority_rows.append([c['status'],*common(g),c['workforce'],c['people'],
                 round(c['percent'],1),fmt_date(c['start']),
                 fmt_date(c['end']),c['days'],short_action(c)])
    summary=[]
    by_group=defaultdict(list)
    for c in cases:by_group[c['group']].append(c)
    for g,people in group_staff.items():
        relevant=[c for c in by_group[g] if c['flag']]
        if not relevant:continue
        active=[c for c in relevant if c['status']!='منتهية']
        # Peak counted from any overlapping 2+ cohort; also show peak 50% qualifying.
        best=max(relevant,key=lambda c:(c['percent'],c['days'], -c['start'].toordinal()))
        worst_status='جارية/قادمة' if active else 'منتهية'
        summary.append([*common(g),len(people),len(overlap_people[g]),len(relevant),len(active),
                        best['people'],round(best['percent'],1),fmt_date(best['start']),
                        fmt_date(best['end']),best['days'],worst_status])
    summary.sort(key=lambda r:(-r[12],-r[11],r[0])) # % then number of people, independent for each port
    qa_rows=[['اسم التقرير',REPORT_TITLE],['الإدارة',DEPARTMENT],['الإعداد',PREPARED_BY],['الإشراف',SUPERVISED_BY],['تاريخ إعداد التقرير (تاريخ اليوم)',AS_OF],['تاريخ تصنيف الحالات',AS_OF],['الفترة',PERIOD_START+' إلى '+PERIOD_END],
             ['مدراء مستهدفون',len(managers_by_id)],['موظفون ضمن مجموعات المدير والموقع والمستويات',len(staff_by_id)],
             ['سجلات إجازات مطابقة',matched],['مجموعات مدير/منفذ/مستويات تنظيمية',len(group_staff)],
             ['حالات التقاطع الكلية (بنفس الموظفين)',len(cases)],
             ['حالات تحقق 50% أو أكثر',len(priority)],
             ['جميع حالات التقاطع بغض النظر عن النسبة',len(all_cases)],
             ['منهجية الحالة','نفس الموظفين المجازين بالضبط في أيام متصلة، على مستوى المدير والمنفذ والمستويات التنظيمية 1-3'],
             ['تفسير النسبة','عدد المجازين المتطابقين في الفترة ÷ جميع موظفي المدير في الموقع والمسار التنظيمي نفسه'],
             ['تعريف الفترة','اليوم الأول إلى اليوم الأخير اللذين غاب خلالهما نفس الأشخاص معًا'],
             ['تاريخ النهاية','يُحتسب ضمن أيام الإجازة؛ يلزم تأكيد تعريفه في نظام الموارد البشرية'],
             ['تنبيه مصدر البيانات','قد تختلف تبعية الموظفين بين الملفين؛ الحالات غير المتطابقة في الموقع مستبعدة'],
             ['الحدود','لا توجد بيانات مناوبات أو حد أدنى للتغطية التشغيلية؛ النسبة ليست دليل عجز تشغيلي'],
             ['المسميات التنظيمية','كل مسار تنظيمي فريد للمستويات 1-3 يمثل مجموعة مستقلة حتى مع تطابق المدير والموقع؛ لا تستخدم متعدد'],
             ['الترجمة','تطبق الترجمات المعتمدة في OFFICIAL_ORG_TRANSLATIONS؛ المصطلحات الشائعة تُعرض بترجمة وصفية مع الأصل الإنجليزي وليست أسماء رسمية معتمدة'],
             ['الانتباه للتبعية','يعتمد عدد الفريق على الموظفين الذين يظهر مديرهم المباشر في شيت الموظفين؛ لا يشمل جميع المرؤوسين غير المباشرين']]
    qa_rows += [['تنبيه', 'الموظفون ناقصو مستويات تنظيمية قد يظهرون ضمن مجموعة غير محددة؛ يلزم استكمال المسميات قبل اعتمادها']]
    qa_rows += [[k,v] for k,v in sorted(problems.items())]
    qa_rows += [['مثال استبعاد: '+k, f'{eid} | {old} | {new}'] for k,eid,old,new in ambiguous[:30]]

    out=Path(OUTPUT_FILE) if OUTPUT_FILE else Path(LEAVES_FILE).resolve().parent/'تقرير_تقاطع_الإجازات_حسب_الوحدات_التنظيمية.xlsx'
    out.parent.mkdir(parents=True,exist_ok=True)
    with pd.ExcelWriter(out,engine='openpyxl') as wr:
        pd.DataFrame(priority_rows,columns=['حالة الفترة','اسم المدير','المسمى الوظيفي','المنفذ','المستوى الأول','المستوى الثاني','المستوى الثالث','نوع التجميع','إجمالي موظفي المجموعة التنظيمية في الموقع','أعلى عدد مجازين معًا','نسبة الغياب المجدول %','بداية التقاطع','نهاية التقاطع','مدة غياب المجموعة نفسها (يوم)','الإجراء المقترح']).to_excel(wr,sheet_name='حالات 50 فأكثر',index=False)
        pd.DataFrame(summary,columns=['اسم المدير','المسمى الوظيفي','المنفذ','المستوى الأول','المستوى الثاني','المستوى الثالث','نوع التجميع','إجمالي موظفي المجموعة','الموظفون الذين ظهر لهم تقاطع','حالات 50 فأكثر','حالات جارية أو قادمة','أعلى عدد مجازين معًا','أعلى نسبة غياب %','بداية أشد فترة','نهاية أشد فترة','مدة أشد فترة (يوم)','وضع الحالات']).to_excel(wr,sheet_name='ملخص المدير والمنفذ',index=False)
        pd.DataFrame(all_rows,columns=['حالة الفترة','اسم المدير','المسمى الوظيفي','المنفذ','المستوى الأول','المستوى الثاني','المستوى الثالث','نوع التجميع','إجمالي موظفي المجموعة التنظيمية في الموقع','أعلى عدد مجازين معًا','نسبة الغياب المجدول %','بداية التقاطع','نهاية التقاطع','مدة غياب المجموعة نفسها (يوم)','الإجراء المقترح']).to_excel(wr,sheet_name='جميع حالات التقاطع',index=False)
        pd.DataFrame(qa_rows,columns=['المؤشر','القيمة']).to_excel(wr,sheet_name='ضبط الجودة',index=False)

    wb=load_workbook(out)
    for ws in wb:
        # هوية التقرير والتاريخ على كل صفحة.
        ws.insert_rows(1,amount=3)
        end_col=ws.max_column
        ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=end_col)
        headline=ws.cell(1,1)
        headline.value=REPORT_TITLE
        headline.fill=PatternFill('solid',fgColor='DCEBEA')
        headline.font=Font(name='Arial',size=12,bold=True,color='123752')
        headline.alignment=Alignment(horizontal='right',vertical='center')
        ws.row_dimensions[1].height=33
        for line, textline in [(2, DEPARTMENT),(3,f'{PREPARED_BY}  |  {SUPERVISED_BY}  |  تاريخ إعداد التقرير: {asof:%d/%m/%Y}')]:
            ws.merge_cells(start_row=line,start_column=1,end_row=line,end_column=end_col)
            c=ws.cell(line,1);c.value=textline
            c.fill=PatternFill('solid',fgColor='EFF5F7')
            c.font=Font(name='Arial',size=11,bold=line==2,color='123752')
            c.alignment=Alignment(horizontal='right',vertical='center')
            ws.row_dimensions[line].height=27
        ws.sheet_view.rightToLeft=True
        ws.freeze_panes='D5' if ws.title!='ضبط الجودة' else 'A5'
        ws.sheet_properties.pageSetUpPr.fitToPage=True
        ws.sheet_format.defaultRowHeight=23
        if ws.max_row >= 5:
            ws.auto_filter.ref=f'A4:{get_column_letter(ws.max_column)}{ws.max_row}'
        ws.row_dimensions[4].height=44
        ws.sheet_view.showGridLines=False
        headers=[c.value for c in ws[4]]
        for i,h in enumerate(headers,1):
            letter=get_column_letter(i)
            max_width=40 if any(t in str(h) for t in ['المدير','الإجراء','المسمى','المستوى']) else 23
            if h=='المنفذ':max_width=35
            ws.column_dimensions[letter].width=min(max_width,max(17,len(str(h))*1.05+4))
        for c in ws[4]:
            c.fill=PatternFill('solid',fgColor='123752')
            c.font=Font(name='Arial',bold=True,color='FFFFFF',size=10)
            c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        for r in ws.iter_rows(min_row=5):
            status=str(r[0].value or '') if ws.title in ['حالات 50 فأكثر','جميع حالات التقاطع'] else ''
            for c in r:
                c.font=Font(name='Arial',size=10,color='58636E' if status=='منتهية' else '17304A')
                c.fill=PatternFill('solid',fgColor='EAECF0' if status=='منتهية' else ('F3F7F8' if c.row%2==0 else 'FFFFFF'))
                c.alignment=Alignment(vertical='center',wrap_text=False)
                if c.column>1 and isinstance(c.value,(int,float)):c.number_format='#,##0.0' if '%' in str(headers[c.column-1]) else '#,##0'
            ws.row_dimensions[r[0].row].height=22
    wb.save(out)
    print(f'5/5 تم إنشاء التقرير: {out}',flush=True)
    print(f'  حالات 50% فأكثر: {len(priority)} | مجموعات مستهدفة: {len(summary)} | جميع حالات التقاطع: {len(all_cases)}',flush=True)
    if out_of_port:print(f'تنبيه: {out_of_port} سجل إجازة بموقع مختلف تم استبعاده وإظهاره في ضبط الجودة.',flush=True)

if __name__=='__main__':build()

# -*- coding: utf-8 -*-
"""
ports_report.py — تحليل إجازات مدراء المنافذ وإنشاء HTML في تشغيل واحد.
المدخلات: ملف المدراء والموظفين + ملف سجل الإجازات.
المخرجات: ports_dashboard_READY.html (بياناتك الفعلية مضمّنة)، ports_dashboard_data.json (نسخة مراجعة).
المرجع الزمني: 18-11-2026؛ داخل اللوحة يمكن تغيير اليوم لإعادة عرض الحالات القادمة.
لا يحتاج تنزيل مكتبات Python، ويعمل دون اتصال؛ لا يغير ملف Excel الأصلي.
"""
from pathlib import Path
from collections import defaultdict, Counter
from datetime import date, datetime, timedelta
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import json, re, html

# ===================== عدلي المسارات فقط =====================
MANAGERS_WORKBOOK = r'C:\Users\rammutairi\Downloads\Employee Master.xlsx'
LEAVES_WORKBOOK   = r'C:\Users\rammutairi\Downloads\تقرير تدقيق تداخل الإجازات.xlsx'
MANAGERS_SHEET    = 'مدراء المنافذ'
ROSTER_SHEET      = 'الموظفين تحت مدراء المنافذ'
LEAVES_SHEET      = 'البيانات الخام'
REFERENCE_DATE    = '2026-11-18'  # يوم العرض؛ يعاد حساب القريب/المنتهي من داخـل HTML أيضًا
OUTPUT_FOLDER     = r''  # فارغ = مجلد ports_executive_report بجوار ملف الإجازات
# ============================================================
START, END = date(2026,10,1), date(2026,12,31)
N={'x':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

def norm(x):
    s=str(x or '').strip().lower()
    s=re.sub('[\u064b-\u065f\u0670\u0640]','',s).replace('أ','ا').replace('إ','ا').replace('آ','ا').replace('ى','ي').replace('ة','ه')
    return re.sub(r'\s+',' ',s)

def emp_id(x):
    s=str(x or '').strip()
    return s[:-2] if s.endswith('.0') and s[:-2].isdigit() else s

def excel_date(s):
    if isinstance(s,date):return s
    if not s:return None
    st=str(s).strip()
    if re.match(r'^\d+(\.\d+)?$',st):
        try:return date(1899,12,30)+timedelta(days=int(float(st)))
        except:return None
    for f in ['%Y-%m-%d','%Y-%m-%dT%H:%M:%S','%d/%m/%Y','%m/%d/%Y','%Y/%m/%d','%d-%m-%Y']:
        try:return datetime.strptime(st[:19] if 'T' in f else st.split(' ')[0],f).date()
        except:pass
    return None

def read_xlsx(path, sheet):
    path=Path(path)
    if not path.exists():raise FileNotFoundError('لم أجد ملف Excel: '+str(path))
    with ZipFile(path) as z:
        ss=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            root=ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall('x:si',N):ss.append(''.join(t.text or '' for t in si.findall('.//x:t',N)))
        wb=ET.fromstring(z.read('xl/workbook.xml'))
        sheets={e.get('name'):e.get('{'+N['r']+'}id') for e in wb.findall('.//x:sheet',N)}
        if sheet not in sheets:raise ValueError(f'لم أجد شيت «{sheet}» في {path.name}؛ الموجود: {list(sheets)}')
        rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        relmap={r.get('Id'):r.get('Target') for r in rels}
        target=relmap[sheets[sheet]].lstrip('/')
        if not target.startswith('xl/'):target='xl/'+target
        xml=ET.fromstring(z.read(target))
        output=[]
        for row in xml.findall('.//x:sheetData/x:row',N):
            cells={}
            for c in row.findall('x:c',N):
                ref=c.get('r',''); match=re.match(r'[A-Z]+',ref)
                if not match:continue
                i=0
                for letter in match.group(0):i=i*26+ord(letter)-64
                value=c.find('x:v',N);inline=c.find('x:is',N)
                if c.get('t')=='s' and value is not None:
                    txt=ss[int(value.text)]
                elif c.get('t')=='inlineStr' and inline is not None:
                    txt=''.join(t.text or '' for t in inline.findall('.//x:t',N))
                else:txt=value.text if value is not None else ''
                cells[i-1]=txt
            if cells:output.append([cells.get(i,'') for i in range(max(cells)+1)])
        if not output:return []
        headers=[str(x).strip() for x in output[0]]
        if not any('اسم' in h or 'Manager' in h or 'Person' in h for h in headers):
            for i,row in enumerate(output[:20]):
                if sum(bool(x) for x in row)>=3 and any('Manager' in str(x) or 'الاسم الرسمي' in str(x) or 'الرقم الوظيفي' in str(x) for x in row):
                    headers=[str(x).strip() for x in row];output=output[i:];break
        return [{headers[i]:r[i] if i<len(r) else '' for i in range(len(headers)) if headers[i]} for r in output[1:] if any(r)]

def get(r,aliases):
    keys={norm(k):k for k in r}
    for a in aliases:
        if norm(a) in keys:return str(r[keys[norm(a)]] or '').strip()
    return ''

NAME=['الاسم الرسمي بالعربي','الاسم بالعربي','اسم الموظف','Employee Name','(Employee) Full Name','Full Name']
ID=['الرقم الوظيفي','Employee ID','Person ID','(Employee) Person ID','الرقم']
MGR=['المدير','اسم المدير','اسم المدير المباشر','المدير المباشر','Manager','Manager Name','Direct Manager','(Employee) Manager (Name)']
PORT=['المنفذ','اسم المنفذ','موقع العمل','مقر العمل','Location (Name)','(Employee) Location (Name)','Location']
SUPERVISOR_ID=['الرقم الوظيفي للمدير المباشر','Manager Employee ID','Supervisor ID','Manager ID']
SUPERVISOR_FIRST=['المدير المباشر - الاسم الاول','Supervisor First Name']
SUPERVISOR_LAST=['المدير المباشر - الاسم الاخير','Supervisor Last Name']
START_COL=['Actual Start Date','تاريخ البداية','تاريخ بدء الإجازة','تاريخ بداية الإجازة','Start Date']
END_COL=['Actual End Date','تاريخ النهاية','تاريخ نهاية الإجازة','End Date']
STATUS=['Approval Status','Approval Status (Name)','Approval Status (Label)','حالة الاعتماد']
APPROVED={'approved','معتمد','معتمده','موافق عليه','موافق','تم الاعتماد','approved request'}


def build():
    masters=read_xlsx(MANAGERS_WORKBOOK,MANAGERS_SHEET)
    roster=read_xlsx(MANAGERS_WORKBOOK,ROSTER_SHEET)
    leaves=read_xlsx(LEAVES_WORKBOOK,LEAVES_SHEET)
    if not masters or not roster or not leaves:raise ValueError('إحدى الأوراق فارغة')
    selected={norm(get(r,NAME)):get(r,NAME) for r in masters if get(r,NAME)}
    manager_by_id={emp_id(get(r,ID)):get(r,NAME) for r in masters if emp_id(get(r,ID)) and get(r,NAME)}
    # Fallback by first and last name ONLY if unambiguous, otherwise supervisor ID is mandatory.
    manager_short=defaultdict(set)
    for r in masters:
        full=get(r,NAME)
        parts=norm(full).split()
        if len(parts)>=2:
            manager_short[(parts[0],parts[-1])].add(full)
    if not selected:raise ValueError('لم أعثر على «الاسم الرسمي بالعربي» في شيت مدراء المنافذ.')

    roster_map={};roster_names={};unmatched_roster=0;ambiguous_roster=0
    resolved_roster=[]
    for r in roster:
        ident=emp_id(get(r,ID));nm=norm(get(r,NAME))
        supervisor_id=emp_id(get(r,SUPERVISOR_ID))
        manager_name=manager_by_id.get(supervisor_id)
        if not manager_name:
            direct=get(r,MGR)
            manager_name=selected.get(norm(direct))
        if not manager_name:
            short=(norm(get(r,SUPERVISOR_FIRST)), norm(get(r,SUPERVISOR_LAST)))
            options=manager_short.get(short,set())
            if len(options)==1:manager_name=next(iter(options))
            elif len(options)>1:ambiguous_roster+=1
        p=get(r,PORT)
        if not manager_name or not p or not (ident or nm):
            unmatched_roster+=1;continue
        info=(manager_name,p)
        resolved_roster.append((ident or 'name:'+nm,manager_name,p))
        if ident:
            if ident in roster_map and roster_map[ident]!=info:
                raise ValueError(f'الرقم الوظيفي {ident} له أكثر من تبعية في شيت الموظفين؛ يلزم تحديد السجل المعتمد.')
            roster_map[ident]=info
        if nm and nm not in roster_names:roster_names[nm]=info
    if not resolved_roster:
        raise ValueError('لم أستطع ربط الموظفين بالمدراء المستهدفين: تحققي من الرقم الوظيفي للمدير المباشر وLocation (Name).')
    print(f'تم ربط {len(resolved_roster)} موظفًا بمدراء المنافذ باستخدام الرقم الوظيفي للمدير المباشر.')
    valid=[]; unmatched_leaves=[]
    for r in leaves:
        st=excel_date(get(r,START_COL));en=excel_date(get(r,END_COL))
        if not st or not en or en<START or st>END or en<st:continue
        status=norm(get(r,STATUS));
        if status and status not in APPROVED:continue
        ident=emp_id(get(r,ID));nm=norm(get(r,NAME));manager_name=get(r,MGR);p=get(r,PORT)
        # Match selected master manager; use roster to resolve if raw manager naming differs.
        k=norm(manager_name)
        if k in selected:manager_name=selected[k]
        else:
            manager_name=manager_by_id.get(emp_id(get(r,SUPERVISOR_ID)),manager_name)
            if norm(manager_name) not in selected:
                info=roster_map.get(ident) or roster_names.get(nm)
                if info:manager_name=info[0]
        if norm(manager_name) not in selected:
            unmatched_leaves.append((ident,manager_name,p));continue
        if not p:
            p=(roster_map.get(ident) or roster_names.get(nm) or ('',''))[1]
        if not p:
            unmatched_leaves.append((ident,manager_name,p));continue
        if not ident:ident='name:'+nm
        if ident=='name:':unmatched_leaves.append((ident,manager_name,p));continue
        valid.append(dict(manager=manager_name,port=p,emp=ident,start=max(st,START),end=min(en,END)))
    if not valid:
        raise ValueError('لم تتم مطابقة أي إجازات لمدراء المنافذ. راجعي أعمدة اسم المدير والرقم الوظيفي ومسميات المنافذ. لا يمكن إنشاء نسب غير مستندة لبيانات.')
    # Workforce: combine roster and raw leave employees, dedupe inside manager + port.
    workforce=defaultdict(set);display={}
    for ident,manager_name,p in resolved_roster:
        k2=(manager_name,norm(p));display[k2]=p
        workforce[k2].add(ident)
    # Prefer latest raw leave assignment for people whose older roster location changed.
    current_assignment={}
    for r in sorted(valid,key=lambda x:(x['start'],x['end'])):
        current_assignment[r['emp']]=(r['manager'],norm(r['port']))
    for group,employees in workforce.items():
        employees.difference_update({emp for emp in employees if emp in current_assignment and current_assignment[emp]!=group})
    for r in valid:
        k=(r['manager'],norm(r['port']));display[k]=r['port'];workforce[k].add(r['emp'])
    # use all target manager + port groups; keep IDs unique per group
    daily=defaultdict(set);raw_records=defaultdict(list)
    for r in valid:
        k=(r['manager'],norm(r['port']))
        raw_records[k].append(r)
        d=r['start']
        while d<=r['end']:
            daily[(k,d)].add(r['emp']);d+=timedelta(days=1)
    groups=[];days=[];exact=[];warnings=[]
    for k,workers in sorted(workforce.items()):
        mgr,p=k;key=mgr+' || '+p
        count=len(workers)
        if not count:continue
        groups.append(dict(key=key,manager=mgr,port=display[k],workforce=count))
        for (group,d),people in daily.items():
            if group!=k:continue
            # A collision means >=2 people in same team same date
            if len(people)<2:continue
            days.append(dict(key=key,date=d.isoformat(),on_leave=len(people),ratio=(len(people)/count if count else None)))
        same=defaultdict(set)
        for r in raw_records.get(k,[]):same[(r['start'],r['end'])].add(r['emp'])
        for (st,en),emps in same.items():
            if len(emps)>=2:
                exact.append(dict(manager=mgr,port=display[k],start=st.isoformat(),end=en.isoformat(),duration=(en-st).days+1,employee_count=len(emps)))
    if unmatched_leaves:warnings.append(f'{len(unmatched_leaves)} إجازة لم تطابق قائمة المديرين المستهدفين أو لم يتوفر منفذها.')
    if unmatched_roster:warnings.append(f'{unmatched_roster} صفًا من ملف موظفي المنافذ لم يُربط بمدير مستهدف ومنفذ؛ يُراجع اكتمال إجمالي القوى العاملة.')
    # Each continuous run of intersecting approved leaves is a distinct actionable case.
    # Its start may be before the evaluation date, in which case it is ongoing.
    grouped_days=defaultdict(list)
    for item in days:grouped_days[item['key']].append(item)
    cases=[]
    group_index={g['key']:g for g in groups}
    for key,series in grouped_days.items():
        series.sort(key=lambda x:x['date'])
        run=[]
        def save_run(run):
            if not run:return
            peak=max(run,key=lambda r:r['on_leave'])
            g=group_index[key]
            cases.append(dict(key=key,manager=g['manager'],port=g['port'],start=run[0]['date'],end=run[-1]['date'],duration=len(run),peak_date=peak['date'],peak_on_leave=peak['on_leave'],ratio=peak['ratio'],workforce=g['workforce']))
        for item in series:
            if run and (date.fromisoformat(item['date'])-date.fromisoformat(run[-1]['date'])).days>1:
                save_run(run);run=[]
            run.append(item)
        save_run(run)
    result=dict(groups=groups,daily=days,cases=cases,exact=exact,warnings=warnings,source='مدراء المنافذ + الموظفون تحت مدراء المنافذ + البيانات الخام للإجازات',as_of=REFERENCE_DATE)
    out=Path(OUTPUT_FOLDER) if OUTPUT_FOLDER else Path(LEAVES_WORKBOOK).resolve().parent / 'ports_executive_report'
    out.mkdir(parents=True,exist_ok=True)
    (out/'ports_dashboard_data.json').write_text(json.dumps(result,ensure_ascii=False),encoding='utf-8')
    src=HTML_TEMPLATE
    src=src.replace('<script id="embedded" type="application/json">null</script>', '<script id="embedded" type="application/json">'+json.dumps(result,ensure_ascii=False).replace('</','<\\/')+'</script>')
    (out/'ports_dashboard_READY.html').write_text(src,encoding='utf-8')
    print(f'تم إنشاء: {out / "ports_dashboard_READY.html"}')
    print(f'المدراء والمنفذ: {len(groups)}، الأيام المتقاطعة: {len(days)}، مجموعات تطابق كاملة: {len(exact)}')
    for w in warnings:print('تنبيه:',w)

HTML_TEMPLATE = '<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>تقاطع إجازات المنافذ | Q4 2026</title><style>\n*{box-sizing:border-box}body{margin:0;background:#f4f7f9;color:#1b3145;font:14px/1.65 Tahoma,Arial,sans-serif}header{background:#102c3d;color:white;padding:22px 4%}header h1{margin:0;font-size:24px}header small{color:#bbd3da}.wrap{max-width:1380px;margin:auto;padding:24px}h2{font-size:17px;margin:0 0 12px}h3{font-size:15px;margin:0}.top{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:16px}.sub{color:#71818e;font-size:12px}.btn,input,select{background:#fff;border:1px solid #dce5ec;border-radius:8px;padding:9px;color:#1b3145}.btn{cursor:pointer}.primary{background:#087d80;color:white;border:0}.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin:0 0 20px}.panel{background:#fff;border:1px solid #e2e9ee;border-radius:12px;padding:19px;box-shadow:0 2px 9px #102b3910}.metric small{color:#6c7e8b}.metric b{display:block;font-size:32px;font-weight:750}.metric em{font-style:normal;color:#85939d;font-size:11px}.split{display:grid;grid-template-columns:1.3fr 1fr;gap:14px;margin-bottom:18px}.section{margin-bottom:18px}.sectionhead{display:flex;justify-content:space-between;gap:10px;align-items:center;margin-bottom:10px}.tablewrap{overflow:auto;max-height:430px}table{border-collapse:collapse;width:100%;white-space:nowrap;text-align:right}th,td{padding:12px;border-bottom:1px solid #ebf0f3;font-size:12px}th{background:#f5f8fa;color:#576f7c;position:sticky;top:0}tr:hover td{background:#f8fbfa}.pill{font-size:11px;border-radius:5px;padding:4px 8px;background:#eff3f5}.red{background:#fae8e5;color:#b34b49}.orange{background:#fef0d8;color:#9b671c}.green{background:#def3ec;color:#137c64}.note{background:#edf6f6;border-right:3px solid #168b8a;padding:10px 13px;color:#45606c;font-size:12px;margin:10px 0 20px}.dim{color:#7a8b97}.chart{overflow-x:auto}.chart svg{width:100%;height:auto;min-width:570px}.legend{display:flex;gap:20px;font-size:12px;margin:5px 0 12px;color:#536d7a}.swatch{display:inline-block;width:10px;height:10px;border-radius:3px;margin-left:5px}.noblock{text-align:center;padding:50px 20px;color:#8293a0}.noblock strong{display:block;color:#2c5061}.barback{height:10px;background:#edf2f5;border-radius:8px;overflow:hidden}.barfill{height:100%;background:#0e9790}.rowbar{margin:0 0 16px}.rowbar .line{display:flex;justify-content:space-between;gap:8px;font-size:12px}#exactCards{display:flex;gap:12px;margin:10px 0 14px}#exactCards>div{flex:1;background:#f4f8f9;border-radius:8px;padding:12px}#exactCards b{display:block;font-size:25px}.foot{font-size:11px;color:#778895;margin-top:15px}#summaryLines p{padding:13px;border-right:3px solid #0b9188;background:#f7fafb;margin:8px 0}#summaryLines strong{color:#0b7573}@media(max-width:900px){.grid{grid-template-columns:repeat(2,1fr)}.split{grid-template-columns:1fr}.wrap{padding:12px}}@media print{.buttons,.note{display:none}.panel{box-shadow:none}.section{break-inside:avoid}}\n</style></head><body><header><h1>تقاطع إجازات موظفي المنافذ</h1><small>لوحة تنفيذية لاتخاذ الإجراء · الربع الرابع 2026 · مرجع التقييم: 18 نوفمبر 2026</small></header><div class="wrap"><div class="top"><div><h2>أين تحتاج جدولة الإجازات إلى مراجعة؟</h2><div class="sub">مدراء المنافذ المستهدفون فقط · فترة مستقبلية قابلة للتدخل</div></div><div class="buttons"><label>تاريخ التقييم <input type="date" id="asof" value="2026-11-18"></label> <button class="btn primary" onclick="document.getElementById(\'file\').click()">تحميل بيانات التحليل</button><input type="file" id="file" accept=".json" hidden> <button class="btn" onclick="window.print()">طباعة</button></div></div><div class="note" id="alert">لم يتم إدخال بيانات Excel بعد؛ الأرقام غير معروضة حتى لا نستخدم نتائج افتراضية. استخدمي سكربت تجهيز البيانات ثم حمّلي ملف JSON الناتج هنا.</div><div class="grid"><div class="panel metric"><small>مدراء المنافذ المستهدفون</small><b id="m1">—</b><em>من القائمة المعتمدة</em></div><div class="panel metric"><small>فرق لديها تقاطع قادم</small><b id="m2">—</b><em>في بقية الربع</em></div><div class="panel metric"><small>حالات قريبة عالية الغياب</small><b id="m3">—</b><em>خلال 7 أيام أو جارية؛ 30% فأكثر</em></div><div class="panel metric"><small>أعلى غياب في يوم ذروة قادم</small><b id="m4">—</b><em id="m4s">من إجمالي فريق المدير والمنفذ</em></div></div><section class="section"><div class="sectionhead"><h2>الإجراء المطلوب الآن — أعلى الحالات القادمة</h2><select id="days" class="btn"><option value="999">حتى نهاية الربع</option><option value="7">خلال 7 أيام</option><option value="14">خلال 14 يومًا</option><option value="30">خلال 30 يومًا</option></select></div><div class="panel tablewrap" id="priority"></div><div class="note">ترتيب الحالات بحسب أعلى نسبة غياب متوقعة في يوم واحد. «مراجعة عاجلة» مؤشر فرز أولي وليس حكمًا بوجود عجز تشغيلي.</div></section><section class="split"><div class="panel"><div class="sectionhead"><h2>أسابيع الربع — التقاطعات القابلة للمعالجة</h2><span class="dim">الأسابيع الماضية باهتة</span></div><div class="legend"><span><i class="swatch" style="background:#058c89"></i>مجموعات المدير والمنفذ التي شهدت تقاطعًا</span><span><i class="swatch" style="background:#c87d56"></i>ذروة المجازين في مجموعة واحدة</span></div><div class="chart" id="weeks"></div><div class="foot">عدد المجموعات يُحسب مرة لكل أسبوع، والذروة هي أكبر عدد مجازين في يوم واحد ضمن مدير ومنفذ واحد، وليس مجموع جميع الموظفين.</div></div><div class="panel"><h2>أعلى نسب الغياب المستقبلية</h2><div style="margin-top:19px" id="risk"></div><div class="foot">النسبة = أعلى عدد مجازين في يوم ÷ إجمالي موظفي المدير بالمنفذ نفسه.</div></div></section><section class="split"><div class="panel"><h2>الإجازات المتطابقة بالكامل</h2><div class="sub">موظفون في نفس فريق المدير والمنفذ، يبدأون وينتهون في التاريخين نفسيهما</div><div id="exactCards"></div><div class="tablewrap" id="exact"></div></div><div class="panel"><h2>ماذا يحتاج القرار؟</h2><div id="summaryLines"></div></div></section><section class="section"><div class="sectionhead"><h2>جميع المدراء والمنافذ المشمولين</h2><input id="search" placeholder="بحث باسم مدير أو منفذ"></div><div class="panel tablewrap" id="all"></div></section><div class="foot">المنهجية: يُسجل تقاطع إذا كان هناك موظفان مختلفان بإجازة في اليوم نفسه، تحت المدير والمنفذ نفسيهما. تُستخدم بيانات الإجازات الأحدث لتحديد التبعية عند اختلافها عن ملف الموظفين. نسب القوة العاملة تقديرية عند اختلاف تاريخ تحديث المصادر. لا يمكن الجزم بكفاية التغطية دون بيانات المناوبات والحد الأدنى المطلوب للتشغيل.</div></div><script id="embedded" type="application/json">null</script><script>\nlet D=null;const E=s=>String(s??\'\').replace(/[&<>"\']/g,c=>({\'&\':\'&amp;\',\'<\':\'&lt;\',\'>\':\'&gt;\',\'"\':\'&quot;\',"\'":\'&quot;\'}[c])),N=n=>Number(n||0).toLocaleString(\'en-US\'),P=n=>n==null||!isFinite(n)?\'—\':(n*100).toFixed(1)+\'%\',DAY=(a,b)=>Math.round((Date.parse(a+\'T12:00:00\')-Date.parse(b+\'T12:00:00\'))/864e5),A=()=>document.getElementById(\'asof\').value;function nodata(){return \'<div class="noblock"><strong>لا توجد بيانات حالياً</strong>حمّلي ملف نتائج التحليل لعرض الأرقام الفعلية</div>\'}function T(h,r){return r.length?`<table><thead><tr>${h.map(x=>`<th>${x}</th>`).join(\'\')}</tr></thead><tbody>${r.map(a=>`<tr>${a.map(x=>`<td>${x}</td>`).join(\'\')}</tr>`).join(\'\')}</tbody></table>`:nodata()}\nfunction model(){let ds=(D.daily||[]).filter(x=>x.date>=A()),groups=D.groups||[],gmap=new Map(groups.map(x=>[x.key,x])),peaks=new Map();for(let d of ds){let g=gmap.get(d.key);if(!g||d.on_leave<2)continue;let old=peaks.get(d.key);if(!old||d.ratio>old.ratio||(d.ratio===old.ratio&&d.date<old.date))peaks.set(d.key,{...d,manager:g.manager,port:g.port,workforce:g.workforce})}return {ds,groups,peaks:[...peaks.values()],cases:(D.cases||[]).filter(x=>x.end>=A())}}\nfunction render(){if(!D){[\'priority\',\'weeks\',\'risk\',\'exact\',\'summaryLines\',\'all\'].forEach(x=>document.getElementById(x).innerHTML=nodata());return}let m=model();document.getElementById(\'alert\').innerText=\'البيانات محملة من ملفات العمل. يمكن تغيير تاريخ التقييم لإعادة احتساب الحالات القادمة والأسابيع الماضية.\';document.getElementById(\'m1\').innerText=N(m.groups.length);document.getElementById(\'m2\').innerText=N(new Set(m.cases.map(x=>x.key)).size);document.getElementById(\'m3\').innerText=N(m.cases.filter(x=>DAY(x.start,A())<=7&&x.ratio>=.3).length);let max=m.peaks.filter(x=>x.ratio!=null).sort((a,b)=>b.ratio-a.ratio)[0];document.getElementById(\'m4\').innerText=max?P(max.ratio):\'—\';document.getElementById(\'m4s\').innerText=max?max.manager+\' · \'+max.port:\'من إجمالي موظفي الفريق\';\nlet filtered=m.cases.filter(x=>DAY(x.start,A())<=+document.getElementById(\'days\').value).sort((a,b)=>(b.ratio??-1)-(a.ratio??-1)||a.start.localeCompare(b.start));document.getElementById(\'priority\').innerHTML=T([\'المدير\',\'المنفذ\',\'بداية التقاطع\',\'قرب البداية\',\'مدة الحالة\',\'ذروة المجازين\',\'تاريخ الذروة\',\'إجمالي الفريق\',\'غير المجازين\',\'نسبة الغياب\',\'التصنيف\'],filtered.slice(0,30).map(x=>[E(x.manager),E(x.port),x.start,DAY(x.start,A())<0?\'<span class="pill red">جارية</span>\':DAY(x.start,A())===0?\'اليوم\':\'بعد \'+N(DAY(x.start,A()))+\' أيام\',x.duration+\' يوم\',N(x.peak_on_leave),x.peak_date,N(x.workforce),N(Math.max(0,x.workforce-x.peak_on_leave)),`<b>${P(x.ratio)}</b>`,x.ratio==null?\'<span class="pill">تحقق من القوة العاملة</span>\':x.ratio>=.5?\'<span class="pill red">عاجلة</span>\':x.ratio>=.3?\'<span class="pill orange">مرتفعة</span>\':\'<span class="pill green">متابعة</span>\']));\nlet weeks=[];for(let ts=Date.parse(\'2026-09-28T12:00:00\');ts<=Date.parse(\'2026-12-31T12:00:00\');ts+=7*864e5){let s=new Date(ts).toISOString().slice(0,10),e=new Date(ts+6*864e5).toISOString().slice(0,10),items=(D.daily||[]).filter(x=>x.date>=s&&x.date<=e&&x.on_leave>=2);weeks.push({s,e,g:new Set(items.map(x=>x.key)).size,p:Math.max(0,...items.map(x=>x.on_leave))})}let width=860,height=250,min=25,maxY=200,top=15,scale=Math.max(1,...weeks.map(x=>Math.max(x.g,x.p))),col=width/weeks.length;let svg=`<svg viewBox="0 0 ${width} ${height}" aria-label="توزيع أسبوعي للتقاطعات">`;for(let j=0;j<=4;j++){let y=maxY-j*maxY/4;svg+=`<line x1="10" x2="${width-4}" y1="${y}" y2="${y}" stroke="#dce5eb"/><text x="${width-5}" y="${y-5}" text-anchor="end" font-size="9" fill="#7d8b92">${Math.round(scale*j/4)}</text>`}weeks.forEach((w,i)=>{let x=i*col+10,dim=w.e<A()?.22:1,h1=w.g/scale*(maxY-top),h2=w.p/scale*(maxY-top);svg+=`<rect x="${x}" y="${maxY-h1}" width="${col*.28}" height="${h1}" fill="#06928c" opacity="${dim}"><title>الأسبوع ${i+1}: ${w.g} مجموعة</title></rect><rect x="${x+col*.31}" y="${maxY-h2}" width="${col*.28}" height="${h2}" fill="#cc8260" opacity="${dim}"><title>ذروة ${w.p} مجازاً</title></rect><text x="${x+col*.3}" y="${height-23}" text-anchor="middle" font-size="10" fill="#536774">${i+1}</text>`});document.getElementById(\'weeks\').innerHTML=svg+\'</svg>\';\nlet peaks=m.peaks.filter(x=>x.ratio!=null).sort((a,b)=>b.ratio-a.ratio);document.getElementById(\'risk\').innerHTML=peaks.length?peaks.slice(0,7).map(x=>`<div class="rowbar"><div class="line"><b>${E(x.manager)} · ${E(x.port)}</b><strong>${P(x.ratio)}</strong></div><div class="barback"><div class="barfill" style="width:${Math.min(100,x.ratio*100)}%;background:${x.ratio>=.5?\'#c46f5a\':\'#0e9790\'}"></div></div><div class="dim" style="font-size:11px">${N(x.on_leave)} من ${N(x.workforce)} · ${x.date}</div></div>`).join(\'\'):nodata();\nlet exact=(D.exact||[]).filter(x=>x.end>=A()),one=exact.filter(x=>x.duration===1),long=exact.filter(x=>x.duration>1);document.getElementById(\'exactCards\').innerHTML=`<div>تطابق ليوم واحد<b>${N(one.length)}</b></div><div>تطابق لأكثر من يوم<b>${N(long.length)}</b></div>`;document.getElementById(\'exact\').innerHTML=T([\'المدير\',\'المنفذ\',\'البداية\',\'النهاية\',\'عدد الموظفين\'],exact.sort((a,b)=>b.duration-a.duration).slice(0,8).map(x=>[E(x.manager),E(x.port),x.start,x.end,N(x.employee_count)]));let high=peaks.filter(x=>x.ratio>=.5),soon=m.cases.filter(x=>DAY(x.start,A())<=7),same=exact.filter(x=>x.duration>=5);document.getElementById(\'summaryLines\').innerHTML=`<p><strong>${N(high.length)} فرق</strong> تبلغ ذروة الغياب المتوقعة فيها 50% أو أكثر من القوة العاملة.</p><p><strong>${N(soon.length)} حالات تقاطع</strong> تبدأ خلال أسبوع من تاريخ التقييم أو ما زالت جارية.</p><p><strong>${N(same.length)} مجموعة تطابق كامل</strong> لإجازات مدتها 5 أيام فأكثر وما زالت جارية أو قادمة.</p><div class="note">التوصية: مراجعة هذه الفرق أولاً مع مسؤولي المنافذ، والتأكد من خطة التغطية قبل تاريخ بدء الإجازات.</div>`;\nlet q=document.getElementById(\'search\').value.trim(),g=m.groups.filter(x=>(x.manager+\' \'+x.port).includes(q));document.getElementById(\'all\').innerHTML=T([\'المدير\',\'المنفذ\',\'عدد موظفي الفريق\',\'أعلى مجازين قادم\',\'تاريخ الذروة\',\'النسبة\'],g.map(x=>{let p=m.peaks.find(t=>t.key===x.key);return [E(x.manager),E(x.port),N(x.workforce),p?N(p.on_leave):\'0\',p?.date||\'—\',p?P(p.ratio):\'—\']}))}\n[\'asof\',\'days\',\'search\'].forEach(id=>document.getElementById(id).addEventListener(id===\'search\'?\'input\':\'change\',render));document.getElementById(\'file\').addEventListener(\'change\',async e=>{try{let x=JSON.parse(await e.target.files[0].text());if(!Array.isArray(x.groups)||!Array.isArray(x.daily))throw Error(\'صيغة الملف غير صحيحة\');D=x;render()}catch(err){alert(err.message)}});try{let a=JSON.parse(document.getElementById(\'embedded\').textContent);if(a?.groups)D=a}catch{}render();</script></body></html>\n'

if __name__=='__main__':build()

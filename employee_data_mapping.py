# -*- coding: utf-8 -*-
from pathlib import Path
from datetime import date, datetime
from calendar import monthrange
from collections import defaultdict, Counter
import re
import unicodedata
import sys
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel
from openpyxl.worksheet.table import Table, TableStyleInfo

# ======================= الإعدادات =======================
TARGET_FILE = r'C:\Users\YOUR_USER\Downloads\Employees.xlsx'
TARGET_SHEET = 'Sheet1'
MASTER_FILE = r'C:\Users\YOUR_USER\Downloads\Employee Master.xlsx'
MASTER_SHEET = 'Sheet1'
OUTPUT_FILE = ''  # فارغ = بجانب الملف الأساسي
REPORT_DATE = ''  # فارغ = تاريخ الجهاز؛ أو مثلاً '2026-11-24'
TEXT_DATE_FORMAT = '%m/%d/%Y'  # للتواريخ النصية غير ISO فقط؛ غيّريه إذا كان المصدر يوم/شهر/سنة
# =========================================================

MISSING = {'', 'na', 'n/a', '#n/a', 'nan', 'none', 'null', '-', 'غير متاح', 'غير محدد', 'not applicable'}

def val(v):
    if v is None: return ''
    if isinstance(v, str):
        v = v.strip()
        return '' if v.casefold() in MISSING else v
    return v

def norm(v):
    s = unicodedata.normalize('NFKC', str(val(v))).casefold()
    s = re.sub(r'[\u064b-\u065f\u0670\u0640\u200b-\u200f\u061c\ufeff]', '', s)
    return re.sub(r'\s+', ' ', s).strip()

def namekey(v):
    return norm(v).translate(str.maketrans('أإآٱ','اااا'))

def idkey(v):
    s = norm(v)
    s = s.translate(str.maketrans('٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹','01234567890123456789'))
    if re.fullmatch(r'\d+\.0+', s): s = s.split('.')[0]
    # لا تُحذف الأصفار في بداية الرقم الوظيفي النصي.
    return s

def read(path, sheet):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[sheet]
        it = ws.iter_rows(values_only=True)
        headers = list(next(it))
        while headers and headers[-1] is None: headers.pop()
        rows = [(i, tuple(r[:len(headers)])) for i,r in enumerate(it,2) if any(x is not None and str(x).strip() for x in r[:len(headers)])]
        return headers, rows, wb.epoch
    finally: wb.close()

def column(headers, aliases, required=True):
    for alias in aliases:
        hits = [i for i,h in enumerate(headers) if norm(h)==norm(alias)]
        if len(hits)>1: raise ValueError(f'عنوان عمود مكرر: {alias}')
        if hits: return hits[0]
    if required: raise ValueError('عمود غير موجود: '+ ' / '.join(aliases))
    return None

def get(r,i):
    return val(r[i]) if i is not None and i<len(r) else ''

def parse_date(v, epoch):
    if v == '': return None
    if isinstance(v,datetime): return v.date()
    if isinstance(v,date): return v
    if isinstance(v,(int,float)) and not isinstance(v,bool):
        try:
            d=from_excel(v,epoch)
            return d.date() if isinstance(d,datetime) else None
        except (ValueError,OverflowError): return None
    s=str(v).strip()
    try: return date.fromisoformat(s)
    except ValueError: pass
    try: return datetime.strptime(s,TEXT_DATE_FORMAT).date()
    except ValueError: return None

def elapsed(start,end):
    if start is None or start>end: return None
    months=(end.year-start.year)*12+end.month-start.month
    y,m=divmod(start.year*12+start.month-1+months,12)
    anniversary=date(y,m+1,min(start.day,monthrange(y,m+1)[1]))
    if anniversary>end: months-=1
    return divmod(months,12)

def duration_text(parts):
    if parts is None: return ''
    y,m=parts
    ys='سنة' if y==1 else 'سنتان' if y==2 else f'{y} سنوات' if 3<=y<=10 else f'{y} سنة'
    ms='شهر' if m==1 else 'شهران' if m==2 else f'{m} أشهر' if 3<=m<=10 else f'{m} شهرًا'
    s=' و'.join(x for x in [ys if y else '',ms if m else ''] if x) or 'أقل من شهر'
    # اتجاه عربي ثابت + أرقام عربية هندية + مسافة غير قابلة للكسر.
    return '\u200f'+s.translate(str.maketrans('0123456789','٠١٢٣٤٥٦٧٨٩')).replace(' ','\u00a0')+'\u200f'

FIELDS=['القطاع','الإدارة العامة / الإدارة','المسمى الوظيفي','سنوات الخبرة','العمر','سنوات الخبرة في المنصب الحالي']

def main():
    asof=date.fromisoformat(REPORT_DATE) if REPORT_DATE else date.today()
    th,tr,_=read(TARGET_FILE,TARGET_SHEET)
    mh,mr,epoch=read(MASTER_FILE,MASTER_SHEET)
    ti=column(th,['الرقم الوظيفي','Employee ID']); tn=column(th,['الموظف','الاسم الرسمي بالعربي','اسم الموظف'])
    mi=column(mh,['الرقم الوظيفي','Employee ID','Person ID']); mn=column(mh,['الاسم الرسمي بالعربي','اسم الموظف','الموظف'])
    sector=column(mh,['اسم القطاع باللغة العربية','اسم القطاع باللغه العربية','اسم القطاع باللغة العربيه'],False)
    sector_fallback=column(mh,['المستوى الأول','Level 1 (Label)'],False)
    if sector is None and sector_fallback is None: raise ValueError('لم يوجد عمود القطاع')
    job_ar=column(mh,['المسمى الوظيفي عربي','المسمى الوظيفي بالعربي','المسمى الوظيفي باللغة العربية','المسمى الوظيفي'],False)
    job=column(mh,['Job Title'],job_ar is None)
    levels=[]
    for n in range(6,1,-1):
        ar=column(mh,[f'المستوى {n} عربي',f'Level {n} Arabic',f'Level {n} (Arabic)'],False)
        raw=column(mh,[f'Level {n} (Label)'],False)
        if ar is not None or raw is not None: levels.append((n,ar,raw))
    if not levels: raise ValueError('لم توجد أعمدة المستويات التنظيمية 2–6')
    hire=column(mh,['تاريخ المباشرة']); birth=column(mh,['Date of Birth','تاريخ الميلاد']); entry=column(mh,['Job Entry Date'],False)
    if entry is None: print('تنبيه: Job Entry Date غير موجود؛ سيترك عمود خبرة المنصب فارغًا.')
    byid=defaultdict(list); byname=defaultdict(list)
    for n,r in mr:
        if idkey(get(r,mi)): byid[idkey(get(r,mi))].append((n,r))
        if namekey(get(r,mn)): byname[namekey(get(r,mn))].append((n,r))
    counts=Counter(); issues=[]; output=[]; targetids=Counter(idkey(get(r,ti)) for _,r in tr if idkey(get(r,ti)))
    def note(n,msg): issues.append(f'صف {n}: {msg}')
    def project(r):
        sec=get(r,sector) or get(r,sector_fallback)
        department=next((get(r,a) or get(r,b) for _,a,b in levels if get(r,a) or get(r,b)), '')
        title=get(r,job_ar) or get(r,job)
        return (sec,department,title,parse_date(get(r,hire),epoch),parse_date(get(r,birth),epoch),parse_date(get(r,entry),epoch))
    for n,r in tr:
        eid=idkey(get(r,ti)); name=namekey(get(r,tn)); result=['']*6
        matches=[]; method=''
        if eid:
            if targetids[eid]>1: note(n,'رقم وظيفي مكرر في القائمة الأساسية؛ تم الاحتفاظ بالسجل.')
            hits=byid.get(eid,[])
            if not hits: note(n,'الرقم الوظيفي غير موجود في المصدر؛ لم تتم المطابقة بالاسم بدلًا منه.')
            elif not name: note(n,'الاسم فارغ؛ تعذر التحقق الثنائي.')
            elif any(namekey(get(x,mn))!=name for _,x in hits): note(n,'الرقم الوظيفي موجود ولكن الاسم مختلف أو فارغ في أحد سجلات المصدر؛ لم تُنقل البيانات.')
            else: matches=hits; method='مطابقة الرقم والاسم'
        elif name:
            hits=byname.get(name,[])
            ids={idkey(get(x,mi)) for _,x in hits}
            if len(ids)==1 and '' not in ids and hits and all(namekey(get(x,mn))==name for _,x in byid[next(iter(ids))]):
                matches=byid[next(iter(ids))]; method='مطابقة الاسم فقط'; note(n,'الرقم الوظيفي مفقود؛ تمت مطابقة الاسم الفريد، يلزم التأكيد قبل الاعتماد.')
            else: note(n,'لا يوجد رقم وظيفي والاسم غير موجود أو غير فريد/متعارض.')
        else: note(n,'الرقم الوظيفي والاسم غير متاحين.')
        if matches:
            projections={project(x) for _,x in matches}
            if len(projections)!=1:
                note(n,'سجلات متعددة لنفس الموظف بقيم مختلفة؛ لم تُنقل البيانات.')
            else:
                p=next(iter(projections)); result=list(p[:3])+[duration_text(elapsed(d,asof)) for d in p[3:]]
                counts[method]+=1
                if len(matches)>1: note(n,'سجلات مصدر مكررة متفقة في البيانات المطلوبة؛ استُخدمت مرة واحدة.')
                if not get(matches[0][1],sector) and result[0]: note(n,'استُخدم المستوى الأول للقطاع لعدم توفر الاسم العربي؛ راجعيه.')
                for field,v in zip(FIELDS,result):
                    if v=='': note(n,f'{field}: غير متاح أو تاريخ غير صالح/مستقبلي.')
                if all(v!='' for v in result): counts['مكتملة']+=1
                else: counts['مطابقة ببيانات ناقصة']+=1
        if not any(v!='' for v in result) and not (matches and len({project(x) for _,x in matches})==1): counts['غير محسومة']+=1
        output.append(list(r)+result)
    if len(output)!=len(tr): raise RuntimeError('عدد الصفوف غير مطابق')
    out=Path(OUTPUT_FILE) if OUTPUT_FILE else Path(TARGET_FILE).resolve().parent/'Employee_Data_Completed.xlsx'
    if out.resolve() in {Path(TARGET_FILE).resolve(),Path(MASTER_FILE).resolve()}: raise ValueError('مسار الناتج يجب أن يختلف عن المصادر')
    if out.suffix.lower()!='.xlsx': raise ValueError('الناتج يجب أن يكون xlsx')
    headers=[str(h or '').strip() for h in th]+FIELDS
    if any(not h for h in headers) or len(set(headers))!=len(headers): raise ValueError('عناوين الملف الأساسي فارغة أو مكررة')
    wb=Workbook(); ws=wb.active; ws.title='بيانات الموظفين'; ws.sheet_view.rightToLeft=True
    ws.append(headers)
    for row in output:
        ws.append(row)
        for cell in ws[ws.max_row]:
            if isinstance(cell.value,str): cell.data_type='s'
            cell.alignment=Alignment(horizontal='right',vertical='center',readingOrder=2)
    for c in ws[1]: c.font=Font(bold=True); c.alignment=Alignment(horizontal='right',readingOrder=2)
    if output:
        table=Table(displayName='EmployeeData',ref=f'A1:{get_column_letter(len(headers))}{ws.max_row}')
        table.tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=False,showColumnStripes=False)
        ws.add_table(table)
    ws.freeze_panes='A2'
    for i,h in enumerate(headers,1): ws.column_dimensions[get_column_letter(i)].width=min(48,max(18,len(h)+3, max((len(str(row[i-1] or ''))+2 for row in output),default=0)))
    wb.save(out)
    print('\n===== نتيجة المطابقة =====')
    print('تاريخ الاحتساب:',asof)
    print('السجلات الأصلية:',len(tr),'| الناتج:',len(output))
    for k in ['مطابقة الرقم والاسم','مطابقة الاسم فقط','مكتملة','مطابقة ببيانات ناقصة','غير محسومة']: print(f'{k}: {counts[k]}')
    print('الملاحظات:',len(issues))
    for msg in issues: print('-',msg)
    print('تم حفظ التقرير:',out)

if __name__=='__main__':
    try: main()
    except Exception as e:
        print('خطأ:',e,file=sys.stderr); sys.exit(1)

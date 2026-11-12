# -*- coding: utf-8 -*-
"""
تحليل جانبي مستقل لإجازات الموظفين تحت مدراء «وحدة التفتيش الجمركي».
لا يغيّر التقرير التنفيذي الأصلي.
التشغيل: pip install pandas openpyxl
ثم: python customs_inspection_leave_check.py
"""
from pathlib import Path
from datetime import date, datetime, timedelta
from collections import defaultdict, Counter
import re, unicodedata
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# غيّري مسارات الملفات فقط إذا اختلفت عندك
MANAGERS_FILE = r"C:\Users\rammutairi\Downloads\Employee Master.xlsx"
LEAVES_FILE = r"C:\Users\rammutairi\Downloads\Ports Leaves.xlsx"
MANAGERS_SHEET = "مدراء المنافذ"
EMPLOYEES_SHEET = "الموظفين تحت مدراء المنافذ"
LEAVES_SHEET = "البيانات الخام"
OUTPUT_FILE = ""  # فارغ = الحفظ بجانب ملف الإجازات

# اتركيها فارغة لتحليل كل التواريخ المتاحة، أو حددي نطاقًا YYYY-MM-DD
PERIOD_START = ""
PERIOD_END = ""
END_DATE_INCLUSIVE = True  # أكدي أن تاريخ نهاية الإجازة هو آخر يوم إجازة
ALLOW_NO_APPROVAL_COLUMN = True  # استخدمي ملفًا يحوي الإجازات المعتمدة فقط إن لم يوجد عمود حالة
JOB_TITLE_TARGET = "مدير وحدة التفتيش الجمركي"

ALIASES = {
 "id": ["الرقم الوظيفي", "Employee ID", "(Employee) Person ID", "Person ID", "Employee Number", "Emp ID"],
 "name": ["الاسم الرسمي بالعربي", "اسم الموظف", "Employee Name", "(Employee) Full Name"],
 "manager_id": ["الرقم الوظيفي للمدير المباشر", "Supervisor ID", "Manager Employee ID"],
 "location": ["Location (Name)", "(Employee) Location (Name)", "موقع العمل", "Location"],
 "job": ["Job Title", "المسمى الوظيفي"],
 "start": ["Actual Start Date", "Start Date", "تاريخ بداية الإجازة", "تاريخ البداية"],
 "end": ["Actual End Date", "End Date", "تاريخ نهاية الإجازة", "تاريخ النهاية"],
 "approval": ["Approval Status", "Approval Status (Name)", "حالة الاعتماد", "Approval Status (Label)"],
 "lvl1": ["المستوى الأول", "Level 1", "Level 1 (Name)"],
 "lvl2": ["المستوى الثاني", "Level 2", "Level 2 (Name)"],
 "lvl3": ["المستوى الثالث", "Level 3", "Level 3 (Name)"],
}
APPROVED = {"approved", "approved request", "معتمد", "معتمدة", "موافق", "موافق عليه", "تم الاعتماد"}

def normalize(x):
    if x is None or pd.isna(x): return ""
    s = unicodedata.normalize("NFKC", str(x)).strip().casefold()
    s = re.sub(r"[\u064b-\u065f\u0670\u0640\u200e\u200f\u200b]", "", s)
    for a,b in (("أ","ا"),("إ","ا"),("آ","ا"),("ى","ي"),("ة","ه")): s=s.replace(a,b)
    return re.sub(r"\s+", " ", s).strip()

def emp_id(x):
    if x is None or pd.isna(x): return ""
    s = str(x).strip()
    return s[:-2] if re.fullmatch(r"\d+\.0",s) else s

def getcol(df, key, required=False):
    names = {normalize(c):c for c in df.columns}
    for candidate in ALIASES[key]:
        if normalize(candidate) in names: return names[normalize(candidate)]
    if required: raise ValueError(f"العمود المطلوب غير موجود ({key}). الأعمدة الموجودة: {list(df.columns)}")
    return None

def value(row, col):
    if not col: return ""
    v=row[col]
    return "" if pd.isna(v) else str(v).strip()

def parse_date(x):
    if x is None or pd.isna(x) or str(x).strip()=="":return None
    if isinstance(x,(pd.Timestamp,datetime,date)):return x.date() if isinstance(x,(pd.Timestamp,datetime)) else x
    if isinstance(x,(float,int)) and 25000<float(x)<90000:
        return (datetime(1899,12,30)+timedelta(days=int(x))).date()
    s=str(x).strip()
    if re.fullmatch(r"\d{5}(?:\.0)?",s) and 25000<float(s)<90000:
        return (datetime(1899,12,30)+timedelta(days=int(float(s)))).date()
    d=pd.to_datetime(s,errors="coerce",dayfirst=not bool(re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}",s)))
    return None if pd.isna(d) else d.date()

def read_excel(path,sheet):
    if not Path(path).is_file():raise FileNotFoundError(f"لم أجد الملف: {path}")
    df=pd.read_excel(path,sheet_name=sheet,dtype=object,engine="openpyxl")
    df.columns=[str(c).strip() for c in df.columns]
    return df.dropna(how="all")

def main():
    managers=read_excel(MANAGERS_FILE,MANAGERS_SHEET)
    employees=read_excel(MANAGERS_FILE,EMPLOYEES_SHEET)
    leaves=read_excel(LEAVES_FILE,LEAVES_SHEET)
    mi=getcol(managers,"id",True); mn=getcol(managers,"name",True); mj=getcol(managers,"job",True)
    ei=getcol(employees,"id",True); em=getcol(employees,"manager_id",True)
    en=getcol(employees,"name"); el=getcol(employees,"location",True)
    levels=[getcol(employees,f"lvl{i}") for i in (1,2,3)]
    li=getcol(leaves,"id",True); ls=getcol(leaves,"start",True); le=getcol(leaves,"end",True)
    la=getcol(leaves,"approval"); ll=getcol(leaves,"location")
    if not la and not ALLOW_NO_APPROVAL_COLUMN:raise ValueError("لا يوجد عمود حالة اعتماد في مصدر الإجازات")
    selected={}
    for _,r in managers.iterrows():
        if normalize(value(r,mj)) == normalize(JOB_TITLE_TARGET):
            mid=emp_id(value(r,mi))
            if mid: selected[mid]={"name":value(r,mn),"job":value(r,mj)}
    if not selected:
        titles=sorted(set(str(x) for x in managers[mj].dropna()))
        raise ValueError("لم أجد المسمى المطلوب بشكل مطابق. راجعي JOB_TITLE_TARGET. المسميات المتاحة:\n"+"\n".join(titles[:70]))
    print(f"مدراء التفتيش الجمركي المطابقون للمسمى: {len(selected)}")
    # المجموعة = مدير + موقع + المستويات 1–3. لا ندمج الوحدات التنظيمية المختلفة.
    groups=defaultdict(set); staff={}; conflicts=[]; qc=Counter()
    for _,r in employees.iterrows():
        mid=emp_id(value(r,em))
        if mid not in selected:continue
        eid=emp_id(value(r,ei)); loc=value(r,el)
        if not eid or not loc:
            qc["موظف بلا رقم أو موقع"]+=1;continue
        org=tuple(value(r,c) for c in levels)
        key=(mid,normalize(loc),*(normalize(v) for v in org))
        if eid in staff and staff[eid]["key"]!=key:
            conflicts.append(["تبعية متضاربة",eid,staff[eid]["location"],loc])
            qc["موظف بتبعية متضاربة"]+=1
            # لا نعد الموظف مرتين؛ يستبعد من كل المجموعات لحين التصحيح
            groups[staff[eid]["key"]].discard(eid)
            staff[eid]["invalid"]=True
            continue
        if eid in staff and staff[eid].get("invalid"):continue
        staff[eid]={"key":key,"name":value(r,en) or eid,"location":loc,"org":org}
        groups[key].add(eid)
    staff={eid:info for eid,info in staff.items() if not info.get("invalid")}
    groups={g:ids for g,ids in groups.items() if ids}
    if not groups:raise ValueError("لم أجد موظفين تابعين مباشرةً لهؤلاء المدراء. راجعي رقم المدير المباشر.")
    print(f"الموظفون المطابقون: {len(staff)} | المجموعات: {len(groups)}")
    pstart=date.fromisoformat(PERIOD_START) if PERIOD_START else date.min
    pend=date.fromisoformat(PERIOD_END) if PERIOD_END else date.max
    daily=defaultdict(set); by_employee=defaultdict(list); seen=set()
    for _,r in leaves.iterrows():
        eid=emp_id(value(r,li))
        if eid not in staff:continue
        st=parse_date(r[ls]); end=parse_date(r[le])
        if not st or not end or end<st:
            qc["إجازة بتاريخ غير صالح"]+=1;continue
        if la:
            approval=normalize(value(r,la))
            if approval and approval not in {normalize(x) for x in APPROVED}:
                qc["إجازة غير معتمدة/حالة غير معروفة"]+=1;continue
        if ll and value(r,ll) and normalize(value(r,ll))!=staff[eid]["key"][1]:
            qc["اختلاف موقع الموظف بين الملفين"]+=1
            conflicts.append(["اختلاف موقع",eid,staff[eid]["location"],value(r,ll)])
            continue
        actual_end=end-timedelta(days=1) if not END_DATE_INCLUSIVE else end
        a=max(st,pstart);b=min(actual_end,pend)
        if b<a:continue
        sig=(eid,st,end)
        if sig in seen:
            qc["إجازة مكررة تمامًا"]+=1;continue
        seen.add(sig)
        by_employee[eid].append((st,end,a,b))
        for offset in range((b-a).days+1):
            daily[(staff[eid]["key"],a+timedelta(days=offset))].add(eid)
    # فترات التقاطع: يومان أو أكثر من الموظفين (لا نجمع إجازات منفصلة)
    # فترة متصلة مع تغير التشكيلة تُعد فترة تقاطع واحدة، وتُحسب الذروة داخلها منفصلة.
    allcases=[];summary=[];employee_rows=[]
    today=date.today()
    for g,ids in sorted(groups.items()):
        mid=g[0]; manager=selected[mid]; sample=staff[sorted(ids)[0]]
        events=sorted((d,s) for (key,d),s in daily.items() if key==g)
        active=[(d,s) for d,s in events if len(s)>=2]
        runs=[]
        for d,s in active:
            if runs and d==runs[-1][-1][0]+timedelta(days=1):
                runs[-1].append((d,s))
            else:runs.append([(d,s)])
        overlap_ids=set().union(*(s for _,s in active)) if active else set()
        leave_ids={eid for eid in ids if by_employee.get(eid)}
        peak=max((len(s) for _,s in events),default=0)
        if not leave_ids: status="لا توجد إجازات مسجلة"
        elif not active:status="إجازات غير متقاطعة"
        else:status="توجد إجازات متقاطعة"
        base=[manager["name"],mid,manager["job"],sample["location"],*sample["org"]]
        summary.append(base+[len(ids),len(leave_ids),len(ids-leave_ids),len(overlap_ids),len(runs),peak,round(100*peak/len(ids),2),status])
        for eid in sorted(ids):
            intervals=by_employee.get(eid,[])
            if not intervals:
                employee_rows.append(base+[staff[eid]["name"],eid,"لا توجد إجازة مسجلة",None,None,0])
            for st,en,a,b in sorted(intervals):
                is_overlap=any((eid in daily.get((g,a+timedelta(days=i)),set()) and len(daily.get((g,a+timedelta(days=i)),set()))>=2) for i in range((b-a).days+1))
                employee_rows.append(base+[staff[eid]["name"],eid,"متقاطعة" if is_overlap else "غير متقاطعة",st,en,(b-a).days+1])
        for run in runs:
            start=run[0][0];end=run[-1][0]
            maxabs=max(len(s) for _,s in run)
            peak_dates=[d for d,s in run if len(s)==maxabs]
            union_ids=set().union(*(s for _,s in run))
            if end<today:period_state="منتهية"
            elif start>today:period_state="قادمة"
            else:period_state="جارية"
            names="\n".join(f"{i}- {staff[eid]['name']} ({eid})" for i,eid in enumerate(sorted(union_ids),1))
            allcases.append(base+[period_state,start,end,(end-start).days+1,maxabs,round(100*maxabs/len(ids),2),
                                  min(peak_dates),max(peak_dates),len(union_ids),names])
    summary.sort(key=lambda r:(r[-1]!="توجد إجازات متقاطعة",-r[-2],r[0],r[3]))
    allcases.sort(key=lambda r:({"جارية":0,"قادمة":1,"منتهية":2}[r[7]],-r[12],-r[11],r[8]))
    out=Path(OUTPUT_FILE) if OUTPUT_FILE else Path(LEAVES_FILE).resolve().parent/"تحليل_جانبي_مدراء_وحدات_التفتيش_الجمركي.xlsx"
    out.parent.mkdir(parents=True,exist_ok=True)
    with pd.ExcelWriter(out,engine="openpyxl") as writer:
        pd.DataFrame(summary,columns=["اسم المدير","الرقم الوظيفي للمدير","المسمى الوظيفي","موقع العمل","المستوى الأول","المستوى الثاني","المستوى الثالث","إجمالي الموظفين","موظفون لديهم إجازات","موظفون دون إجازات مسجلة","موظفون شاركوا في تقاطع","عدد فترات التقاطع","أعلى عدد مجازين في يوم","أعلى نسبة غياب يومي %","نتيجة التحليل"]).to_excel(writer,sheet_name="ملخص مدراء التفتيش",index=False)
        pd.DataFrame(allcases,columns=["اسم المدير","الرقم الوظيفي للمدير","المسمى الوظيفي","موقع العمل","المستوى الأول","المستوى الثاني","المستوى الثالث","حالة الفترة","بداية التقاطع","نهاية التقاطع","مدة التقاطع بالأيام","ذروة المجازين في يوم","نسبة الذروة %","أول يوم ذروة","آخر يوم ذروة","موظفون شاركوا خلال الفترة","قائمة الموظفين المشاركين"]).to_excel(writer,sheet_name="فترات التقاطع",index=False)
        pd.DataFrame(employee_rows,columns=["اسم المدير","الرقم الوظيفي للمدير","المسمى الوظيفي","موقع العمل","المستوى الأول","المستوى الثاني","المستوى الثالث","اسم الموظف","الرقم الوظيفي","حالة الإجازة","بداية الإجازة","نهاية الإجازة","أيام الإجازة ضمن النطاق"]).to_excel(writer,sheet_name="جميع الموظفين والإجازات",index=False)
        pd.DataFrame([["تاريخ إعداد التقرير",today.isoformat()],["المسمى المستهدف",JOB_TITLE_TARGET],
          ["عدد المدراء المستهدفين",len(selected)],["الموظفون المطابقون",len(staff)],
          ["المجموعات المستقلة",len(groups)],["إجمالي سجلات الإجازات المطابقة",len(seen)],
          ["مجموعات بها تقاطع",sum(r[-1]=="توجد إجازات متقاطعة" for r in summary)],
          ["مجموعات بإجازات دون تقاطع",sum(r[-1]=="إجازات غير متقاطعة" for r in summary)],
          ["مجموعات بلا إجازات",sum(r[-1]=="لا توجد إجازات مسجلة" for r in summary)],
          ["تعريف المجموعة","مدير مباشر + موقع العمل + المستوى الأول والثاني والثالث"],
          ["تعريف التقاطع","موظفان مختلفان أو أكثر بإجازة في اليوم نفسه داخل المجموعة"],
          ["مدة التقاطع","كل الأيام المتصلة التي يوجد فيها اثنان أو أكثر، حتى لو تغيّر الأشخاص"],
          ["الذروة","أعلى عدد موظفين فريدين مجازين في يوم واحد ضمن الفترة"],
          ["تنبيه","نسبة الغياب المجدول لا تعني بالضرورة عجزًا تشغيليًا"],
          ["تنبيه اعتماد","عند غياب عمود الاعتماد يفترض أن ملف المصدر يضم إجازات معتمدة فقط"]]+
          [[k,v] for k,v in sorted(qc.items())]+
          [["ملاحظة سجل", " | ".join(map(str,r))] for r in conflicts[:200]],
          columns=["المؤشر","القيمة"]).to_excel(writer,sheet_name="ضبط الجودة",index=False)
    wb=load_workbook(out)
    for ws in wb:
        ws.sheet_view.rightToLeft=True
        ws.freeze_panes="D2" if ws.max_column>3 else "A2"
        ws.auto_filter.ref=ws.dimensions
        ws.row_dimensions[1].height=38
        for c in ws[1]:
            c.fill=PatternFill("solid",fgColor="173B5B")
            c.font=Font(name="Arial",size=11,bold=True,color="FFFFFF")
            c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        for row in ws.iter_rows(min_row=2):
            gray=(ws.title=="فترات التقاطع" and row[7].value=="منتهية")
            for c in row:
                c.font=Font(name="Arial",size=10,color="5C6773" if gray else "172A3A")
                c.fill=PatternFill("solid",fgColor="E6E8EB" if gray else ("F0F5F8" if c.row%2==0 else "FFFFFF"))
                c.alignment=Alignment(vertical="top",wrap_text=True)
            if ws.title=="فترات التقاطع":
                n=str(row[16].value or "").count("\n")+1
                ws.row_dimensions[row[0].row].height=min(390,max(28,n*16))
        for j in range(1,ws.max_column+1):
            header=str(ws.cell(1,j).value or "")
            width=56 if "قائمة الموظفين" in header else (38 if "المسمى" in header or "المؤشر" in header else 25)
            ws.column_dimensions[get_column_letter(j)].width=width
        for row in ws.iter_rows(min_row=2):
            for c in row:
                if isinstance(c.value,(datetime,date)):c.number_format="dd/mm/yyyy"
    wb.save(out)
    print("تم إنشاء:",out)
    print("المدراء:",len(selected),"| المجموعات:",len(groups),"| فترات التقاطع:",len(allcases))
    print("مجموعات بها تقاطع:",sum(r[-1]=="توجد إجازات متقاطعة" for r in summary))
if __name__=="__main__":
    main()

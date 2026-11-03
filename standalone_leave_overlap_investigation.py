# -*- coding: utf-8 -*-
"""Standalone leave-overlap investigation. Requires: pandas, openpyxl.
Edit INPUT_FILE and optionally SHEET_NAME. No imports from other scripts.
"""
from pathlib import Path
from collections import defaultdict
from datetime import datetime
from itertools import combinations
import pandas as pd
from openpyxl.styles import Font, PatternFill

# ===== CHANGE ONLY THIS PATH =====
INPUT_FILE = r"C:\Users\rammutairi\Downloads\Ports Leaves.xlsx"
SHEET_NAME = 0  # Excel tab index or name
Q4_START = "2026-10-01"
Q4_END = "2026-12-31"
# If you have verified that end date is the RETURN date, change official methodology separately.
# This program compares both possibilities; it does not assume either is correct.

COL = {
    "id": "(Employee) Person ID",
    "first": "(Employee) First Name",
    "last": "(Employee) Last Name",
    "manager": "Manager",
    "location": "(Employee) Location (Name)",
    "start": "Actual Start Date",
    "end": "Actual End Date",
    "approval": "Approval Status",
}
START, END = pd.Timestamp(Q4_START), pd.Timestamp(Q4_END)

def tidy(x):
    return " ".join(str(x).strip().split()) if pd.notna(x) else ""

def employee_id(x):
    if pd.isna(x): return ""
    s = str(x).strip()
    return s[:-2] if s.endswith(".0") and s[:-2].isdigit() else s

def dates(series):
    # Dates already stored as Excel date values are handled directly.
    # Strings in dd/mm/yyyy interpreted day-first.
    return pd.to_datetime(series, errors="coerce", dayfirst=True).dt.normalize()

def read_input():
    p=Path(INPUT_FILE)
    if not p.is_file():
        raise FileNotFoundError(f"Excel file not found: {p}")
    if p.suffix.lower() == ".csv":
        raw=pd.read_csv(p,encoding="utf-8-sig")
    else:
        raw=pd.read_excel(p,sheet_name=SHEET_NAME)
    raw.columns=[tidy(c) for c in raw.columns]
    lookup={c.casefold():c for c in raw.columns}
    missing=[v for v in COL.values() if v.casefold() not in lookup]
    if missing:
        raise ValueError(f"Missing columns: {missing}\nAvailable columns: {list(raw.columns)}")
    get=lambda key:raw[lookup[COL[key].casefold()]]
    df=pd.DataFrame({
        "record":range(1,len(raw)+1),
        "id":get("id").map(employee_id),
        "name":(get("first").map(tidy)+" "+get("last").map(tidy)).str.strip(),
        "manager":get("manager").map(tidy),
        "location":get("location").map(tidy),
        "start":dates(get("start")),
        "end":dates(get("end")),
        "approval":get("approval").map(tidy).str.upper(),
    })
    df["reason"]=""
    df.loc[df["approval"]!="APPROVED","reason"]="Not approved"
    df.loc[df["id"].eq("")|df["manager"].eq("")|df["location"].eq(""),"reason"]="Missing ID / manager / location"
    df.loc[df["start"].isna()|df["end"].isna(),"reason"]="Invalid date"
    df.loc[df["end"]<df["start"],"reason"]="End before start"
    df.loc[(df["end"]<START)|(df["start"]>END),"reason"]="Outside Q4 2026"
    valid=df[df["reason"].eq("")].copy()
    return df,valid

def calendar(valid, exclusive=False):
    # day => unique employees; records retained for record-level counts
    day_people=defaultdict(set)
    day_records=defaultdict(set)
    for r in valid.itertuples(index=False):
        first=max(r.start,START)
        last=min(r.end-pd.Timedelta(days=1) if exclusive else r.end,END)
        if last<first: continue
        for day in pd.date_range(first,last):
            key=(r.manager,r.location,day)
            day_people[key].add(r.id)
            day_records[key].add(r.record)
    return day_people,day_records

def measure(day_people, day_records, min_people=2, min_days=1, sunday_thursday=False):
    employee_days=defaultdict(set)
    overlapping_records=set()
    hot_days=[]
    for (mgr,loc,day),people in day_people.items():
        if sunday_thursday and day.weekday() not in (6,0,1,2,3):
            continue
        if len(people)<min_people: continue
        hot_days.append((mgr,loc,day,len(people)))
        for pid in people:
            employee_days[pid].add(day)
        overlapping_records.update(day_records[(mgr,loc,day)])
    affected={pid for pid,ds in employee_days.items() if len(ds)>=min_days}
    return affected,employee_days,overlapping_records,hot_days

def cases(hot_days):
    # Same definition as original report: consecutive hot calendar days per manager/location.
    bygroup=defaultdict(list)
    for mgr,loc,day,_ in hot_days:
        bygroup[(mgr,loc)].append(day)
    count=0
    for dates in bygroup.values():
        prev=None
        for d in sorted(set(dates)):
            if prev is None or (d-prev).days>1: count+=1
            prev=d
    return count

def investigate():
    all_rows, valid=read_input()
    if valid.empty: raise ValueError("No valid APPROVED Q4 2026 leave records.")
    inc, inc_rec=calendar(valid,False)
    exc, exc_rec=calendar(valid,True)
    a,ad,ar,ah=measure(inc,inc_rec)
    b,bd,br,bh=measure(exc,exc_rec)
    denominator=valid["id"].nunique()
    scenarios=[]
    for label,dp,dr in [("Inclusive end date",inc,inc_rec),("Exclusive end date (hypothesis only)",exc,exc_rec)]:
        for working in [False,True]:
            for n in [2,3,5,10]:
                for d in [1,2,3,5]:
                    affected,_,records,hot=measure(dp,dr,n,d,working)
                    scenarios.append({
                        "End-date rule":label,
                        "Day assumption":"Sun–Thu ONLY (unverified for shift workers)" if working else "All calendar days",
                        "Minimum simultaneous employees":n,
                        "Minimum overlap days per employee (not necessarily consecutive)":d,
                        "Affected employees":len(affected),
                        "Affected / leave-taking employees":len(affected)/denominator,
                        "Overlapping leave records":len(records),
                        "Manager-location-days with overlap":len(hot),
                    })
    scenarios=pd.DataFrame(scenarios)
    # How much of the result is sensitive to the end-date convention?
    sensitive=a-b
    details=[]
    for pid in sorted(valid["id"].unique()):
        rows=valid[valid["id"]==pid]
        details.append({
            "Employee ID":pid,"Employee name":rows.iloc[0]["name"],
            "Inclusive overlap days":len(ad.get(pid,set())),
            "Exclusive overlap days":len(bd.get(pid,set())),
            "Affected inclusive":pid in a,"Affected exclusive":pid in b,
            "Only affected when end date included":pid in sensitive,
            "Distinct approved leave records":len(rows),
        })
    details=pd.DataFrame(details)
    # Pairwise boundary dates are a diagnostic, not the original report's "cases".
    boundary=[]
    for (mgr,loc),g in valid.groupby(["manager","location"]):
        starts=defaultdict(list)
        ends=defaultdict(list)
        for r in g.itertuples(index=False):
            starts[r.start].append(r)
            ends[r.end].append(r)
        for day in set(starts)&set(ends):
            if not START<=day<=END: continue
            for x in ends[day]:
                for y in starts[day]:
                    if x.id==y.id or x.record==y.record: continue
                    boundary.append({
                        "Manager":mgr,"Location":loc,"Shared boundary date":day,
                        "Ending employee":x.id,"Starting employee":y.id,
                        "Ending record":x.record,"Starting record":y.record,
                        "Overlap on shared date if end inclusive":True
                    })
    boundary=pd.DataFrame(boundary)
    # Record-level participation
    records=valid[["record","id","name","manager","location","start","end"]].copy()
    records["Overlaps with inclusive end"] = records["record"].isin(ar)
    records["Overlaps with exclusive end"] = records["record"].isin(br)
    daily=pd.DataFrame([{
        "Manager":mgr,"Location":loc,"Date":day,"Simultaneously absent employees":n,
        "Intensity band":("2–4" if n<=4 else "5–9" if n<=9 else "10–15" if n<=15 else "16+")
    } for mgr,loc,day,n in ah]).sort_values(["Manager","Location","Date"]) if ah else pd.DataFrame(
        columns=["Manager","Location","Date","Simultaneously absent employees","Intensity band"])
    dist=[]
    for band in ["2–4","5–9","10–15","16+"]:
        n=int((daily["Intensity band"]==band).sum())
        dist.append({"Concurrent employees":band,"Manager-location-days":n,
                     "Share of overlap manager-location-days":n/len(daily) if len(daily) else 0})
    dist=pd.DataFrame(dist)
    # Distinguish daily count, leave records, affected employees and case counts.
    duplicates=valid.duplicated(["id","manager","location","start","end"],keep=False).sum()
    single=(valid["start"]==valid["end"]).sum()
    multi_groups=int((valid.groupby("id")[["manager","location"]].apply(
        lambda x:len(set(map(tuple,x.to_numpy())))>1)).sum())
    summary=pd.DataFrame([
        ("Unique leave-taking employees",denominator,"Denominator for this analysis; not all 6193 staff"),
        ("Affected employees - inclusive",len(a),"At least one shared calendar day"),
        ("Affected share - inclusive",len(a)/denominator,"Compare with original 89.8%; if different, check filters and definitions"),
        ("Affected employees - exclusive",len(b),"Hypothetical: end date is NOT a leave day"),
        ("Affected share - exclusive",len(b)/denominator,"Do not adopt without HRIS verification"),
        ("Percentage-point reduction",(len(a)-len(b))/denominator,"Sensitivity to end-date interpretation"),
        ("Unique employees losing affected status",len(sensitive),"These employees no longer qualify under exclusive end"),
        ("Original-style consecutive overlap cases",cases(ah),"Per manager/location, consecutive days with 2+ unique employees absent"),
        ("Overlap cases under exclusive end",cases(bh),"Same case definition, hypothetical end-date rule"),
        ("Approved leave records with overlap - inclusive",len(ar),"Each record counted once"),
        ("Approved leave records with overlap - exclusive",len(br),"Each record counted once"),
        ("All valid approved Q4 leave records",len(valid),"Includes leaves crossing Q4 boundary"),
        ("Single-day leave records",int(single),"These disappear completely under exclusive-end interpretation"),
        ("Records with identical ID/group/start/end",int(duplicates),"Possible duplicates; investigate before deletion"),
        ("Employees appearing in multiple manager/location groups",multi_groups,"May be valid reassignment or inconsistent grouping"),
        ("Manager-location-days with 2+ absent",len(ah),"One day in two groups counts twice"),
        ("Matching end/start boundary pairs",len(boundary),"Pairs, not unique people or original-style cases"),
    ],columns=["Metric","Value","Interpretation"])
    notes=pd.DataFrame([
        ("Cannot confirm Actual End Date meaning","Check HRIS: last leave day or return-to-work date?"),
        ("Cannot establish staffing shortage","Need shift rosters, minimum staffing and actual attendance."),
        ("Weekend test is hypothetical","Ports may operate on Fridays/Saturdays; Sunday–Thursday is not necessarily correct."),
        ("89.8% may differ from this result","Source filters, ID/group cleaning, or earlier script definitions may differ."),
        ("Exclusive end date removes single-day leave","This is a strong warning against changing the rule without HRIS confirmation."),
        ("Minimum overlap days is cumulative","2+ days here need not be consecutive."),
        ("Record counts vs employee counts","One employee can have multiple overlapping leave records."),
        ("Consecutive overlap cases vs employee counts","A case is a continuous run of 2+ absent within a manager/location group."),
        ("Employee master count 6193","Not used here; source is leave records only, so do not call denominator total headcount."),
        ("No definitive causal proof","Scenario changes quantify effects; validate definitions and operational context."),
    ],columns=["Limitation / question","Explanation"])
    src=Path(INPUT_FILE)
    output=src.parent/f"تحليل_اسباب_تزامن_الاجازات_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    with pd.ExcelWriter(output,engine="openpyxl",datetime_format="dd/mm/yyyy") as w:
        for name,df in [
            ("Executive findings",summary),
            ("Scenario comparison",scenarios),
            ("Employee impact",details),
            ("Boundary-date pairs",boundary),
            ("Overlapping leave records",records),
            ("Daily concurrency",daily),
            ("Concurrency distribution",dist),
            ("Data limitations",notes),
            ("Excluded source records",all_rows[all_rows["reason"].ne("")]),
        ]:
            df.to_excel(w,sheet_name=name,index=False)
            ws=w.sheets[name]
            ws.freeze_panes="B2"
            ws.auto_filter.ref=ws.dimensions
            for cell in ws[1]:
                cell.fill=PatternFill("solid",fgColor="17365D")
                cell.font=Font(bold=True,color="FFFFFF")
            for col in ws.columns:
                width=max([len(str(cell.value or "")) for cell in list(col)[:75]]+[12])
                ws.column_dimensions[col[0].column_letter].width=min(65,width+3)
            if name=="Executive findings":
                for row in [4,6,7]:
                    ws.cell(row,2).number_format="0.0%"
            elif name=="Scenario comparison":
                for c in list(ws.columns)[5][1:]:c.number_format="0.0%"
            elif name=="Concurrency distribution":
                for c in list(ws.columns)[2][1:]:c.number_format="0.0%"
    print("\nDONE:",output)
    print("Inclusive:",len(a),f"({len(a)/denominator:.1%})")
    print("Exclusive:",len(b),f"({len(b)/denominator:.1%})")
    print("Boundary-sensitive employees:",len(sensitive))
    print("Single-day leaves lost if end date excluded:",int(single))
    print("IMPORTANT: Do not change official KPI until HRIS confirms end-date meaning.")

if __name__=="__main__":
    try:
        investigate()
    except Exception as exc:
        print("\nERROR:",type(exc).__name__,str(exc))
        print("Check INPUT_FILE, column names, file permissions, and required packages.")
        raise

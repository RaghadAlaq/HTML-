# Creates a private, offline drag-and-drop HTML board from an Excel workbook.
# Requires openpyxl (read only). Does not modify the source workbook.
from pathlib import Path
import json, datetime, webbrowser
from openpyxl import load_workbook

# ======================= SETTINGS =======================
EXCEL_PATH = r"C:\Users\rammutairi\Downloads\axon_rules.xlsx"
SHEET_NAME = "Final Report"
HEADER_ROW = None  # Auto-detect Ref. header; or set Excel row number (e.g. 1)
# ========================================================

ORDER = [[20,58,49,59,23,46,21], [44,51,50,45,28,29,26,1,5,11,13,43,39,36,40,12,17,7,33,37,52,42,38,19,41,60,34,57,18], [6,8,16,2,10,14]]
LABELS = ['أهمية عالية','أهمية متوسطة','أهمية منخفضة']

def clean(v):
    if isinstance(v,(datetime.datetime,datetime.date,datetime.time)): return v.isoformat(sep=' ') if isinstance(v,datetime.datetime) else v.isoformat()
    return v if v is None or isinstance(v,(str,int,float,bool)) else str(v)

def main():
    path=Path(EXCEL_PATH.strip().strip('"').strip("'")).expanduser().resolve()
    if not path.is_file(): raise SystemExit('File not found: '+str(path))
    wb=load_workbook(path,read_only=True,data_only=True)
    s=SHEET_NAME.strip()
    if s not in wb.sheetnames:
        available=', '.join(wb.sheetnames)
        wb.close()
        raise SystemExit(f'Sheet not found: {s!r}. Available sheets: {available}')
    ws=wb[s]
    raw=list(ws.values); wb.close()
    def isref(v): return str(v or '').strip().lower().replace('.','') in ('ref','reference','reference number','رقم المرجع','الرقم المرجعي')
    h=HEADER_ROW-1 if HEADER_ROW is not None else next((i for i,r in enumerate(raw[:50]) if any(isref(v) for v in r)),None)
    if h is None: raise SystemExit('Could not find Ref. header in the first 50 rows.')
    if not 0 <= h < len(raw): raise SystemExit('HEADER_ROW is outside the sheet.')
    # Include every populated column, even columns with blank headers.
    end=max((i+1 for r in raw[h:] for i,v in enumerate(r) if v is not None and v!=''),default=0)
    headers=[str(v) if v is not None and v!='' else f'Column {i+1}' for i,v in enumerate(raw[h][:end])]
    ref=next((i for i,v in enumerate(headers) if isref(v)),None)
    if ref is None: raise SystemExit('Could not find the reference column in the selected header row.')
    pri=next((i for i,v in enumerate(headers) if v.strip().lower() in ('الأهمية','الاهمية','priority','criticality','criticality level','importance')),None)
    lookup={f'DQ_HR_{n:03}':(b,j) for b,ns in enumerate(ORDER) for j,n in enumerate(ns)}
    groups=[[],[],[]]; warnings=[]
    for r in raw[h+1:]:
        vals=[clean(v) for v in r[:end]]
        if not any(v is not None and v!='' for v in vals): continue
        vals += [None]*(end-len(vals))
        if vals[ref] is None or str(vals[ref]).strip()=='':
            warnings.append('A nonempty row has no reference; included in Medium.'); key=''
        else: key=str(vals[ref]).strip().upper()
        b,j=lookup.get(key,(1,99999))
        if pri is not None:
            b=1  # Existing Excel priority is authoritative.
            matched=False
            p=str(vals[pri] or '').strip().lower().replace('أ','ا').replace('ة','ه')
            for k,words in enumerate([('عالي','high','p1'),('متوسط','medium','p2'),('ضعيف','منخفض','low','p3')]):
                if any(w in p for w in words): b=k; matched=True; break
            if not matched: warnings.append(f'{key or "Blank reference"}: missing or unrecognized importance; placed in Medium for review.')
        if pri is None and key not in lookup: warnings.append(f'{key or "Blank reference"}: no prior ranking; check priority.')
        groups[b].append({'id':sum(map(len,groups)), 'v':vals,'rank':j})
    had_priority = pri is not None
    if pri is None:
        headers.insert(ref,'الأهمية'); pri=ref; ref+=1
        for group in groups:
            for row in group: row['v'].insert(pri,None)
    # Keep existing report order within its explicit priority groups; otherwise apply previous ranking.
    if not had_priority:
        for group in groups: group.sort(key=lambda row: row['rank'])
    payload={'headers':headers,'ref':ref,'priority':pri,'groups':groups,'source':path.name+' / '+s,'warnings':warnings,'labels':LABELS,'hiddenColumns':[]}
    if not any(groups): raise SystemExit('No data rows found.')
    safe=json.dumps(payload,ensure_ascii=False,allow_nan=False).replace('<','\\u003c').replace('\u2028','\\u2028').replace('\u2029','\\u2029')
    out=path.with_name(path.stem+'_priority_board.html')
    out.write_text(HTML.replace('__PAYLOAD__',safe),encoding='utf-8')
    print(f'Created {out}\n{sum(map(len,groups))} rows. Original workbook unchanged.')
    webbrowser.open(out.as_uri())

HTML=r'''<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>HR Data Quality | Priority review</title>
<style>
*{box-sizing:border-box}body{margin:0;background:#f6f7f9;color:#223044;font:14px Arial,Tahoma,sans-serif}header{position:sticky;top:0;z-index:5;background:white;border-bottom:1px solid #d8dfe7;padding:16px 22px}h1{font-size:20px;margin:0 0 5px}small{color:#596778}button,select,input{font:inherit}button,select{background:white;border:1px solid #cbd4df;border-radius:5px;padding:7px 10px;color:#223044;cursor:pointer}button:hover{background:#edf2f7}button:disabled{opacity:.4;cursor:default}.primary{background:#243950;color:white;border-color:#243950}.primary:hover{background:#39536f}.toolbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-top:12px}input{padding:8px 10px;border:1px solid #cbd4df;border-radius:5px;min-width:200px}.spacer{flex:1}main{padding:16px 22px 32px}.bucket{background:white;border:1px solid #d6dde6;border-radius:7px;overflow:hidden;margin-bottom:17px}.bucket h2{font-size:15px;margin:0;padding:12px 14px;display:flex;align-items:center;gap:12px}.b0 h2{background:#f4dada;color:#833535}.b1 h2{background:#fff0c7;color:#785717}.b2 h2{background:#dceee0;color:#326442}.scroll{overflow:auto;max-height:430px;min-height:70px}table{border-collapse:separate;border-spacing:0;width:100%;min-width:950px;table-layout:fixed}th,td{border-right:1px solid #e0e5eb;border-bottom:1px solid #e0e5eb;text-align:left;padding:8px;vertical-align:top;overflow-wrap:anywhere;white-space:pre-wrap}th{background:#e9edf2;position:sticky;top:0;z-index:1;font-size:12px}td{font-size:12px;line-height:1.45;background:white}tr:hover td{background:#f5f8fb}.grip{cursor:grab;font-size:18px;display:inline-block;padding:0 6px}.actions{display:flex;gap:3px;align-items:center;flex-wrap:wrap}.actions button{padding:3px 7px}.actions select{font-size:11px;padding:4px;max-width:95px}.dropBefore td{border-top:3px solid #3977b1}.dragging{opacity:.45}.endzone{padding:9px 14px;color:#657587;border-top:1px dashed #d8dfe7;font-size:12px;background:#fafbfd}.endzone.over{background:#dceaf8}.count{font-weight:normal;font-size:12px}.warn{padding:10px 14px;background:#fff4dd;border:1px solid #ead5ad;border-radius:5px;margin:0 0 14px}.hidden{display:none}#status{font-size:12px;color:#536273}a{color:inherit} .priority{font-weight:bold;text-align:center}.b0 .priority{background:#f8e7e7}.b1 .priority{background:#fff5dc}.b2 .priority{background:#e7f3e9} @media print{header .toolbar,.actions,.endzone,.grip{display:none}header{position:static}.scroll{max-height:none;overflow:visible}main{padding:0}table{min-width:0}}
#columnMenu{position:relative}#columnMenu summary{list-style:none;cursor:pointer;background:white;border:1px solid #cbd4df;border-radius:5px;padding:7px 10px}#columnMenu summary::-webkit-details-marker{display:none}#columnPanel{position:absolute;top:calc(100% + 6px);left:0;width:290px;max-width:85vw;background:white;border:1px solid #cbd4df;border-radius:7px;box-shadow:0 8px 24px #22304420;padding:10px;z-index:20}.columnTools{display:flex;gap:8px;margin-bottom:8px}#columnList{max-height:320px;overflow:auto}#columnList label{display:flex;align-items:center;gap:8px;padding:7px 4px;overflow-wrap:anywhere;cursor:pointer}#columnList input{min-width:0;flex-shrink:0}#columnList label:hover{background:#f2f5f8}table{min-width:0} .toolbar{position:relative}
</style><header><h1>HR Data Quality Rules — Priority Review</h1><small id="source"></small><div class="toolbar"><input id="search" placeholder="Search reference, attribute, rule…"><button id="clear">Clear</button><button id="undo">Undo</button><button id="reset">Reset order</button><details id="columnMenu"><summary>Columns ▾</summary><div id="columnPanel"><div class="columnTools"><button id="showAll">Show all</button><button id="hideAll">Hide all</button></div><div id="columnList"></div></div></details><span class="spacer"></span><button id="save">Save board</button><button class="primary" id="export">Export Excel (.xlsx)</button></div><div class="toolbar"><small>Drag the handle ⠿ to any row or section. Use ↑ ↓ or the priority menu as an alternative.</small><span class="spacer"></span><span id="status"></span></div></header><main id="main"></main>
<script id="payload" type="application/json">__PAYLOAD__</script><script>
'use strict';
const initial=JSON.parse(document.getElementById('payload').textContent), labels=initial.labels||['أهمية عالية','أهمية متوسطة','أهمية منخفضة'], english=['High','Medium','Low'];
const clone=x=>JSON.parse(JSON.stringify(x));let groups=clone(initial.groups), history=[],drag=null,hidden=new Set(initial.hiddenColumns||[]);const $=s=>document.querySelector(s);const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const total=()=>groups.reduce((n,g)=>n+g.length,0);const locate=id=>{for(let b=0;b<3;b++){let i=groups[b].findIndex(r=>r.id===id);if(i>=0)return[b,i]}return null};
function snapshot(){history.push(clone(groups));if(history.length>100)history.shift()}
function move(id,b,before=null){let loc=locate(id);if(!loc||before===id)return;let [old,i]=loc;snapshot();let [row]=groups[old].splice(i,1);let j=before===null?groups[b].length:groups[b].findIndex(r=>r.id===before);groups[b].splice(j<0?groups[b].length:j,0,row);render();}
function step(id,d){let [b,i]=locate(id);let j=i+d;if(j<0||j>=groups[b].length)return;snapshot();let a=groups[b];[a[i],a[j]]=[a[j],a[i]];render();document.getElementById('r'+id)?.scrollIntoView({block:'nearest'});}
function visibleCols(){return initial.headers.map((_,i)=>i).filter(i=>!hidden.has(i))}
function updateColumns(){ $('#columnList').innerHTML=initial.headers.map((h,i)=>`<label><input type="checkbox" data-col="${i}" ${hidden.has(i)?'':'checked'}><span dir="auto">${esc(h)}</span></label>`).join('');$('#columnMenu summary').textContent=`Columns (${initial.headers.length-hidden.size}/${initial.headers.length}) ▾`; }
function render(){const cols=visibleCols();const scrolls=[...document.querySelectorAll('.scroll')].map(e=>[e.scrollTop,e.scrollLeft]);let q=$('#search').value.trim().toLowerCase();$('#source').textContent=initial.source+'  •  '+total()+' records';$('#status').textContent=groups.map((g,i)=>english[i]+': '+g.length).join('  |  ');$('#undo').disabled=!history.length;let html=initial.warnings.length?'<details class="warn"><summary>Review notes ('+initial.warnings.length+')</summary>'+initial.warnings.map(esc).join('<br>')+'</details>':'';for(let b=0;b<3;b++){let visible=groups[b].filter(r=>!q||r.v.some((v,i)=>i!==initial.priority&&String(v??'').toLowerCase().includes(q))||labels[b].includes(q));html+=`<section class="bucket b${b}"><h2>${english[b]} — ${labels[b]} <span class="count">${visible.length}${q?' / '+groups[b].length:''} rules</span></h2><div class="scroll"><table><colgroup><col style="width:145px">${cols.map(i=>{let h=initial.headers[i];return `<col style="width:${i===initial.priority?105:i===initial.ref?110:/description|comment|justification|مبررات|وصف|ملاحظ|تعليقات/i.test(h)?310:/name|attribute/i.test(h)?190:135}px">`}).join('')}</colgroup><thead><tr><th>Move</th>${cols.map(i=>'<th dir="auto">'+esc(initial.headers[i])+'</th>').join('')}</tr></thead><tbody>${visible.map(r=>`<tr id="r${r.id}" data-id="${r.id}" data-b="${b}"><td><div class="actions"><span class="grip" draggable="true" data-id="${r.id}" title="Drag row">⠿</span><button data-step="-1" data-id="${r.id}" title="Move up">↑</button><button data-step="1" data-id="${r.id}" title="Move down">↓</button><select data-move="${r.id}" title="Move to priority">${english.map((x,i)=>`<option value="${i}" ${i===b?'selected':''}>${x}</option>`).join('')}</select></div></td>${cols.map(i=>`<td dir="auto" ${i===initial.priority?'class="priority"':''}>${esc(i===initial.priority?labels[b]:r.v[i])}</td>`).join('')}</tr>`).join('')}</tbody></table></div><div class="endzone" data-end="${b}">Drop here to move to the end of ${english[b]} priority</div></section>`}$('#main').innerHTML=html;[...document.querySelectorAll('.scroll')].forEach((e,i)=>{if(scrolls[i]){e.scrollTop=scrolls[i][0];e.scrollLeft=scrolls[i][1]}})}
$('#columnList').addEventListener('change',e=>{if(e.target.dataset.col===undefined)return;let i=+e.target.dataset.col;if(e.target.checked)hidden.delete(i);else hidden.add(i);updateColumns();render()});$('#showAll').onclick=()=>{hidden.clear();updateColumns();render()};$('#hideAll').onclick=()=>{hidden=new Set(initial.headers.map((_,i)=>i));updateColumns();render()};document.addEventListener('click',e=>{if(!e.target.closest('#columnMenu'))$('#columnMenu').open=false});updateColumns();
$('#search').oninput=render;$('#clear').onclick=()=>{$('#search').value='';render()};$('#undo').onclick=()=>{if(history.length){groups=history.pop();render()}};$('#reset').onclick=()=>{if(confirm('Restore the order from the original import?')){snapshot();groups=clone(initial.groups);render()}};
$('#main').addEventListener('click',e=>{let el=e.target.closest('[data-step]');if(el)step(+el.dataset.id,+el.dataset.step)});$('#main').addEventListener('change',e=>{if(e.target.dataset.move!==undefined)move(+e.target.dataset.move,+e.target.value)});
$('#main').addEventListener('dragstart',e=>{let el=e.target.closest('.grip');if(!el)return;drag=+el.dataset.id;e.dataTransfer.setData('text/plain',String(drag));e.dataTransfer.effectAllowed='move';document.getElementById('r'+drag)?.classList.add('dragging')});function clearDrop(){document.querySelectorAll('.dropBefore,.over,.dragging').forEach(e=>e.classList.remove('dropBefore','over','dragging'))}
$('#main').addEventListener('dragover',e=>{if(drag===null)return;let el=e.target.closest('tr[data-id],.endzone');if(!el)return;e.preventDefault();e.dataTransfer.dropEffect='move';document.querySelectorAll('.dropBefore,.over').forEach(x=>x.classList.remove('dropBefore','over'));el.classList.add(el.matches('tr')?'dropBefore':'over')});$('#main').addEventListener('drop',e=>{if(drag===null)return;let el=e.target.closest('tr[data-id],.endzone');if(!el)return;e.preventDefault();let id=drag;drag=null;clearDrop();if(el.matches('tr'))move(id,+el.dataset.b,+el.dataset.id);else move(id,+el.dataset.end)});document.addEventListener('dragend',()=>{drag=null;clearDrop()});
function download(data,name,type){let u=URL.createObjectURL(new Blob([data],{type}));let a=document.createElement('a');a.href=u;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(u),1000)}
$('#save').onclick=()=>{let d=document.documentElement.cloneNode(true);let p=d.querySelector('#payload');p.textContent=JSON.stringify({...initial,groups,hiddenColumns:[...hidden]}).replace(/</g,'\\u003c');d.querySelector('#main').innerHTML='';d.querySelector('#columnMenu').removeAttribute('open');d.querySelector('#search').removeAttribute('value');download('<!doctype html>'+d.outerHTML,'axon_priority_review_saved.html','text/html;charset=utf-8')};
// Minimal offline XLSX writer, using uncompressed ZIP and inline strings. No external libraries.
const enc=new TextEncoder();const xml=s=>String(s??'').replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g,'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
function crc(bytes){let c=0xffffffff;for(let b of bytes){c^=b;for(let k=0;k<8;k++)c=(c>>>1)^((c&1)?0xedb88320:0)}return(c^0xffffffff)>>>0}
function zip(files){let chunks=[],central=[],offset=0,cs=0;for(let [name,content] of Object.entries(files)){let n=enc.encode(name),data=enc.encode(content),c=crc(data);let l=new Uint8Array(30+n.length),v=new DataView(l.buffer);v.setUint32(0,0x04034b50,true);v.setUint16(4,20,true);v.setUint32(14,c,true);v.setUint32(18,data.length,true);v.setUint32(22,data.length,true);v.setUint16(26,n.length,true);l.set(n,30);chunks.push(l,data);let h=new Uint8Array(46+n.length),w=new DataView(h.buffer);w.setUint32(0,0x02014b50,true);w.setUint16(4,20,true);w.setUint16(6,20,true);w.setUint32(16,c,true);w.setUint32(20,data.length,true);w.setUint32(24,data.length,true);w.setUint16(28,n.length,true);w.setUint32(42,offset,true);h.set(n,46);central.push(h);cs+=h.length;offset+=l.length+data.length}let e=new Uint8Array(22),v=new DataView(e.buffer);v.setUint32(0,0x06054b50,true);v.setUint16(8,central.length,true);v.setUint16(10,central.length,true);v.setUint32(12,cs,true);v.setUint32(16,offset,true);return new Blob([...chunks,...central,e])}
const prefix='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';const rel='http://schemas.openxmlformats.org/officeDocument/2006/relationships';const pkg='http://schemas.openxmlformats.org/package/2006/relationships';
function col(i){let s='';for(i++;i;i=Math.floor((i-1)/26))s=String.fromCharCode(65+(i-1)%26)+s;return s}
function cell(v,i,r,style){let a=`r="${col(i)}${r}" s="${style}"`;if(v===null||v===undefined)return `<c ${a}/>`;if(typeof v==='number'&&Number.isFinite(v))return `<c ${a}><v>${v}</v></c>`;if(typeof v==='boolean')return `<c ${a} t="b"><v>${v?1:0}</v></c>`;return `<c ${a} t="inlineStr"><is><t xml:space="preserve">${xml(v)}</t></is></c>`}
function makeExcel(){let rows=[`<row r="1">${initial.headers.map((v,i)=>cell(v,i,1,1)).join('')}</row>`],r=2;groups.forEach((g,b)=>g.forEach(item=>{rows.push(`<row r="${r}">${item.v.map((v,i)=>cell(i===initial.priority?labels[b]:v,i,r,i===initial.priority?b+2:0)).join('')}</row>`);r++}));let files={};files['[Content_Types].xml']=prefix+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';files['_rels/.rels']=prefix+`<Relationships xmlns="${pkg}"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`;files['xl/workbook.xml']=prefix+`<workbook xmlns="${ns}" xmlns:r="${rel}"><sheets><sheet name="Priority Review" sheetId="1" r:id="rId1"/></sheets></workbook>`;files['xl/_rels/workbook.xml.rels']=prefix+`<Relationships xmlns="${pkg}"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rel}/styles" Target="styles.xml"/></Relationships>`;files['xl/styles.xml']=prefix+`<styleSheet xmlns="${ns}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>${['243950','F4DADA','FFF0C7','DCEEE0'].map(c=>`<fill><patternFill patternType="solid"><fgColor rgb="FF${c}"/><bgColor indexed="64"/></patternFill></fill>`).join('')}</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>${[2,3,4,5].map((f,i)=>`<xf numFmtId="0" fontId="${i===0?1:0}" fillId="${f}" borderId="0" xfId="0" applyFill="1"/>`).join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;files['xl/worksheets/sheet1.xml']=prefix+`<worksheet xmlns="${ns}"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>${initial.headers.map((h,i)=>`<col min="${i+1}" max="${i+1}" width="${/description|comment|justification|مبررات/i.test(h)?65:22}" customWidth="1"/>`).join('')}</cols><sheetData>${rows.join('')}</sheetData><autoFilter ref="A1:${col(initial.headers.length-1)}${r-1}"/></worksheet>`;return zip(files)}
$('#export').onclick=()=>download(makeExcel(),'axon_priority_review.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');render();
</script></html>'''
if __name__=='__main__': main()

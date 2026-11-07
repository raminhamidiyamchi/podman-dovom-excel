<!DOCTYPE html>
<html lang="fa">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>پودمان دوم Excel</title>
<style>
:root{--excel:#217346;--excel-dark:#185c37;--border:#d7ded9;--panel:#f7f9f7;--text:#202a24;--muted:#66736b;--accent:#e8f3ec}
*{box-sizing:border-box}body{margin:0;background:#f3f6f4;color:var(--text);font-family:Arial,Helvetica,sans-serif;direction:ltr}
button,input,select{font:inherit}button{cursor:pointer}
.topbar{height:54px;background:var(--excel);color:white;display:flex;align-items:center;padding:0 14px;gap:13px}
.brand{font-size:19px;font-weight:700;white-space:nowrap}.bookname{width:175px;padding:7px 9px;border:1px solid #ffffff66;border-radius:4px;background:#ffffff18;color:white}.bookname::placeholder{color:#eee}
.topright{margin-left:auto;display:flex;gap:7px}.topbtn{background:#ffffff18;color:white;border:1px solid #ffffff55;border-radius:4px;padding:7px 10px;font-size:12px}.topbtn:hover{background:#ffffff30}
.tabs{display:flex;align-items:stretch;gap:0;background:#fff;border-bottom:1px solid var(--border);padding:0 10px;overflow-x:auto;direction:ltr}
.tab{border:0;border-radius:0;background:white;padding:12px 17px;font-size:13px;color:#303b34;white-space:nowrap;border-bottom:3px solid transparent}
.tab:hover{background:#f1f6f2}.tab.active{color:var(--excel);font-weight:700;border-bottom-color:var(--excel);background:#f7fbf8}
.ribbon{display:flex;gap:0;align-items:stretch;min-height:112px;background:#fff;border-bottom:1px solid #c9cdd1;padding:5px 8px 3px;overflow-x:auto;direction:ltr;scrollbar-width:thin}
.rgroup{display:flex;flex-direction:column;justify-content:space-between;padding:1px 8px 0;border-right:1px solid #d7dadd;min-width:max-content;position:relative}.rgroup:last-child{border:0}
.tools{display:flex;align-items:flex-start;gap:2px;flex-wrap:nowrap}.tool{min-width:37px;min-height:68px;border:1px solid transparent;background:transparent;border-radius:3px;padding:4px 3px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;color:#40464b;font-size:10px;white-space:nowrap}
.tool:hover{background:#f1f6fb;border-color:#d8e5f1}.tool:active{background:#e2edf8}.tool .ico{font-size:19px;line-height:22px;color:#4b535a}.tool.wide{min-width:58px}.tool.compact{min-width:29px;min-height:27px;padding:1px 3px}.groupname{text-align:center;font-size:10px;color:#555d64;padding-top:4px;margin-top:2px;white-space:nowrap}
.ctrl{height:27px;border:1px solid #bfc5ca;border-radius:3px;background:white;padding:3px 6px;font-size:12px;color:#27333b}
.ribbonrow{display:flex;gap:3px;align-items:center;margin-bottom:4px}.rgroup .ribbonrow:last-of-type{margin-bottom:0}
.home-group{height:100%;display:flex;flex-direction:column;justify-content:space-between}.home-tools{display:flex;gap:3px;align-items:flex-start}.home-stack{display:flex;flex-direction:column;gap:3px}.home-select{height:27px;font-size:12px}.home-iconbtn{height:28px;min-width:27px;border:1px solid transparent;background:transparent;border-radius:3px;color:#3e454b;font-size:17px;padding:2px 5px}.home-iconbtn:hover{background:#edf4fa;border-color:#d7e4ef}.home-label{font-size:10px;color:#555d64;text-align:center;border-top:1px solid #e2e4e6;padding-top:4px;margin-top:3px}.paste-main{height:69px;min-width:54px;border:0;background:transparent;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;color:#444;font-size:11px}.paste-main .ico{font-size:31px;line-height:34px}.paste-main:hover{background:#f0f5fa;border-radius:3px}.font-mini{display:flex;align-items:center;gap:1px}.font-color-mark{border-bottom:4px solid #e53935;padding-bottom:2px}.fill-color-mark{border-bottom:4px solid #a6a6a6;padding-bottom:2px}.ribbon select:focus,.ribbon button:focus-visible{outline:1px solid #5b9bd5}
.formulabar{height:42px;background:#fff;display:flex;align-items:center;gap:8px;padding:6px 10px;border-bottom:1px solid var(--border)}
.namebox{width:78px;text-align:center;font-weight:bold}.fx{font-weight:bold;color:var(--excel);font-style:italic}.formula-input{flex:1;min-width:80px}
.workspace{padding:10px;overflow:auto}.inserted-objects{display:flex;flex-wrap:wrap;gap:12px;margin-bottom:10px}.inserted-object{position:relative;background:#fff;border:1px solid #cbd5ce;border-radius:5px;padding:10px;min-width:220px;max-width:100%;box-shadow:0 1px 3px #0000000d}.inserted-object img{max-width:100%;max-height:260px;object-fit:contain;display:block}.inserted-object .remove-object{position:absolute;top:4px;right:5px;border:0;background:#f4f4f4;border-radius:4px;color:#555}.inserted-object svg{max-width:100%;height:auto;display:block}.inserted-object h4{margin:0 24px 8px 0;font-size:13px}.inserted-object p{white-space:pre-wrap;margin:4px 0;font-size:13px}.gridwrap{background:white;border:1px solid #cbd5ce;border-radius:5px;overflow:auto;max-height:60vh;direction:ltr}
table{border-collapse:collapse;table-layout:fixed;direction:ltr;min-width:100%}th,td{border:1px solid #dfe5e0;width:105px;min-width:105px;height:29px;padding:3px 6px;font-size:13px;white-space:nowrap;overflow:hidden}td{vertical-align:middle}
th{background:#edf2ee;color:#4c5b51;text-align:center;font-weight:600;position:sticky;top:0;z-index:2}.col-head{position:relative;user-select:none}.col-resizer{position:absolute;right:-2px;top:0;width:7px;height:100%;cursor:col-resize;z-index:8}.col-resizer:hover,.col-resizer.active{background:#16834a;opacity:.65}body.resizing,body.resizing *{cursor:col-resize!important;user-select:none!important}th.corner{left:0;z-index:4;width:42px;min-width:42px}th.rowhead{left:0;z-index:1;width:42px;min-width:42px}
td{background:white;outline:none}td.selected{outline:2px solid #16834a;outline-offset:-2px;background:#f5fbf6}td:focus{outline:2px solid #16834a;outline-offset:-2px}
.sheetbar{display:flex;align-items:center;gap:7px;background:#fff;border-top:1px solid var(--border);padding:8px 12px}.sheet{border:0;border-bottom:2px solid transparent;border-radius:0;padding:7px 13px;font-weight:bold;color:var(--excel);background:#f5faf6}.sheet.active-sheet{border-bottom:2px solid var(--excel);background:#f5faf6;color:var(--excel)}.addsheet{border:1px solid var(--border);background:white;border-radius:4px;padding:5px 10px}
.status{display:flex;gap:18px;background:#fff;border-top:1px solid var(--border);padding:7px 12px;font-size:11px;color:var(--muted)}
.status{position:relative;justify-content:space-between;min-height:42px}.designer-credit{position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);color:#185c37;font-weight:800;font-size:16px;white-space:nowrap;text-align:center}.help{font-size:11px;color:#64736a;margin-left:auto;align-self:center}.modal{position:fixed;inset:0;background:#0006;display:none;align-items:center;justify-content:center;z-index:10}.modal.show{display:flex}.dialog{background:#fff;border-radius:9px;padding:20px;width:min(420px,92vw);box-shadow:0 12px 40px #0003}.dialog h3{margin:0 0 12px}.dialog p{font-size:13px;color:#56645a;line-height:1.6}.dialog-actions{display:flex;justify-content:flex-end;gap:8px;margin-top:15px}
.primary{background:var(--excel);color:white;border:0;border-radius:4px;padding:8px 13px}.secondary{background:white;border:1px solid var(--border);border-radius:4px;padding:8px 13px}
@media(max-width:650px){.topbar{height:48px;padding:0 8px;gap:7px}.brand{font-size:15px}.bookname{width:100px}.topbtn{padding:6px;font-size:11px}.tab{padding:11px 13px}.ribbon{min-height:105px;padding:6px}.tool{min-width:44px}.workspace{padding:5px}th,td{min-width:90px;width:90px}.status{gap:8px;flex-wrap:wrap;padding-bottom:30px}.designer-credit{top:auto;bottom:5px;left:50%;transform:translateX(-50%);font-size:15px}.help{display:none}}

.insert-tool{min-width:57px;min-height:78px;font-size:10px;white-space:normal;max-width:76px;line-height:1.15}.insert-tool .ico{font-size:24px;line-height:28px;color:#2876b9}.insert-tool.wide{min-width:58px}.insert-chart-grid{display:grid;grid-template-columns:repeat(3,48px);grid-template-rows:repeat(2,39px);gap:1px 2px;align-content:start}.insert-mini{border:1px solid transparent;background:transparent;border-radius:3px;display:flex;align-items:center;justify-content:center;gap:2px;flex-direction:column;color:#3d4b55;font-size:17px;min-width:45px;min-height:38px;padding:1px}.insert-mini small{font-size:8px;white-space:nowrap}.insert-mini:hover{background:#eef5fb;border-color:#d8e5f1}.insert-tool:hover{background:#f1f6fb;border-color:#d8e5f1}
.chart-object{min-width:560px;max-width:100%;}.chart-object .ctrl{height:30px;font-size:12px}.chart-object h4{padding-right:4px}.chart-object svg{display:block;max-width:100%;height:auto}@media(max-width:650px){.chart-object{min-width:0}.chart-object>div:first-of-type{overflow-x:auto}}\n</style>
</head>
<body>
<header class="topbar">
 <div class="brand">▦ پودمان دوم Excel</div>
 <input class="bookname" id="bookName" value="Book1" aria-label="Workbook name">
 <div class="topright">
  <button class="topbtn" id="newBtn">＋ New</button>
  <button class="topbtn" id="openBtn">📂 Open CSV</button>
  <button class="topbtn" id="saveBtn">💾 Save</button>
  <button class="topbtn" id="downloadBtn">⬇ Export CSV</button>
  <input type="file" id="fileInput" accept=".csv,text/csv" hidden>
 </div>
</header>
<nav class="tabs" id="tabs">
 <button class="tab" data-tab="File">File</button>
 <button class="tab active" data-tab="Home">Home</button>
 <button class="tab" data-tab="Insert">Insert</button>
 <button class="tab" data-tab="Page Layout">Page Layout</button>
 <button class="tab" data-tab="Formulas">Formulas</button>
 <button class="tab" data-tab="Data">Data</button>
 <button class="tab" data-tab="Review">Review</button>
 <button class="tab" data-tab="View">View</button>
</nav>
<section class="ribbon" id="ribbon"></section>
<section class="formulabar">
 <input class="ctrl namebox" id="cellName" value="A1" aria-label="Cell address">
 <span class="fx">fx</span>
 <input class="ctrl formula-input" id="formulaInput" placeholder="Enter a value or formula, e.g. =SUM(A1:A5)">
 <button class="secondary" id="applyBtn">✓ Enter</button>
</section>
<main class="workspace"><div id="insertedObjects" class="inserted-objects"></div><div class="gridwrap" id="gridwrap"><table id="grid"></table></div></main>
<div class="sheetbar"><button class="sheet" id="sheet1">Sheet1</button><button class="addsheet" id="addSheet" title="Add worksheet">＋</button><span class="help">Tip: Drag across cells to select a range; use Merge & Center to combine them.</span></div>
<footer class="status"><span id="statusCell">Ready</span><span id="statusValue">Value: —</span><span id="statusSelection">Selection: A1</span><span class="designer-credit">طراح: رامین حمیدی</span></footer>
<div class="modal" id="modal"><div class="dialog"><h3 id="modalTitle">Mini Excel</h3><p id="modalText"></p><div class="dialog-actions"><button class="secondary" id="modalClose">Close</button></div></div></div>
<script>
const ROWS=60,COLS=20;
let cells={},colWidths={},sheets=[{name:'Sheet1',cells:{},colWidths:{}}],activeSheetIndex=0,selected='A1',selectedRefs=['A1'],activeTab='Home',zoom=100,clipboard='',clipboardCut=false,formatPainter=null,dragAnchor=null,isDragging=false;
const $=id=>document.getElementById(id),grid=$('grid'),formula=$('formulaInput'),cellName=$('cellName');
const colName=n=>{let s='';while(n){let r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26)}return s};
function refParts(ref){let m=String(ref).toUpperCase().match(/^([A-Z]+)(\d+)$/);if(!m)return null;let c=0;for(let x of m[1])c=c*26+x.charCodeAt(0)-64;return{r:+m[2],c}}
function raw(ref){return cells[ref]?.value??''}
function rangeRefs(range){
 let parts=String(range).toUpperCase().split(':');if(parts.length>2)return[];
 let p=refParts(parts[0]),q=refParts(parts[1]||parts[0]);if(!p||!q||p.r<1||q.r<1||p.c<1||q.c<1||p.r>ROWS||q.r>ROWS||p.c>COLS||q.c>COLS)return[];
 let arr=[];for(let r=Math.min(p.r,q.r);r<=Math.max(p.r,q.r);r++)for(let c=Math.min(p.c,q.c);c<=Math.max(p.c,q.c);c++)arr.push(colName(c)+r);return arr;
}
function splitFormulaArgs(text){
 let out=[],buf='',depth=0,inQ=false;
 for(let ch of text){if(ch==='"')inQ=!inQ;if(!inQ){if(ch==='(')depth++;if(ch===')')depth--;if(ch===','&&depth===0){out.push(buf.trim());buf='';continue;}}buf+=ch;}
 if(buf.trim())out.push(buf.trim());return out;
}
function formulaArg(token,stack){
 token=token.trim();
 if(/^".*"$/.test(token))return token.slice(1,-1).replace(/""/g,'"');
 if(/^[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?$/i.test(token)){
  const refs=rangeRefs(token.toUpperCase());return refs.map(r=>calc(r,stack));
 }
 const n=Number(token);if(token!==''&&Number.isFinite(n))return n;
 const m=token.match(/^(TRUE|FALSE)$/i);if(m)return m[1].toUpperCase()==='TRUE';
 return token;
}
function calc(ref,stack=[]){
 if(stack.includes(ref))return '#CYCLE!';
 let x=String(raw(ref)??'');if(!x.startsWith('=')){if(x.trim()!==''&&!isNaN(Number(x)))return Number(x);return x}
 let f=x.slice(1).trim(),m=f.match(/^([A-Z]+)\((.*)\)$/i);
 if(m){
  const fn=m[1].toUpperCase(),args=splitFormulaArgs(m[2]).map(t=>formulaArg(t,[...stack,ref]));
  const flat=args.flat(Infinity),nums=flat.filter(v=>typeof v==='number'&&Number.isFinite(v));
  switch(fn){
   case 'SUM':return nums.reduce((a,b)=>a+b,0);
   case 'AVERAGE':return nums.length?nums.reduce((a,b)=>a+b,0)/nums.length:0;
   case 'MIN':return nums.length?Math.min(...nums):0;
   case 'MAX':return nums.length?Math.max(...nums):0;
   case 'COUNT':return nums.length;
   case 'COUNTA':return flat.filter(v=>v!==''&&v!==null&&v!==undefined).length;
   case 'PRODUCT':return nums.reduce((a,b)=>a*b,1);
   case 'ABS':return Math.abs(Number(args[0]));
   case 'SQRT':return Number(args[0])>=0?Math.sqrt(Number(args[0])):'#NUM!';
   case 'ROUND':{let d=Math.max(0,Math.min(10,Number(args[1]||0)));return Number(Number(args[0]).toFixed(d));}
   case 'ROUNDUP':{let d=Math.max(0,Math.min(10,Number(args[1]||0))),p=10**d;return Math.sign(Number(args[0]))*Math.ceil(Math.abs(Number(args[0]))*p)/p;}
   case 'ROUNDDOWN':{let d=Math.max(0,Math.min(10,Number(args[1]||0))),p=10**d;return Math.sign(Number(args[0]))*Math.floor(Math.abs(Number(args[0]))*p)/p;}
   case 'IF':return args[0]?args[1]:args[2];
   case 'AND':return args.every(Boolean);
   case 'OR':return args.some(Boolean);
   case 'NOT':return !args[0];
   case 'CONCAT':case 'CONCATENATE':return flat.map(v=>v??'').join('');
   case 'LEFT':return String(args[0]??'').slice(0,Math.max(0,Number(args[1]??1)));
   case 'RIGHT':return String(args[0]??'').slice(-Math.max(0,Number(args[1]??1)));
   case 'LEN':return String(args[0]??'').length;
   case 'UPPER':return String(args[0]??'').toUpperCase();
   case 'LOWER':return String(args[0]??'').toLowerCase();
   case 'PROPER':return String(args[0]??'').toLowerCase().replace(/(^|\\s)\\S/g,c=>c.toUpperCase());
   default:return '#NAME?';
  }
 }
 let expr=f.replace(/\b([A-Z]{1,3}\d+)\b/gi,z=>{let v=calc(z.toUpperCase(),[...stack,ref]);return typeof v==='number'?v:0});
 if(!/^[\d\s.+\-*/()%<>!=]+$/.test(expr))return '#NAME?';
 try{let n=Function('"use strict";return ('+expr+')')();return Number.isFinite(n)?n:'#VALUE!'}catch{return '#VALUE!'}
}
function display(ref){
 if(showFormulaMode&&String(raw(ref)||'').startsWith('='))return raw(ref);
 let v=calc(ref),o=cells[ref]||{};
 if(typeof v!=='number')return v;
 if(o.format==='percent')return (v*100).toFixed(o.decimals??0)+'%';
 if(o.format==='currency')return new Intl.NumberFormat('en-US',{minimumFractionDigits:o.decimals??2,maximumFractionDigits:o.decimals??2}).format(v);
 if(o.format==='number')return new Intl.NumberFormat('en-US',{minimumFractionDigits:o.decimals??0,maximumFractionDigits:o.decimals??0}).format(v);
 return v;
}
function esc(s){return String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function styleFor(o){
 return `${o.bold?'font-weight:bold;':''}${o.italic?'font-style:italic;':''}${o.underline?'text-decoration:underline;':''}${o.color?'color:'+o.color+';':''}${o.bg?'background:'+o.bg+';':''}${o.align?'text-align:'+o.align+';':''}${o.valign?'vertical-align:'+o.valign+';':''}${o.wrap?'white-space:normal;overflow-wrap:anywhere;':''}${o.fontFamily?'font-family:'+JSON.stringify(o.fontFamily)+';':''}${o.fontSize?'font-size:'+o.fontSize+'px;':''}${o.indent?'padding-left:'+(6+o.indent*12)+'px;':''}${o.orientation==='vertical'?'writing-mode:vertical-rl;transform:rotate(180deg);':''}${o.border?'border:'+o.border+';':''}`;
}
function render(){
 let h='<colgroup><col style="width:42px">';for(let c=1;c<=COLS;c++)h+=`<col data-col="${c}" style="width:${colWidths[c]||105}px">`;h+='</colgroup><tr><th class="corner"></th>';for(let c=1;c<=COLS;c++)h+=`<th class="col-head" data-col="${c}">${colName(c)}<span class="col-resizer" data-resize="${c}" title="برای تغییر پهنا بکشید؛ برای تنظیم خودکار دوبار کلیک کنید"></span></th>`;h+='</tr>';
 for(let r=1;r<=ROWS;r++){h+=`<tr><th class="rowhead">${r}</th>`;for(let c=1;c<=COLS;c++){let ref=colName(c)+r,o=cells[ref]||{},isSel=selectedRefs.includes(ref);if(o.mergedInto)continue;let span=o.merge||{};h+=`<td contenteditable="${sheetProtected ? "false" : "true"}" spellcheck="${spellCheckEnabled}" data-ref="${ref}" ${span.cols>1?`colspan="${span.cols}"`:''} ${span.rows>1?`rowspan="${span.rows}"`:''} style="${styleFor(o)}" class="${isSel?'selected':''}">${esc(display(ref))}</td>`}h+='</tr>'}
 grid.innerHTML=h;
 grid.querySelectorAll('.col-resizer').forEach(handle=>{
  handle.addEventListener('mousedown',e=>{
   e.preventDefault();e.stopPropagation();const c=Number(handle.dataset.resize),startX=e.clientX,startW=colWidths[c]||105;
   handle.classList.add('active');document.body.classList.add('resizing');
   const moveEv=ev=>{const w=Math.max(45,Math.min(500,startW+ev.clientX-startX));colWidths[c]=w;const col=grid.querySelector(`col[data-col="${c}"]`);if(col)col.style.width=w+'px';grid.querySelectorAll(`tr > :nth-child(${c+1})`).forEach(x=>{x.style.width=w+'px';x.style.minWidth=w+'px'});};
   const upEv=()=>{handle.classList.remove('active');document.body.classList.remove('resizing');document.removeEventListener('mousemove',moveEv);document.removeEventListener('mouseup',upEv);};
   document.addEventListener('mousemove',moveEv);document.addEventListener('mouseup',upEv);
  });
  handle.addEventListener('dblclick',e=>{e.preventDefault();e.stopPropagation();const c=Number(handle.dataset.resize);let max=colName(c).length*9+22;for(let r=1;r<=ROWS;r++){const v=String(display(colName(c)+r)??'');max=Math.max(max,Math.min(300,v.length*8+18));}colWidths[c]=Math.max(55,Math.min(300,max));const col=grid.querySelector(`col[data-col="${c}"]`);if(col)col.style.width=colWidths[c]+'px';grid.querySelectorAll(`tr > :nth-child(${c+1})`).forEach(x=>{x.style.width=colWidths[c]+'px';x.style.minWidth=colWidths[c]+'px'});});
 });
 grid.querySelectorAll('td').forEach(td=>{
  td.addEventListener('focus',()=>{if(!isDragging&&!selectedRefs.includes(td.dataset.ref))select(td.dataset.ref)});
  td.addEventListener('mousedown',e=>{
   if(e.button!==0)return;
   dragAnchor=td.dataset.ref;isDragging=true;select(dragAnchor);
   // Prevent accidental text selection while dragging across cells.
   e.preventDefault();
  });
  td.addEventListener('mouseenter',()=>{
   if(!isDragging||!dragAnchor)return;
   select(dragAnchor+':'+td.dataset.ref);
  });
  td.addEventListener('input',()=>{cells[td.dataset.ref]={...(cells[td.dataset.ref]||{}),value:td.innerText};formula.value=td.innerText;status()});
  td.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key==='Tab'){e.preventDefault();commit(td);move(e.key==='Enter'?1:0,e.key==='Tab'?1:0)}});
  td.addEventListener('blur',()=>{commit(td);refresh()});
 });
}
document.addEventListener('mouseup',()=>{
 if(!isDragging)return;
 isDragging=false;
 const active=grid.querySelector(`[data-ref="${selected}"]`);
 if(active)active.focus({preventScroll:true});
});
function select(ref){
 let refs=rangeRefs(ref);if(!refs.length){cellName.value=selected;return}
 selected=refs[0];selectedRefs=refs;cellName.value=ref.toUpperCase();formula.value=raw(selected);
 if(formatPainter){const keys=['bold','italic','underline','color','bg','align','valign','wrap','fontFamily','fontSize','border','format','decimals','orientation','indent'];refs.forEach(r=>{cells[r]={...(cells[r]||{})};keys.forEach(k=>{if(formatPainter[k]!==undefined)cells[r][k]=formatPainter[k];else delete cells[r][k]})});formatPainter=null;render();}
 grid.querySelectorAll('td.selected').forEach(x=>x.classList.remove('selected'));
 refs.forEach(r=>grid.querySelector(`[data-ref="${r}"]`)?.classList.add('selected'));
 status();
}
function commit(td){if(td?.dataset?.ref)cells[td.dataset.ref]={...(cells[td.dataset.ref]||{}),value:td.innerText}}
function refresh(){let focus=document.activeElement?.dataset?.ref;grid.querySelectorAll('td').forEach(td=>{if(td.dataset.ref!==focus)td.textContent=display(td.dataset.ref)});status()}
function status(){
 $('statusCell').textContent='Cell: '+selected;
 $('statusValue').textContent='Value: '+display(selected);
 $('statusSelection').textContent='Selection: '+cellName.value;
}
function move(dr,dc){let p=refParts(selected),r=Math.max(1,Math.min(ROWS,p.r+dr)),c=Math.max(1,Math.min(COLS,p.c+dc)),ref=colName(c)+r;select(ref);grid.querySelector(`[data-ref="${ref}"]`)?.focus()}
function apply(){if(sheetProtected){tell('Protected Sheet','ابتدا محافظت کاربرگ را غیرفعال کن.');return;}cells[selected]={...(cells[selected]||{}),value:formula.value};render();select(cellName.value||selected);grid.querySelector(`[data-ref="${selected}"]`)?.focus()}
$('applyBtn').onclick=apply;formula.addEventListener('keydown',e=>{if(e.key==='Enter')apply()});
cellName.addEventListener('change',()=>{let r=cellName.value.toUpperCase().replace(/\s/g,'');if(rangeRefs(r).length)select(r);else{cellName.value=selected;tell('Cell reference','Enter a valid cell or range, for example A1 or A1:B5.')}});
function btn(icon,label,act,cls=''){return `<button class="tool ${cls}" data-act="${act}" title="${label}"><span class="ico">${icon}</span><span>${label}</span></button>`}
function group(name,tools){return `<div class="rgroup"><div class="tools">${tools}</div><div class="groupname">${name}</div></div>`}
function renderRibbon(){let h='';
 if(activeTab==='Home'){
  h+=`<div class="rgroup"><div class="home-group"><div class="home-tools"><div class="home-stack"><button class="paste-main" data-act="paste" title="Paste"><span class="ico">📋</span><span>Paste</span><span>▾</span></button></div><div class="home-stack"><button class="home-iconbtn" data-act="cut" title="Cut">✂</button><button class="home-iconbtn" data-act="copy" title="Copy">▢</button><button class="home-iconbtn" data-act="formatpaint" title="Format Painter">🖌</button></div></div><div class="home-label">Clipboard</div></div></div>`;
  h+=`<div class="rgroup"><div class="home-group"><div class="home-stack"><div class="ribbonrow"><select class="ctrl home-select" id="fontFamily" style="width:142px"><option>Aptos Narrow</option><option>Arial</option><option>Calibri</option><option>Tahoma</option><option>Times New Roman</option><option>Verdana</option><option>Georgia</option></select><select class="ctrl home-select" id="fontSize" style="width:58px"><option>8</option><option>9</option><option>10</option><option selected>11</option><option>12</option><option>14</option><option>16</option><option>18</option><option>20</option><option>24</option><option>28</option><option>36</option></select><button class="home-iconbtn" data-act="growfont" title="Increase Font Size">A⁺</button><button class="home-iconbtn" data-act="shrinkfont" title="Decrease Font Size">A⁻</button></div><div class="home-tools"><button class="home-iconbtn" data-act="bold" title="Bold"><b>B</b></button><button class="home-iconbtn" data-act="italic" title="Italic"><i>I</i></button><button class="home-iconbtn" data-act="underline" title="Underline"><u>U</u>⌄</button><button class="home-iconbtn" data-act="border" title="Borders">▦⌄</button><button class="home-iconbtn" title="Fill Color" data-act="fillmenu"><span class="fill-color-mark">▰</span>⌄</button><button class="home-iconbtn" title="Font Color" data-act="fontmenu"><span class="font-color-mark">A</span>⌄</button><label class="home-iconbtn" title="Choose Font Color" style="display:flex;align-items:center;gap:2px">A<input type="color" id="fontColor" value="#e53935" style="width:17px;height:18px;border:0;padding:0"></label><label class="home-iconbtn" title="Choose Fill Color" style="display:flex;align-items:center;gap:2px">▰<input type="color" id="fillColor" value="#fff2cc" style="width:17px;height:18px;border:0;padding:0"></label></div></div><div class="home-label">Font</div></div></div>`;
  h+=`<div class="rgroup"><div class="home-group"><div class="home-stack"><div class="home-tools"><button class="home-iconbtn" data-act="top" title="Top Align">⇡</button><button class="home-iconbtn" data-act="middle" title="Middle Align">↕</button><button class="home-iconbtn" data-act="bottom" title="Bottom Align">⇣</button><button class="home-iconbtn" data-act="orientation" title="Orientation">↗⌄</button><button class="home-iconbtn" data-act="indent" title="Decrease Indent">⇤</button><button class="home-iconbtn" data-act="indentmore" title="Increase Indent">⇥</button></div><div class="home-tools"><button class="home-iconbtn" data-act="left" title="Align Left">☰</button><button class="home-iconbtn" data-act="center" title="Center">≡</button><button class="home-iconbtn" data-act="right" title="Align Right">☷</button><button class="home-iconbtn" data-act="wrap" title="Wrap Text" style="font-size:12px">↪ Wrap Text</button><button class="home-iconbtn" data-act="merge" title="Merge & Center" style="font-size:12px">▣ Merge & Center⌄</button></div></div><div class="home-label">Alignment</div></div></div>`;
  h+=`<div class="rgroup"><div class="home-group"><div class="home-stack"><select class="ctrl home-select" id="numberFormat" style="width:178px"><option value="general">General</option><option value="number">Number</option><option value="currency">Currency</option><option value="percent">Percentage</option></select><div class="home-tools"><button class="home-iconbtn" data-act="currency" title="Accounting Number Format">$⌄</button><button class="home-iconbtn" data-act="percent" title="Percent Style">%</button><button class="home-iconbtn" data-act="thousands" title="Comma Style">,</button><button class="home-iconbtn" data-act="decdec" title="Decrease Decimal">.0 ←</button><button class="home-iconbtn" data-act="incdec" title="Increase Decimal">.0 →</button></div></div><div class="home-label">Number</div></div></div>`;
  h+=`<div class="rgroup"><div class="home-group"><div class="home-tools"><button class="tool" data-act="conditional" title="Conditional Formatting"><span class="ico">▦</span><span>Conditional<br>Formatting⌄</span></button><button class="tool" data-act="styles" title="Format as Table"><span class="ico">▤</span><span>Format as<br>Table⌄</span></button><button class="tool" data-act="styles" title="Cell Styles"><span class="ico">▧</span><span>Cell<br>Styles⌄</span></button></div><div class="home-label">Styles</div></div></div>`;
  h+=`<div class="rgroup"><div class="home-group"><div class="home-tools"><button class="tool" data-act="insert" title="Insert Cells"><span class="ico">▤↔</span><span>Insert⌄</span></button><button class="tool" data-act="delete" title="Delete Cells"><span class="ico">▤×</span><span>Delete⌄</span></button><button class="tool" data-act="colwidth" title="Format Cells"><span class="ico">▤</span><span>Format⌄</span></button></div><div class="home-label">Cells</div></div></div>`;
  h+=`<div class="rgroup"><div class="home-group"><div class="home-tools"><div class="home-stack"><button class="home-iconbtn" data-act="autosum" title="AutoSum">∑⌄</button><button class="home-iconbtn" data-act="clear" title="Clear">◇⌄</button></div><div class="home-stack"><button class="home-iconbtn" data-act="sortaz" title="Sort & Filter">A↧⌄</button><button class="home-iconbtn" data-act="find" title="Find & Select">⌕ Find & Select⌄</button></div></div><div class="home-label">Editing</div></div></div>`;
 }
 if(activeTab==='Insert'){
  const ib=(icon,label,act,cls='')=>`<button class="tool insert-tool ${cls}" data-act="${act}" title="${label}"><span class="ico">${icon}</span><span>${label}</span></button>`;
  const mini=(icon,label,act)=>`<button class="insert-mini" data-act="${act}" title="${label}"><span>${icon}</span><small>${label}</small></button>`;
  h+=group('Tables',ib('▦','PivotTable','pivot','wide')+ib('▤','Recommended<br>PivotTables','recommendedpivot','wide')+ib('▧','Table','table','wide'));
  h+=group('Illustrations',ib('▧','Pictures','picture','wide')+ib('◯','Shapes','shapes','wide')+ib('♧','Icons','icons','wide')+ib('◇','3D Models','models','wide')+ib('▤','SmartArt','smartart','wide')+ib('▣','Screenshot','screenshot','wide'));
  h+=group('Charts',`<div class="insert-chart-grid">${mini('▥','Recommended','recommendedcharts')}${mini('▥','Column','chart')}${mini('⌁','Line','linechart')}${mini('◔','Pie','piechart')}${mini('▤','Bar','barchart')}${mini('⌖','Scatter','scatterchart')}</div>`);
  h+=group('Sparklines',ib('⌁','Line','sparkline','wide')+ib('▥','Column','sparkcolumn','wide')+ib('▰','Win/Loss','sparkwinloss','wide'));
  h+=group('Filters',ib('▽','Slicer','slicer','wide')+ib('▤','Timeline','timeline','wide'));
  h+=group('Links',ib('🔗','Link','link','wide'));
  h+=group('Text',ib('T','Text Box','textbox','wide')+ib('▤','Header & Footer','headerfooter','wide')+ib('W','WordArt','wordart','wide')+ib('✎','Signature Line','signature','wide')+ib('▣','Object','object','wide'));
  h+=group('Symbols',ib('π','Equation','equation','wide')+ib('Ω','Symbol','symbol','wide'));
}
 if(activeTab==='Page Layout'){
 h+=group('Themes',btn('🎨','Themes','themes')+btn('🎨','Colors','themecolors')+btn('🔤','Fonts','themefonts'));
 h+=group('Page Setup',btn('📄','Margins','margins')+btn('↔','Orientation','orientation')+btn('▤','Size','papersize')+btn('▧','Print Area','printarea')+btn('↕','Breaks','pagebreaks')+btn('🖨','Print Titles','printtitles'));
 h+=group('Scale to Fit',btn('↔','Width','width')+btn('↕','Height','height')+btn('％','Scale','scale'));
 h+=group('Sheet Options',btn('▦','Gridlines','gridlines')+btn('▣','Headings','headings')+btn('⇄','Sheet Direction','sheetdirection'));
 h+=group('Arrange',btn('↥','Bring Forward','bringfront')+btn('↧','Send Backward','sendback')+btn('◫','Selection Pane','selectionpane'));
 }
 if(activeTab==='Formulas'){
 h+=group('Function Library',btn('∑','AutoSum','autosum')+btn('ƒx','Insert Function','function')+btn('Σ','Recently Used','recent'));
 h+=group('Categories',btn('＋','Math & Trig','math')+btn('▤','Statistical','stat')+btn('⌕','Logical','logical')+btn('▦','Text','textfn')+btn('📅','Date & Time','datefn'));
 h+=group('Defined Names',btn('◇','Define Name','definename')+btn('▤','Name Manager','namemanager'));
 h+=group('Formula Auditing',btn('⇢','Trace Precedents','precedents')+btn('⇠','Trace Dependents','dependents')+btn('✓','Check Formula','checkformula')+btn('ƒ','Show Formulas','showformulas'));
}
 if(activeTab==='Data'){h+=group('Sort & Filter',btn('⬇','Sort A to Z','sortaz')+btn('⬆','Sort Z to A','sortza')+btn('▽','Filter','filter')+btn('✕','Clear Filter','clearfilter'));h+=group('Data Tools',btn('◈','Remove Duplicates','duplicates')+btn('⌁','Text to Columns','textcolumns')+btn('✓','Data Validation','validation'));h+=group('Outline',btn('＋','Group','group')+btn('－','Ungroup','ungroup'))}
 if(activeTab==='File'){h+=group('Workbook',btn('＋','New Workbook','newfile')+btn('📂','Open CSV','openfile')+btn('💾','Save','savefile')+btn('⬇','Export CSV','exportfile'));h+=group('Print',btn('🖨','Print','printfile')+btn('ℹ','Workbook Info','fileinfo'))}
 if(activeTab==='Review'){h+=group('Proofing',btn('✓','Spelling','spelling')+btn('🌐','Language','language'));h+=group('Comments',btn('💬','New Comment','comment')+btn('🗨','Show Comments','showcomments')+btn('✕','Delete Comment','deletecomment'));h+=group('Protection',btn('🔒','Protect Sheet','protect')+btn('🔑','Protect Workbook','protectbook')+btn('🔓','Unprotect','unprotect'))}
 if(activeTab==='View'){h+=group('Workbook Views',btn('▦','Normal','normal')+btn('▤','Page Layout','pagelayout')+btn('▧','Page Break','pagebreak'));h+=group('Show',btn('☑','Gridlines','gridlines')+btn('☑','Headings','headings')+btn('☰','Formula Bar','formulabar'));h+=group('Zoom',btn('－','Zoom Out','zoomout')+btn('＋','Zoom In','zoomin')+btn('100%','100%','zoom100'))}
 $('ribbon').innerHTML=h;$('ribbon').querySelectorAll('[data-act]').forEach(b=>b.onclick=()=>act(b.dataset.act));
 let fc=$('fontColor'),bc=$('fillColor'),nf=$('numberFormat'),ff=$('fontFamily'),fs=$('fontSize');
 if(fc)fc.oninput=()=>fmt('color',fc.value);if(bc)bc.oninput=()=>fmt('bg',bc.value);if(nf){nf.value=cells[selected]?.format||'general';nf.onchange=()=>fmt('format',nf.value);}
 if(ff){ff.value=cells[selected]?.fontFamily||'Aptos Narrow';ff.onchange=()=>fmt('fontFamily',ff.value);}if(fs){fs.value=String(cells[selected]?.fontSize||11);fs.onchange=()=>fmt('fontSize',Number(fs.value));}
}
function fmt(k,v){selectedRefs.forEach(ref=>cells[ref]={...(cells[ref]||{}),[k]:v});render();select(cellName.value||selected)}
function tell(title,msg){$('modalTitle').textContent=title;$('modalText').textContent=msg;$('modal').classList.add('show')}
$('modalClose').onclick=()=>$('modal').classList.remove('show');
function objectCard(title,content){
 const wrap=$('insertedObjects'),card=document.createElement('div');card.className='inserted-object';
 const heading=document.createElement('h4');heading.textContent=title;card.appendChild(heading);
 const close=document.createElement('button');close.className='remove-object';close.textContent='×';close.title='Remove';close.onclick=()=>card.remove();card.appendChild(close);
 if(typeof content==='string')card.insertAdjacentHTML('beforeend',content);else card.appendChild(content);
 wrap.appendChild(card);card.scrollIntoView({block:'nearest',behavior:'smooth'});
}

function chartFromSelection(kind){
 const refs=selectedRefs.slice(),ps=refs.map(refParts);
 if(!refs.length){tell('ساخت نمودار','ابتدا محدوده داده‌ها را انتخاب کن.');return;}
 const minR=Math.min(...ps.map(x=>x.r)),maxR=Math.max(...ps.map(x=>x.r)),minC=Math.min(...ps.map(x=>x.c)),maxC=Math.max(...ps.map(x=>x.c));
 if(maxC-minC<1||maxR-minR<1){tell('ساخت نمودار','حداقل دو ستون و دو ردیف انتخاب کن؛ ردیف اول عنوان ستون‌ها باشد.');return;}
 // Read the selected rectangle as a true 2D table. First row = series names,
 // first column = category labels. Keep every selected data row and data column.
 const matrix=[];
 for(let r=minR;r<=maxR;r++){
  const row=[];
  for(let c=minC;c<=maxC;c++)row.push(String(raw(colName(c)+r)??'').trim());
  matrix.push(row);
 }
 const toNumber=v=>{const rawv=String(v??'').replace(/[,٬\s]/g,'').replace('٫','.').trim();return rawv!==''&&Number.isFinite(Number(rawv))?Number(rawv):null;};
 let labels=[],series=[];
 // Handle the common classroom layout where subjects/categories are in the first
 // row and their values are in the second row (e.g. ریاضی | عربی / 14 | 12).
 const rowValues=matrix.length===2?matrix[1].map(toNumber):[];
 const rowLayout=matrix.length===2 && matrix[0].length>=2 &&
   matrix[0].every(v=>v!=='') && rowValues.filter(v=>v!==null).length>=2;
 if(rowLayout && kind!=='scatterchart'){
  labels=matrix[0].slice();
  series=[{name:'نمره',values:rowValues}];
 }else{
  labels=matrix.slice(1).map((row,i)=>row[0]||('ردیف '+(i+1)));
  for(let c=1;c<matrix[0].length;c++){
   const name=matrix[0][c]||('مجموعه '+c);
   const values=matrix.slice(1).map(row=>toNumber(row[c]));
   if(values.some(v=>v!==null))series.push({name,values});
  }
 }
 if(!series.length){tell('ساخت نمودار','در ستون‌های انتخاب‌شده عددی پیدا نشد. محدوده باید مانند A1:C5 باشد: عنوان‌ها در ردیف اول و نام دسته‌ها در ستون اول.');return;}
 if(!series.length){tell('ساخت نمودار','در محدوده انتخاب‌شده داده عددی پیدا نشد.');return;}
 const card=document.createElement('div');card.className='inserted-object chart-object';
 card.style.cssText='position:relative;background:#fff;border:1px solid #cfd8d1;border-radius:8px;padding:12px;margin:12px 0;resize:both;overflow:auto;min-width:360px;max-width:100%;';
 const head=document.createElement('div');head.style.cssText='display:flex;align-items:center;gap:8px;cursor:move;padding:4px 2px;border-bottom:1px solid #e5eae6;margin-bottom:8px';
 const heading=document.createElement('strong');heading.textContent='نمودار جدید';head.appendChild(heading);
 const hint=document.createElement('small');hint.textContent='برای جابه‌جایی، این نوار را بکشید';hint.style.color='#77827a';head.appendChild(hint);
 const remove=document.createElement('button');remove.className='remove-object';remove.textContent='×';remove.title='حذف نمودار';remove.onclick=()=>card.remove();head.appendChild(remove);card.appendChild(head);
 let drag=false,dx=0,dy=0;
 head.addEventListener('pointerdown',e=>{if(e.target===remove)return;drag=true;dx=e.clientX;dy=e.clientY;head.setPointerCapture(e.pointerId)});
 head.addEventListener('pointermove',e=>{if(!drag)return;const wrap=$('insertedObjects');const r=card.getBoundingClientRect();card.style.transform=`translate(${e.clientX-dx}px,${e.clientY-dy}px)`;dx=e.clientX;dy=e.clientY;});
 head.addEventListener('pointerup',()=>{drag=false;card.style.transform='';});
 const controls=document.createElement('div');controls.style.cssText='display:flex;gap:8px;flex-wrap:wrap;align-items:center;background:#f7f9f7;padding:9px;border-radius:6px;margin:8px 0';
 const mkSelect=(label,opts)=>{const l=document.createElement('label');l.style.cssText='font-size:12px;display:flex;gap:4px;align-items:center';l.append(label);const sel=document.createElement('select');sel.className='ctrl';opts.forEach(([v,t])=>{const o=document.createElement('option');o.value=v;o.textContent=t;sel.appendChild(o)});l.appendChild(sel);controls.appendChild(l);return sel};
 const type=kind==='linechart'?'line':kind==='piechart'?'pie':kind==='barchart'?'bar':'column';
 const typeSel=mkSelect('نوع نمودار', [['column','ستونی خوشه‌ای'],['stackedColumn','ستونی انباشته'],['bar','میله‌ای'],['line','خطی'],['area','مساحتی'],['pie','دایره‌ای'],['doughnut','حلقه‌ای'],['scatter','پراکندگی']]);typeSel.value=type;
 const titleInput=document.createElement('input');titleInput.className='ctrl';titleInput.placeholder='عنوان نمودار';titleInput.value='عنوان نمودار';titleInput.style.width='150px';controls.appendChild(titleInput);
 const xTitle=document.createElement('input');xTitle.className='ctrl';xTitle.placeholder='عنوان محور افقی';xTitle.style.width='125px';controls.appendChild(xTitle);
 const yTitle=document.createElement('input');yTitle.className='ctrl';yTitle.placeholder='عنوان محور عمودی';yTitle.style.width='125px';controls.appendChild(yTitle);
 const check=(label,checked)=>{const l=document.createElement('label');l.style.cssText='font-size:12px;display:flex;gap:4px;align-items:center';const i=document.createElement('input');i.type='checkbox';i.checked=checked;l.append(i,label);controls.appendChild(l);return i};
 const legend=check('راهنما',true),labelsOn=check('برچسب داده',false),gridOn=check('خطوط شبکه',true);
 const palette=mkSelect('رنگ', [['0','آبی و نارنجی'],['1','سبز و آبی'],['2','رنگارنگ'],['3','خاکستری']]);
 const legendPos=mkSelect('جایگاه راهنما',[['right','راست'],['bottom','پایین'],['top','بالا'],['left','چپ']]);
 const switchBtn=document.createElement('button');switchBtn.className='secondary';switchBtn.textContent='تعویض سطر/ستون';switchBtn.title='اگر داده‌ها به‌جای ستون‌ها در ردیف‌ها قرار دارند، این گزینه را بزنید.';controls.appendChild(switchBtn);
 const orientationNote=document.createElement('small');orientationNote.textContent=(matrix.length===2&&rowLayout)?'چیدمان تشخیص‌داده‌شده: عنوان دسته‌ها در ردیف اول و مقادیر در ردیف دوم. برای مقایسه این داده‌ها، نمودار ستونی یا میله‌ای مناسب است.':'حالت فعلی: هر ستون یک مجموعه داده است؛ ردیف اول عنوان مجموعه‌ها و ستون اول نام دسته‌هاست.';orientationNote.style.cssText='font-size:11px;color:#5d6b61';controls.appendChild(orientationNote);
 const canvas=document.createElement('div');canvas.style.cssText='width:100%;min-height:300px;overflow:auto;background:white;border:1px solid #e0e6e1;border-radius:5px;padding:8px;direction:ltr';card.append(controls,canvas);
 let swapped=false;
 const palettes=[['#4472C4','#ED7D31','#A5A5A5','#FFC000','#70AD47'],['#217346','#5B9BD5','#ED7D31','#8064A2','#C0504D'],['#5B9BD5','#70AD47','#FFC000','#8064A2','#F4B183'],['#606060','#909090','#b8b8b8','#404040','#d0d0d0']];
 function draw(){
  let cats=labels.slice(),ss=series.map(z=>({name:z.name,values:z.values.slice()}));
  if(swapped){const n=ss.map(z=>z.name);ss=cats.map((cat,i)=>({name:cat,values:ss.map(z=>z.values[i]??null)}));cats=n;}
  let t=typeSel.value;
  if(t==='scatter'){
   const count=ss.reduce((n,z)=>n+z.values.filter(v=>v!==null).length,0);
   if(count<2){canvas.innerHTML='<div style="padding:30px;text-align:center;color:#8a4b00">برای نمودار پراکندگی حداقل دو جفت داده عددی لازم است. برای مقایسه نمره‌های ریاضی و عربی، نوع نمودار را روی «ستونی خوشه‌ای» قرار بده.</div>';return;}
  }
  if((t==='pie'||t==='doughnut')&&ss.length>1)ss=ss.slice(0,1);
  const colors=palettes[Number(palette.value)||0],W=760,H=410,legendSide=legend.checked?legendPos.value:'none';
  const left=72,right=legendSide==='right'?165:30,top=legendSide==='top'?72:52,bottom=legendSide==='bottom'?105:75;
  const pw=W-left-right,ph=H-top-bottom,all=ss.flatMap(z=>z.values.filter(v=>v!==null));
  const min=Math.min(0,...all),max=Math.max(0,...all),span=max-min||1;
  const y=v=>top+(max-v)/span*ph,base=y(0),colorsFor=i=>colors[i%colors.length];
  let svg=`<svg viewBox="0 0 ${W} ${H}" width="100%" style="min-width:560px;font-family:Arial,sans-serif"><text x="${W/2}" y="27" text-anchor="middle" font-size="19" font-weight="bold" fill="#26352b">${esc(titleInput.value||'نمودار')}</text>`;
  if(t==='pie'||t==='doughnut'){
   const vals=ss[0].values.map(v=>Math.max(0,v||0)),total=vals.reduce((a,b)=>a+b,0),cx=left+pw/2,cy=top+ph/2,rad=Math.min(pw,ph)*.38;
   if(total<=0)svg+=`<text x="${W/2}" y="${H/2}" text-anchor="middle" fill="#a33">برای نمودار دایره‌ای اعداد مثبت وارد کنید</text>`;
   else{let a=-Math.PI/2;vals.forEach((v,i)=>{const a2=a+v/total*Math.PI*2,x1=cx+rad*Math.cos(a),y1=cy+rad*Math.sin(a),x2=cx+rad*Math.cos(a2),y2=cy+rad*Math.sin(a2),large=a2-a>Math.PI?1:0;svg+=`<path d="M ${cx} ${cy} L ${x1} ${y1} A ${rad} ${rad} 0 ${large} 1 ${x2} ${y2} Z" fill="${colorsFor(i)}" stroke="#fff" stroke-width="2"/>`;if(labelsOn.checked){const mid=(a+a2)/2;svg+=`<text x="${cx+rad*.68*Math.cos(mid)}" y="${cy+rad*.68*Math.sin(mid)}" text-anchor="middle" font-size="12" fill="#fff">${Math.round(v/total*100)}%</text>`;}a=a2;});if(t==='doughnut')svg+=`<circle cx="${cx}" cy="${cy}" r="${rad*.53}" fill="#fff"/><text x="${cx}" y="${cy+5}" text-anchor="middle" font-size="13" fill="#444">${total}</text>`;}
  }else{
   const groupW=pw/Math.max(1,cats.length),n=ss.length,bw=Math.max(4,groupW*.76/(t==='stackedColumn'?1:n));
   if(gridOn.checked){for(let i=0;i<=5;i++){const v=min+span*i/5,yy=y(v);svg+=`<line x1="${left}" y1="${yy}" x2="${left+pw}" y2="${yy}" stroke="#e6ebe7"/><text x="${left-9}" y="${yy+4}" text-anchor="end" font-size="11" fill="#66736b">${Number(v.toFixed(1))}</text>`;}}
   ss.forEach((ser,j)=>{
    if(t==='line'||t==='area'||t==='scatter'){
     const pts=ser.values.map((v,i)=>v===null?null:[left+i*groupW+groupW/2,y(v)]).filter(Boolean);
     if(t==='area'&&pts.length){const poly=`${pts[0][0]},${base} ${pts.map(p=>p.join(',')).join(' ')} ${pts[pts.length-1][0]},${base}`;svg+=`<polygon points="${poly}" fill="${colorsFor(j)}" opacity=".16"/>`;}
     if(t==='scatter'){pts.forEach(p=>svg+=`<circle cx="${p[0]}" cy="${p[1]}" r="5" fill="${colorsFor(j)}"/>`);}
     else {svg+=`<polyline points="${pts.map(p=>p.join(',')).join(' ')}" fill="none" stroke="${colorsFor(j)}" stroke-width="3"/>`;pts.forEach(p=>svg+=`<circle cx="${p[0]}" cy="${p[1]}" r="4" fill="${colorsFor(j)}"/>`);}
     if(labelsOn.checked)ser.values.forEach((v,i)=>{if(v!==null)svg+=`<text x="${left+i*groupW+groupW/2}" y="${y(v)-9}" text-anchor="middle" font-size="10">${v}</text>`;});
    }else{
     ser.values.forEach((v,i)=>{if(v===null)return;let x=left+i*groupW+groupW*.12+(t==='stackedColumn'?0:j*bw),yy=y(v),height=Math.abs(base-yy);if(t==='stackedColumn'&&j>0){let prev=ss.slice(0,j).reduce((sum,z)=>sum+(z.values[i]||0),0);yy=y(prev+v);height=Math.abs(y(prev)-yy);}
      const horizontal=t==='bar';if(horizontal){/* horizontal bars are rendered as rotated category bars below */}else svg+=`<rect x="${x}" y="${Math.min(yy,base)}" width="${bw-3}" height="${Math.max(1,height)}" fill="${colorsFor(j)}" rx="2"/>`;
      if(labelsOn.checked)svg+=`<text x="${x+bw/2}" y="${Math.max(top+12,Math.min(yy,base)-5)}" text-anchor="middle" font-size="10">${v}</text>`;
     });
    }
   });
   if(t==='bar'){svg=svg.replace(/<svg[^>]*>/,'<svg viewBox="0 0 '+W+' '+H+'" width="100%" style="min-width:560px;font-family:Arial,sans-serif">');ss.forEach((ser,j)=>ser.values.forEach((v,i)=>{if(v===null)return;const rowH=ph/(cats.length*Math.max(1,ss.length)),yy=top+i*rowH*ss.length+j*rowH;svg+=`<rect x="${left}" y="${yy+2}" width="${Math.max(1,(v-min)/(span||1)*pw)}" height="${Math.max(4,rowH-4)}" fill="${colorsFor(j)}" rx="2"/>`;}));}
   if(t==='bar'){
    cats.forEach((cat,i)=>{const rowH=ph/Math.max(1,cats.length);svg+=`<text x="${left-8}" y="${top+(i+.5)*rowH+4}" text-anchor="end" font-size="11" fill="#4b574f">${esc(cat.slice(0,18))}</text>`;});
   }else{
    cats.forEach((cat,i)=>svg+=`<text x="${left+i*groupW+groupW/2}" y="${H-45}" text-anchor="middle" font-size="11" fill="#4b574f">${esc(cat.slice(0,16))}</text>`);
   }
   if(xTitle.value)svg+=`<text x="${left+pw/2}" y="${H-10}" text-anchor="middle" font-size="12">${esc(xTitle.value)}</text>`;
   if(yTitle.value)svg+=`<text transform="translate(16 ${top+ph/2}) rotate(-90)" text-anchor="middle" font-size="12">${esc(yTitle.value)}</text>`;
  }
  if(legend.checked){if(legendSide==='right'||legendSide==='left'){const lx=legendSide==='right'?W-right+12:8;ss.forEach((z,i)=>svg+=`<rect x="${lx}" y="${top+i*24}" width="12" height="12" fill="${colorsFor(i)}"/><text x="${lx+18}" y="${top+11+i*24}" font-size="11" fill="#333">${esc(z.name.slice(0,20))}</text>`);}else{const y0=legendSide==='top'?48:H-22;ss.forEach((z,i)=>{const x=left+i*120;svg+=`<rect x="${x}" y="${y0-10}" width="10" height="10" fill="${colorsFor(i)}"/><text x="${x+15}" y="${y0}" font-size="11">${esc(z.name.slice(0,15))}</text>`;});}}
  svg+='</svg>';canvas.innerHTML=svg;
 }
 [typeSel,titleInput,xTitle,yTitle,legend,labelsOn,gridOn,palette,legendPos].forEach(el=>{el.addEventListener('change',draw);el.addEventListener('input',draw)});
 switchBtn.onclick=()=>{swapped=!swapped;orientationNote.textContent=swapped?'حالت فعلی: هر ردیف یک مجموعه داده است.':'حالت فعلی: هر ستون یک مجموعه داده است؛ ردیف اول عنوان مجموعه‌ها و ستون اول نام دسته‌هاست.';draw()};
 $('insertedObjects').appendChild(card);card.scrollIntoView({block:'nearest',behavior:'smooth'});draw();
}
function insertPicture(){let input=document.createElement('input');input.type='file';input.accept='image/*';input.onchange=()=>{let f=input.files[0];if(!f)return;let reader=new FileReader();reader.onload=()=>{let img=document.createElement('img');img.src=reader.result;img.alt=f.name;objectCard('تصویر: '+f.name,img)};reader.readAsDataURL(f)};input.click();}

function insertFileObject(accept,title){
 const input=document.createElement('input');input.type='file';input.accept=accept;
 input.onchange=()=>{const f=input.files&&input.files[0];if(!f)return;
  const reader=new FileReader();reader.onload=()=>{const box=document.createElement('div');
   if(f.type.startsWith('image/')){const img=document.createElement('img');img.src=reader.result;img.alt=f.name;box.appendChild(img);}
   else {const a=document.createElement('a');a.href=reader.result;a.download=f.name;a.textContent='دریافت فایل: '+f.name;box.appendChild(a);}
   objectCard(title+': '+f.name,box);
  };reader.readAsDataURL(f);
 };input.click();
}
function insertSparkline(kind){
 const ps=selectedRefs.map(refParts);if(ps.length<2){tell('نمودار درون‌سلولی','ابتدا حداقل دو سلول عددی را انتخاب کن.');return;}
 const vals=selectedRefs.map(r=>Number(raw(r))).filter(Number.isFinite);
 if(vals.length<2){tell('نمودار درون‌سلولی','محدوده انتخابی باید حداقل دو عدد داشته باشد.');return;}
 const max=Math.max(...vals,1),min=Math.min(...vals,0),w=220,h=65;
 let svg='';
 if(kind==='sparkline'){const pts=vals.map((v,i)=>`${8+i*(w-16)/(vals.length-1)},${h-8-(v-min)/(max-min||1)*(h-16)}`).join(' ');svg=`<polyline points="${pts}" fill="none" stroke="#217346" stroke-width="3"/>`;}
 else {const bw=(w-16)/vals.length;svg=vals.map((v,i)=>{let bh=(v-min)/(max-min||1)*(h-16)+3;return `<rect x="${8+i*bw}" y="${h-8-bh}" width="${Math.max(2,bw-3)}" height="${bh}" fill="${kind==='sparkwinloss'?(v>=0?'#217346':'#c0504d'):'#5b9bd5'}"/>`}).join('');}
 const el=document.createElement('div');el.innerHTML=`<svg viewBox="0 0 ${w} ${h}" width="${w}" height="${h}">${svg}</svg>`;objectCard('نمودار کوچک '+(kind==='sparkline'?'خطی':kind==='sparkcolumn'?'ستونی':'برد/باخت'),el.firstElementChild);
}
function insertSmartArt(){
 const text=prompt('عبارت‌ها را با ویرگول جدا کن (حداکثر ۶ مورد):','برنامه‌ریزی، اجرا، ارزیابی');if(!text)return;
 const items=text.split(/[،,]/).map(x=>x.trim()).filter(Boolean).slice(0,6);if(!items.length)return;
 const style=prompt('نوع چیدمان:\\n1 = فرایندی افقی\\n2 = چرخه‌ای\\n3 = فهرست عمودی','1')||'1';
 const el=document.createElement('div');el.style.cssText='display:flex;gap:9px;align-items:center;justify-content:center;flex-wrap:wrap;padding:22px;max-width:760px';
 if(style==='2'){
  el.style.cssText+=';display:grid;grid-template-columns:repeat(3,minmax(110px,1fr));';
  items.forEach((t,i)=>{const b=document.createElement('div');b.textContent=t;b.style.cssText=`min-height:75px;display:grid;place-items:center;text-align:center;padding:12px;border-radius:50%;background:${['#e2f0d9','#ddebf7','#fff2cc','#fce4d6'][i%4]};border:3px solid ${['#70ad47','#4472c4','#bf9000','#c65911'][i%4]};font-weight:bold`;el.appendChild(b);});
 }else if(style==='3'){
  el.style.cssText+=';display:flex;flex-direction:column;align-items:stretch;max-width:420px';
  items.forEach((t,i)=>{const b=document.createElement('div');b.textContent=(i+1)+'. '+t;b.style.cssText='padding:12px 18px;background:#e8f3ec;border-right:5px solid #217346;font-weight:bold';el.appendChild(b);});
 }else{
  items.forEach((t,i)=>{const b=document.createElement('div');b.textContent=t;b.style.cssText='padding:15px 20px;min-width:100px;text-align:center;background:#e8f3ec;border:2px solid #217346;border-radius:6px;font-weight:bold';el.appendChild(b);if(i<items.length-1){const arrow=document.createElement('span');arrow.textContent='➜';arrow.style.fontSize='24px';el.appendChild(arrow);}});
 }
 objectCard('SmartArt آموزشی',el);
}
function insertIcons(){
 const choices=[['★','ستاره'],['✓','تأیید'],['❤','قلب'],['☀','خورشید'],['☎','تلفن'],['✉','ایمیل'],['⌂','خانه'],['⚙','تنظیمات'],['♟','مهره'],['♫','موسیقی'],['➜','پیکان'],['◆','لوزی'],['●','دایره'],['■','مربع'],['⚑','پرچم'],['♻','بازیافت'],['✚','به‌علاوه'],['⌕','جست‌وجو']];
 const el=document.createElement('div');el.style.cssText='display:grid;grid-template-columns:repeat(6,minmax(55px,1fr));gap:8px;padding:12px';
 choices.forEach(([symbol,name])=>{const b=document.createElement('button');b.type='button';b.className='secondary';b.title=name;b.style.cssText='min-height:68px;display:flex;flex-direction:column;align-items:center;justify-content:center;font-size:25px';const big=document.createElement('span');big.textContent=symbol;const small=document.createElement('small');small.textContent=name;small.style.fontSize='10px';b.append(big,small);b.onclick=()=>{const chosen=document.createElement('div');chosen.textContent=symbol;chosen.style.cssText='font-size:56px;color:#217346;text-align:center;padding:12px';objectCard('آیکون: '+name,chosen);};el.appendChild(b);});
 objectCard('گالری آیکون‌ها – برای درج روی نماد کلیک کن',el);
}
function insertScatter(){
 const ps=selectedRefs.map(refParts);if(!ps.length)return;
 const minR=Math.min(...ps.map(x=>x.r)),maxR=Math.max(...ps.map(x=>x.r)),minC=Math.min(...ps.map(x=>x.c)),maxC=Math.max(...ps.map(x=>x.c));
 if(maxC-minC<1){tell('نمودار پراکندگی','دو ستون عددی را انتخاب کن: ستون اول X و ستون دوم Y.');return;}
 const pts=[];for(let r=minR+(maxR>minR?1:0);r<=maxR;r++){const x=Number(raw(colName(minC)+r)),y=Number(raw(colName(minC+1)+r));if(Number.isFinite(x)&&Number.isFinite(y))pts.push([x,y]);}
 if(pts.length<2){tell('نمودار پراکندگی','حداقل دو جفت عددی لازم است.');return;}
 const xs=pts.map(p=>p[0]),ys=pts.map(p=>p[1]),xmin=Math.min(...xs),xmax=Math.max(...xs),ymin=Math.min(...ys),ymax=Math.max(...ys);
 const dots=pts.map(([x,y])=>`<circle cx="${35+(x-xmin)/(xmax-xmin||1)*360}" cy="${210-(y-ymin)/(ymax-ymin||1)*175}" r="5" fill="#217346"/>`).join('');
 const el=document.createElement('div');el.innerHTML=`<svg viewBox="0 0 420 235" width="420"><line x1="30" y1="210" x2="405" y2="210" stroke="#777"/><line x1="30" y1="15" x2="30" y2="210" stroke="#777"/>${dots}</svg>`;objectCard('نمودار پراکندگی',el.firstElementChild);
}

let pageSettings={margin:'normal',orientation:'portrait',paper:'A4',fitWidth:'1',fitHeight:'0',scale:100,printArea:'',printTitles:''};
let sheetDirection='ltr';
function applySheetDirection(){
 const rtl=sheetDirection==='rtl';
 grid.style.direction=rtl?'rtl':'ltr';
 grid.style.textAlign=rtl?'right':'left';
 const wrap=$('gridwrap');
 if(wrap)wrap.style.direction=rtl?'rtl':'ltr';
}

function pageMessage(msg){tell('تنظیمات صفحه',msg);}
function setPageCss(){
 let old=document.getElementById('pageSettingsStyle');if(old)old.remove();
 const st=document.createElement('style');st.id='pageSettingsStyle';
 const m={normal:'15mm',narrow:'7mm',wide:'25mm'}[pageSettings.margin]||'15mm';
 st.textContent=`@media print{@page{size:${pageSettings.paper} ${pageSettings.orientation};margin:${m}}body{background:white!important}.topbar,.tabs,.ribbon,.formulabar,.sheetbar,.status,.help{display:none!important}.workspace{padding:0!important;overflow:visible!important}.gridwrap{max-height:none!important;overflow:visible!important;border:0!important}table{min-width:0!important;width:100%!important}th,td{break-inside:avoid}}`;
 document.head.appendChild(st);
}
function printPreview(){setPageCss();pageMessage('تنظیمات چاپ اعمال شد. پنجره چاپ مرورگر باز می‌شود؛ می‌توانی پیش‌نمایش را ببینی یا چاپ را لغو کنی.');setTimeout(()=>window.print(),150);}

let recentFormulas=[];
let definedNames={};
let showFormulaMode=false;
function formulaWizard(category='all'){
 const catalog={
  all:[['SUM','جمع اعداد','SUM(A1:A5)'],['AVERAGE','میانگین','AVERAGE(A1:A5)'],['MIN','کمترین مقدار','MIN(A1:A5)'],['MAX','بیشترین مقدار','MAX(A1:A5)'],['COUNT','شمارش اعداد','COUNT(A1:A5)'],['COUNTA','شمارش سلول‌های پر','COUNTA(A1:A5)'],['IF','شرطی','IF(A1>=10,"قبول","مردود")'],['ROUND','گرد کردن','ROUND(A1,2)'],['ABS','قدر مطلق','ABS(A1)'],['SQRT','جذر','SQRT(A1)'],['CONCAT','ترکیب متن','CONCAT(A1," ",B1)'],['LEN','تعداد نویسه','LEN(A1)']],
  math:[['SUM','جمع','SUM(A1:A5)'],['PRODUCT','ضرب','PRODUCT(A1:A5)'],['ABS','قدر مطلق','ABS(A1)'],['SQRT','جذر','SQRT(A1)'],['ROUND','گرد کردن','ROUND(A1,2)'],['ROUNDUP','گرد کردن به بالا','ROUNDUP(A1,0)'],['ROUNDDOWN','گرد کردن به پایین','ROUNDDOWN(A1,0)']],
  stat:[['AVERAGE','میانگین','AVERAGE(A1:A5)'],['COUNT','شمارش اعداد','COUNT(A1:A5)'],['COUNTA','شمارش سلول‌های پر','COUNTA(A1:A5)'],['MIN','کمترین','MIN(A1:A5)'],['MAX','بیشترین','MAX(A1:A5)']],
  logical:[['IF','تابع شرطی','IF(A1>=10,"قبول","مردود")'],['AND','و','AND(A1>0,B1>0)'],['OR','یا','OR(A1>0,B1>0)'],['NOT','نقیض','NOT(A1>0)']],
  textfn:[['CONCAT','ترکیب متن','CONCAT(A1," ",B1)'],['LEFT','چند نویسه از چپ','LEFT(A1,3)'],['RIGHT','چند نویسه از راست','RIGHT(A1,3)'],['LEN','طول متن','LEN(A1)'],['UPPER','حروف بزرگ','UPPER(A1)'],['LOWER','حروف کوچک','LOWER(A1)'],['PROPER','حروف ابتدای کلمات','PROPER(A1)']],
  datefn:[['TODAY','تاریخ امروز','TODAY()'],['NOW','تاریخ و زمان','NOW()']]
 };
 const list=catalog[category]||catalog.all;
 const menu=list.map((x,i)=>`${i+1}. ${x[0]} — ${x[1]} — ${x[2]}`).join('\n');
 const pick=prompt('تابع موردنظر را انتخاب کن؛ شماره را وارد کن:\n'+menu,'1');
 if(!pick)return;const item=list[Number(pick)-1];if(!item){tell('Insert Function','شماره انتخاب‌شده معتبر نیست.');return;}
 const formulaText=prompt('فرمول را بررسی و در صورت نیاز ویرایش کن:', '='+item[2]);
 if(formulaText===null)return;formula.value=formulaText.startsWith('=')?formulaText:'='+formulaText;apply();
 recentFormulas.unshift(item[0]);recentFormulas=[...new Set(recentFormulas)].slice(0,8);
}
function showFormulaHelp(){
 tell('توابع قابل استفاده','ریاضی: SUM, PRODUCT, ABS, SQRT, ROUND, ROUNDUP, ROUNDDOWN\nآماری: AVERAGE, MIN, MAX, COUNT, COUNTA\nمنطقی: IF, AND, OR, NOT\nمتنی: CONCAT, LEFT, RIGHT, LEN, UPPER, LOWER, PROPER\nنمونه: =IF(A1>=10,"قبول","مردود")');
}
function traceFormula(which){
 const f=String(raw(selected)||'');if(!f.startsWith('=')){tell('Formula Auditing','سلول انتخاب‌شده فرمول ندارد.');return;}
 const refs=[...f.matchAll(/\b([A-Z]{1,3}\d+)(?::([A-Z]{1,3}\d+))?\b/g)].flatMap(m=>rangeRefs(m[2]?m[1]+':'+m[2]:m[1]));
 const unique=[...new Set(refs)].filter(r=>r!==selected);
 if(!unique.length){tell('Formula Auditing','ارجاع سلولی در این فرمول پیدا نشد.');return;}
 if(which==='precedents'){select(unique[0]);tell('Trace Precedents','سلول‌های مؤثر در فرمول: '+unique.join(', ')+'. نخستین سلول انتخاب شد.');}
 else {const dependents=Object.keys(cells).filter(r=>r!==selected&&String(raw(r)).startsWith('=')&&new RegExp('\\\\b'+selected+'\\\\b').test(String(raw(r))));tell('Trace Dependents',dependents.length?'سلول‌های وابسته: '+dependents.join(', '):'سلول وابسته‌ای پیدا نشد.');}
}
let cellComments={};
let sheetProtected=false,workbookProtected=false,spellCheckEnabled=true;
function showComments(){
 const entries=Object.entries(cellComments);
 tell('Comments',entries.length?entries.map(([r,c])=>r+': '+c).join('\\n'):'هیچ یادداشتی ثبت نشده است.');
}
function fileInfo(){
 tell('Workbook Info','نام فایل: '+($('bookName').value||'Book1')+'\\nتعداد برگه‌ها: '+sheets.length+'\\nبرگه فعال: '+sheets[activeSheetIndex].name+'\\nطراح: رامین حمیدی');
}
function act(a){


 let o=cells[selected]||{},p=refParts(selected);
 if(a==='newfile'){if(confirm('یک کارپوشه جدید ایجاد شود؟ تغییرات ذخیره‌نشده از بین می‌روند.'))$('newBtn').click();return;}
 if(a==='openfile'){$('openBtn').click();return;}
 if(a==='savefile'){$('saveBtn').click();return;}
 if(a==='exportfile'){$('downloadBtn').click();return;}
 if(a==='printfile'){window.print();return;}
 if(a==='fileinfo'){fileInfo();return;}
 if(a==='spelling'){spellCheckEnabled=!spellCheckEnabled;grid.querySelectorAll('td[contenteditable="true"]').forEach(td=>td.setAttribute('spellcheck',String(spellCheckEnabled)));tell('Spelling',spellCheckEnabled?'بررسی املایی مرورگر فعال شد.':'بررسی املایی غیرفعال شد.');return;}
 if(a==='language'){let v=prompt('زبان بررسی املایی مرورگر (مثلاً fa-IR یا en-US):','fa-IR');if(v){grid.setAttribute('lang',v);grid.querySelectorAll('td').forEach(td=>td.setAttribute('lang',v));tell('Language','زبان بررسی املایی روی '+v+' تنظیم شد. پشتیبانی به مرورگر وابسته است.');}return;}
 if(a==='comment'){let c=prompt('یادداشت برای سلول '+selected+':',cellComments[selected]||'');if(c!==null){if(c.trim())cellComments[selected]=c.trim();else delete cellComments[selected];render();select(cellName.value||selected);tell('Comment','یادداشت سلول '+selected+' ثبت شد.');}return;}
 if(a==='showcomments'){showComments();return;}
 if(a==='deletecomment'){if(cellComments[selected]){delete cellComments[selected];tell('Comment','یادداشت حذف شد.');}else tell('Comment','برای این سلول یادداشتی وجود ندارد.');return;}
 if(a==='protect'){sheetProtected=!sheetProtected;render();tell('Protect Sheet',sheetProtected?'محافظت فعال شد؛ ویرایش مستقیم سلول‌ها غیرفعال است.':'محافظت غیرفعال شد.');return;}
 if(a==='protectbook'){workbookProtected=!workbookProtected;tell('Protect Workbook',workbookProtected?'محافظت ساختار کارپوشه فعال شد (محدودیت آموزشی).':'محافظت ساختار کارپوشه غیرفعال شد.');return;}
 if(a==='unprotect'){sheetProtected=false;workbookProtected=false;render();tell('Unprotect','محافظت کاربرگ و کارپوشه غیرفعال شد.');return;}

 if(a==='function'){formulaWizard('all');return;}
 if(a==='math'){formulaWizard('math');return;}
 if(a==='stat'){formulaWizard('stat');return;}
 if(a==='logical'){formulaWizard('logical');return;}
 if(a==='textfn'){formulaWizard('textfn');return;}
 if(a==='datefn'){tell('Date & Time','در این نسخه توابع TODAY و NOW هنوز محاسبه نمی‌شوند. برای تاریخ ثابت، تاریخ را مستقیماً در سلول وارد کن.');return;}
 if(a==='recent'){if(!recentFormulas.length){tell('Recently Used','هنوز تابعی استفاده نشده است.');return;}let n=prompt('توابع اخیر:\\n'+recentFormulas.map((x,i)=>(i+1)+'. '+x).join('\\n'),'1');if(n&&recentFormulas[Number(n)-1]){const fn=recentFormulas[Number(n)-1];const examples={SUM:'SUM(A1:A5)',AVERAGE:'AVERAGE(A1:A5)',IF:'IF(A1>=10,"قبول","مردود")'};formula.value='='+(examples[fn]||fn+'(A1)');apply();}return;}
 if(a==='precedents'){traceFormula('precedents');return;}
 if(a==='dependents'){traceFormula('dependents');return;}
 if(a==='checkformula'){const f=String(raw(selected)||'');if(!f.startsWith('=')){tell('Check Formula','سلول انتخاب‌شده فرمول ندارد.');return;}const result=calc(selected);tell('Check Formula','فرمول: '+f+'\\nنتیجه فعلی: '+result+'\\nوضعیت: '+(String(result).startsWith('#')?'نیازمند بررسی':'محاسبه بدون خطای آشکار'));return;}
 if(a==='showformulas'){showFormulaMode=!showFormulaMode;render();select(cellName.value||selected);tell('Show Formulas',showFormulaMode?'نمایش فرمول‌ها فعال شد.':'نمایش نتیجه فرمول‌ها فعال شد.');return;}
 if(a==='definename'){const name=prompt('نام برای سلول یا محدوده انتخاب‌شده:','MyRange');if(name&&/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)){definedNames[name]=cellName.value; tell('Define Name',name+' به '+cellName.value+' اختصاص یافت.');}else if(name)tell('Define Name','نام باید با حرف یا زیرخط شروع شود و فاصله نداشته باشد.');return;}
 if(a==='namemanager'){const list=Object.entries(definedNames);tell('Name Manager',list.length?list.map(([n,r])=>n+' = '+r).join('\\n'):'هنوز نامی تعریف نشده است.');return;}

 if(a==='sheetdirection'){
  const choice=prompt('جهت نمایش کاربرگ را انتخاب کن:\\n1 = راست به چپ (مناسب فارسی)\\n2 = چپ به راست (مناسب انگلیسی)',sheetDirection==='rtl'?'1':'2');
  if(choice==='1'||choice==='2'){
   sheetDirection=choice==='1'?'rtl':'ltr';
   applySheetDirection();
   pageMessage(sheetDirection==='rtl'?'جهت کاربرگ راست به چپ شد؛ ستون A در سمت راست قرار می‌گیرد.':'جهت کاربرگ چپ به راست شد؛ ستون A در سمت چپ قرار می‌گیرد.');
   try{localStorage.setItem('miniExcelSheetDirection',sheetDirection);}catch(e){}
  }
  return;
 }

 if(a==='margins'){let v=prompt('حاشیه صفحه:\\n1 = Normal\\n2 = Narrow\\n3 = Wide',pageSettings.margin==='normal'?'1':pageSettings.margin==='narrow'?'2':'3');if(v&&['1','2','3'].includes(v)){pageSettings.margin=({'1':'normal','2':'narrow','3':'wide'})[v];pageMessage('حاشیه صفحه روی '+pageSettings.margin+' تنظیم شد.');}return;}
 if(a==='orientation'){let v=prompt('جهت صفحه:\\n1 = Portrait (عمودی)\\n2 = Landscape (افقی)',pageSettings.orientation==='portrait'?'1':'2');if(v==='1'||v==='2'){pageSettings.orientation=v==='1'?'portrait':'landscape';pageMessage('جهت صفحه تغییر کرد. برای مشاهده نتیجه از Print Preview استفاده کن.');}return;}
 if(a==='papersize'){let v=prompt('اندازه کاغذ:\\n1 = A4\\n2 = A3\\n3 = Letter\\n4 = Legal',({'A4':'1','A3':'2','Letter':'3','Legal':'4'})[pageSettings.paper]||'1');pageSettings.paper=({'1':'A4','2':'A3','3':'Letter','4':'Legal'})[v]||pageSettings.paper;pageMessage('اندازه کاغذ: '+pageSettings.paper);return;}
 if(a==='width'||a==='height'){let v=prompt((a==='width'?'Fit to width':'Fit to height')+' (0 = Automatic, 1 = one page, 2 = two pages, ...)',a==='width'?pageSettings.fitWidth:pageSettings.fitHeight);if(v!==null&&/^\\d+$/.test(v)){pageSettings[a==='width'?'fitWidth':'fitHeight']=v;pageMessage('مقیاس‌بندی چاپ تنظیم شد. توجه: مرورگر ممکن است بعضی گزینه‌های مقیاس را مستقل از تنظیمات چاپ خود کنترل کند.');}return;}
 if(a==='scale'){let v=Number(prompt('مقیاس چاپ (10 تا 400 درصد):',String(pageSettings.scale)));if(v>=10&&v<=400){pageSettings.scale=v;pageMessage('مقیاس چاپ روی '+v+'٪ تنظیم شد.');}return;}
 if(a==='printarea'){let refs=selectedRefs.slice();if(!refs.length){pageMessage('ابتدا محدوده موردنظر را انتخاب کن.');return;}let ps=refs.map(refParts),r1=Math.min(...ps.map(x=>x.r)),r2=Math.max(...ps.map(x=>x.r)),c1=Math.min(...ps.map(x=>x.c)),c2=Math.max(...ps.map(x=>x.c));pageSettings.printArea=colName(c1)+r1+':'+colName(c2)+r2;pageMessage('محدوده چاپ تعیین شد: '+pageSettings.printArea+'. برای چاپ از Print Preview استفاده کن.');return;}
 if(a==='pagebreaks'){let v=prompt('مدیریت شکست صفحه:\\n1 = پاک‌کردن تنظیمات دستی\\n2 = نمایش راهنمای شکست صفحه','1');if(v==='1')pageMessage('شکست‌های دستی پاک شدند (در این نسخه شکست صفحه به‌صورت خودکار توسط مرورگر تعیین می‌شود).');else if(v==='2')pageMessage('برای کنترل شکست صفحه، از Print Preview و تنظیمات چاپ مرورگر استفاده کن.');return;}
 if(a==='printtitles'){let v=prompt('ردیف عنوانی که باید در چاپ تکرار شود (مثلاً 1):','1');if(v!==null&&/^\\d+$/.test(v)){pageSettings.printTitles=v;pageMessage('ردیف '+v+' به‌عنوان ردیف عنوان چاپ ثبت شد. تکرار خودکار در چندصفحه‌ای در نسخه بعدی تکمیل می‌شود.');}return;}
 if(a==='themes'||a==='themecolors'||a==='themefonts'){pageMessage('تم‌های آماده در حال توسعه هستند. برای تغییر ظاهر سلول‌ها از ابزارهای Home استفاده کن.');return;}
 if(a==='bringfront'||a==='sendback'||a==='selectionpane'){pageMessage('Arrange برای اشیای شناور درج‌شده در شیت در نسخه بعدی توسعه داده می‌شود.');return;}
 if(a==='gridlines'){let visible=grid.dataset.gridlines!=='off';grid.dataset.gridlines=visible?'off':'on';grid.querySelectorAll('td,th').forEach(x=>x.style.borderColor=visible?'transparent':'');pageMessage(visible?'خطوط شبکه مخفی شدند.':'خطوط شبکه نمایش داده شدند.');return;}
 if(a==='headings'){let visible=grid.dataset.headings!=='off';grid.dataset.headings=visible?'off':'on';grid.querySelectorAll('th.corner,th.rowhead').forEach(x=>x.style.display=visible?'none':'');pageMessage(visible?'سرستون‌ها و شماره ردیف‌ها مخفی شدند.':'سرستون‌ها و شماره ردیف‌ها نمایش داده شدند.');return;}
 if(a==='printpreview'||a==='print') {printPreview();return;}

 if(['bold','italic','underline'].includes(a)){let val=!o[a];selectedRefs.forEach(ref=>cells[ref]={...(cells[ref]||{}),[a]:val});render();select(cellName.value)}
 else if(['left','center','right'].includes(a))fmt('align',a);
 else if(a==='middle')fmt('valign','middle');
 else if(a==='top')fmt('valign','top');
 else if(a==='bottom')fmt('valign','bottom');
 else if(a==='growfont'||a==='shrinkfont'){let size=Number(cells[selected]?.fontSize||11)+(a==='growfont'?1:-1);fmt('fontSize',Math.max(8,Math.min(72,size)));}
 else if(a==='currency')fmt('format','currency');
 else if(a==='thousands')fmt('format','number');
 else if(a==='formatpaint'){formatPainter={...(cells[selected]||{})};tell('Format Painter','قالب سلول انتخاب‌شده کپی شد. اکنون روی سلول یا محدوده مقصد کلیک کن تا همان قالب اعمال شود.');}
 else if(a==='fillmenu'){$('fillColor')?.click();}
 else if(a==='fontmenu'){$('fontColor')?.click();}
 else if(a==='orientation'){let choice=prompt('جهت متن: 1 = افقی، 2 = عمودی','2');fmt('orientation',choice==='2'?'vertical':'horizontal');}
 else if(a==='indent'||a==='indentmore'){selectedRefs.forEach(ref=>{let old=cells[ref]?.indent||0;cells[ref]={...(cells[ref]||{}),indent:Math.max(0,Math.min(5,old+(a==='indentmore'?1:-1)))}});render();select(cellName.value);}
 else if(a==='merge'){
  let ps=selectedRefs.map(refParts),minR=Math.min(...ps.map(x=>x.r)),maxR=Math.max(...ps.map(x=>x.r)),minC=Math.min(...ps.map(x=>x.c)),maxC=Math.max(...ps.map(x=>x.c));
  if(selectedRefs.length===1&&cells[selected]?.merge){let m=cells[selected].merge;for(let r=minR;r<minR+m.rows;r++)for(let c=minC;c<minC+m.cols;c++){let ref=colName(c)+r;if(ref!==selected&&cells[ref])delete cells[ref].mergedInto;}delete cells[selected].merge;render();select(selected);return;}
  if(selectedRefs.length<2){tell('Merge & Center','ابتدا محدوده‌ای مانند A1:C1 را انتخاب کن.');return;}
  let anchor=colName(minC)+minR;cells[anchor]={...(cells[anchor]||{}),merge:{rows:maxR-minR+1,cols:maxC-minC+1},align:'center',valign:'middle'};
  for(let r=minR;r<=maxR;r++)for(let c=minC;c<=maxC;c++){let ref=colName(c)+r;if(ref!==anchor)cells[ref]={...(cells[ref]||{}),mergedInto:anchor};}
  render();select(anchor);tell('Merge & Center','سلول‌ها ادغام شدند و محتوای سلول بالا-چپ حفظ شد. برای لغو ادغام، سلول ادغام‌شده را انتخاب و Merge & Center را دوباره بزن.');
 }
 else if(a==='conditional'){
  let threshold=Number(prompt('مقدار آستانه؛ سلول‌های بزرگ‌تر از این مقدار رنگی می‌شوند:','50'));if(!Number.isFinite(threshold))return;
  let count=0;selectedRefs.forEach(ref=>{let v=calc(ref);if(typeof v==='number'&&v>threshold){cells[ref]={...(cells[ref]||{}),bg:'#c6efce',color:'#006100'};count++;}});render();select(cellName.value);tell('Conditional Formatting',`${count} سلول بزرگ‌تر از ${threshold} با رنگ سبز مشخص شد.`);
 }
 else if(a==='wrap')fmt('wrap',!o.wrap);
 else if(a==='percent')fmt('format','percent');
 else if(a==='autosum'){let col=colName(p.c),from=Math.max(1,p.r-4),choice=prompt('تابع موردنظر: 1=SUM، 2=AVERAGE، 3=COUNT، 4=MAX، 5=MIN','1'),fn=({'1':'SUM','2':'AVERAGE','3':'COUNT','4':'MAX','5':'MIN'})[choice||'1']||'SUM';formula.value=`=${fn}(${col}${from}:${col}${p.r-1})`;apply()}
 else if(['function','math','stat','logical','textfn','recent'].includes(a))tell('Function Library','Available functions: =SUM(A1:A5), =AVERAGE(A1:A5), =MAX(A1:A5), =MIN(A1:A5), =COUNT(A1:A5). Select a cell and enter a formula in the formula bar.');
 else if(a==='clear'){
  let choice=prompt('Clear: 1 = Contents, 2 = Formats, 3 = All','1');
  if(choice==='1'||choice==='3')selectedRefs.forEach(ref=>cells[ref]={...(choice==='3'?{}:cells[ref]),value:''});
  if(choice==='2'||choice==='3')selectedRefs.forEach(ref=>{let v=cells[ref]?.value;cells[ref]=choice==='2'?{value:v}:{}});
  render();select(cellName.value);
 }
 else if(a==='copy'||a==='cut'){
  let matrix=selectedRefs.map(ref=>raw(ref)),p0=refParts(selectedRefs[0]),p1=refParts(selectedRefs[selectedRefs.length-1]),rows=[];
  for(let r=p0.r;r<=p1.r;r++){let row=[];for(let c=p0.c;c<=p1.c;c++)row.push(raw(colName(c)+r));rows.push(row.join('\t'))}
  clipboard=rows.join('\n');clipboardCut=a==='cut';window._miniClip=clipboard;
  if(a==='cut'){selectedRefs.forEach(ref=>cells[ref]={...(cells[ref]||{}),value:''});render();select(cellName.value)}
  tell(a==='copy'?'Copied':'Cut',`${selectedRefs.length} cell(s) ${a==='copy'?'copied':'cut'}. Choose a destination cell and select Paste.`);
 }
 else if(a==='paste'){
  let text=clipboard||window._miniClip||'';
  if(!text){tell('Paste','First copy or cut cells inside this workbook.');return}
  let rows=text.split(/\r?\n/);rows.forEach((line,ri)=>line.split('\t').forEach((v,ci)=>{let rr=p.r+ri,cc=p.c+ci;if(rr<=ROWS&&cc<=COLS){let ref=colName(cc)+rr;cells[ref]={...(cells[ref]||{}),value:v}}}));
  clipboardCut=false;render();select(selected);
 }
 else if(a==='find'){
  let mode=prompt('1 = Find text, 2 = Go To cell','1');if(mode==='2'){let target=prompt('Cell address (e.g. D12):',selected)?.toUpperCase();if(target&&rangeRefs(target).length){select(target);grid.querySelector(`[data-ref="${selected}"]`)?.scrollIntoView({block:'center',inline:'center'});}else tell('Go To','آدرس سلول معتبر نیست.');return;}
  let q=prompt('Find what?');if(q){let found=Object.keys(cells).find(k=>String(raw(k)).toLowerCase().includes(q.toLowerCase()));if(found){select(found);grid.querySelector(`[data-ref="${found}"]`)?.scrollIntoView({block:'center',inline:'center'})}else tell('Find','No match found.')}
 }
 else if(a==='sortaz'||a==='sortza'){
  let refs=Array.from({length:ROWS},(_,i)=>colName(p.c)+(i+1)),vals=refs.map(r=>raw(r)).filter(v=>v!=='').sort((x,y)=>String(x).localeCompare(String(y),undefined,{numeric:true}));
  if(a==='sortza')vals.reverse();refs.forEach((r,i)=>{if(i<vals.length)cells[r]={...(cells[r]||{}),value:vals[i]};else if(cells[r])cells[r]={...cells[r],value:''}});render();select(selected)
 }
 else if(a==='filter')tell('Filter','برای مرتب‌سازی داده‌ها از گزینه‌های Sort A to Z و Sort Z to A استفاده کنید. فیلترهای پیشرفته در مرحله بعد اضافه می‌شوند.');
 else if(a==='clearfilter')tell('Filter','No active filter.');
 else if(a==='colwidth'){let choice=prompt('Format:\n1 = Column Width\n2 = AutoFit Column Width\n3 = Number Format\n4 = Font Size','1');if(choice==='3'){let f=prompt('Number format: general / number / currency / percent','number');if(['general','number','currency','percent'].includes(f))fmt('format',f);return;}if(choice==='4'){let z=Number(prompt('Font size (8–72):','12'));if(z>=8&&z<=72)fmt('fontSize',z);return;}let cols=[...new Set(selectedRefs.map(x=>refParts(x).c))];if(choice==='2'){cols.forEach(c=>{let max=colName(c).length*9+22;for(let r=1;r<=ROWS;r++){let v=String(display(colName(c)+r)??'');max=Math.max(max,Math.min(300,v.length*8+18));}colWidths[c]=Math.max(55,Math.min(300,max));});}else{let n=Number(prompt('Column width in pixels (45–500):',String(colWidths[p.c]||105)));if(n>=45&&n<=500)cols.forEach(c=>colWidths[c]=n);else return;}render();select(cellName.value||selected)}
 else if(a==='insert'||a==='delete'){
  let kind=prompt('Type R for row or C for column:','R')?.toUpperCase();if(!['R','C'].includes(kind))return;
  let idx=kind==='R'?p.r:p.c;
  if(a==='insert'){
   if(kind==='R'){for(let r=ROWS;r>idx;r--)for(let c=1;c<=COLS;c++){let from=colName(c)+(r-1),to=colName(c)+r;if(cells[from])cells[to]={...cells[from]};else delete cells[to]}for(let c=1;c<=COLS;c++)delete cells[colName(c)+idx]}
   else{for(let c=COLS;c>idx;c--)for(let r=1;r<=ROWS;r++){let from=colName(c-1)+r,to=colName(c)+r;if(cells[from])cells[to]={...cells[from]};else delete cells[to]}for(let r=1;r<=ROWS;r++)delete cells[colName(idx)+r]}
  }else{
   if(kind==='R'){for(let r=idx;r<ROWS;r++)for(let c=1;c<=COLS;c++){let from=colName(c)+(r+1),to=colName(c)+r;if(cells[from])cells[to]={...cells[from]};else delete cells[to]}for(let c=1;c<=COLS;c++)delete cells[colName(c)+ROWS]}
   else{for(let c=idx;c<COLS;c++)for(let r=1;r<=ROWS;r++){let from=colName(c+1)+r,to=colName(c)+r;if(cells[from])cells[to]={...cells[from]};else delete cells[to]}for(let r=1;r<=ROWS;r++)delete cells[colName(COLS)+r]}
  }render();select(kind==='R'?`A${Math.min(idx,ROWS)}`:`${colName(Math.min(idx,COLS))}1`);
 }
 else if(a==='chart'||a==='linechart'||a==='piechart'||a==='barchart')chartFromSelection(a==='barchart'?'chart':a);
 else if(a==='shapes'){
  let choice=prompt('Choose a shape:\n1 = Rectangle\n2 = Rounded rectangle\n3 = Circle\n4 = Arrow\n5 = Line','1');
  if(!['1','2','3','4','5'].includes(choice||''))return;
  let label=prompt('Shape text (optional):','');if(label===null)return;
  let svg='';
  if(choice==='1')svg=`<svg viewBox="0 0 320 150"><rect x="35" y="25" width="250" height="100" fill="#d9ead3" stroke="#217346" stroke-width="3"/><text x="160" y="82" text-anchor="middle" font-size="16" fill="#202a24">${esc(label)}</text></svg>`;
  if(choice==='2')svg=`<svg viewBox="0 0 320 150"><rect x="35" y="25" width="250" height="100" rx="18" fill="#d9eaf7" stroke="#4472c4" stroke-width="3"/><text x="160" y="82" text-anchor="middle" font-size="16" fill="#202a24">${esc(label)}</text></svg>`;
  if(choice==='3')svg=`<svg viewBox="0 0 320 150"><circle cx="160" cy="75" r="55" fill="#fff2cc" stroke="#bf9000" stroke-width="3"/><text x="160" y="81" text-anchor="middle" font-size="15" fill="#202a24">${esc(label)}</text></svg>`;
  if(choice==='4')svg=`<svg viewBox="0 0 320 150"><path d="M35 55 H230 V25 L290 75 L230 125 V95 H35 Z" fill="#d9eaf7" stroke="#4472c4" stroke-width="3"/><text x="145" y="82" text-anchor="middle" font-size="14" fill="#202a24">${esc(label)}</text></svg>`;
  if(choice==='5')svg='<svg viewBox="0 0 320 150"><line x1="35" y1="115" x2="285" y2="35" stroke="#217346" stroke-width="5"/></svg>';
  let box=document.createElement('div');box.innerHTML=svg;objectCard('شکل',box.firstElementChild);
 }
 else if(a==='pivot'){
  let ps=selectedRefs.map(refParts);if(!ps.length)return;
  let minR=Math.min(...ps.map(x=>x.r)),maxR=Math.max(...ps.map(x=>x.r)),minC=Math.min(...ps.map(x=>x.c)),maxC=Math.max(...ps.map(x=>x.c));
  if(maxR<=minR||maxC<=minC){tell('PivotTable','برای ساخت گزارش، محدوده‌ای شامل ردیف عنوان‌ها و حداقل یک ستون دسته‌بندی و یک ستون عددی انتخاب کن.');return;}
  let categoryCol=minC,valueCol=minC+1,totals={};for(let r=minR+1;r<=maxR;r++){let k=String(raw(colName(categoryCol)+r)||'(خالی)'),v=Number(raw(colName(valueCol)+r));if(Number.isFinite(v))totals[k]=(totals[k]||0)+v;}
  let rows=Object.entries(totals);if(!rows.length){tell('PivotTable','در ستون دوم داده عددی پیدا نشد.');return;}
  let tbl='<table style="border-collapse:collapse;min-width:260px"><tr><th style="background:#e8f1eb;padding:7px;border:1px solid #b8c8bc">'+esc(String(raw(colName(categoryCol)+minR)||'دسته'))+'</th><th style="background:#e8f1eb;padding:7px;border:1px solid #b8c8bc">جمع '+esc(String(raw(colName(valueCol)+minR)||'مقدار'))+'</th></tr>'+rows.map(([k,v])=>'<tr><td style="padding:7px;border:1px solid #d7ded9">'+esc(k)+'</td><td style="padding:7px;border:1px solid #d7ded9;text-align:right">'+v+'</td></tr>').join('')+'<tr><th style="padding:7px;border:1px solid #d7ded9">جمع کل</th><th style="padding:7px;border:1px solid #d7ded9">'+rows.reduce((a,x)=>a+x[1],0)+'</th></tr></table>';
  objectCard('گزارش خلاصه (PivotTable ساده)',tbl);tell('PivotTable','یک گزارش خلاصه بر اساس دسته‌بندی ستون اول و جمع مقادیر ستون دوم ساخته شد. این نسخه آموزشی از PivotTable پیشرفته اکسل ساده‌تر است.');
 }
 else if(a==='table'){let refs=selectedRefs.slice(),ps=refs.map(refParts),minR=Math.min(...ps.map(x=>x.r)),maxR=Math.max(...ps.map(x=>x.r)),minC=Math.min(...ps.map(x=>x.c)),maxC=Math.max(...ps.map(x=>x.c));if(maxR===minR){tell('Format as Table','برای ساخت جدول، محدوده‌ای با حداقل دو ردیف انتخاب کن.');return}let palette=['#d9ead3','#d9eaf7','#fce4d6'];let head=prompt('رنگ جدول را انتخاب کن: 1 سبز، 2 آبی، 3 نارنجی','1');let bg=palette[(Math.max(1,Math.min(3,Number(head)||1)))-1];for(let r=minR;r<=maxR;r++)for(let c=minC;c<=maxC;c++){let ref=colName(c)+r;cells[ref]={...(cells[ref]||{}),border:'1px solid #aab7ae',bg:r===minR?bg:(r%2===0?'#f7faf8':'#ffffff'),bold:r===minR,align:'center'}}render();select(cellName.value);tell('جدول ایجاد شد','ردیف اول به عنوان سرستون قالب‌بندی شد. برای نتیجه بهتر، پیش از درج جدول محدوده داده‌ها را همراه با عنوان ستون‌ها انتخاب کن.')}
 else if(a==='picture')insertPicture();
 else if(a==='link'){let url=prompt('آدرس پیوند را وارد کن:','https://');if(url&&/^https?:\/\//i.test(url)){let label=prompt('متن نمایشی پیوند:','مشاهده وب‌سایت')||url;let el=document.createElement('p');let link=document.createElement('a');link.href=url;link.target='_blank';link.rel='noopener';link.textContent=label;el.appendChild(link);objectCard('پیوند',el)}else if(url)tell('پیوند','آدرس باید با http:// یا https:// شروع شود.');}
 else if(a==='wordart'){let t=prompt('متن WordArt را وارد کن:','عنوان من');if(t){let el=document.createElement('div');el.textContent=t;el.style.cssText='font-size:30px;font-weight:800;color:#217346;text-shadow:2px 2px 0 #d9eaf7;padding:18px;text-align:center';objectCard('WordArt',el)}}
 else if(a==='headerfooter'){let hd=prompt('متن سربرگ:','گزارش کلاسی');if(hd!==null){let ft=prompt('متن پابرگ:','نام هنرجو');const el=document.createElement('div');el.innerHTML='<div style="border-bottom:1px solid #aaa;padding:8px;text-align:center">'+esc(hd)+'</div><div style="border-top:1px solid #aaa;padding:8px;text-align:center;margin-top:20px">'+esc(ft||'')+'</div>';objectCard('پیش‌نمایش سربرگ و پابرگ',el);}}
 else if(a==='equation'){let t=prompt('عبارت ریاضی را وارد کن:','a² + b² = c²');if(t){let el=document.createElement('div');el.textContent=t;el.style.cssText='font-size:24px;text-align:center;padding:15px;font-family:serif';objectCard('معادله',el)}}
 else if(a==='symbol'){let t=prompt('نماد موردنظر را انتخاب یا وارد کن:','©  ®  ™  €  £  ¥  Ω  π  ±  ÷  ×');if(t){let el=document.createElement('div');el.textContent=t;el.style.cssText='font-size:28px;padding:15px;text-align:center';objectCard('نماد',el)}}
 else if(a==='recommendedpivot'){act('pivot');}
 else if(a==='recommendedcharts'){let k=prompt('نوع نمودار پیشنهادی را انتخاب کن:\\n1 ستونی، 2 خطی، 3 دایره‌ای','1');if(k==='1')chartFromSelection('chart');else if(k==='2')chartFromSelection('linechart');else if(k==='3')chartFromSelection('piechart');}
 else if(a==='barchart')chartFromSelection('chart');
 else if(a==='scatterchart')insertScatter();
 else if(['sparkline','sparkcolumn','sparkwinloss'].includes(a))insertSparkline(a);
 else if(a==='icons')insertIcons();
 else if(a==='smartart')insertSmartArt();
 else if(a==='models')tell('3D Models','نمایش و ویرایش مدل سه‌بعدی در این نسخه سبک HTML پشتیبانی نمی‌شود؛ می‌توانی تصویر مدل را از Pictures وارد کنی.');
 else if(a==='screenshot')tell('Screenshot','مرورگر اجازه گرفتن تصویر از کل صفحه را بدون مجوز ویژه نمی‌دهد. برای تمرین، تصویر صفحه را با ابزار Snipping Tool ویندوز بگیر و از Pictures وارد کن.');
 else if(a==='slicer'||a==='timeline'){let f=prompt('نام فیلد برای فیلتر آموزشی را وارد کن:','دسته‌بندی');if(f){const el=document.createElement('div');el.innerHTML='<label>'+esc(f)+' : <select><option>همه</option><option>گروه اول</option><option>گروه دوم</option></select></label>';objectCard(a==='slicer'?'فیلتر برشی آموزشی':'فیلتر زمانی آموزشی',el);}}
 else if(a==='signature'){let t=prompt('نام امضاکننده:','نام و نام خانوادگی');if(t){const el=document.createElement('div');el.innerHTML='<div style="width:260px;border-bottom:1px solid #333;padding:25px 5px 5px;text-align:center">'+esc(t)+'</div><small>امضا</small>';objectCard('خط امضا',el);}}
 else if(a==='object')insertFileObject('*/*','شیء پیوست‌شده');
 else if(a==='textbox'){let text=prompt('متن کادر را وارد کن:','یادداشت من');if(text!==null&&text.trim()){let box=document.createElement('div');box.contentEditable='true';box.textContent=text;box.title='برای ویرایش متن کلیک کن';box.style.cssText='min-width:220px;min-height:70px;padding:14px;border:1px dashed #217346;background:#fff;border-radius:5px;outline:none';objectCard('کادر متنی – برای ویرایش کلیک کن',box)}}
 else if(a==='function'){let f=prompt('تابع را وارد کن. مثال: =SUM(A1:A5)','=SUM(A1:A5)');if(f!==null&&f.trim()){formula.value=f.startsWith('=')?f:'='+f;apply()}}
 else if(a==='zoomin'){zoom=Math.min(160,zoom+10);$('gridwrap').style.zoom=zoom+'%'}
 else if(a==='zoomout'){zoom=Math.max(60,zoom-10);$('gridwrap').style.zoom=zoom+'%'}
 else if(a==='zoom100'){zoom=100;$('gridwrap').style.zoom='100%'}
 else if(a==='gridlines'){grid.querySelectorAll('td,th').forEach(x=>x.style.borderColor=x.style.borderColor==='transparent'?'':'transparent')}
 else if(a==='border'){
  let choice=prompt('Border style: 1 = All borders, 2 = Outside border, 3 = Remove borders','1');
  let val=choice==='3'?'':choice==='2'?'2px solid #52665a':'1px solid #87968b';
  selectedRefs.forEach(ref=>cells[ref]={...(cells[ref]||{}),border:val});render();select(cellName.value);
 }
 else if(a==='styles'){
  let choice=prompt('Cell style: 1 = Heading, 2 = Total, 3 = Highlight, 4 = Normal','1');
  let preset=choice==='1'?{bold:true,bg:'#d9ead3',align:'center',border:'1px solid #87968b'}:choice==='2'?{bold:true,bg:'#d9eaf7',border:'2px solid #52665a'}:choice==='3'?{bg:'#fff2cc',bold:true}: {};
  selectedRefs.forEach(ref=>cells[ref]={...(cells[ref]||{}),...preset});render();select(cellName.value);
 }
 else if(a==='incdec'||a==='decdec'){
  selectedRefs.forEach(ref=>{let old=cells[ref]?.decimals??0;cells[ref]={...(cells[ref]||{}),decimals:Math.max(0,Math.min(6,old+(a==='incdec'?1:-1)))}});render();select(cellName.value);
 }
 else if(a==='paste')tell('Paste','Choose a cell and press Ctrl+V to paste text from your device.');
 else if(['shapes','pivot','pagelayout','pagebreak','normal','formulabar','duplicates','textcolumns','validation','group','ungroup','spelling','comment','protect','protectbook','precedents','checkformula'].includes(a))tell('Mini Excel','This tool will be implemented in a later development stage.');
 else tell('Mini Excel','This tool is part of the classroom interface.');
}
document.querySelectorAll('.tab').forEach(t=>t.onclick=()=>{document.querySelector('.tab.active')?.classList.remove('active');t.classList.add('active');activeTab=t.dataset.tab;renderRibbon()});
function storeActiveSheet(){if(!sheets[activeSheetIndex])return;sheets[activeSheetIndex].cells=cells;sheets[activeSheetIndex].colWidths=colWidths;}
function renderSheetTabs(){const bar=document.querySelector('.sheetbar');bar.querySelectorAll('.sheet-tab').forEach(x=>x.remove());sheets.forEach((sh,i)=>{const b=document.createElement('button');b.className='sheet'+(i===activeSheetIndex?' active-sheet':'')+' sheet-tab';b.textContent=sh.name;b.title='دوبار کلیک برای تغییر نام؛ کلیک راست برای حذف';b.onclick=()=>switchSheet(i);b.ondblclick=()=>{const n=prompt('نام برگه جدید:',sh.name);if(n&&n.trim()){sh.name=n.trim().slice(0,31);renderSheetTabs();}};b.oncontextmenu=e=>{e.preventDefault();if(sheets.length===1){tell('Worksheets','حداقل یک برگه باید در فایل باقی بماند.');return;}if(confirm('برگه «'+sh.name+'» حذف شود؟')){sheets.splice(i,1);if(activeSheetIndex>=sheets.length)activeSheetIndex=sheets.length-1;loadActiveSheet();}};bar.insertBefore(b,$('addSheet'));});}
function loadActiveSheet(){const sh=sheets[activeSheetIndex];cells=sh.cells||{};colWidths=sh.colWidths||{};selected='A1';selectedRefs=['A1'];render();select('A1');renderSheetTabs();}
function switchSheet(i){if(i===activeSheetIndex)return;storeActiveSheet();activeSheetIndex=i;loadActiveSheet();}
function addNewSheet(){storeActiveSheet();let base='Sheet'+(sheets.length+1),name=base,n=2;while(sheets.some(x=>x.name.toLowerCase()===name.toLowerCase()))name=base+' ('+(n++)+')';sheets.push({name,cells:{},colWidths:{}});activeSheetIndex=sheets.length-1;loadActiveSheet();}
$('saveBtn').onclick=()=>{storeActiveSheet();localStorage.setItem('miniExcelWorkbook',JSON.stringify({sheets,activeSheetIndex,name:$('bookName').value,cells,colWidths}));tell('Saved','Workbook and all worksheets saved in this browser on this device.')};
$('newBtn').onclick=()=>{if(confirm('Create a new workbook? Unsaved changes may be lost.')){cells={};colWidths={};sheets=[{name:'Sheet1',cells:{},colWidths:{}}];activeSheetIndex=0;$('bookName').value='Book1';loadActiveSheet();}};
$('openBtn').onclick=()=>$('fileInput').click();
$('fileInput').onchange=async e=>{let f=e.target.files[0];if(!f)return;let text=await f.text(),lines=text.replace(/^\uFEFF/,'').split(/\r?\n/);cells={};lines.slice(0,ROWS).forEach((line,r)=>{let vals=line.match(/("([^"]|"")*"|[^,]*)(,|$)/g)||[];vals.forEach((v,c)=>{let s=v.replace(/,$/,'');if(s.startsWith('"')&&s.endsWith('"'))s=s.slice(1,-1).replace(/""/g,'"');if(s!=='')cells[colName(c+1)+(r+1)]={value:s}})});$('bookName').value=f.name.replace(/\.csv$/i,'');render();select('A1');e.target.value=''};
$('downloadBtn').onclick=()=>{let lines=[];for(let r=1;r<=ROWS;r++){let vals=[];for(let c=1;c<=COLS;c++)vals.push('"'+String(calc(colName(c)+r)).replace(/"/g,'""')+'"');lines.push(vals.join(','))}let blob=new Blob(['\ufeff'+lines.join('\r\n')],{type:'text/csv;charset=utf-8'}),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=($('bookName').value||'Book1')+'.csv';a.click();URL.revokeObjectURL(a.href)};
$('addSheet').onclick=addNewSheet;
grid.addEventListener('paste',e=>{if(sheetProtected){e.preventDefault();tell('Protected Sheet','ابتدا محافظت کاربرگ را غیرفعال کن.');return;}let text=e.clipboardData?.getData('text');if(!text)return;e.preventDefault();let rows=text.replace(/\r/g,'').split('\n').filter((line,i,arr)=>line!==''||i<arr.length-1);let p0=refParts(selected);rows.forEach((line,ri)=>line.split('\t').forEach((v,ci)=>{let rr=p0.r+ri,cc=p0.c+ci;if(rr<=ROWS&&cc<=COLS){let ref=colName(cc)+rr;cells[ref]={...(cells[ref]||{}),value:v};}}));render();select(selected);});
document.addEventListener('keydown',e=>{
 if((e.ctrlKey||e.metaKey)&&['c','x','v'].includes(e.key.toLowerCase())&&document.activeElement?.tagName!=='INPUT'){
  e.preventDefault();act(({c:'copy',x:'cut',v:'paste'})[e.key.toLowerCase()]);
 }
});
try{let saved=JSON.parse(localStorage.getItem('miniExcelWorkbook'));if(saved?.sheets?.length){sheets=saved.sheets;activeSheetIndex=Math.min(saved.activeSheetIndex||0,sheets.length-1);cells=sheets[activeSheetIndex].cells||{};colWidths=sheets[activeSheetIndex].colWidths||{};$('bookName').value=saved.name||'Book1'}else if(saved?.cells){cells=saved.cells;colWidths=saved.colWidths||{};sheets=[{name:'Sheet1',cells,colWidths}];$('bookName').value=saved.name||'Book1'}}catch{}
try{sheetDirection=localStorage.getItem('miniExcelSheetDirection')||'ltr';}catch(e){}
renderRibbon();render();select('A1');renderSheetTabs();
</script>
</body></html>

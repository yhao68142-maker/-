(()=>{
'use strict';
const enc=new TextEncoder(),dec=new TextDecoder(),NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const xml=s=>{const d=new DOMParser().parseFromString(s,'application/xml');if(d.querySelector('parsererror'))throw Error('Excel 内部 XML 无法读取');return d;};
const serialize=d=>new XMLSerializer().serializeToString(d);
const nodes=(d,n)=>Array.from(d.getElementsByTagNameNS('*',n));
const colNum=s=>[...s.replace(/\$/g,'').toUpperCase()].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0);
function colName(n){let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s}
function normalize(s){return String(s??'').normalize('NFKC').replace(/\s+/g,'').toUpperCase()}
function normalizeName(s){return normalize(s).replace(/GEN(\d+)\.0(?!\d)/g,'GEN$1').replace(/[-‐‑‒–—_]/g,'')}
function normalizeAliasName(s){return normalizeName(s).replace(/通用周转烘烤载条/g,'通用周转烘烤料盒').replace(/\(?返修\)?/g,'')}
const requesterHeaders=['需求人','申请人','申请人姓名','需求申请人','申请者','申请人员'];
const designerHeaders=['设计人','设计负责人','设计担当','设计者','设计人员','设计工程师'];
function findHeader(items,names){const normalized=new Set(names.map(normalize));return items.find(([,v])=>normalized.has(normalize(v)))}
function personCols(items){const r=findHeader(items,requesterHeaders),d=findHeader(items,designerHeaders);return{requesterCol:r?+r[0]:null,designerCol:d?+d[0]:null}}

async function unzip(buffer){const a=new Uint8Array(buffer),v=new DataView(a.buffer);let end=a.length-22;for(;end>=Math.max(0,a.length-65557);end--)if(v.getUint32(end,true)===0x06054b50)break;if(end<0)throw Error('不是有效的 .xlsx 文件，请确认文件未加密');const count=v.getUint16(end+10,true);let p=v.getUint32(end+16,true),total=0;const entries=new Map();for(let i=0;i<count;i++){if(v.getUint32(p,true)!==0x02014b50)throw Error('Excel 压缩目录损坏');const flags=v.getUint16(p+8,true),method=v.getUint16(p+10,true),size=v.getUint32(p+20,true),rawSize=v.getUint32(p+24,true),nl=v.getUint16(p+28,true),el=v.getUint16(p+30,true),cl=v.getUint16(p+32,true),local=v.getUint32(p+42,true),name=dec.decode(a.slice(p+46,p+46+nl));total+=rawSize;if(total>150*1024*1024||count>10000)throw Error('文件内容过大，请使用小于 20 MB 的普通审核表');if(flags&1)throw Error('不支持加密 Excel，请先另存为未加密 .xlsx');const start=local+30+v.getUint16(local+26,true)+v.getUint16(local+28,true);let data=a.slice(start,start+size);if(method===8){if(!globalThis.DecompressionStream)throw Error('请使用新版 Chrome 或 Edge 浏览器');data=new Uint8Array(await new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer())}else if(method!==0)throw Error('此 Excel 使用了不支持的压缩格式');entries.set(name,data);p+=46+nl+el+cl}return entries}
const crcTable=Array.from({length:256},(_,n)=>{for(let k=0;k<8;k++)n=n&1?0xedb88320^(n>>>1):n>>>1;return n>>>0});
function crc(a){let c=0xffffffff;for(const n of a)c=crcTable[(c^n)&255]^(c>>>8);return(c^0xffffffff)>>>0}
function zip(entries){let offset=0,centralSize=0;const pieces=[],central=[];for(const [name,data]of entries){const n=enc.encode(name),c=crc(data),h=new Uint8Array(30+n.length),v=new DataView(h.buffer);v.setUint32(0,0x04034b50,true);v.setUint16(4,20,true);v.setUint16(6,0x800,true);v.setUint32(14,c,true);v.setUint32(18,data.length,true);v.setUint32(22,data.length,true);v.setUint16(26,n.length,true);h.set(n,30);pieces.push(h,data);const ch=new Uint8Array(46+n.length),cv=new DataView(ch.buffer);cv.setUint32(0,0x02014b50,true);cv.setUint16(4,20,true);cv.setUint16(6,20,true);cv.setUint16(8,0x800,true);cv.setUint32(16,c,true);cv.setUint32(20,data.length,true);cv.setUint32(24,data.length,true);cv.setUint16(28,n.length,true);cv.setUint32(42,offset,true);ch.set(n,46);central.push(ch);centralSize+=ch.length;offset+=h.length+data.length}const e=new Uint8Array(22),v=new DataView(e.buffer);v.setUint32(0,0x06054b50,true);v.setUint16(8,entries.size,true);v.setUint16(10,entries.size,true);v.setUint32(12,centralSize,true);v.setUint32(16,offset,true);return new Blob([...pieces,...central,e],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})}
async function readWorkbook(file){if(file.size>20*1024*1024)throw Error('文件超过 20 MB，请拆分后再试');if(!/\.xlsx$/i.test(file.name))throw Error('请选择 .xlsx 文件，旧版 .xls 请先另存为 .xlsx');const entries=await unzip(await file.arrayBuffer());if(!entries.has('xl/workbook.xml'))throw Error('文件不是受支持的 Excel 工作簿');const wb=xml(dec.decode(entries.get('xl/workbook.xml'))),rels=xml(dec.decode(entries.get('xl/_rels/workbook.xml.rels'))),shared=entries.has('xl/sharedStrings.xml')?nodes(xml(dec.decode(entries.get('xl/sharedStrings.xml'))),'si').map(si=>nodes(si,'t').map(t=>t.textContent).join('')):[];const sheets=nodes(wb,'sheet').map(s=>{const id=s.getAttribute('r:id')||s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id'),rel=nodes(rels,'Relationship').find(r=>r.getAttribute('Id')===id);let path=rel?.getAttribute('Target');if(!path)throw Error('工作表路径无法识别');path=path.startsWith('/')?path.slice(1):'xl/'+path;const d=xml(dec.decode(entries.get(path))),rows=nodes(d,'row').map(r=>{const cells={};nodes(r,'c').forEach(c=>{const ref=c.getAttribute('r'),k=colNum(ref.match(/[A-Z]+/)[0]),t=c.getAttribute('t'),v=nodes(c,'v')[0]?.textContent??'';cells[k]=t==='s'?shared[Number(v)]??'':t==='inlineStr'?nodes(c,'t').map(t=>t.textContent).join(''):v});return{number:Number(r.getAttribute('r')),cells}});return{name:s.getAttribute('name'),path,doc:d,rows}});return{entries,wb,sheets,name:file.name}}

function detect(sheet,kind){
 for(const row of sheet.rows.slice(0,30)){
  const items=Object.entries(row.cells),people=personCols(items);
  if(kind==='source'){
   const nc=items.find(([,v])=>normalize(v)==='名称说明'),cc=items.find(([,v])=>['治工具编号','治工具编码'].includes(normalize(v)));
   if(nc&&cc)return{header:row.number,nameCol:+nc[0],codeCol:+cc[0],...people};
  }else{
   const n=items.find(([,v])=>normalize(v)==='治具名称+型号');
   if(n){const q=items.find(([,v])=>normalize(v)==='数量');return{header:row.number,nameCol:+n[0],quantityCol:q?+q[0]:null,...people}}
  }
 }
 if(kind==='source'&&sheet.rows[0]?.cells[42]&&sheet.rows[0]?.cells[24]){
  return{header:sheet.rows[0].number,nameCol:42,codeCol:24,...personCols(Object.entries(sheet.rows[0].cells)),fallback:true};
 }
 throw Error(kind==='source'?'未找到「治工具编号」及「名称说明」，请检查工作表。':'未找到「治具名称+型号」，请检查工作表。');
}
function records(sheet,cfg){return sheet.rows.filter(r=>r.number>cfg.header&&String(r.cells[cfg.nameCol]??'').trim()).map(r=>({row:r.number,name:String(r.cells[cfg.nameCol]),code:cfg.codeCol?String(r.cells[cfg.codeCol]??'').trim():'',quantity:parseQuantity(cfg.quantityCol?r.cells[cfg.quantityCol]:null),requester:cfg.requesterCol?String(r.cells[cfg.requesterCol]??'').trim():'',designer:cfg.designerCol?String(r.cells[cfg.designerCol]??'').trim():''}))}
function parseQuantity(value){const s=String(value??'').normalize('NFKC').trim();if(!/^\d+(?:\.0+)?$/.test(s))return null;const n=Number(s);return Number.isSafeInteger(n)&&n>0?n:null}
function codeCount(code){
 const text=String(code??'').normalize('NFKC').trim();
 const parts=text.split(/\s*[~～〜至]\s*/);
 if(parts.length===1)return /^[A-Za-z][A-Za-z0-9]*\d$/.test(text)?1:null;
 if(parts.length!==2)return null;
 const a=parts[0].match(/^([A-Za-z][A-Za-z0-9]*?)(\d+)$/),b=parts[1].match(/^([A-Za-z][A-Za-z0-9]*?)(\d+)$/);
 if(!a)return null;
 let end;
 if(b){if(a[1].toUpperCase()!==b[1].toUpperCase())return null;end=b[2]}
 else if(/^\d+$/.test(parts[1])&&parts[1].length<=a[2].length){end=a[2].slice(0,a[2].length-parts[1].length)+parts[1]}
 else return null;
 const count=BigInt(end)-BigInt(a[2])+1n;
 return count>0n&&count<=BigInt(Number.MAX_SAFE_INTEGER)?Number(count):null;
}
function mergePeople(a,b){const values=[...String(a??'').split('、'),...String(b??'').split('、')].map(s=>s.trim()).filter(Boolean);return[...new Set(values)].join('、')}
function nameMatchMode(targetName,sourceName){
 const targetKey=normalizeName(targetName),sourceKey=normalizeName(sourceName);
 if(targetKey===sourceKey)return'exact';
 if(targetKey.length>=8&&sourceKey.endsWith(targetKey))return'suffix';
 if(normalizeAliasName(targetName)===normalizeAliasName(sourceName))return'alias';
 return null;
}
function matchWeight(mode){return mode==='exact'?1003:mode==='suffix'?1002:mode==='alias'?1001:0}
function groupCandidates(hits){
 const groups=[];
 for(const hit of hits){const s=hit.source;let g=groups.find(g=>g.code===s.code&&g.name===s.name);
  if(g){g.rows.push(s.row);g.sourceIndices.push(hit.index);g.requester=mergePeople(g.requester,s.requester);g.designer=mergePeople(g.designer,s.designer)}
  else groups.push({...s,rows:[s.row],sourceIndices:[hit.index],count:codeCount(s.code)})
 }
 return groups;
}
function matchRows(target,source){
 const m=target.length,n=source.length,modes=Array.from({length:m},()=>Array(n).fill(null)),dp=Array.from({length:m+1},()=>Array(n+1).fill(0));
 for(let i=0;i<m;i++)for(let j=0;j<n;j++)modes[i][j]=nameMatchMode(target[i].name,source[j].name);
 for(let i=1;i<=m;i++)for(let j=1;j<=n;j++){
  let best=Math.max(dp[i-1][j],dp[i][j-1]),mode=modes[i-1][j-1];
  if(mode)best=Math.max(best,dp[i-1][j-1]+matchWeight(mode));
  dp[i][j]=best;
 }
 const ordered=new Map(),usedSource=new Set();let i=m,j=n;
 while(i>0&&j>0){
  const mode=modes[i-1][j-1],score=mode?dp[i-1][j-1]+matchWeight(mode):-1;
  if(mode&&dp[i][j]===score){ordered.set(i-1,{sourceIndex:j-1,mode});usedSource.add(j-1);i--;j--;continue}
  if(dp[i-1][j]>=dp[i][j-1])i--;else j--;
 }
 const results=target.map(t=>({...t,candidates:[],selected:null,reason:null,matchMode:null,status:'missing'}));
 for(const [ti,pair]of ordered){
  const s=source[pair.sourceIndex],candidate={...s,rows:[s.row],sourceIndices:[pair.sourceIndex],count:codeCount(s.code)},hasCode=!!s.code;
  results[ti]={...target[ti],candidates:hasCode?[candidate]:[],selected:hasCode?0:null,reason:hasCode?'order':null,matchMode:pair.mode,status:hasCode?'matched':'noCode'};
 }
 for(let ti=0;ti<m;ti++){
  if(ordered.has(ti))continue;
  const t=target[ti],available=source.map((s,index)=>({source:s,index})).filter(x=>!usedSource.has(x.index));
  let matchMode='exact',hits=available.filter(x=>nameMatchMode(t.name,x.source.name)==='exact');
  if(!hits.length){matchMode='suffix';hits=available.filter(x=>nameMatchMode(t.name,x.source.name)==='suffix')}
  if(!hits.length){matchMode='alias';hits=available.filter(x=>nameMatchMode(t.name,x.source.name)==='alias')}
  const groups=groupCandidates(hits),candidates=groups.filter(g=>g.code);let selected=candidates.length===1?0:null,reason=selected===0?'fallback':null;
  if(candidates.length>1&&t.quantity!==null&&t.quantity!==undefined){const fitting=candidates.map((candidate,index)=>candidate.count===t.quantity?index:-1).filter(index=>index>=0);if(fitting.length===1){selected=fitting[0];reason='quantity'}}
  if(selected!==null)for(const sourceIndex of candidates[selected].sourceIndices)usedSource.add(sourceIndex);
  results[ti]={...t,candidates,selected,reason,matchMode:hits.length?matchMode:null,status:candidates.length===1?'matched':candidates.length>1?'ambiguous':hits.length?'noCode':'missing'};
 }
 return results;
}

function shiftRef(s,at,delta=1){return s.replace(/(\$?)([A-Z]{1,3})(\$?\d+)/g,(m,a,b,c)=>colNum(b)>=at?a+colName(colNum(b)+delta)+c:m)}
function shiftFormula(s,at,delta,target,current){return s.replace(/"(?:[^"]|"")*"|(?:(?:'((?:[^']|'')+)'|([\p{L}\p{N}_.\[\]]+))!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?![\p{L}\p{N}_]|\s*\()/gu,(m,quoted,plain,ref)=>{if(m.startsWith('"'))return m;const scope=quoted?.replace(/''/g,"'")??plain;if(scope?scope!==target:current!==target)return m;return m.slice(0,m.length-ref.length)+shiftRef(ref,at,delta)})}
function exportColumns(book,sheet,cfg,columns,append=false){
 if(!Array.isArray(columns)||!columns.length)throw Error('没有可导出的列');
 const entries=new Map(book.entries),d=xml(serialize(sheet.doc)),delta=columns.length;
 const at=append?Math.max(...sheet.rows.flatMap(r=>Object.keys(r.cells).map(Number)),0)+1:cfg.nameCol+1;
 const headerRow=sheet.rows.find(r=>r.number===cfg.header);
 let prefix=0;
 if(!append)while(prefix<delta&&headerRow?.cells[at+prefix]===columns[prefix].header)prefix++;
 const insertDelta=append?delta:delta-prefix,insertAt=at+prefix,shift=!append&&insertDelta>0;
 if(shift){
  if(nodes(d,'tablePart').length||nodes(d,'drawing').length||nodes(d,'legacyDrawing').length)throw Error('该工作表含结构化表格或绘图，暂不支持保留这些对象插列。请使用普通审核表副本。');
  for(const el of Array.from(d.getElementsByTagName('*'))){
   for(const attr of ['r','ref','sqref','activeCell','topLeftCell'])if(el.hasAttribute(attr))el.setAttribute(attr,shiftRef(el.getAttribute(attr),insertAt,insertDelta));
   if(el.localName==='row'&&el.hasAttribute('spans'))el.removeAttribute('spans');
   if(el.localName==='col'){
    const min=+el.getAttribute('min'),max=+el.getAttribute('max');
    if(min>=insertAt){el.setAttribute('min',min+insertDelta);el.setAttribute('max',max+insertDelta)}
    else if(max>=insertAt){const tail=el.cloneNode();tail.setAttribute('min',insertAt+insertDelta);tail.setAttribute('max',max+insertDelta);el.setAttribute('max',insertAt-1);el.parentNode.append(tail)}
   }
   if(el.localName==='pane'&&+el.getAttribute('xSplit')>=insertAt-1)el.setAttribute('xSplit',+el.getAttribute('xSplit')+insertDelta);
  }
  for(const other of book.sheets){const sd=other.path===sheet.path?d:xml(serialize(other.doc));for(const f of [...nodes(sd,'f'),...nodes(sd,'formula'),...nodes(sd,'formula1'),...nodes(sd,'formula2')])f.textContent=shiftFormula(f.textContent,insertAt,insertDelta,sheet.name,other.name);if(other.path!==sheet.path)entries.set(other.path,enc.encode(serialize(sd)))}
  const wd=xml(dec.decode(entries.get('xl/workbook.xml')));for(const n of nodes(wd,'definedName'))n.textContent=shiftFormula(n.textContent,insertAt,insertDelta,sheet.name,book.sheets[Number(n.getAttribute('localSheetId'))]?.name??'');entries.set('xl/workbook.xml',enc.encode(serialize(wd)));
 }
 let cols=nodes(d,'cols')[0];if(!cols){cols=d.createElementNS(NS,'cols');d.documentElement.insertBefore(cols,nodes(d,'sheetData')[0])}
 const createFrom=append?0:prefix;
 for(let i=createFrom;i<delta;i++){const c=d.createElementNS(NS,'col');c.setAttribute('min',at+i);c.setAttribute('max',at+i);c.setAttribute('width',columns[i].width??(columns[i].header==='编码+名称'?'78':'18'));c.setAttribute('customWidth','1');cols.append(c)}
 if(createFrom<delta){const sorted=nodes(cols,'col').sort((a,b)=>+a.getAttribute('min')-+b.getAttribute('min'));sorted.forEach(c=>cols.append(c))}
 let stylesDoc=null,yellowFillId=null;const yellowStyles=new Map();
 const yellowStyle=baseStyle=>{if(!entries.has('xl/styles.xml'))return baseStyle;stylesDoc??=xml(dec.decode(entries.get('xl/styles.xml')));const fills=nodes(stylesDoc,'fills')[0],cellXfs=nodes(stylesDoc,'cellXfs')[0];if(!fills||!cellXfs)return baseStyle;
  if(yellowFillId===null){const directFills=Array.from(fills.children).filter(n=>n.localName==='fill');yellowFillId=directFills.length;const fill=stylesDoc.createElementNS(NS,'fill'),pattern=stylesDoc.createElementNS(NS,'patternFill'),fg=stylesDoc.createElementNS(NS,'fgColor'),bg=stylesDoc.createElementNS(NS,'bgColor');pattern.setAttribute('patternType','solid');fg.setAttribute('rgb','FFFFFF00');bg.setAttribute('indexed','64');pattern.append(fg,bg);fill.append(pattern);fills.append(fill);fills.setAttribute('count',String(yellowFillId+1))}
  const key=String(baseStyle??0);if(yellowStyles.has(key))return yellowStyles.get(key);const xfs=Array.from(cellXfs.children).filter(n=>n.localName==='xf'),base=xfs[Number(key)]??xfs[0];if(!base)return baseStyle;const xf=base.cloneNode(true),styleId=xfs.length;xf.setAttribute('fillId',String(yellowFillId));xf.setAttribute('applyFill','1');cellXfs.append(xf);cellXfs.setAttribute('count',String(styleId+1));yellowStyles.set(key,styleId);return styleId};
 const writeCell=(rn,col,value,highlight=false)=>{const row=nodes(d,'row').find(r=>+r.getAttribute('r')===rn);if(!row)return;const ref=colName(col)+rn,old=nodes(row,'c').find(c=>c.getAttribute('r')===ref);if(old)old.remove();const c=d.createElementNS(NS,'c');c.setAttribute('r',ref);c.setAttribute('t','inlineStr');const style=nodes(row,'c').find(c=>c.getAttribute('r')===colName(cfg.nameCol)+rn)?.getAttribute('s')??'0';c.setAttribute('s',String(highlight?yellowStyle(style):style));const is=d.createElementNS(NS,'is'),t=d.createElementNS(NS,'t');t.setAttribute('xml:space','preserve');t.textContent=value??'';is.append(t);c.append(is);const next=nodes(row,'c').find(c=>colNum(c.getAttribute('r').match(/[A-Z]+/)[0])>col);row.insertBefore(c,next??null)};
 columns.forEach((column,i)=>{const data=new Map(column.values??[]),highlightRows=new Set(column.highlightRows??[]);data.set(cfg.header,column.header);for(const[rn,value]of data)writeCell(rn,at+i,value,rn!==cfg.header&&highlightRows.has(rn))});
 if(stylesDoc)entries.set('xl/styles.xml',enc.encode(serialize(stylesDoc)));
 const dim=nodes(d,'dimension')[0];if(dim){const parts=dim.getAttribute('ref').split(':'),last=parts.at(-1),endCol=at+delta-1;if(colNum(last.match(/[A-Z]+/)[0])<endCol)dim.setAttribute('ref',(parts.length===2?parts[0]:'A1')+':'+colName(endCol)+Math.max(...sheet.rows.map(r=>r.number)))}
 entries.set(sheet.path,enc.encode(serialize(d)));return zip(entries);
}
function exportColumn(book,sheet,cfg,values,header='编码+名称',append=false){return exportColumns(book,sheet,cfg,[{header,values}],append)}

function applyRowFills(book,sheet,rowColorEntries){
 const entries=new Map(book.entries),d=xml(serialize(sheet.doc));
 if(!entries.has('xl/styles.xml'))throw Error('该 Excel 缺少样式文件，无法直接上色');
 const stylesDoc=xml(dec.decode(entries.get('xl/styles.xml'))),fills=nodes(stylesDoc,'fills')[0],cellXfs=nodes(stylesDoc,'cellXfs')[0];
 if(!fills||!cellXfs)throw Error('Excel 样式结构无法识别');
 const fillIds=new Map(),styleIds=new Map();
 const getFillId=color=>{
  color=String(color||'').replace(/^#/,'').toUpperCase();
  if(color.length===6)color='FF'+color;
  if(!/^[0-9A-F]{8}$/.test(color))throw Error('无效的 Excel 填充颜色');
  if(fillIds.has(color))return fillIds.get(color);
  const directFills=Array.from(fills.children).filter(n=>n.localName==='fill'),fillId=directFills.length;
  const fill=stylesDoc.createElementNS(NS,'fill'),pattern=stylesDoc.createElementNS(NS,'patternFill'),fg=stylesDoc.createElementNS(NS,'fgColor'),bg=stylesDoc.createElementNS(NS,'bgColor');
  pattern.setAttribute('patternType','solid');fg.setAttribute('rgb',color);bg.setAttribute('indexed','64');pattern.append(fg,bg);fill.append(pattern);fills.append(fill);fills.setAttribute('count',String(fillId+1));fillIds.set(color,fillId);return fillId;
 };
 const styleFor=(baseStyle,color)=>{
  const key=String(baseStyle??0)+'|'+color;
  if(styleIds.has(key))return styleIds.get(key);
  const xfs=Array.from(cellXfs.children).filter(n=>n.localName==='xf'),base=xfs[Number(baseStyle??0)]??xfs[0];
  if(!base)return baseStyle??0;
  const xf=base.cloneNode(true),id=xfs.length;xf.setAttribute('fillId',String(getFillId(color)));xf.setAttribute('applyFill','1');cellXfs.append(xf);cellXfs.setAttribute('count',String(id+1));styleIds.set(key,id);return id;
 };
 const rowsByNumber=new Map(nodes(d,'row').map(r=>[Number(r.getAttribute('r')),r]));
 for(const [rnRaw,color] of rowColorEntries){
  const rn=Number(rnRaw),row=rowsByNumber.get(rn);if(!row)continue;
  for(const cell of nodes(row,'c')){const base=cell.getAttribute('s')??'0';cell.setAttribute('s',String(styleFor(base,color)))}
 }
 entries.set('xl/styles.xml',enc.encode(serialize(stylesDoc)));entries.set(sheet.path,enc.encode(serialize(d)));return zip(entries);
}
window.FixtureExcel={codeCount,parseQuantity,readWorkbook,detect,records,matchRows,exportColumn,exportColumns,applyRowFills,colName};
})();

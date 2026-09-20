(()=>{
'use strict';
if(!window.FixtureExcel){const el=document.getElementById('message');el.textContent='Excel 组件未加载。请完整解压源码包后打开 index.html，或刷新在线网页。';el.classList.add('error');return;}
const{readWorkbook,detect,records,matchRows,exportColumn,exportColumns,colName}=window.FixtureExcel;
const $=id=>document.getElementById(id);let source,target,sourceSheet,targetSheet,sourceCfg,targetCfg,sourceRows=[],results=[];let batchVersion=0;
function message(s,error=false){$('message').textContent=s;$('message').classList.toggle('error',error)}
function invalidate(){results=[];$('results').hidden=true;$('empty').hidden=false;$('match').disabled=!(source&&target)}
function resetDetected(){
 source=target=sourceSheet=targetSheet=null;sourceCfg=targetCfg=null;sourceRows=[];results=[];
 for(const kind of ['source','target']){const sel=$(kind+'Sheet');sel.replaceChildren(new Option('自动识别后显示',''));sel.disabled=true;$(kind+'Label').textContent='等待识别'}
 invalidate();
}
function setDetected(kind,book,sheetIndex){
 const sel=$(kind+'Sheet');sel.replaceChildren(...book.sheets.map((s,i)=>new Option(s.name,i)));sel.value=String(sheetIndex);sel.disabled=false;
 $(kind+'Label').textContent=book.name;
 if(kind==='source'){source=book;sourceSheet=book.sheets[sheetIndex]}else{target=book;targetSheet=book.sheets[sheetIndex]}
}
function setSheet(kind){
 const book=kind==='source'?source:target;if(!book)return;const sel=$(kind+'Sheet'),sheet=book.sheets[+sel.value];
 if(kind==='source')sourceSheet=sheet;else targetSheet=sheet;invalidate();message('工作表已切换，点击「开始匹配」重新核对。')
}
function classify(book){
 const sourceSheets=[],targetSheets=[];
 book.sheets.forEach((sheet,i)=>{try{detect(sheet,'source');sourceSheets.push(i)}catch{}try{detect(sheet,'target');targetSheets.push(i)}catch{}});
 return{sourceSheets,targetSheets};
}
function roleScore(file,cls,role){
 const name=file.name.toUpperCase();let score=10;
 if(role==='source'){if(!cls.targetSheets.length)score+=5;if(/流程|统计|编码名称/.test(name))score+=3;if(/审核|TX-/.test(name))score-=1}
 else{if(!cls.sourceSheets.length)score+=5;if(/审核|申请|TX-/.test(name))score+=3;if(/流程统计/.test(name))score-=1}
 return score;
}
async function loadPair(fileList){
 const files=Array.from(fileList??[]);
 if(files.length!==2){message('需要两份 .xlsx 文件。你可以一次拖入两份，也可以先拖一份、再拖第二份。',true);return}
 if(files.some(f=>!/\.xlsx$/i.test(f.name))){message('两份文件都必须是 .xlsx 格式。',true);return}
 const version=++batchVersion;resetDetected();$('bothLabel').textContent='正在读取并自动识别…';message('正在读取两份 Excel，并根据表头自动判断文件类型…');
 try{
  const books=await Promise.all(files.map(readWorkbook));if(version!==batchVersion)return;
  const classes=books.map(classify),combos=[];
  if(classes[0].sourceSheets.length&&classes[1].targetSheets.length)combos.push({s:0,t:1,si:classes[0].sourceSheets[0],ti:classes[1].targetSheets[0],score:roleScore(files[0],classes[0],'source')+roleScore(files[1],classes[1],'target')});
  if(classes[1].sourceSheets.length&&classes[0].targetSheets.length)combos.push({s:1,t:0,si:classes[1].sourceSheets[0],ti:classes[0].targetSheets[0],score:roleScore(files[1],classes[1],'source')+roleScore(files[0],classes[0],'target')});
  if(!combos.length)throw Error('无法从这两份文件中同时识别出「流程统计表」和「申请审核表」。请确认一份包含“治工具编号/名称说明”，另一份包含“治具名称+型号”。');
  combos.sort((x,y)=>y.score-x.score);const pick=combos[0];
  setDetected('source',books[pick.s],pick.si);setDetected('target',books[pick.t],pick.ti);
  $('bothLabel').textContent='两份文件已自动识别';
  invalidate();message(`识别完成：流程统计表「${source.name}」；申请审核表「${target.name}」。可直接点击「开始匹配」。`);
 }catch(e){if(version!==batchVersion)return;resetDetected();$('bothLabel').textContent='重新选择两份 Excel';message(e.message,true)}
}
const input=$('bothFiles'),box=$('dropZone');let dragDepth=0,pendingFiles=[];
function sameFile(a,b){return a.name===b.name&&a.size===b.size&&a.lastModified===b.lastModified}
function acceptFiles(fileList){
 const incoming=Array.from(fileList??[]);
 if(!incoming.length){message('没有检测到可读取的文件。请从文件资源管理器拖入 .xlsx 文件。',true);return}
 const invalid=incoming.filter(f=>!/\.xlsx$/i.test(f.name));
 if(invalid.length){message('只支持 .xlsx 文件。请确认拖入的是 Excel 文件本身，而不是快捷方式或工作表内容。',true);return}
 if(incoming.length>=2){
  pendingFiles=[];loadPair(incoming.slice(0,2));return;
 }
 const file=incoming[0];
 if(source&&target&&pendingFiles.length===0){resetDetected()}
 if(!pendingFiles.some(f=>sameFile(f,file)))pendingFiles.push(file);
 if(pendingFiles.length===1){
  $('bothLabel').textContent=`已收到 1 份：${file.name}`;
  message('已收到第 1 份 Excel。请再拖入或选择第 2 份，系统会自动识别两份表的类型。');
  return;
 }
 const pair=pendingFiles.slice(0,2);pendingFiles=[];loadPair(pair);
}
input.addEventListener('click',e=>{e.stopPropagation();input.value=''});
input.addEventListener('change',e=>acceptFiles(e.target.files));
box.addEventListener('click',e=>{if(e.target!==input)input.click()});
box.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();input.click()}});
for(const kind of ['source','target'])$(kind+'Sheet').addEventListener('change',()=>setSheet(kind));
box.addEventListener('dragenter',e=>{if(!Array.from(e.dataTransfer?.types??[]).includes('Files'))return;e.preventDefault();e.stopPropagation();dragDepth++;box.classList.add('dragover');$('bothLabel').textContent='松开鼠标即可加入 Excel'});
box.addEventListener('dragover',e=>{if(!Array.from(e.dataTransfer?.types??[]).includes('Files'))return;e.preventDefault();e.stopPropagation();if(e.dataTransfer)e.dataTransfer.dropEffect='copy';box.classList.add('dragover')});
box.addEventListener('dragleave',e=>{e.preventDefault();e.stopPropagation();dragDepth=Math.max(0,dragDepth-1);if(!dragDepth){box.classList.remove('dragover');if(pendingFiles.length===1)$('bothLabel').textContent=`已收到 1 份：${pendingFiles[0].name}`;else if(!(source&&target))$('bothLabel').textContent='选择或拖入两份 Excel'}});
box.addEventListener('drop',e=>{e.preventDefault();e.stopPropagation();dragDepth=0;box.classList.remove('dragover');acceptFiles(e.dataTransfer?.files)});
window.addEventListener('dragover',e=>{if(Array.from(e.dataTransfer?.types??[]).includes('Files'))e.preventDefault()});
window.addEventListener('drop',e=>{if(Array.from(e.dataTransfer?.types??[]).includes('Files'))e.preventDefault()});
window.addEventListener('dragend',()=>{dragDepth=0;box.classList.remove('dragover')});
resetDetected();message('请拖入或选择两份 Excel。可以一次加入两份，也可以分两次加入；系统会自动识别文件类型。原文件不会被修改。');
window.fixtureAppReady=true;

function people(r){const c=r.selected===null?null:r.candidates[r.selected];return{requester:String(r.requester||c?.requester||'').trim(),designer:String(r.designer||c?.designer||'').trim()}}
function personMapping(cfg,label){const out=[];if(cfg.requesterCol)out.push(`需求人 ${colName(cfg.requesterCol)}列`);if(cfg.designerCol)out.push(`设计人 ${colName(cfg.designerCol)}列`);return out.length?`${label}${out.join(' / ')}`:''}
function run(){if(!source||!target)throw Error('请先选择两份 Excel');sourceCfg=detect(sourceSheet,'source');targetCfg=detect(targetSheet,'target');sourceRows=records(sourceSheet,sourceCfg);const ts=records(targetSheet,targetCfg);if(!sourceRows.length||!ts.length)throw Error('选择的工作表没有可匹配的数据');results=matchRows(ts,sourceRows);$('results').hidden=false;$('empty').hidden=true;const personInfo=[personMapping(targetCfg,'审核表：'),personMapping(sourceCfg,'流程表：')].filter(Boolean).join('；'),orderCount=results.filter(r=>r.reason==='order').length,fallbackCount=results.filter(r=>r.reason==='fallback'||r.reason==='quantity').length;$('mapping').textContent=`先按原始顺序对齐，再对异常项补充匹配；顺序对应 ${orderCount} 条，补充匹配 ${fallbackCount} 条。来源 ${colName(sourceCfg.codeCol)} 列 + ${colName(sourceCfg.nameCol)} 列 → 审核表 ${colName(targetCfg.nameCol)} 列${personInfo?'；'+personInfo:'；未识别到人员表头，需求人/设计人将留空'}`;message(`已读取 ${sourceRows.length} 条流程记录，核对 ${results.length} 条审核项目。系统优先保持两份表的原始顺序，同一流程记录不会自动重复分配。${sourceCfg.fallback?'未识别到标准标题，当前使用 X / AP 列，请核对。':''}`);render();return summary()}
function summary(){return{total:results.length,matched:results.filter(r=>r.selected!==null).length,ambiguous:results.filter(r=>r.status==='ambiguous'&&r.selected===null).length,missing:results.filter(r=>!r.candidates.length).length}}
$('match').onclick=()=>{try{run()}catch(e){message(e.message,true)}};
function appendTextCell(tr,value,empty='—'){const td=document.createElement('td');td.textContent=value||empty;tr.append(td)}
function render(){const counts=summary();for(const[k,v]of Object.entries(counts))$(k).textContent=v;const search=$('search').value.toLowerCase(),filter=$('filter').value;$('rows').replaceChildren();for(const r of results){const done=r.selected!==null,p=people(r),haystack=`${r.name} ${p.requester} ${p.designer} ${r.candidates.map(c=>`${c.code} ${c.requester} ${c.designer}`).join(' ')}`.toLowerCase();if(filter==='todo'&&done||filter==='done'&&!done)continue;if(search&&!haystack.includes(search))continue;const tr=document.createElement('tr');appendTextCell(tr,String(r.row),'');appendTextCell(tr,r.name,'');appendTextCell(tr,r.quantity===null?'—':String(r.quantity),'—');appendTextCell(tr,p.requester);appendTextCell(tr,p.designer);const st=document.createElement('td'),badge=document.createElement('span');badge.className='badge'+(!done?(r.candidates.length?' todo':' miss'):'');badge.textContent=done?(r.reason==='order'?'顺序对应':r.reason==='fallback'?'补充匹配':r.reason==='quantity'?'数量匹配':r.reason==='alias'?'近似名称匹配':r.reason==='suffix'?'名称简写匹配':r.status==='ambiguous'?'已确认':'已匹配'):r.status==='ambiguous'?'同名待确认':r.status==='noCode'?'来源无编码':'未找到';st.append(badge);tr.append(st);const value=document.createElement('td');if(done){const c=r.candidates[r.selected];const line=document.createElement('div');line.textContent=c.code+' '+c.name;value.append(line);const note=document.createElement('small');const who=[c.requester&&`需求人：${c.requester}`,c.designer&&`设计人：${c.designer}`].filter(Boolean).join(' · ');note.textContent=`流程表第 ${c.rows.join('、')} 行 · 编码个数：${c.count??'无法识别'}${r.reason==='quantity'?`，与审核数量 ${r.quantity} 一致`:''}${r.quantity!==null&&c.count!==null&&r.quantity!==c.count?` · ⚠ 审核数量 ${r.quantity} 与编码个数 ${c.count} 不同`:''}${who?' · '+who:''}`;value.append(note)}else if(!r.candidates.length)value.textContent=r.status==='noCode'?'名称可匹配，但流程表编码为空。':'流程表没有找到可匹配的名称。';if(r.candidates.length>1){const sel=document.createElement('select');sel.className='candidate';sel.setAttribute('aria-label',`审核表第 ${r.row} 行选择编码`);sel.append(new Option('请选择对应编码（未选时编码、人名按可用信息导出）',''));r.candidates.forEach((c,i)=>{const who=[c.requester,c.designer].filter(Boolean).join(' / ');sel.append(new Option(`第 ${c.rows.join('、')} 行 · ${c.count===null?'个数无法识别':c.count+' 个编码'} · ${c.code} ${c.name}${who?' · '+who:''}`,i))});sel.value=r.selected===null?'':String(r.selected);sel.onchange=()=>{r.selected=sel.value===''?null:+sel.value;r.reason=r.selected===null?null:'manual';render()};value.append(sel)}tr.append(value);$('rows').append(tr)}if(!$('rows').children.length){const tr=document.createElement('tr'),td=document.createElement('td');td.colSpan=7;td.textContent='没有符合条件的项目。';tr.append(td);$('rows').append(tr)}}
$('filter').onchange=render;$('search').oninput=render;
function download(blob,name){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),30000)}
$('exportTarget').onclick=()=>{try{const codeValues=[],requesterValues=[],designerValues=[],problemRows=[];for(const r of results){const p=people(r),candidate=r.selected===null?null:r.candidates[r.selected],code=candidate?candidate.code+' '+candidate.name:'',quantityMismatch=!!(candidate&&r.quantity!==null&&candidate.count!==null&&r.quantity!==candidate.count);codeValues.push([r.row,code]);requesterValues.push([r.row,p.requester]);designerValues.push([r.row,p.designer]);if(r.selected===null||quantityMismatch)problemRows.push(r.row)}download(exportColumns(target,targetSheet,targetCfg,[{header:'编码+名称',values:codeValues,width:78,highlightRows:problemRows},{header:'需求人',values:requesterValues,width:18},{header:'设计人',values:designerValues,width:18}]),target.name.replace(/\.xlsx$/i,'_编码人员匹配.xlsx'));const n=results.filter(r=>r.selected===null).length;message(`已导出审核结果，并新增「编码+名称 / 需求人 / 设计人」。有问题的「编码+名称」单元格已标黄色。${n?`${n} 个未确认或未找到的项目编码保留空白。`:'所有项目均已填入编码。'}`)}catch(e){message(e.message,true)}};
$('exportSource').onclick=()=>{try{download(exportColumn(source,sourceSheet,sourceCfg,sourceRows.map(r=>[r.row,r.code+' '+r.name]),'编码+名称',true),source.name.replace(/\.xlsx$/i,'_编码名称.xlsx'));message('已导出流程拼接表，在末尾添加「编码+名称」。')}catch(e){message(e.message,true)}};
if(document.modelContext?.registerTool){try{Promise.resolve(document.modelContext.registerTool({name:'match_uploaded_workbooks',description:'对用户已选择的两份 Excel 先按原始顺序对齐，再对异常项执行名称和数量匹配，并显示结果；不导出文件。',inputSchema:{type:'object',properties:{},additionalProperties:false},annotations:{readOnlyHint:false,untrustedContentHint:true},execute(input){if(!input||typeof input!=='object'||Object.keys(input).length)throw Error('不接受参数');return run()}})).catch(()=>{})}catch{}}
})();

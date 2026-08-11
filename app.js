(()=>{"use strict";
const C=window.STORE_STOCK_CONFIG,$=id=>document.getElementById(id);
let products=[],stores=[],sales=new Map(),baseMonday=monday(new Date()),through=addDays(new Date(),-1),pendingRows=null,uploadInProgress=false;
function d0(d){return new Date(d.getFullYear(),d.getMonth(),d.getDate())}
function monday(d){const x=d0(d),day=x.getDay();x.setDate(x.getDate()+(day===0?-6:1-day));return x}
function addDays(d,n){const x=d0(d);x.setDate(x.getDate()+n);return x}
function iso(d){return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`}
function fmt(d){return iso(d).replaceAll("-","/")}
function txt(v){return String(v??"").trim()}
function jan(v){let s=txt(v).replace(/^"|"$/g,"");return /^\d+\.0$/.test(s)?s.slice(0,-2):s}
function qty(v){const n=Number(String(v??"").replace(/,/g,"").trim());return Number.isFinite(n)?Math.trunc(n):0}
function parseDate(v){if(v==null||v==="")return null;if(v instanceof Date&&!isNaN(v))return d0(v);if(typeof v==="number"&&window.XLSX){const p=XLSX.SSF.parse_date_code(v);if(p)return new Date(p.y,p.m-1,p.d)}const m=txt(v).replace(/^"|"$/g,"").match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);return m?new Date(+m[1],+m[2]-1,+m[3]):null}
function key(j,s){return `${j}|||${s}`}
function esc(v){return String(v??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]))}
function notice(m,t=""){const e=$("notice");if(!m){e.className="notice hidden";e.textContent="";return}e.className=`notice ${t}`;e.textContent=m}
function hidx(h,w){const z=txt(w).toLowerCase();return h.findIndex(x=>txt(x).toLowerCase()===z)}
function canon(h){const z=txt(h);for(const [k,a] of Object.entries(C.STORE_ALIASES||{}))if(a.some(x=>txt(x)===z))return k;return z}
function detectStores(headers){const valid=new Set(C.SALES_SHEETS.map(x=>x.store)),out=[];headers.forEach((h,i)=>{const c=canon(h);if(valid.has(c))out.push({store:c,index:i})});return out}
function parseCSV(s){const rows=[];let row=[],f="",q=false;for(let i=0;i<s.length;i++){const c=s[i];if(q){if(c==='"'){if(s[i+1]==='"'){f+='"';i++}else q=false}else f+=c}else{if(c==='"')q=true;else if(c===","){row.push(f);f=""}else if(c==="\n"){row.push(f);rows.push(row);row=[];f=""}else if(c!=="\r")f+=c}}row.push(f);if(row.some(v=>v!==""))rows.push(row);return rows}
function csvUrl(id,gid){return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(id)}/gviz/tq?tqx=out:csv&gid=${encodeURIComponent(gid)}&t=${Date.now()}`}
async function fetchRows(id,gid,label){const r=await fetch(csvUrl(id,gid),{cache:"no-store"});if(!r.ok)throw new Error(`${label}: HTTP ${r.status}`);const t=await r.text();if(/^\s*</.test(t)||t.toLowerCase().includes("<html"))throw new Error(`${label}: Googleシートを取得できません。`);return parseCSV(t)}
function jsonp(params){return new Promise((resolve,reject)=>{if(!C.API_URL)return reject(new Error("API URLが未設定です。"));const cb=`__ss_${Date.now()}_${Math.floor(Math.random()*1e6)}`;let done=false;const sc=document.createElement("script"),timer=setTimeout(()=>finish(new Error("月曜在庫の読込がタイムアウトしました。")),20000);function finish(err,data){if(done)return;done=true;clearTimeout(timer);delete window[cb];sc.remove();err?reject(err):resolve(data)}window[cb]=d=>finish(null,d);sc.onerror=()=>finish(new Error("月曜在庫へ接続できません。"));sc.src=`${C.API_URL}?${new URLSearchParams({...params,callback:cb,_:Date.now()})}`;document.head.appendChild(sc)})}
async function loadStock(){const data=await jsonp({action:"stock"});if(!data?.ok)throw new Error(data?.message||"月曜在庫を読み込めません。");if(data.baseDate!==iso(baseMonday))throw new Error(`今週の月曜在庫が未更新です。保存中: ${data.baseDate||"なし"} / 必要: ${iso(baseMonday)}`);const rows=data.rows;if(!Array.isArray(rows)||rows.length<2)throw new Error("月曜在庫データがありません。");const h=rows[0].map(txt),ix={jan:hidx(h,"JAN"),sku:hidx(h,"品番"),name:hidx(h,"品名")};if(Object.values(ix).some(i=>i<0))throw new Error("月曜在庫の必須見出し JAN / 品番 / 品名 が不足しています。");const sc=detectStores(h);if(!sc.length)throw new Error("対象店舗列を検出できません。");stores=sc.map(x=>x.store);products=[];for(let r=1;r<rows.length;r++){const row=rows[r],j=jan(row[ix.jan]),sku=txt(row[ix.sku]);if(!j||!sku)continue;const stock={};sc.forEach(s=>stock[s.store]=qty(row[s.index]));products.push({jan:j,sku,name:txt(row[ix.name]),stock})}if(!products.length)throw new Error("月曜在庫の商品を1件も読み込めませんでした。")}
async function loadSales(){sales=new Map();if(through<baseMonday)return;const rs=await Promise.allSettled(C.SALES_SHEETS.map(async s=>{const rows=await fetchRows(s.id,s.gid,s.store);for(const row of rows){if(row.length<6)continue;const j=jan(row[0]),sold=qty(row[2]),dt=parseDate(row[5]);if(!j||!dt||dt<baseMonday||dt>through)continue;const k=key(j,s.store);sales.set(k,(sales.get(k)||0)+sold)}}));const bad=[];rs.forEach((r,i)=>{if(r.status==="rejected")bad.push(C.SALES_SHEETS[i].store)});if(bad.length)throw new Error(`売上シート取得NG: ${bad.join("、")}`)}
function cur(p,s){const b=p.stock[s]||0,v=sales.get(key(p.jan,s))||0;return{base:b,sold:v,current:Math.max(0,b-v)}}
function render(){const q=$("searchInput").value.trim().toUpperCase();if(!q){$("results").innerHTML='<div class="empty">品番を入力してください。</div>';return}let m=products.filter(p=>p.sku.toUpperCase().includes(q));m=m.slice(0,C.MAX_RESULTS||200);$("results").innerHTML=m.length?m.map(p=>`<article class="item"><div class="item-head"><div><div class="sku">${esc(p.sku)}</div><div class="name">${esc(p.name)}</div></div><div class="jan">JAN ${esc(p.jan)}</div></div><div class="table-wrap"><table><thead><tr>${stores.map(s=>`<th>${esc(s.replace("ららぽーと",""))}</th>`).join("")}</tr></thead><tbody><tr>${stores.map(s=>{const x=cur(p,s);return `<td class="${x.current>0?"":"zero"}"><div class="current">${x.current}</div><div class="detail">月曜 ${x.base} − 売れ ${x.sold}</div></td>`}).join("")}</tr></tbody></table></div></article>`).join(""):'<div class="empty">一致する品番がありません。</div>'}
async function loadAll(){$("searchInput").disabled=true;notice("");baseMonday=monday(new Date());through=addDays(new Date(),-1);$("baseDate").textContent=`${fmt(baseMonday)}（月）`;$("salesThrough").textContent=through<baseMonday?"差引なし":`${fmt(through)}まで`;try{await loadStock();await loadSales();$("searchInput").disabled=false;notice("");render()}catch(e){notice(e.message,"error");$("results").innerHTML='<div class="empty">月曜在庫を取り込んでください。</div>'}}
function compactCell(v){return txt(v).replace(/\s+/g," ")}
function uploadMatches(sourceRows,serverRows){
  if(!Array.isArray(sourceRows)||!Array.isArray(serverRows))return false;
  if(sourceRows.length!==serverRows.length||sourceRows.length<2)return false;
  const picks=[0,1,Math.floor((sourceRows.length-1)/2),sourceRows.length-1];
  for(const r of [...new Set(picks)]){
    const a=sourceRows[r]||[],b=serverRows[r]||[],width=Math.min(a.length,b.length,9);
    for(let c=0;c<width;c++)if(compactCell(a[c])!==compactCell(b[c]))return false;
  }
  return true
}
function sleep(ms){return new Promise(resolve=>setTimeout(resolve,ms))}
async function confirmUpload(sourceRows){
  await sleep(6000);
  for(let attempt=1;attempt<=8;attempt++){
    try{
      const data=await jsonp({action:"stock"});
      if(data?.ok&&data.baseDate===iso(baseMonday)&&uploadMatches(sourceRows,data.rows))return data;
    }catch(_){}
    if(attempt<8){
      $("uploadState").textContent=`更新確認中… ${attempt}/8`;
      await sleep(2000);
    }
  }
  throw new Error("更新完了を確認できませんでした。再読込して在庫が反映されているか確認してください。");
}
async function chooseAndUpload(file){
  if(uploadInProgress)return;
  try{
    const b=await file.arrayBuffer(),wb=XLSX.read(b,{type:"array",cellDates:true}),ws=wb.Sheets[wb.SheetNames[0]],rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:false});
    if(rows.length<2)throw new Error("Excelにデータがありません。");
    const h=rows[0].map(txt);
    for(const req of ["JAN","品番","品名"])if(hidx(h,req)<0)throw new Error(`必要な見出しがありません: ${req}`);
    const sc=detectStores(h);if(!sc.length)throw new Error("対象店舗列を検出できません。");
    if(!confirm(`${file.name}\n${rows.length.toLocaleString()}行\n${sc.map(x=>x.store).join(" / ")}\n\nこの内容で月曜在庫を更新しますか？`))return;
    pendingRows=rows;
    await uploadRows();
  }catch(e){
    finishUploadUI();
    notice(`取込NG: ${e.message}`,"error");
  }
}
function startUploadUI(){
  uploadInProgress=true;
  $("selectStockBtn").disabled=true;
  $("reloadBtn").disabled=true;
  $("uploadState").textContent="更新中…";
}
function finishUploadUI(){
  uploadInProgress=false;
  $("selectStockBtn").disabled=false;
  $("reloadBtn").disabled=false;
  $("uploadState").textContent="";
}
function submitUploadForm(rows){
  const payload=JSON.stringify({baseDate:iso(baseMonday),rows});
  const form=document.createElement("form");
  form.method="POST";form.action=C.API_URL;form.target="uploadFrame";form.style.display="none";
  for(const [name,value] of Object.entries({action:"upload",origin:location.origin,payload})){
    const i=document.createElement("input");i.type="hidden";i.name=name;i.value=value;form.appendChild(i);
  }
  document.body.appendChild(form);form.submit();setTimeout(()=>form.remove(),1500);
}
async function uploadRows(){
  if(!pendingRows)return;
  if(!C.API_URL){notice("API URLが未設定です。","error");return}
  const sourceRows=pendingRows;
  startUploadUI();
  submitUploadForm(sourceRows);
  try{
    const confirmed=await confirmUpload(sourceRows);
    pendingRows=null;
    finishUploadUI();
    notice(`月曜在庫を更新しました。${confirmed.rows.length.toLocaleString()}行`,"success");
    await loadAll();
  }catch(e){
    finishUploadUI();
    notice(`月曜在庫更新確認NG: ${e.message}`,"error");
  }
}
window.addEventListener("message",e=>{
  if(!uploadInProgress||!e.data||e.data.type!=="STORESTOCK_UPLOAD")return;
  if(e.data.ok===false){
    finishUploadUI();
    notice(`月曜在庫更新NG: ${e.data.message||"不明なエラー"}`,"error");
  }
});
$("reloadBtn").addEventListener("click",()=>{if(!uploadInProgress)loadAll()});$("searchInput").addEventListener("input",render);$("selectStockBtn").addEventListener("click",()=>{if(!uploadInProgress)$("stockFileInput").click()});$("stockFileInput").addEventListener("change",e=>{const f=e.target.files?.[0];if(f)chooseAndUpload(f);e.target.value=""});loadAll();
})();
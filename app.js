/* Today's Basket V4.6 — persistent employee workbooks + Supabase authentication. */
const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const n=v=>{if(typeof v==='number')return Number.isFinite(v)?v:0;let x=Number(String(v??'').replace(/,/g,'').replace(/[₹ ]/g,''));return Number.isFinite(x)?x:0};
const fmt=v=>new Intl.NumberFormat('en-IN',{maximumFractionDigits:2}).format(v||0),money=v=>'₹'+fmt(v),key=v=>String(v??'').trim(),norm=v=>key(v).toLowerCase();
const round=v=>Math.ceil((v-1e-9)*100)/100;
let page='overview', tab='all', filter='',uploads={}, meta={}, records=[], matches=[], stats={}, edits={},covers={}, mappings={}, approved=false, revision=1, history=[], activeId=null;
let analysisBrand='all',analysisSubcategory='all',analysisCategoryQuery='',grnData=null,grnLevel='months',grnMonth='',grnDay='',grnNumber='',grnSearch='',grnImportError='',grnLoading=false,analyticsImportError='';
let salesDays=60,baseCover=30,db=null,indentCategory='all',analysisCategory='all',analysisFloor='all',analysisSearch='',analyticsRows=[],analyticsMeta=null,analyticsMode='value',analysisMonth='all',selectedSalesDay=null,currentAnalyticsView=null;
function flash(x){const t=$('#toast');t.textContent=x;t.classList.remove('hidden');clearTimeout(window.tm);window.tm=setTimeout(()=>t.classList.add('hidden'),4500)}
function today(){return new Date().toISOString().slice(0,10)}
function dateOf(v){if(v instanceof Date)return new Date(v.getFullYear(),v.getMonth(),v.getDate());if(typeof v==='number'&&v>20000&&v<90000){const t=XLSX.SSF.parse_date_code(v);return t?new Date(t.y,t.m-1,t.d):null}let s=key(v);if(!s)return null;let m=s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/);if(m)return new Date(+m[3],+m[2]-1,+m[1]);let d=new Date(s);return isNaN(d.getTime())?null:new Date(d.getFullYear(),d.getMonth(),d.getDate())}
function compareDates(d){let x=new Date();x.setHours(0,0,0,0);return d?Math.round((d-x)/86400000):null}
function isNumeric(v){return v!==null&&v!==undefined&&String(v).trim()!==''&&Number.isFinite(Number(String(v).replace(/,/g,'')))}
function headerIndex(rows,required){for(let i=0;i<Math.min(rows.length,35);i++){let h=rows[i].map(x=>norm(x).replace(/\s+/g,' '));if(required.every(x=>h.includes(x)))return i}return -1}
function col(headers,name,fallback){const at=headers.findIndex(v=>norm(v).replace(/\s+/g,' ')===name);return at<0?fallback:at}
function parseStock(rows){const hi=headerIndex(rows,['particulars']);if(hi<0)throw Error('Could not find Particulars column. Please check this stock balance report.');const h=rows[hi];let out=[];for(const r of rows.slice(hi+1)){let name=key(r[col(h,'particulars',0)]),sku=key(r[col(h,'item code',1)]);if(!sku||!name||/^(grand total|total|today.s basket|arunachala logistics)/i.test(name))continue;let stock=r[col(h,'quantity',5)];if(!isNumeric(stock))continue;out.push({sku,name,cat:key(r[col(h,'category name',3)])||'Uncategorized',qty:n(stock),uom:key(r[col(h,'uom',8)])||key(r[8])||'NOS',shelf:n(r[col(h,'shelf life in days',9)])})}return out}
function parseSales(rows){let hi=headerIndex(rows,['item','total quantity','total amount']);if(hi<0)throw Error('Sales sheet must contain Item, Total Quantity and Total Amount.');let out=[];for(const r of rows.slice(hi+1)){let name=key(r[0]);if(!name||/^(total|grand total|category|group)/i.test(name)||!isNumeric(r[1])||!isNumeric(r[2]))continue;out.push({name,qty:n(r[1]),amount:n(r[2])})}return out}
function parseAnalytics(rows){
  let hi=-1, headers=[];
  for(let i=0;i<Math.min(rows.length,35);i++){
    let h=(rows[i]||[]).map(x=>norm(x).replace(/[^a-z0-9]+/g,''));
    if(h.some(x=>['item','itemname','product','productname','particulars','sku'].includes(x)) && h.some(x=>['amount','totalamount','salesamount','netamount','saleamount'].includes(x))) {hi=i;headers=h;break}
  }
  if(hi<0)throw Error('Floor/day report needs ItemName (or Item) and TotalAmount headers.');
  const ix=names=>headers.findIndex(x=>names.includes(x));
  const item=ix(['item','itemname','product','productname','particulars','sku']), amt=ix(['amount','totalamount','salesamount','netamount','saleamount']);
  const qty=ix(['quantity','totalquantity','qty','totalqty','salesqty']);
  const floor=ix(['profitcentre','profitcenter','floor','floorname','floorwise','profitcentrename']);
  const date=ix(['date','createddate','billdate','invoicedate','salesdate','transactiondate']);
  const cat=ix(['category','categoryname','department']);
  const subcat=ix(['subcategory','subcategoryname','subgroup']);
  const skuIx=ix(['itemcode','sku','skucode']);
   const brandIx=ix(['brand','brandname','itembrand','productbrand','manufacturerbrand','brands','branddescription','brandlabel']);
  if(floor<0 && date<0)throw Error('Floor/day report must also contain a Floor/Profit Centre column or a Date column.');
  let out=[];
  for(const r of rows.slice(hi+1)){
    const name=key(r[item]);if(!name||/^(total|grand total|subtotal|group)/i.test(name)||!isNumeric(r[amt]))continue;
    const d=date>=0?dateOf(r[date]):null;
    const day=d?`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`:null;
    out.push({name,sku:skuIx>=0?key(r[skuIx]):'',amount:n(r[amt]),qty:qty>=0?n(r[qty]):null,floor:floor>=0?key(r[floor]):'',day,cat:cat>=0?key(r[cat]):'',subcat:subcat>=0?key(r[subcat]):'',brand:brandIx>=0?(key(r[brandIx])||'Unspecified brand'):'Unspecified brand'});
  }
  if(!out.length)throw Error('No valid sales rows detected in the floor/day report.');
  return {rows:out,hasFloor:floor>=0,hasDay:date>=0};
}
function parsePending(rows){let hi=headerIndex(rows,['indent no']);if(hi<0)throw Error('Could not find Indent No in the previous-indent report.');const h=rows[hi];let out=[];for(const r of rows.slice(hi+1)){let sku=key(r[col(h,'itemcode',3)]||r[3]);if(!sku)continue;let exp=dateOf(r[col(h,'expirydate',2)]),pending=n(r[col(h,'pendingqty',7)]);if(!Number.isFinite(pending)||pending<=0)continue;out.push({sku,indent:key(r[col(h,'indent no',0)]),exp,days:compareDates(exp),pending,status:key(r[col(h,'indent status',11)]),isExpired:exp?compareDates(exp)<0:false,expiryUnknown:!exp})}return out}
function datasets(){return {store:uploads.store||[],warehouse:uploads.warehouse||[],sales:uploads.sales||[],pending:uploads.pending||[]}}
function complete(){return ['store','warehouse','sales','pending'].every(t=>uploads[t])}
function review(){const d=datasets(),wm=new Map(),pm=new Map(),sm=new Map(),sn=new Map(),duplicates=[];for(const r of d.warehouse){if(wm.has(r.sku))duplicates.push('Warehouse duplicate Item Code '+r.sku);wm.set(r.sku,(wm.get(r.sku)||0)+r.qty)}for(const r of d.pending){const p=pm.get(r.sku)||{valid:0,expired:0,unknown:0};if(r.isExpired)p.expired+=r.pending;else if(r.expiryUnknown)p.unknown+=r.pending;else p.valid+=r.pending;pm.set(r.sku,p)}for(const r of d.sales){let name=key(r.name),prev=sm.get(name)||{qty:0,amount:0};prev.qty+=r.qty;prev.amount+=r.amount;sm.set(name,prev)}const nameCount=new Map();for(const r of d.store)nameCount.set(r.name,(nameCount.get(r.name)||0)+1);let rows=[];for(const s of d.store){if(sn.has(s.sku))duplicates.push('Store duplicate Item Code '+s.sku);sn.set(s.sku,true);let mapped=mappings[s.sku],salesName=mapped||s.name,sale=sm.get(salesName);let ambiguous=nameCount.get(s.name)>1&&!mapped;let pend=pm.get(s.sku)||{valid:0,expired:0,unknown:0};let warehouse=wm.has(s.sku)?wm.get(s.sku):null;rows.push({...s,warehouse,salesName,salesQty:sale?.qty??null,amount:sale?.amount??null,matched:!!sale&&!ambiguous,ambiguous,pending:pend.valid,expiredPending:pend.expired,unknownPending:pend.unknown})}const used=new Set(rows.filter(r=>r.matched).map(r=>r.salesName)),unmatchedSales=[...sm.keys()].filter(k=>!used.has(k));matches=rows;stats={duplicates,unmatchedSales,store:rows.length,warehouse:d.warehouse.length,sales:d.sales.length,pending:d.pending.length,matched:rows.filter(r=>r.matched).length,missingSales:rows.filter(r=>r.salesQty===null).length,missingWarehouse:rows.filter(r=>r.warehouse===null).length,expiredLines:d.pending.filter(p=>p.isExpired).length,unknownExpiry:d.pending.filter(p=>p.expiryUnknown).length};return rows}
function calculate(){if(!matches.length)review();records=matches.map(r=>{const cv=covers[r.cat]??baseCover,rate=r.matched?r.salesQty/salesDays:null,target=rate===null?null:round(rate*cv),need=target===null?null:Math.max(0,round(target-r.qty-r.pending)),suggest=need===null||r.warehouse===null?0:Math.min(need,Math.max(r.warehouse,0)),cover=rate===null?null:rate>0?r.qty/rate:null;let flags=[];if(r.salesQty===null)flags.push('No Sales Data');else if(r.salesQty===0)flags.push('Zero sales');if(r.ambiguous)flags.push('Ambiguous name');if(r.warehouse===null)flags.push('Warehouse unavailable');else if(need!==null&&need>r.warehouse)flags.push('Warehouse shortage');if(r.expiredPending>0)flags.push('Expired previous indent');if(r.unknownPending>0)flags.push('PO expiry unverified');if(cover!==null&&cover>cv*2)flags.push('High stock');let edit=Object.hasOwn(edits,r.sku)?edits[r.sku]:suggest;return {...r,rate,target,need,suggest,cover,coverDays:cv,final:edit,flags,remark:''}});approved=false;return records}
function summary(){let d=records.length?records:calculate(),revenue=d.reduce((a,r)=>a+(r.amount||0),0),req=d.reduce((a,r)=>a+(r.need||0),0),final=d.reduce((a,r)=>a+n(r.final),0);return {revenue,req,final,low:d.filter(r=>r.qty===0).length,short:d.filter(r=>r.flags.includes('Warehouse shortage')).length,issues:d.filter(r=>r.flags.length).length}}
function heading(title,sub='',buttons=''){return `<div class="page-head"><div><h1>${esc(title)}</h1>${sub?`<p>${esc(sub)}</p>`:''}</div><div class="links">${buttons}</div></div>`}
function cards(arr){return `<div class="cards">${arr.map(x=>`<div class="metric"><small>${esc(x[0])}</small><strong>${esc(x[1])}</strong><small>${esc(x[2]||'')}</small></div>`).join('')}</div>`}
function onlyDemo(){return !complete()?`<div class="notice compact-note"><b>Report data pending.</b> The four indent datasets must be published or uploaded for verified calculations.</div>`:''}
function categories(){let x=new Map();for(let r of records){let c=x.get(r.cat)||{cat:r.cat,amount:0,qty:0,need:0,count:0};c.amount+=r.amount||0;c.qty+=r.salesQty||0;c.need+=r.need||0;c.count++;x.set(r.cat,c)}return [...x.values()].sort((a,b)=>b.amount-a.amount)}
function overview(){const s=summary(),cats=categories(),max=Math.max(1,...cats.map(x=>x.amount));return heading('Executive Dashboard','See what sold, what’s running low, and what needs attention today.',`<button class="btn" data-page="upload">Import reports ↗</button>`)+onlyDemo()+`<div class="hero"><div><h2>Know what to indent. Know why.</h2></div><button class="btn orange" data-page="indent">Open Indent Workspace →</button></div>`+cards([['Sales revenue',money(s.revenue),'Actual sales Total Amount'],['Calculated requirement',fmt(s.req),'Before warehouse cap'],['Warehouse shortage SKUs',fmt(s.short),'Demand exceeds available stock'],['Review exceptions',fmt(s.issues),'Includes missing sales and expiry issues']])+`<div class="grid2"><section class="panel"><h3>Category-wise sales</h3><p class="muted small">Actual sales report amounts, not inferred prices</p>${cats.slice(0,9).map(c=>`<div class="row-item"><div style="width:100%"><div class="flexline"><strong>${esc(c.cat)}</strong><span>${money(c.amount)}</span></div><div class="bar" style="margin-top:9px"><span style="width:${c.amount/max*100}%"></span></div></div></div>`).join('')}</section><section class="panel"><div class="panel-head"><h3>Action center</h3><button class="ghost" data-page="exceptions">View alerts →</button></div>${records.filter(r=>r.flags.length).slice(0,9).map(r=>`<div class="row-item"><div><b>${esc(r.name)}</b><div class="muted small">${esc(r.sku)} · ${esc(r.flags.join(', '))}</div></div><span class="pill warn">Review</span></div>`).join('')||'<p>No current alerts</p>'}</section></div><div class="footer">${complete()?'Calculated using the four uploaded indent reports.':'Illustrative figures · Import reports to view actual results.'}</div>`}
function upload(){return heading('Import & Match','Upload reports one at a time. Check the matching preview before calculating.')+`<div class="upload-grid">${[['store','1 · Store stock','Column B Item Code, D Category, F Quantity'],['warehouse','2 · Warehouse stock','Column B Item Code, F Quantity'],['sales','3 · Sales report','Only Item, Total Quantity, Total Amount'],['pending','4 · Previous indents','ItemCode, PendingQty, ExpiryDate']].map(([k,title,desc])=>`<div class="upload-card ${uploads[k]?'loaded':''}"><div class="glyph">${uploads[k]?'✓':'⇧'}</div><h3>${title}</h3><p>${desc}</p><label class="file-upload-action">⇧ &nbsp; Choose report<input type="file" accept=".xlsx,.xls,.csv" data-upload="${k}"></label><p>${uploads[k]?fmt(uploads[k].length)+' records ready':'No data uploaded'}</p><div class="report-stamp">${reportDateBadge(k)}</div></div>`).join('')}</div><section class="panel"><div class="panel-head"><div><h3>Total Sales Report · Data Analysis only</h3><p class="muted small">Separate from the four indent reports. Never changes required or final indent quantities.</p></div><span class="pill">ANALYTICS</span></div><div class="analytics-upload"><div><strong>⇧ Upload Total Sales Report</strong><p class="muted small">Reads CreatedDate, ItemCode, ItemName, Category, Quantity, TotalAmount and ProfitCentre.</p><p class="muted small">${analyticsMeta?fmt(analyticsRows.length)+' sales records ready':'No sales data loaded'}</p><div class="report-stamp">${reportDateBadge('analytics')}</div></div><label class="file-upload-action">⇧ &nbsp; Choose sales report<input type="file" data-upload="analytics" accept=".xlsx,.xls,.csv"></label></div><button class="ghost" data-page="dataAnalysis">Open Data Analysis →</button></section><div class="panel"><div class="panel-head"><div><h3>Analysis settings</h3><p class="muted small">Editable before each calculation</p></div><span class="pill">30-day default stock cover</span></div><div class="filters"><label>Sales period <select id="salesDays"><option value="30" ${salesDays===30?'selected':''}>30 days</option><option value="60" ${salesDays===60?'selected':''}>60 days</option></select></label><label>Global stock cover <input id="baseCover" type="number" min="1" max="365" value="${baseCover}"> days</label><button class="ghost" id="categoryCover">Configure category targets</button></div><p class="muted small">Use category overrides whenever store conditions differ. Sales report filenames do not control the selected days.</p></div><div class="panel"><div class="panel-head"><h3>Match diagnostics</h3><span class="pill ${complete()?'':'warn'}">${complete()?'All four reports present':'Demo / incomplete reports'}</span></div><div class="statusline"><span>Store SKUs: <b>${fmt(stats.store)}</b></span><span>Exact-name matches: <b>${fmt(stats.matched)}</b></span><span>No sales match: <b>${fmt(stats.missingSales)}</b></span><span>Warehouse missing: <b>${fmt(stats.missingWarehouse)}</b></span><span>Expired PO lines: <b>${fmt(stats.expiredLines)}</b></span><span>Unknown PO expiry: <b>${fmt(stats.unknownExpiry)}</b></span></div>${stats.duplicates?.length?`<div class="notice">Duplicate Item Codes detected: ${esc(stats.duplicates.slice(0,5).join('; '))} — review before approval.</div>`:''}<div class="filters"><input id="matchSearch" placeholder="Search store product / Item Code"><select id="matchFilter"><option value="all">All matches</option><option value="unmatched">Unmatched or ambiguous</option><option value="warehouse">Missing warehouse stock</option></select><button class="btn" id="calculateBtn" ${!complete()?'disabled title="Upload all four real files first"':''}>Calculate indents →</button></div><div class="table-wrap" id="matchTable">${matchTable()}</div><p class="muted small">Select an unmatched store product to manually map its sales name. Exact product-name matching is otherwise required.</p></div>`}
function matchTable(){const q=norm($('#matchSearch')?.value),f=$('#matchFilter')?.value||'all',rows=matches.filter(r=>(!q||(r.name+' '+r.sku).toLowerCase().includes(q))&&(f==='all'||f==='unmatched'&&!r.matched||f==='warehouse'&&r.warehouse===null)).slice(0,300);return `<table><thead><tr><th>Item Code</th><th>Product (Store)</th><th>Category</th><th>Store Qty</th><th>Warehouse Qty</th><th>Sales Qty</th><th>Match</th><th>Action</th></tr></thead><tbody>${rows.map(r=>`<tr><td>${esc(r.sku)}</td><td><span class="product">${esc(r.name)}</span></td><td>${esc(r.cat)}</td><td>${fmt(r.qty)} ${esc(r.uom)}</td><td>${r.warehouse===null?'—':fmt(r.warehouse)}</td><td>${r.salesQty===null?'—':fmt(r.salesQty)}</td><td><span class="pill ${r.matched?'':'warn'}">${r.matched?'Exact match':r.ambiguous?'Ambiguous':'Unmatched'}</span></td><td><button class="ghost" data-map="${esc(r.sku)}">Map</button></td></tr>`).join('')}</tbody></table>${rows.length===300?'<p class="muted small">Showing first 300 matching records; refine search to inspect others.</p>':''}`}
function tableRows(list,editable){return `<div class="table-wrap"><table><thead><tr><th>Item Code</th><th>Product</th><th>Category</th><th>60/30d Sales</th><th>Sales ₹</th><th>Store Qty</th><th>Valid Pending</th><th>WH Qty</th><th>Cover Days</th><th>Requirement</th><th>Suggested</th><th>Final Qty</th><th>Flags</th><th>Review</th></tr></thead><tbody>${list.slice(0,800).map(r=>`<tr><td>${esc(r.sku)}</td><td class="indent-name-cell"><span class="indent-product-name" title="${esc(r.name)}">${esc(r.name)}</span><span class="indent-uom">${esc(r.uom)}</span></td><td>${esc(r.cat)}</td><td>${r.salesQty===null?'No data':fmt(r.salesQty)}</td><td>${r.amount===null?'—':money(r.amount)}</td><td>${fmt(r.qty)}</td><td>${fmt(r.pending)}</td><td>${r.warehouse===null?'Unavailable':fmt(r.warehouse)}</td><td>${fmt(r.coverDays)}</td><td>${r.need===null?'Review':fmt(r.need)}</td><td>${fmt(r.suggest)}</td><td>${editable?`<input class="editqty" type="number" min="0" step="any" data-edit="${esc(r.sku)}" value="${r.final}" ${approved?'disabled':''}>`:fmt(r.final)}</td><td>${r.flags.length?r.flags.map(f=>`<span class="pill warn" style="display:inline-block;margin:2px">${esc(f)}</span>`).join(''):'<span class="pill">OK</span>'}</td><td><button class="ghost" data-detail="${esc(r.sku)}">Details</button></td></tr>`).join('')}</tbody></table>${list.length>800?'<div class="notice">Showing first 800 rows; use search and category filters to inspect remaining SKUs. All records are included in export.</div>':''}</div>`}
function filtered(){return records.filter(r=>(!filter||(r.name+' '+r.sku).toLowerCase().includes(norm(filter)))&&(indentCategory==='all'||r.cat===indentCategory)&&(tab==='all'||tab==='alerts'&&r.flags.length||tab==='positive'&&r.need>0||tab==='nosales'&&r.salesQty===null||tab==='zero'&&r.salesQty===0||tab==='shortage'&&r.flags.includes('Warehouse shortage')))}
function indent(){const s=summary();return heading('Indent Workspace','Review sales-based requirements and final replenishment quantities.',`<button class="btn" id="exportXlsx">Export detailed Excel ↓</button>`)+sourceStrip([['store','Store'],['warehouse','Warehouse'],['sales','30/60-day Sales'],['pending','Indents']])+onlyDemo()+`<div class="dev-notice"><strong>Under active development.</strong> This workspace is usable for calculations and review. Please verify suggested quantities against your store process and share feedback for improvements.</div>`+cards([['Total required',fmt(s.req),'Calculated requirement'],['Warehouse-capped suggested',fmt(records.reduce((a,r)=>a+r.suggest,0)),'Before manager adjustments'],['Final indent quantity',fmt(s.final),'Editable by manager'],['Shortage SKU count',fmt(s.short),'Warehouse below requirement']])+`<section class="panel"><div class="panel-head"><h3>Indent controls</h3><span class="pill ${approved?'':'warn'}">${approved?'Approved · manager-unlock required':'Draft · editable'}</span></div><div class="filters"><label>Sales window <select id="salesDays"><option value="30" ${salesDays===30?'selected':''}>30 days</option><option value="60" ${salesDays===60?'selected':''}>60 days</option></select></label><label>Default cover <input type="number" id="baseCover" min="1" max="365" value="${baseCover}"> days</label><button class="ghost" id="categoryCover">Category cover settings</button><button class="ghost" id="recalc">Recalculate</button><button class="btn" id="approve" ${approved?'disabled':''}>Approve & Finalize</button>${approved?'<button class="ghost" id="unlock">Manager unlock</button>':''}</div><div class="filters"><input id="indentSearch" value="${esc(filter)}" placeholder="Search Item Code or product"><span class="muted small">${fmt(filtered().length)} products</span></div><div class="indent-category-control"><label for="indentCategorySelect">Filter by category</label><select id="indentCategorySelect" aria-label="Filter indents by category"><option value="all" ${indentCategory==='all'?'selected':''}>All categories (${fmt(records.length)})</option>${[...new Set(records.map(r=>r.cat||'Uncategorised'))].sort((a,b)=>a.localeCompare(b,'en',{sensitivity:'base'})).map(c=>`<option value="${esc(c)}" ${indentCategory===c?'selected':''}>${esc(c)} (${fmt(records.filter(r=>(r.cat||'Uncategorised')===c).length)})</option>`).join('')}</select><span class="muted small">${fmt(filtered().length)} matching SKUs</span></div><div class="tabs">${[['all','All SKUs'],['positive','Positive requirement'],['alerts','Exceptions'],['shortage','Warehouse shortages'],['nosales','No sales data'],['zero','Zero sales']].map(([id,name])=>`<button data-tab="${id}" class="${tab===id?'selected':''}">${name}</button>`).join('')}</div><div id="indentRows">${tableRows(filtered(),true)}</div></section>`}
function monthLabel(month){if(!/^\d{4}-(0[1-9]|1[0-2])$/.test(month))return month;const [y,m]=month.split('-').map(Number);return new Intl.DateTimeFormat('en-IN',{month:'long',year:'numeric'}).format(new Date(y,m-1,1))}
function dataGroup(){
  const rows=analyticsRows;
  const fromReport=rows.length>0;
  const months=[...new Set(rows.map(r=>r.day?.slice(0,7)).filter(x=>/^\d{4}-(0[1-9]|1[0-2])$/.test(x)))].sort();
  const periodRows=rows.filter(r=>analysisMonth==='all'||r.day?.startsWith(analysisMonth+'-'));
  const floors=analyticsMeta?.hasFloor?[...new Set(periodRows.map(r=>r.floor).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'en',{sensitivity:'base'})):[];
  const floorRows=periodRows.filter(r=>analysisFloor==='all'||r.floor===analysisFloor);
  const categoryChoices=[...new Set(floorRows.map(r=>r.cat||'Unmapped category'))].sort((a,b)=>a.localeCompare(b,'en',{sensitivity:'base'}));
  const categoryRows=floorRows.filter(r=>analysisCategory==='all'||(r.cat||'Unmapped category')===analysisCategory);
  const subcategories=[...new Set(categoryRows.map(r=>r.subcat).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'en',{sensitivity:'base'}));
  const brands=[...new Set(categoryRows.map(r=>r.brand||'Unspecified brand'))].sort((a,b)=>a.localeCompare(b,'en',{sensitivity:'base'}));
  if(analysisBrand!=='all'&&!brands.includes(analysisBrand))analysisBrand='all';
  const passesFilter=r=>(analysisFloor==='all'||r.floor===analysisFloor)&&(analysisCategory==='all'||(r.cat||'Unmapped category')===analysisCategory)&&(analysisBrand==='all'||(r.brand||'Unspecified brand')===analysisBrand)&&(!analysisSearch||norm(r.name+' '+r.sku).includes(norm(analysisSearch)));
  const visible=periodRows.filter(passesFilter);
  const summary=new Map();
  for(const r of visible){let cat=r.cat||'Unmapped category',old=summary.get(cat)||{cat,amount:0,qty:0,skus:new Set()};old.amount+=r.amount;old.qty+=r.qty||0;old.skus.add(r.sku||r.name);summary.set(cat,old)}
  const cats=[...summary.values()].sort((a,b)=>b.amount-a.amount);
  const skuMap=new Map();
  for(const r of visible){let id=r.sku||r.name,old=skuMap.get(id)||{sku:r.sku,name:r.name,cat:r.cat||'Unmapped category',qty:0,amount:0};old.amount+=r.amount;old.qty+=r.qty||0;skuMap.set(id,old)}
  const skus=[...skuMap.values()].sort((a,b)=>b.amount-a.amount);
  const dayMap=new Map();for(const r of visible.filter(x=>x.day)){const d=dayMap.get(r.day)||{day:r.day,amount:0,qty:0};d.amount+=r.amount;d.qty+=r.qty||0;dayMap.set(r.day,d)}
  const daily=[...dayMap.values()].sort((a,b)=>a.day.localeCompare(b.day));
  const monthMap=new Map();
  for(const r of rows.filter(r=>r.day&&passesFilter(r))){const m=r.day.slice(0,7),o=monthMap.get(m)||{month:m,amount:0,qty:0,dates:new Set()};o.amount+=r.amount;o.qty+=r.qty||0;o.dates.add(r.day);monthMap.set(m,o)}
  const monthly=[...monthMap.values()].sort((a,b)=>a.month.localeCompare(b.month));
  return {fromReport,rows,visible,cats,skus,floors,months,categoryChoices,subcategories,brands,daily,monthly,dates:daily.map(x=>x.day),total:visible.reduce((x,r)=>x+r.amount,0),qty:visible.reduce((x,r)=>x+(r.qty||0),0)};
}

function dailyDetails(d){
  const available=d.daily.map(r=>r.day);
  const day=available.includes(selectedSalesDay)?selectedSalesDay:available[available.length-1];
  if(!day)return '';
  const dayRows=d.visible.filter(r=>r.day===day);
  const total=dayRows.reduce((sum,r)=>sum+r.amount,0);
  const qty=dayRows.reduce((sum,r)=>sum+(r.qty||0),0);
  const categoryMap=new Map(),skuMap=new Map();
  for(const r of dayRows){
    const cat=r.cat||'Unmapped category',c=categoryMap.get(cat)||{name:cat,amount:0,qty:0};c.amount+=r.amount;c.qty+=r.qty||0;categoryMap.set(cat,c);
    const id=r.sku||r.name,k=skuMap.get(id)||{sku:r.sku,name:r.name,cat,amount:0,qty:0};k.amount+=r.amount;k.qty+=r.qty||0;skuMap.set(id,k);
  }
  const cats=[...categoryMap.values()].sort((a,b)=>b.amount-a.amount);
  const skus=[...skuMap.values()].sort((a,b)=>b.amount-a.amount);
  const label=new Date(day+'T12:00:00').toLocaleDateString('en-IN',{day:'numeric',month:'long',year:'numeric'});
  return `<div id="dailyDetail" class="daily-detail" aria-live="polite"><div class="panel-head"><div><span class="eyebrow">SELECTED TRADING DAY</span><h3>${esc(label)}</h3><p class="muted small">Filtered by selected month, floor and category · ${fmt(dayRows.length)} sales lines</p></div><span class="pill">DAY BREAKDOWN</span></div><div class="daily-summary"><div><small>Sales revenue</small><strong>${money(total)}</strong></div><div><small>Reported quantity</small><strong>${fmt(qty)}</strong></div><div><small>Categories</small><strong>${fmt(cats.length)}</strong></div><div><small>SKUs</small><strong>${fmt(skus.length)}</strong></div></div><div class="grid2"><div><h4>Category breakdown</h4><div class="daily-detail-scroll">${cats.map(c=>`<div class="daily-category"><span>${esc(c.name)}</span><b>${money(c.amount)}</b><div class="daily-mini-track"><i style="width:${total>0?Math.max(0,c.amount/total*100):0}%"></i></div></div>`).join('')||'<p class="muted">No category data</p>'}</div></div><div><h4>Top-selling SKUs on this day</h4><div class="daily-detail-scroll"><table class="daily-sku-table"><thead><tr><th>Product</th><th>Qty</th><th>Sales</th></tr></thead><tbody>${skus.slice(0,30).map(k=>`<tr><td>${esc(k.name)}<small>${esc(k.sku||'')}</small></td><td>${fmt(k.qty)}</td><td>${money(k.amount)}</td></tr>`).join('')}</tbody></table></div>${skus.length>30?`<p class="muted small">Showing top 30 of ${fmt(skus.length)} SKUs</p>`:''}</div></div></div>`;
}
function dailyFocus(d){
  const last=d.daily.length-1;
  const day=d.dates.includes(selectedSalesDay)?selectedSalesDay:d.dates[last];
  const entry=d.daily.find(r=>r.day===day);
  if(!entry)return '';
  const idx=d.dates.indexOf(day);
  const average=d.daily.reduce((sum,r)=>sum+r.amount,0)/Math.max(1,d.daily.length);
  const deviation=average?((entry.amount-average)/Math.abs(average))*100:0;
  const label=new Date(day+'T12:00:00').toLocaleDateString('en-IN',{day:'numeric',month:'long',year:'numeric'});
  const change=(deviation>=0?'+':'')+fmt(deviation)+'%';
  return `<div class="daily-focus" id="dailyFocus" aria-live="polite">
    <div class="focus-date"><span>SELECTED DAY</span><strong>${esc(label)}</strong></div>
    <div class="focus-revenue"><span>DAY REVENUE</span><strong>${money(entry.amount)}</strong></div>
    <div class="focus-delta"><span>VS PERIOD DAILY AVERAGE</span><strong class="${deviation>=0?'up':'down'}">${change}</strong></div>
    <div class="focus-nav" aria-label="Navigate trading days"><button type="button" data-day-move="-1" ${idx<=0?'disabled':''} aria-label="Previous trading day" title="Previous day">&#8592;</button><button type="button" data-day-move="1" ${idx>=last?'disabled':''} aria-label="Next trading day" title="Next day">&#8594;</button></div>
  </div>`;
}
function dailyChart(d){
  const peak=Math.max(1,...d.daily.map(x=>Math.max(0,x.amount)));
  const day=d.dates.includes(selectedSalesDay)?selectedSalesDay:d.dates[d.dates.length-1];
  return `<div class="daily-analytics-chart">${dailyFocus(d)}
    <div class="daily-graph" role="group" aria-label="Daily sales by date; select a bar to inspect the day">
      ${d.daily.map((r,i)=>{
        const active=r.day===day;
        const label=new Date(r.day+'T12:00:00').toLocaleDateString('en-IN',{day:'numeric',month:'short'});
        const h=Math.max(5,Math.min(100,Math.max(0,r.amount)/peak*100)).toFixed(2);
        return `<button type="button" class="daily-point ${active?'active':''}" data-sales-day="${esc(r.day)}" style="--height:${h}%;--delay:${Math.min(i*18,400)}ms" aria-label="${esc(label)}: ${money(r.amount)}. View daily details" aria-pressed="${active?'true':'false'}" title="${esc(label)} · ${money(r.amount)}"><span class="daily-tooltip" aria-hidden="true">${money(r.amount)}</span><span class="daily-stick" aria-hidden="true"></span><span class="daily-date">${esc(label)}</span></button>`;
      }).join('')}
    </div><div class="daily-legend"><span><i></i> Sales by trading day</span><span><i class="selected"></i> Selected day</span><span class="daily-hint">Click a bar or use the arrows to explore</span></div></div>`;
}
function selectSalesDay(day){
  const d=currentAnalyticsView||dataGroup();
  if(!d.dates.includes(day))return;
  const changed=selectedSalesDay!==day;
  selectedSalesDay=day;
  const detail=$('#dailyDetail');if(detail)detail.outerHTML=dailyDetails(d);
  const focus=$('#dailyFocus');if(focus)focus.outerHTML=dailyFocus(d);
  $$('[data-sales-day]').forEach(el=>{
    const active=el.dataset.salesDay===day;
    el.classList.toggle('active',active);
    if(el.classList.contains('daily-point')){
      el.setAttribute('aria-pressed',String(active));
      el.classList.remove('clicked');
      if(active&&changed){void el.offsetWidth;el.classList.add('clicked');}
    }
  });
  $$('.daily-table tr').forEach(tr=>tr.classList.toggle('active',tr.querySelector('[data-sales-day]')?.dataset.salesDay===day));
}
function chartRows(data,max=18){const top=data.slice(0,max),peak=Math.max(1,...top.map(x=>x.amount));return `<div class="chart-bars">${top.map((r,i)=>`<button class="chart-bar-row" data-analytics-cat="${esc(r.cat)}" title="Inspect ${esc(r.cat)}"><div class="chart-label"><span>${esc(r.cat)}</span><b>${money(r.amount)}</b></div><div class="chart-track"><div class="chart-fill" style="--target:${Math.max(1,r.amount/peak*100).toFixed(2)}%;--delay:${Math.min(i*45,600)}ms"></div></div></button>`).join('')}</div>`}
// The Inventory Report contains GRN-linked stock-on-hand, not a complete original receipts ledger.
// Work on the uploaded file only in the browser; never embed private inventory records in GitHub Pages.
function parseInventoryGrn(rows){
  const clean=h=>norm(h).replace(/[^a-z0-9]/g,'');
  let header=-1,index={};
  for(let i=0;i<Math.min(rows.length,60);i++){
    const names=(rows[i]||[]).map(clean);
    const at=(...choices)=>names.findIndex(x=>choices.includes(x));
    const found={no:at('grn','grnnumber','grnno'),date:at('grndate','grncreateddate'),sku:at('itemcode','skucode'),name:at('itemname','productname','particulars'),qty:at('quantity','stockqty','stockquantity')};
    if(Object.values(found).every(x=>x>=0)){
      header=i;index={...found,category:at('category','categoryname'),vendor:at('vendor','vendorname','supplier'),batch:at('batchno','batchnumber'),expiry:at('expirydate','expdate'),uom:at('uom'),subcat:at('subcategory'),floor:at('profitcenter','profitcentre')};break;
    }
  }
  if(header<0)throw Error('Select the Inventory Report with GRN, GRNDate, ItemCode, ItemName and Quantity columns.');
  const lines=[],months=new Map(),allGrns=new Map();let unlinked=0,undated=0,unreadable=0;
  for(const raw of rows.slice(header+1)){
    const val=(k)=>index[k]>=0?raw[index[k]]:null;
    const sku=key(val('sku')),name=key(val('name'));
    if(!sku||!name||/^(total|grand total|subtotal)$/i.test(name))continue;
    const number=key(val('no'));
    if(!number){unlinked++;continue;}
    const date=dateOf(val('date'));
    if(!date||!Number.isFinite(date.getTime())||date.getFullYear()<2000||date.getFullYear()>2100){undated++;continue;}
    const day=[date.getFullYear(),String(date.getMonth()+1).padStart(2,'0'),String(date.getDate()).padStart(2,'0')].join('-');
    const month=day.slice(0,7),expiry=dateOf(val('expiry'));
    const item={no:number,day,month,sku,name,qty:n(val('qty')),category:key(val('category')),subcat:key(val('subcat')),vendor:key(val('vendor')),batch:key(val('batch')),uom:key(val('uom'))||'NOS',floor:key(val('floor')),expiry:expiry?`${expiry.getFullYear()}-${String(expiry.getMonth()+1).padStart(2,'0')}-${String(expiry.getDate()).padStart(2,'0')}`:''};
    lines.push(item);
    let mo=months.get(month);
    if(!mo){mo={month,days:new Map(),grns:new Map(),lines:0,qty:0};months.set(month,mo)}
    let d=mo.days.get(day);
    if(!d){d={day,grns:new Map(),lines:0,qty:0};mo.days.set(day,d)}
    let grn=d.grns.get(number);
    if(!grn){grn={no:number,day,month,items:[],vendors:new Set(),qty:0,skuSet:new Set()};d.grns.set(number,grn)}
    grn.items.push(item);grn.qty+=item.qty;grn.skuSet.add(sku);if(item.vendor)grn.vendors.add(item.vendor);
    d.lines++;d.qty+=item.qty;mo.lines++;mo.qty+=item.qty;
    mo.grns.set(day+'|'+number,grn);allGrns.set(day+'|'+number,grn);
  }
  if(!lines.length)throw Error('No valid GRN-linked inventory lines were found. Check GRN and GRNDate.');
  return {lines,months,allGrns,unlinked,undated,unreadable};
}
function grnDateLabel(day,short=false){const d=new Date(day+'T12:00:00');return Number.isNaN(d.getTime())?day:d.toLocaleDateString('en-IN',short?{day:'numeric',month:'short',year:'numeric'}:{weekday:'long',day:'numeric',month:'long',year:'numeric'});}
function grnMatches(grn,query){if(!query)return true;return norm(grn.no+' '+grn.day+' '+grn.month+' '+monthLabel(grn.month)+' '+grn.items.map(r=>r.sku+' '+r.name+' '+r.vendor+' '+r.category+' '+r.batch).join(' ')).includes(query);}
function grnSummaryCard(title,number,subtitle){return `<div class="grn-metric"><small>${esc(title)}</small><strong>${esc(number)}</strong><span>${esc(subtitle||'')}</span></div>`;}
function grnCrumbs(){
  const parts=[`<button type="button" data-grn-nav="months" class="${grnLevel==='months'?'active':''}">All months</button>`];
  if(grnMonth)parts.push(`<span aria-hidden="true">›</span><button type="button" data-grn-nav="dates" class="${grnLevel==='dates'?'active':''}">${esc(monthLabel(grnMonth))}</button>`);
  if(grnDay)parts.push(`<span aria-hidden="true">›</span><button type="button" data-grn-nav="grns" class="${grnLevel==='grns'?'active':''}">${esc(grnDateLabel(grnDay,true))}</button>`);
  if(grnNumber)parts.push(`<span aria-hidden="true">›</span><span class="grn-current">${esc(grnNumber)}</span>`);
  return `<nav class="grn-breadcrumb" aria-label="GRN report location">${parts.join('')}</nav>`;
}
function grnMonthScreen(q){
  const months=[...grnData.months.values()].sort((a,b)=>b.month.localeCompare(a.month));
  const filtered=months.map(m=>({month:m,grns:[...m.grns.values()].filter(g=>grnMatches(g,q))})).filter(o=>o.grns.length);
  return `<div class="grn-section-heading"><div><span class="eyebrow">STEP 1 OF 4</span><h3>Browse by month</h3><p class="muted small">Select a GRN month to see its dates.</p></div><span class="pill">${fmt(filtered.length)} months</span></div>
    <div class="grn-month-grid">${filtered.map(({month:m,grns})=>{
      const countDays=new Set(grns.map(g=>g.day)).size;
      const sum=grns.reduce((a,g)=>a+g.qty,0);
      return `<button type="button" class="grn-month-card" data-grn-month="${esc(m.month)}"><span class="grn-card-top"><span class="grn-card-icon">▦</span><span class="grn-card-arrow">→</span></span><strong>${esc(monthLabel(m.month))}</strong><span>${fmt(grns.length)} GRNs · ${fmt(countDays)} dates</span><small>${fmt(sum)} inventory units</small></button>`;
    }).join('')||'<div class="grn-no-results">No months match your search.</div>'}</div>`;
}
function grnDateScreen(q){
  const mo=grnData.months.get(grnMonth);if(!mo)return '<div class="grn-no-results">Select a month.</div>';
  const dates=[...mo.days.values()].sort((a,b)=>b.day.localeCompare(a.day)).map(d=>({date:d,grns:[...d.grns.values()].filter(g=>grnMatches(g,q))})).filter(x=>x.grns.length);
  return `<div class="grn-section-heading"><div><span class="eyebrow">STEP 2 OF 4</span><h3>${esc(monthLabel(grnMonth))} · Dates</h3><p class="muted small">Choose the GRN date to see all GRN numbers recorded on it.</p></div><span class="pill">${fmt(dates.length)} dates</span></div>
    <div class="grn-list">${dates.map(({date:d,grns})=>`<button type="button" class="grn-list-row" data-grn-date="${esc(d.day)}"><span class="grn-list-date-icon">${esc(d.day.slice(-2))}</span><span class="grn-list-main"><strong>${esc(grnDateLabel(d.day))}</strong><small>${fmt(grns.length)} GRN numbers · ${fmt(grns.reduce((sum,g)=>sum+g.items.length,0))} inventory lines</small></span><span class="grn-list-num">${fmt(grns.reduce((sum,g)=>sum+g.qty,0))} units</span><span class="grn-card-arrow">→</span></button>`).join('')||'<div class="grn-no-results">No dates match your search.</div>'}</div>`;
}
function grnNumberScreen(q){
  const d=grnData.months.get(grnMonth)?.days.get(grnDay);if(!d)return '<div class="grn-no-results">Choose a GRN date.</div>';
  const all=[...d.grns.values()].filter(g=>grnMatches(g,q)).sort((a,b)=>a.no.localeCompare(b.no,'en',{numeric:true}));
  return `<div class="grn-section-heading"><div><span class="eyebrow">STEP 3 OF 4</span><h3>${esc(grnDateLabel(grnDay))} · GRN Numbers</h3><p class="muted small">Select a GRN to see the products and batches associated with it.</p></div><span class="pill">${fmt(all.length)} GRNs</span></div>
    <div class="grn-list">${all.map(g=>`<button type="button" class="grn-list-row" data-grn-number="${esc(g.no)}"><span class="grn-list-icon">▣</span><span class="grn-list-main"><strong class="grn-number-text">${esc(g.no)}</strong><small>${esc([...g.vendors].slice(0,2).join(', ')||'Vendor not specified')} · ${fmt(g.skuSet.size)} SKUs</small></span><span class="grn-list-num">${fmt(g.items.length)} lines</span><span class="grn-card-arrow">→</span></button>`).join('')||'<div class="grn-no-results">No GRNs match your search.</div>'}</div>`;
}
function grnProductScreen(q){
  const g=grnData.months.get(grnMonth)?.days.get(grnDay)?.grns.get(grnNumber);if(!g)return '<div class="grn-no-results">Select a GRN number.</div>';
  const items=g.items.filter(r=>!q||norm(r.name+' '+r.sku+' '+r.category+' '+r.batch+' '+r.vendor).includes(q));
  return `<div class="grn-section-heading"><div><span class="eyebrow">STEP 4 OF 4</span><h3 class="grn-detail-no">${esc(g.no)}</h3><p class="muted small">GRN date: ${esc(grnDateLabel(g.day))}</p></div><span class="pill">${fmt(items.length)} product lines</span></div>
    <div class="grn-detail-summary"><div><small>Unique products</small><strong>${fmt(new Set(items.map(x=>x.sku)).size)}</strong></div><div><small>Inventory lines</small><strong>${fmt(items.length)}</strong></div><div><small>Inventory quantity</small><strong>${fmt(items.reduce((sum,x)=>sum+x.qty,0))}</strong></div><div><small>Vendor</small><strong class="grn-vendor">${esc([...g.vendors].join(', ')||'Not specified')}</strong></div></div>
    <div class="table-wrap grn-product-table"><table><thead><tr><th>Item Code</th><th>Product Name</th><th>Category</th><th>Batch</th><th>Expiry</th><th>Inventory Qty</th><th>UOM</th><th>Vendor</th></tr></thead><tbody>${items.map(r=>`<tr><td><span class="grn-sku">${esc(r.sku)}</span></td><td><span class="grn-item-name">${esc(r.name)}</span></td><td>${esc(r.category||'—')}</td><td>${esc(r.batch||'—')}</td><td>${r.expiry?esc(grnDateLabel(r.expiry,true)):'—'}</td><td class="grn-qty">${fmt(r.qty)}</td><td>${esc(r.uom)}</td><td>${esc(r.vendor||'—')}</td></tr>`).join('')||'<tr><td colspan="8">No products match your search.</td></tr>'}</tbody></table></div>
    <p class="grn-snapshot-footnote">Quantities shown are the remaining inventory quantities in the uploaded snapshot, grouped by their associated GRN—not verified original received quantities.</p>`;
}
function grn(){
  const data=grnData;
  return heading('Store GRN Reports','GRN history organised by month, date, number and product.')+`
    <section class="panel grn-landing"><div class="panel-head grn-top-head"><div><h3>Inventory-based GRN explorer</h3><p class="muted small">Upload your Inventory Report to browse GRN-linked products.</p></div><span class="pill">STORE GRN</span></div>
      <div class="grn-upload"><div><strong>${data?'Inventory data ready':'Select Inventory Report'}</strong><p class="muted small">${data?fmt(data.lines.length)+' GRN-linked inventory lines loaded':'Reads GRN, GRNDate, ItemCode, ItemName, Quantity, Batch and Vendor.'}</p><div class="report-stamp">${reportDateBadge('grn')}</div></div><label class="file-upload-action">⇧ ${data?'Replace inventory report':'Upload inventory report'}<input type="file" data-upload="grn" accept=".xlsx,.xls,.csv" aria-label="Upload inventory report for GRN analysis" ${grnLoading?'disabled':''}></label></div>
      ${grnImportError?`<div class="grn-error" role="alert">${esc(grnImportError)}</div>`:''}
      ${grnLoading?'<div class="grn-processing" role="status"><span class="grn-small-spinner"></span> Reading inventory and preparing your GRN report…</div>':''}
      ${data?`<div class="grn-overview-metrics">${grnSummaryCard('GRN numbers',fmt(data.allGrns.size),'Unique GRN/date combinations')}${grnSummaryCard('Months',fmt(data.months.size),'Available GRN months')}${grnSummaryCard('Linked inventory lines',fmt(data.lines.length),'Product / batch rows')}${grnSummaryCard('Unlinked lines',fmt(data.unlinked+data.undated),'Missing GRN or valid date')}</div>
      ${(data.unlinked+data.undated)>0?`<div class="grn-info"><strong>Unlinked inventory lines:</strong> ${fmt(data.unlinked)} without GRN number; ${fmt(data.undated)} without a readable GRN date. These cannot be placed in the monthly drill-down.</div>`:''}
      <div class="grn-explorer-head">${grnCrumbs()}<div class="grn-explorer-controls">${grnLevel!=='months'?'<button type="button" class="grn-back" data-grn-back="1">← Back</button>':''}<label class="grn-search-wrap"><span aria-hidden="true">⌕</span><input id="grnSearch" type="search" value="${esc(grnSearch)}" placeholder="Search GRN, SKU, product or vendor" aria-label="Search GRN data"></label></div></div>
      <div id="grnExplorer" aria-live="polite">${grnLevel==='months'?grnMonthScreen(norm(grnSearch)):grnLevel==='dates'?grnDateScreen(norm(grnSearch)):grnLevel==='grns'?grnNumberScreen(norm(grnSearch)):grnProductScreen(norm(grnSearch))}</div>
      <div class="grn-info grn-data-caution"><strong>Inventory snapshot:</strong> GRN numbers and dates come from the uploaded report. Inventory quantity is not necessarily the quantity originally received. DC number is not provided in this report.</div>`:
      `<div class="grn-empty"><span aria-hidden="true">▣</span><strong>Your GRN dashboard is ready</strong><p>Select the Inventory Report exported from Focus.Net. You can then open each month, date and GRN to see its products.</p></div>`}
    </section>`;
}
function grnUpdateSearch(){const box=$('#grnExplorer');if(!box||!grnData)return;const q=norm(grnSearch);box.innerHTML=grnLevel==='months'?grnMonthScreen(q):grnLevel==='dates'?grnDateScreen(q):grnLevel==='grns'?grnNumberScreen(q):grnProductScreen(q);}
function grnNavigate(level){
  grnLevel=level;
  if(level==='months'){grnMonth='';grnDay='';grnNumber='';}
  if(level==='dates'){grnDay='';grnNumber='';}
  if(level==='grns')grnNumber='';
  grnSearch='';render();
}
// Stock Requirements uses the union of actual store and warehouse SKU lists, never analytics data.
let stockReqTab='all',stockReqSearch='';
function stockRequirementsData(){
  if(!uploads.store || !uploads.warehouse)return null;
  const combined=new Map();
  for(const row of uploads.store){
    const sku=key(row.sku);if(!sku)continue;
    const item=combined.get(sku)||{sku,name:row.name||sku,cat:row.cat||'Uncategorized',store:0,warehouse:0,inStore:false,inWarehouse:false};
    item.store+=n(row.qty);item.inStore=true;
    if(row.name)item.name=row.name;
    if(row.cat)item.cat=row.cat;
    combined.set(sku,item);
  }
  for(const row of uploads.warehouse){
    const sku=key(row.sku);if(!sku)continue;
    const item=combined.get(sku)||{sku,name:row.name||sku,cat:row.cat||'Uncategorized',store:0,warehouse:0,inStore:false,inWarehouse:false};
    item.warehouse+=n(row.qty);item.inWarehouse=true;
    if(!item.name&&row.name)item.name=row.name;
    if(!item.cat&&row.cat)item.cat=row.cat;
    combined.set(sku,item);
  }
  return [...combined.values()].map(r=>({...r,status:r.store<=0&&r.warehouse<=0?'both':r.warehouse<=0&&r.store>0?'storeonly':r.store<=0&&r.warehouse>0?'whonly':'available'}))
    .filter(r=>r.status!=='available')
    .sort((a,b)=>a.name.localeCompare(b.name,'en',{sensitivity:'base'})||a.sku.localeCompare(b.sku));
}
function stockRequirements(){
  const data=stockRequirementsData();
  if(!data)return heading('Stock Requirements','Store and warehouse stock availability.')+`<section class="panel"><h3>Upload stock reports to begin</h3><p class="muted small">This section only uses your actual Store Stock Balance and Warehouse Stock Balance Excel files. Both are required.</p><button type="button" class="btn" data-page="upload">Go to Import & Match →</button></section>`;
  const counts={all:data.length,both:0,storeonly:0,whonly:0};
  for(const item of data)counts[item.status]++;
  const tabs=[['all','All requirements'],['both','No stock in either'],['storeonly','Warehouse empty · store has stock'],['whonly','Store empty · warehouse has stock']];
  const visible=data.filter(item=>(stockReqTab==='all'||item.status===stockReqTab)&&(!stockReqSearch||norm(item.name+' '+item.sku+' '+item.cat).includes(norm(stockReqSearch))));
  const labelByStatus={both:'No stock in either',storeonly:'Warehouse empty',whonly:'Store empty'};
  return heading('Stock Requirements','Find items that are unavailable in either the store, warehouse, or both.')+sourceStrip([['store','Store Stock'],['warehouse','Warehouse Stock']])+
    `<div class="cards">${[['Both out of stock',counts.both,'Store 0 · Warehouse 0'],['Warehouse empty',counts.storeonly,'Store has stock'],['Store empty',counts.whonly,'Warehouse has stock'],['Total requirements',counts.all,'Distinct stock-affected SKUs']].map(([title,value,subtitle])=>`<div class="metric"><small>${esc(title)}</small><strong>${fmt(value)}</strong><small>${esc(subtitle)}</small></div>`).join('')}</div>`+
    `<section class="panel"><div class="stock-requirements-head"><div><h3>Stock availability</h3><p class="muted small">Based on your latest imported store and warehouse stock reports.</p></div><input class="stock-requirements-search" id="stockReqSearch" type="search" value="${esc(stockReqSearch)}" placeholder="Search product, item code or category" aria-label="Search stock requirements"></div>`+
    `<div class="stock-req-tabs">${tabs.map(([id,label])=>`<button type="button" class="stock-req-tab ${stockReqTab===id?'selected':''}" data-stock-req-tab="${id}" aria-pressed="${stockReqTab===id}">${esc(label)} <small>${fmt(counts[id])}</small></button>`).join('')}</div>`+
    `<div class="table-wrap"><table class="stock-req-table"><thead><tr><th>Item Code</th><th>Product</th><th>Category</th><th>Store Qty</th><th>Warehouse Qty</th><th>Status</th></tr></thead><tbody id="stockReqRows">${stockRequirementsRows(visible)}</tbody></table></div>`+
    `<p class="stock-req-note">Zero or negative stock is treated as unavailable. If an item is absent from one report, that location is shown as 0 (not listed). SKUs absent from <em>both</em> reports cannot be identified.</p></section>`;
}
function stockRequirementsRows(data){
  return data.slice(0,1500).map(r=>`<tr><td>${esc(r.sku)}</td><td><b>${esc(r.name)}</b></td><td>${esc(r.cat)}</td><td>${fmt(r.store)}${r.inStore?'':' <small class="muted">(not listed)</small>'}</td><td>${fmt(r.warehouse)}${r.inWarehouse?'':' <small class="muted">(not listed)</small>'}</td><td><span class="stock-req-state ${r.status}">${esc({both:'Both empty',storeonly:'Warehouse empty',whonly:'Store empty'}[r.status])}</span></td></tr>`).join('')||'<tr><td colspan="6">No matching stock requirements.</td></tr>';
}

function dataAnalysis(){
  const d=dataGroup(),floorName=analysisFloor==='all'?'Entire store':analysisFloor,catName=analysisCategory==='all'?'All categories':analysisCategory,monthName=analysisMonth==='all'?'All months':monthLabel(analysisMonth);
  currentAnalyticsView=d;
  return heading('Data Analysis')+sourceStrip([['analytics','Total Sales']])+`${!d.fromReport?'<div class="notice compact-note">Upload the Total Sales Report to view actual floor, category and SKU performance.</div>':''}
  <section class="panel smart-filter-panel"><div class="panel-head"><div><h3>Smart Filters</h3><p class="muted small">A single set of filters for all charts and daily sales.</p></div><button class="ghost" id="resetAnalytics" type="button">↺ Reset all</button></div>
  <div class="smart-filter-grid">
    <label class="smart-filter-field"><span>Month</span><select id="smartMonth"><option value="all">All months</option>${d.months.map(m=>`<option value="${esc(m)}" ${analysisMonth===m?'selected':''}>${esc(monthLabel(m))}</option>`).join('')}</select></label>
    <label class="smart-filter-field"><span>Floor</span><select id="smartFloor"><option value="all">Entire store</option>${d.floors.map(f=>`<option value="${esc(f)}" ${analysisFloor===f?'selected':''}>${esc(f)}</option>`).join('')}</select></label>
    <div class="smart-filter-field smart-category-field"><label for="analysisCategorySelect">Category · A–Z</label><input type="search" id="analysisCategoryFilter" placeholder="Search categories" value="${esc(analysisCategoryQuery)}" aria-label="Search categories alphabetically"><select id="analysisCategorySelect" aria-label="Select category"><option value="all">All categories</option>${d.categoryChoices.map(cat=>`<option value="${esc(cat)}" ${analysisCategory===cat?'selected':''}>${esc(cat)}</option>`).join('')}</select></div>
    <label class="smart-filter-field"><span>Brand</span><select id="smartBrand" aria-label="Filter by brand"><option value="all">All brands</option>${d.brands.map(b=>`<option value="${esc(b)}" ${analysisBrand===b?'selected':''}>${esc(b)}</option>`).join('')}</select></label>
    <label class="smart-filter-field"><span>SKU / Product</span><input id="analysisSearch" type="search" value="${esc(analysisSearch)}" placeholder="Name or item code" aria-label="Search SKU or product"></label>
  </div>
  <div class="filter-active-line" aria-live="polite"><span class="filter-active-dot"></span><span>${fmt(d.visible.length)} transactions · ${fmt(d.skus.length)} SKUs · ${fmt(d.cats.length)} categories</span><span class="filter-context">${esc(monthName)} · ${esc(floorName)}</span></div>
  </section>
  ${cards([['Sales value',money(d.total),monthName+' · '+floorName+' · '+catName],['Quantity sold',fmt(d.qty),'Mixed UOM: interpret with care'],['SKUs',fmt(d.skus.length),'Visible distinct product lines'],['Categories',fmt(analysisCategory==='all'?d.cats.length:Math.min(1,d.cats.length)),'Selected month and floor']])}
  <section class="panel"><div class="panel-head"><div><h3>${analysisCategory==='all'?'Category overview · top 8 by revenue':esc(analysisCategory)+' · sales performance'}</h3><p class="muted small">${analysisCategory==='all'?'Select a category from the filter above to inspect its SKU results.':'Revenue and contribution for the selected category.'}</p></div><span class="pill">${analysisCategory==='all'?fmt(d.cats.length)+' available categories':'1 selected category'}</span></div>${analysisCategory==='all'?chartRows(d.cats,8):chartRows(d.cats.filter(x=>x.cat===analysisCategory),1)}${analysisCategory==='all'&&d.cats.length>8?'<p class="muted small">Only the top 8 are shown here to keep the dashboard clean. Use Filter by Category to access every category.</p>':''}</section>
  <section class="panel"><div class="panel-head"><div><h3>Monthly sales performance</h3><p class="muted small">Detected automatically from CreatedDate · Month names and revenue comparisons</p></div><span class="pill">${fmt(d.months.length)} months detected</span></div>${d.monthly.length?`<div class="monthly-grid">${d.monthly.map((r,i)=>{const previous=i?d.monthly[i-1]:null;const growth=previous&&previous.amount!==0?(r.amount-previous.amount)/Math.abs(previous.amount)*100:null;return `<button class="month-stat ${analysisMonth===r.month?'selected':''}" data-analytics-month="${r.month}"><span class="month-heading">${esc(monthLabel(r.month))}</span><strong>${money(r.amount)}</strong><small>${fmt(r.dates.size)} trading dates · ${fmt(r.qty)} units</small><span class="month-change ${growth===null?'neutral':growth>=0?'positive':'negative'}">${growth===null?'First available month':`${growth>=0?'+':''}${fmt(growth)}% vs previous month`}</span></button>`}).join('')}</div><p class="muted small">Comparisons are between available report periods; partial months are not directly comparable to full months.</p>`:'<div class="analytics-missing">Month-wise analysis requires CreatedDate in the Total Sales Report.</div>'}</section>
   <section class="panel"><div class="panel-head"><div><h3>Daily sales trend</h3><p class="muted small">Click any date bar to inspect that day's actual sales, categories and SKUs.</p></div><span class="pill ${d.dates.length?'':'warn'}">${d.dates.length?fmt(d.dates.length)+' dates':'Date data unavailable'}</span></div>${d.daily.length?`${dailyChart(d)}${dailyDetails(d)}<div class="table-wrap daily-table"><table style="min-width:520px"><thead><tr><th>Date</th><th>Sales Value</th><th>Reported Qty</th><th>View</th></tr></thead><tbody>${d.daily.map(r=>`<tr class="${((selectedSalesDay&&d.dates.includes(selectedSalesDay)?selectedSalesDay:d.dates[d.dates.length-1])===r.day)?'active':''}"><td>${esc(r.day)}</td><td>${money(r.amount)}</td><td>${fmt(r.qty)}</td><td><button type="button" class="daily-view-btn" data-sales-day="${esc(r.day)}">Inspect day →</button></td></tr>`).join('')}</tbody></table></div>`:'<div class="analytics-missing">Day-by-day figures cannot be generated from a 30/60-day total. Please upload dated transactions or daily summary data for a valid daily trend.</div>'}</section>
  <section class="panel"><div class="panel-head"><div><h3>SKU performance · ${esc(catName)}</h3><p class="muted small">Sales value is recorded Total Amount, not estimated inventory value.</p></div><span class="pill">${fmt(d.skus.length)} SKUs</span></div><div class="table-wrap"><table style="min-width:700px"><thead><tr><th>Product</th><th>Item Code</th><th>Category</th><th>Qty sold</th><th>Sales Amount</th><th>Sales mix</th></tr></thead><tbody id="analysisSkuRows">${analysisSkuRows(d.skus,d.total)}</tbody></table></div></section>
  <section class="panel"><div class="panel-head"><div><h3>Total Sales Report upload</h3><p class="muted small">Dedicated analytics report · never used for indents</p></div><span class="pill">DATA ANALYSIS ONLY</span></div><div class="analytics-upload"><div><strong>Upload Total Sales Report (.xlsx)</strong><p class="muted small">Reads ProfitCentre, Category, ItemCode, ItemName, Quantity, TotalAmount and CreatedDate.</p>${analyticsMeta?`<p class="muted small">${fmt(analyticsRows.length)} transactions ready</p>`:'<p class="muted small">No sales report loaded.</p>'}<div class="report-stamp">${reportDateBadge('analytics')}</div></div><label class="file-upload-action">⇧ &nbsp; Choose sales report<input type="file" data-upload="analytics" accept=".xlsx,.xls,.csv"></label></div>${analyticsMeta?'<button class="ghost" id="clearAnalytics">Remove Total Sales Report</button>':''}</section>`;
}
function analysisSkuRows(rows,total){return rows.slice(0,700).map(r=>`<tr><td><span class="product">${esc(r.name)}</span></td><td>${esc(r.sku||'Not mapped')}</td><td>${esc(r.cat)}</td><td>${fmt(r.qty)}</td><td><b>${money(r.amount)}</b></td><td>${total?fmt(r.amount/total*100):'0'}%</td></tr>`).join('')||'<tr><td colspan="6">No matching products</td></tr>'}
function exceptions(){let arr=records.filter(r=>r.flags.length),types=[...new Set(arr.flatMap(x=>x.flags))];return heading('Exceptions & Reviews','Missing data, shortages, expired POs, excessive stock and manual decisions.')+onlyDemo()+cards([['Flagged SKUs',fmt(arr.length),'Require attention'],['Missing sales',fmt(records.filter(x=>x.salesQty===null).length),'Not treated as zero sales'],['Warehouse missing',fmt(records.filter(x=>x.warehouse===null).length),'Not silently treated as available'],['Expired PO lines',fmt(stats.expiredLines),'Excluded from pending deduction']])+`<div class="panel"><h3>Exceptions by type</h3>${types.map(f=>`<div class="row-item"><b>${esc(f)}</b><span>${fmt(arr.filter(r=>r.flags.includes(f)).length)} SKUs</span></div>`).join('')||'<p>No exceptions.</p>'}</div><div class="panel"><h3>Flagged products</h3>${tableRows(arr,false)}</div>`}
function hist(){return heading('Daily History','Saved local versions and manager approvals. Restore a prior snapshot.')+`<div class="notice">This local application stores versions in your browser using IndexedDB. Clearing browser data or switching devices may remove them. This is not an off-device daily backup or secure role enforcement.</div><section class="panel"><div class="panel-head"><h3>Saved versions</h3></div>${history.length?history.map(r=>`<div class="row-item"><div><b>${esc(r.date)} · Version ${r.version}</b><div class="muted small">${esc(r.created)} · ${r.approved?'Approved':'Draft'} · ${fmt(r.records.length)} SKUs</div></div><button class="ghost" data-restore="${r.id}">Restore snapshot</button></div>`).join(''):'<div class="empty">No saved versions yet.</div>'}</section>`}
// The publishable key is deliberately browser-safe. The service_role/secret key NEVER belongs here.
const SB_URL=(window.TB_API_GATEWAY_URL || 'https://kggypoblcagwqqtshixd.supabase.co').replace(/\/+$/, '');
const SB_PUBLISHABLE_KEY='sb_publishable_R19AkJP-laREZyeSf6LkWw_8YTb_UW3';
const REPORT_BUCKET='todays-basket-private-reports';
const EMPLOYEE_BUCKET='todays-basket-employee-reports';
const REPORT_KIND={inventory:'grn',annual_sales:null,total_sales:'analytics',sales_30_days:'sales',store_stock:'store',warehouse_stock:'warehouse',pending_indents:'pending'};
let sb=null,staff=null,serverSession=null,latestReports={},reportTypes=[],reportLoading=false,officialUploadInProgress=false,loginInProgress=false,sessionGeneration=0,initialRestoreDone=false;
let personalReports={},reportSources={},personalSaving=false,storageStats=[],cleanupWorking=false,cleanupStatus="",cleanupIssues=0;
const KEY_FOR_KIND=Object.fromEntries(Object.entries(REPORT_KIND).filter(([,k])=>k).map(([r,k])=>[k,r]));
const fmtTimestamp=ts=>{if(!ts)return '—';const d=new Date(ts);return Number.isFinite(d.getTime())?new Intl.DateTimeFormat('en-IN',{timeZone:'Asia/Kolkata',day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit',hour12:true}).format(d)+' IST':'—';};
const reportDateBadge=kind=>{const r=reportSources[kind];return r?`<span class="report-date-tag">${esc(r.origin==='personal'?'My report':'Official report')} · Uploaded ${esc(fmtTimestamp(r.uploaded_at))}</span>`:`<span class="report-date-tag empty">No report loaded</span>`;};
const sourceStrip=kinds=>`<div class="report-source-strip">${kinds.map(([kind,label])=>`<span class="report-source-pair"><b>${esc(label)}:</b> ${reportDateBadge(kind)}</span>`).join('')}</div>`;
const loginPause=ms=>new Promise(resolve=>setTimeout(resolve,ms));
function setupSupabase(){
  if(window.supabase?.createClient){
    sb=window.supabase.createClient(SB_URL,SB_PUBLISHABLE_KEY,{auth:{persistSession:true,autoRefreshToken:true,detectSessionInUrl:true}});
    sb.auth.onAuthStateChange((event,session)=>{
      if(event==='TOKEN_REFRESHED'&&staff)serverSession=session;
      if(event==='PASSWORD_RECOVERY'){passwordRecoverySession=true;queueMicrotask(()=>showPasswordDialog('recovery'));}
      if(event==='SIGNED_OUT'&&initialRestoreDone){
        sessionGeneration++;staff=null;serverSession=null;latestReports={};personalReports={};reportSources={};uploads={};meta={};analyticsRows=[];analyticsMeta=null;grnData=null;
        review();calculate();page='overview';render();applyLoginState();
        const bar=$('#syncBar');if(bar)bar.hidden=true;
      }
    });
    return true;
  }
  const error=document.getElementById('loginError');if(error)error.textContent='Secure login library could not load. Connect to the internet and refresh.';
  return false;
}
function applyLoginState(){
  const active=Boolean(staff&&serverSession);
  document.body.classList.toggle('login-locked',!active);
  const portal=document.getElementById('loginPortal');if(portal)portal.hidden=active;
  const id=document.getElementById('signedInId'),name=document.getElementById('signedInName');
  if(id)id.textContent=staff?.employee_id||'—';
  if(name)name.textContent=staff?.full_name||'Employee';
  const avatar=document.querySelector('.identity-avatar');
  if(avatar)avatar.textContent=(staff?.full_name||'Employee').split(/\s+/).slice(0,2).map(x=>x[0]).join('').toUpperCase();
  const adminNav=document.getElementById('reportManagementNav');if(adminNav)adminNav.hidden=staff?.role!=='admin';const employeeNav=document.getElementById('employeeManagementNav');if(employeeNav)employeeNav.hidden=!isEmployeeAdministrator();
  if(!active)setTimeout(()=>document.getElementById('employeeId')?.focus(),40);
}
async function fetchProfile(user){
  const {data,error}=await sb.from('staff_profiles').select('user_id,employee_id,full_name,role,active').eq('user_id',user.id).single();
  if(error||!data?.active)throw Error('This account is not approved for the store dashboard. Please contact the administrator.');
  return data;
}
async function signInEmployee(event){
  event.preventDefault();if(loginInProgress)return;
  const identifier=$('#employeeId').value.trim().toLowerCase(),password=$('#employeePassword').value,err=$('#loginError'),submit=$('#loginForm .login-submit');
  loginInProgress=true;err.textContent='';submit.disabled=true;submit.classList.add('signing-in');
  submit.querySelector('.login-submit-label').textContent='Signing in';
  try{
    if(!sb)throw Error('Supabase has not loaded. Check your internet connection.');
    if(!/^[a-z0-9._-]{3,64}$/.test(identifier))throw Error('Enter a valid Employee ID or username.');
    // This public endpoint validates the actual password server-side with Supabase Auth.
    // Neither employee emails nor password mappings are exposed in the public HTML.
    const response=await fetch(`${SB_URL}/functions/v1/employee-login`,{
      method:'POST',
      headers:{'Content-Type':'application/json','apikey':SB_PUBLISHABLE_KEY},
      body:JSON.stringify({identifier,password}),
      cache:'no-store'
    });
    let result={};try{result=await response.json()}catch(_){}
    if(!response.ok)throw Error(response.status===429?'Too many login attempts. Please try again in 15 minutes.':response.status===503?'Login service is temporarily unavailable. Please try again later.':'Incorrect Employee ID/username or password.');
    if(!result.access_token||!result.refresh_token)throw Error('Login service did not return a valid session.');
    const {data,error}=await sb.auth.setSession({access_token:result.access_token,refresh_token:result.refresh_token});
    if(error||!data.user||!data.session)throw Error('Secure login could not be completed. Please try again.');
    const profile=await fetchProfile(data.user);
    staff=profile;serverSession=data.session;
    $('#loginForm').reset();applyLoginState();go('overview');
    refreshOfficialReports({autoImport:true}).catch(ex=>{console.error(ex);notifySync('Official reports are unavailable. Try Report Management or refresh.',true)});
  }catch(ex){err.textContent=ex.message||'Secure login failed.';staff=null;serverSession=null;if(sb)await sb.auth.signOut().catch(()=>{});applyLoginState();}
  finally{loginInProgress=false;submit.disabled=false;submit.classList.remove('signing-in');submit.querySelector('.login-submit-label').textContent='Sign in';}
}
async function logoutEmployee(){
  sessionGeneration++;
  if(sb){const {error}=await sb.auth.signOut();if(error){flash('Sign-out failed: '+error.message);return;}}
  staff=null;serverSession=null;latestReports={};personalReports={};reportSources={};uploads={};meta={};analyticsRows=[];analyticsMeta=null;grnData=null;
  review();calculate();page='overview';render();applyLoginState();
  const bar=$('#syncBar');if(bar)bar.hidden=true;
}
let passwordRecoverySession=false;
async function restoreSession(){
  if(!sb)return;
  try{
    const {data,error}=await sb.auth.getSession();if(error)throw error;
    if(data.session?.user && passwordRecoverySession){showPasswordDialog('recovery');return;}
    if(data.session?.user){const profile=await fetchProfile(data.session.user);staff=profile;serverSession=data.session;applyLoginState();go('overview');
      refreshOfficialReports({autoImport:true}).catch(ex=>{console.error(ex);notifySync('Unable to load official reports. Please retry.',true)});
    }else applyLoginState();
  }catch(ex){console.error('Session restoration failed',ex);staff=null;serverSession=null;applyLoginState();}
  finally{initialRestoreDone=true;}
}
function searchCategoryOptions(term){
  const select=document.getElementById('analysisCategorySelect');
  if(!select)return;
  const entries=(currentAnalyticsView?.categoryChoices||[]).slice();
  const search=term.trim().toLocaleLowerCase();
  const filtered=entries.filter(c=>c.toLocaleLowerCase().includes(search));
  const selected=analysisCategory;
  select.innerHTML='';
  const all=document.createElement('option');
  all.value=search?'':'all';
  all.textContent=search?'Choose a matching category':'All categories';
  all.disabled=!!search;
  select.appendChild(all);
  for(const row of filtered){
    const opt=document.createElement('option');
    opt.value=row;
    opt.textContent=row;
    select.appendChild(opt);
  }
  select.value=filtered.includes(selected)?selected:search?'':'all';
  const counter=document.getElementById('categoryMatchCount');
  if(counter)counter.textContent=fmt(filtered.length)+' categories · A–Z';
}
const views={overview,stockRequirements,grn,upload,indent,dataAnalysis,exceptions,reportManagement,myReports,employeeManagement},names={overview:'Executive Dashboard',stockRequirements:'Stock Requirements',grn:'GRN Report',upload:'Import & Match',indent:'Indent Workspace',dataAnalysis:'Data Analysis',exceptions:'Exceptions',reportManagement:'Report Management',myReports:'My Reports',employeeManagement:'Employee Management'};
function setOverviewMenu(open){const menu=$('#overviewMenu'),panel=$('#overviewSubmenu');if(!menu||!panel)return;panel.hidden=!open;menu.setAttribute('aria-expanded',String(open));menu.classList.toggle('expanded',open)}
function setGrnMenu(open){const menu=$('#grnMenu'),panel=$('#grnSubmenu');if(!menu||!panel)return;panel.hidden=!open;menu.setAttribute('aria-expanded',String(open));menu.classList.toggle('expanded',open)}
function render(){if(!records.length)calculate();if(page==='grn')setGrnMenu(true);if(page==='stockRequirements')setOverviewMenu(true);$$('nav button').forEach(x=>x.classList.toggle('selected',x.dataset.page===page));$('#breadcrumb').textContent=names[page];const idNode=$('#signedInId');if(idNode)idNode.textContent=staff?.employee_id||'—';$('#view').innerHTML=views[page]();if(page==='dataAnalysis'&&analysisCategoryQuery)searchCategoryOptions(analysisCategoryQuery);document.querySelector('.side').classList.remove('open')}
function go(p){if(p==='reportManagement'&&staff?.role!=='admin')return showAccessRestricted();if(p==='employeeManagement'&&!isEmployeeAdministrator())return showAccessRestricted();if(views[p]){page=p;render();document.querySelector('.main-shell').scrollTop=0}}


function notifySync(message,warning=false){
 const bar=$('#syncBar');if(!bar)return;bar.textContent=message;bar.classList.toggle('sync-warning',warning);bar.hidden=!message;
}
function isEmployeeAdministrator(){return !!(staff?.role==='admin' && staff?.employee_id==='15907' && serverSession?.user?.id && staff.user_id===serverSession.user.id)}
function restrictedPanel(){return `<section class="panel employee-restricted" role="alert"><span class="restricted-symbol" aria-hidden="true">♧</span><h2>Access Restricted</h2><p>You don't have permission to access this section. Please contact your administrator if you need access.</p><button type="button" class="btn" data-page="overview">← Back to Dashboard</button></section>`}
function showAccessRestricted(){page='restrictedAccess';document.getElementById('breadcrumb').textContent='Access Restricted';document.getElementById('view').innerHTML=restrictedPanel();document.querySelector('.side')?.classList.remove('open')}
let teamRows=[],employeeManagementLoading=false,employeeManagementError='',employeeManagementNotice='',employeeSaving=false;
async function loadEmployees(){
 if(!isEmployeeAdministrator())return showAccessRestricted();
 employeeManagementLoading=true;employeeManagementError='';render();
 const {data,error}=await sb.from('staff_profiles').select('user_id,employee_id,full_name,designation,role,active').order('employee_id');
 employeeManagementLoading=false;
 if(error)employeeManagementError='Unable to load employees: '+error.message;
 else teamRows=data||[];
 if(page==='employeeManagement')render();
}
function employeeManagement(){
 if(!isEmployeeAdministrator())return restrictedPanel();
 const list=teamRows||[];
 return heading('Employee Management','Employee accounts and store access.')+`<div class="employee-summary"><div><b>${list.length}</b><span>Total staff</span></div><div><b>${list.filter(x=>x.active).length}</b><span>Active</span></div><div><b>${list.filter(x=>!x.active).length}</b><span>Inactive</span></div></div>
 <section class="panel"><div class="panel-head"><div><h3>Staff directory</h3><p class="muted small">Only Employee 15907 can manage staff accounts.</p></div><button type="button" class="ghost" id="employeeReload">↻ Refresh</button></div>
 ${employeeManagementError?`<div class="notice" role="alert">${esc(employeeManagementError)}</div>`:''}${employeeManagementNotice?`<div class="notice" role="status">${esc(employeeManagementNotice)}</div>`:''}
 ${employeeManagementLoading?'<p class="muted">Loading employees…</p>':`<div class="table-wrap"><table class="employee-directory"><thead><tr><th>Employee</th><th>Employee ID</th><th>Designation</th><th>Permission</th><th>Status</th><th>Action</th></tr></thead><tbody>${list.map(x=>`<tr><td><b>${esc(x.full_name)}</b></td><td>${esc(x.employee_id)}</td><td>${esc(x.designation||'—')}</td><td>${esc(x.role==='admin'?'Administrator':x.role==='manager'?'Manager':'Employee')}</td><td><span class="pill ${x.active?'':'warn'}">${x.active?'Active':'Inactive'}</span></td><td>${x.employee_id==='15907'?'Protected administrator':`<button type="button" class="ghost" data-employee-edit="${esc(x.user_id)}">Manage</button>`}</td></tr>`).join('')||'<tr><td colspan="6">No staff profiles found.</td></tr>'}</tbody></table></div>`}
 </section><section class="panel"><div class="panel-head"><h3>Register an employee</h3></div><p class="muted small">Enter the employee details. Existing Supabase users will be linked; new employees receive an email invitation to set their own password.</p>
 <form id="employeeRegisterForm" class="employee-edit-form"><label>Employee ID<input name="employee_id" inputmode="numeric" maxlength="15" required placeholder="e.g. 15270"></label><label>Employee name<input name="full_name" maxlength="120" required placeholder="Full name"></label><label>Registered email<input name="email" type="email" maxlength="254" required placeholder="employee@example.com"></label><label>Designation<input name="designation" maxlength="100" placeholder="e.g. Store Associate"></label><button type="submit" class="btn" ${employeeSaving?'disabled':''}>${employeeSaving?'Saving…':'Link employee account'}</button></form><p class="muted small">A verified Supabase account is required. An employee password is never saved in this page.</p></section>`;
}
function editEmployee(id){
 if(!isEmployeeAdministrator())return showAccessRestricted();
 const x=teamRows.find(y=>y.user_id===id);if(!x||x.employee_id==='15907')return;
 const box=document.createElement('div');box.className='employee-edit-overlay';box.id='employeeEditOverlay';box.innerHTML=`<form id="employeeEditForm" class="employee-edit-dialog"><h2>Manage employee</h2><p class="muted">${esc(x.full_name)} · ${esc(x.employee_id)}</p><input type="hidden" name="user_id" value="${esc(x.user_id)}"><label>Name<input required name="full_name" maxlength="120" value="${esc(x.full_name)}"></label><label>Designation<input name="designation" maxlength="100" value="${esc(x.designation||'')}"></label><label>Access level<select name="role"><option value="employee" ${x.role==='employee'?'selected':''}>Employee</option><option value="manager" ${x.role==='manager'?'selected':''}>Manager</option></select></label><label class="employee-active"><input type="checkbox" name="active" ${x.active?'checked':''}> Active account</label><p class="muted small">Inactive accounts cannot start new sessions. Existing sessions may need separate revocation in Supabase.</p><div class="employee-edit-actions"><button type="button" class="ghost" id="employeeEditCancel">Cancel</button><button type="submit" class="btn">Save changes</button></div></form>`;
 document.body.appendChild(box);
}
async function saveEmployeeForm(form){
 if(!isEmployeeAdministrator())return showAccessRestricted();
 const fd=new FormData(form),id=String(fd.get('user_id')||'');
 const profile=teamRows.find(x=>x.user_id===id);
 if(!profile||profile.employee_id==='15907')return;
 const full_name=String(fd.get('full_name')||'').trim(),designation=String(fd.get('designation')||'').trim(),role=String(fd.get('role')||'');
 if(!full_name||!['employee','manager'].includes(role))return flash('Check the employee details.');
 employeeSaving=true;
 const {error}=await sb.from('staff_profiles').update({full_name,designation,role,active:fd.has('active')}).eq('user_id',id).eq('employee_id',profile.employee_id);
 employeeSaving=false;
 if(error){flash('Update failed: '+error.message);return;}
 $('#employeeEditOverlay')?.remove();employeeManagementNotice='Employee updated successfully.';await loadEmployees();
}
async function registerEmployee(form){
 if(!isEmployeeAdministrator())return showAccessRestricted();
 const fd=new FormData(form);
 const payload={employee_id:String(fd.get('employee_id')||'').trim(),full_name:String(fd.get('full_name')||'').trim(),email:String(fd.get('email')||'').trim(),designation:String(fd.get('designation')||'').trim()};
 if(!/^\d{3,15}$/.test(payload.employee_id)||!payload.full_name||!payload.email.includes('@'))return flash('Check the employee details.');
 employeeSaving=true;
 try{
   const {data,error}=await sb.functions.invoke('register-employee',{body:payload});
   if(error)throw Error(error.message);
   if(data?.error)throw Error(data.error);
   employeeManagementNotice=data?.invited?'Invitation sent for employee '+payload.employee_id+'.':'Employee '+payload.employee_id+' linked successfully.';
   form.reset();await loadEmployees();
 }catch(e){employeeManagementError='Could not register employee: '+e.message;render();}
 finally{employeeSaving=false;}
}
function reportManagement(){
  if(staff?.role!=='admin')return heading('Access restricted','Only approved administrators can publish official store reports.');
  const available=reportTypes.length?reportTypes:[
   {key:'inventory',title:'Inventory Report'}, {key:'annual_sales',title:'One-Year Sales Report'},
   {key:'total_sales',title:'Total Sales Report'}, {key:'sales_30_days',title:'30-Day Summary Sales by Quantity'},
   {key:'store_stock',title:'Store Stock Balance'}, {key:'warehouse_stock',title:'Warehouse Stock Balance'},
   {key:'pending_indents',title:'Store Indent vs Material Received'}];
  const rows=available.map(row=>{
   const cur=latestReports[row.key],updated=cur?.uploaded_at?fmtTimestamp(cur.uploaded_at):'—';
   const localLoaded=Boolean(row.key==='annual_sales'?false:row.key==='inventory'?grnData:row.key==='total_sales'?analyticsRows.length:uploads[REPORT_KIND[row.key]]);
   return `<tr><td><strong>${esc(row.title)}</strong><small class="manage-description">${esc(row.key==='annual_sales'?'Stored securely; annual analytics is planned for a later version.':row.key==='total_sales'?'Dedicated to Data Analysis—not indent calculations.':row.key==='sales_30_days'?'Used for indent calculations only.':'')}</small></td><td><span class="report-state ${cur?'ok':''}">${cur?'Published':'Not published'}</span></td><td>${esc(updated)}</td><td><span class="report-state ${localLoaded?'ok':''}">${localLoaded?'Loaded':'—'}</span></td></tr>`;
  }).join('');
  return heading('Report Management','Publish official Excel reports securely for all authorized employees.',`<button class="btn" id="refreshReports" type="button">↻ Refresh reports</button>`)+`
    <section class="panel manage-guide"><div class="panel-head"><div><h3>Official store reports</h3><p class="muted small">Files are kept in private Supabase Storage. Never upload confidential reports to GitHub Pages.</p></div><span class="pill">ADMIN ONLY</span></div>
    ${storagePanel()}<div class="manage-table-wrap"><table class="manage-table"><thead><tr><th>Report</th><th>Status</th><th>Uploaded date & time</th><th>Loaded here</th></tr></thead><tbody>${rows}</tbody></table></div></section>
    <section class="panel"><div class="panel-head"><div><h3>Upload and publish a report</h3><p class="muted small">Choose the correct report type. The upload becomes available to other employees only after Publish succeeds.</p></div></div>
      <form id="officialReportForm" class="manage-form">
        <label for="officialReportType">Report type</label>
        <select id="officialReportType" required>${available.map(r=>`<option value="${esc(r.key)}">${esc(r.title)}</option>`).join('')}</select>
        <label for="officialReportFile">Excel workbook</label>
        <input id="officialReportFile" type="file" accept=".xlsx,.xls" required>
        <button type="submit" class="btn orange" id="publishReport" ${officialUploadInProgress?'disabled':''}>${officialUploadInProgress?'Publishing…':'Upload & publish official report'}</button>
      </form>
      <div id="officialReportStatus" class="manage-status" role="status" aria-live="polite">${officialUploadInProgress?'Validating and publishing…':'Your previous report versions are retained for audit.'}</div>
      <div class="manage-footer-note"><strong>Employee uploads:</strong> Files uploaded inside Data Analysis, GRN and Import & Match are saved to each employee's private My Reports area. Only this Admin page publishes shared reports.</div>
    </section>`;
}
async function fetchReportRegistry(){
  const [{data:types,error:typeError},{data:current,error:curError}]=await Promise.all([
    sb.from('report_types').select('key,title,category,instructions').order('title'),
    sb.from('current_reports').select('report_key,upload_id,updated_at')
  ]);
  if(typeError)throw typeError;if(curError)throw curError;
  reportTypes=types||[];const ids=(current||[]).map(r=>r.upload_id);
  let uploadsById={};
  if(ids.length){const {data,error}=await sb.from('report_uploads').select('id,report_key,storage_path,original_filename,bytes,uploaded_at,sha256').in('id',ids);
   if(error)throw error;uploadsById=Object.fromEntries((data||[]).map(x=>[x.id,x]));}
  latestReports=Object.fromEntries((current||[]).filter(c=>uploadsById[c.upload_id]).map(c=>[c.report_key,{...uploadsById[c.upload_id],updated_at:c.updated_at}]));
}
// V4.7: only the newest workbook in each category remains in cloud storage.
const humanBytes=value=>{const n=Number(value)||0;return n>=1048576?(n/1048576).toFixed(1)+' MB':n>=1024?(n/1024).toFixed(1)+' KB':n+' B';};
async function sha256File(file){
 if(!crypto?.subtle?.digest)return null;
 const input=await file.arrayBuffer();
 const digest=await crypto.subtle.digest('SHA-256',input);
 return Array.from(new Uint8Array(digest)).map(v=>v.toString(16).padStart(2,'0')).join('');
}
async function refreshStorageStats(){
 if(!sb||staff?.role!=='admin'){storageStats=[];return;}
 const {data,error}=await sb.rpc('report_storage_usage');
 if(error){console.warn('Storage stats unavailable',error);return;}
 storageStats=data||[];
 if(page==='myReports'||page==='reportManagement')render();
}
function storagePanel(){
 if(staff?.role!=='admin')return '';
 const own=storageStats.find(r=>r.scope==='my_reports');
 const official=storageStats.find(r=>r.scope==='official_reports');
 const all=storageStats.find(r=>r.scope==='all_employee_reports');
 return `<section class="panel storage-monitor"><div class="panel-head"><div><h3>Cloud storage</h3><p class="muted small">Only the latest report file is retained for each report type. Historical timestamps stay in the audit log.</p></div><button type="button" class="ghost" id="runStorageCleanup" ${cleanupWorking?'disabled':''}>${cleanupWorking?'Cleaning…':'Check & clean old files'}</button></div><div class="storage-summary"><span><b>${own?humanBytes(own.used_bytes):'—'}</b><small>My reports · ${own?.object_count??'—'} files</small></span><span><b>${official?humanBytes(official.used_bytes):'—'}</b><small>Official reports · ${official?.object_count??'—'} files</small></span>${staff?.role==='admin'?`<span><b>${all?humanBytes(all.used_bytes):'—'}</b><small>All employee reports · ${all?.object_count??'—'} files</small></span>`:''}</div>${cleanupStatus?`<p class="muted small" role="status">${esc(cleanupStatus)}</p>`:''}${cleanupIssues?`<p class="storage-warning" role="alert">${cleanupIssues} file(s) could not be cleaned. Use Check & clean old files to retry.</p>`:''}</section>`;
}
// Read all versions; the current pointer is checked fresh before any deletion.
async function cleanupOldReports({quiet=false}={}){
 if(!staff||!sb||cleanupWorking)return {deleted:0,failed:0};
 const identity=staff.user_id;cleanupWorking=true;
 let deleted=0,failed=0;
 try{
   const mine=await sb.from('employee_report_versions').select('id,user_id,report_key,storage_path,storage_deleted_at').eq('user_id',identity).is('storage_deleted_at',null);
   const current=await sb.from('employee_current_reports').select('upload_id').eq('user_id',identity);
   if(mine.error||current.error)throw mine.error||current.error;
   const protectedIds=new Set((current.data||[]).map(r=>r.upload_id));
   for(const v of mine.data||[]){
     if(!staff||staff.user_id!==identity)break;
     if(protectedIds.has(v.id))continue;
     // Recheck before deletion in case another session just restored this version.
     const check=await sb.from('employee_current_reports').select('upload_id').eq('user_id',identity).eq('upload_id',v.id);
     if(check.error){failed++;continue}
     if(check.data?.length)continue;
     const result=await sb.storage.from(EMPLOYEE_BUCKET).remove([v.storage_path]);
     if(result.error){failed++;continue}
     const mark=await sb.from('employee_report_versions').update({storage_deleted_at:new Date().toISOString()}).eq('id',v.id).eq('user_id',identity);
     if(mark.error)failed++;else deleted++;
   }
   if(staff?.role==='admin'){
     const [old,live]=await Promise.all([
       sb.from('report_uploads').select('id,storage_path,storage_deleted_at').is('storage_deleted_at',null),
       sb.from('current_reports').select('upload_id')
     ]);
     if(old.error||live.error)throw old.error||live.error;
     const protectedOfficial=new Set((live.data||[]).map(r=>r.upload_id));
     for(const v of old.data||[]){
       if(!staff||staff.user_id!==identity)break;
       if(protectedOfficial.has(v.id))continue;
       const check=await sb.from('current_reports').select('upload_id').eq('upload_id',v.id);
       if(check.error){failed++;continue}
       if(check.data?.length)continue;
       const result=await sb.storage.from(REPORT_BUCKET).remove([v.storage_path]);
       if(result.error){failed++;continue}
       const mark=await sb.from('report_uploads').update({storage_deleted_at:new Date().toISOString()}).eq('id',v.id);
       if(mark.error)failed++;else deleted++;
     }
   }
   cleanupIssues=failed;
   cleanupStatus=deleted?`${deleted} previous file(s) removed; ${failed} cleanup issue(s).`:failed?'Some old files could not be removed.':'No old report files need removal.';
   if(!quiet)flash(cleanupStatus);
   return {deleted,failed};
 }catch(err){cleanupIssues++;cleanupStatus='Storage cleanup could not finish: '+(err.message||err);if(!quiet)flash(cleanupStatus);return {deleted,failed:failed+1};}
 finally{cleanupWorking=false;await refreshStorageStats().catch(()=>{});}
}
// Every employee owns an independent private workbook library, separate from official reports.
async function fetchPersonalRegistry(){
 if(!sb||!staff)return;
 const {data:current,error:currentError}=await sb.from('employee_current_reports')
   .select('report_key,upload_id,updated_at').eq('user_id',staff.user_id);
 if(currentError)throw currentError;
 const ids=(current||[]).map(x=>x.upload_id);
 let versions={};
 if(ids.length){
   const {data,error}=await sb.from('employee_report_versions')
     .select('id,report_key,storage_path,original_filename,bytes,uploaded_at,sha256')
     .eq('user_id',staff.user_id).in('id',ids);
   if(error)throw error;
   versions=Object.fromEntries((data||[]).map(row=>[row.id,row]));
 }
 personalReports=Object.fromEntries((current||[]).filter(c=>versions[c.upload_id]).map(c=>[c.report_key,{...versions[c.upload_id],updated_at:c.updated_at}]));
}
async function refreshOfficialReports(options={autoImport:false}){
 if(!sb||!staff||reportLoading)return;
 reportLoading=true;const generation=++sessionGeneration;
 try{
   notifySync('Checking your saved reports…');
   await Promise.all([fetchReportRegistry(),fetchPersonalRegistry()]);
   refreshStorageStats().catch(console.error);
   if(generation!==sessionGeneration||!staff)return;
   if(options.autoImport){
     // Personal reports override official reports only for this employee; never publish them.
     const registered=Object.entries(REPORT_KIND).filter(([reportKey,kind])=>kind&&(personalReports[reportKey]||latestReports[reportKey]));
     let successful=0,failed=[];
     for(let i=0;i<registered.length;i++){
       if(generation!==sessionGeneration||!staff)break;
       const [reportKey,kind]=registered[i];
       const personal=personalReports[reportKey];const remote=personal||latestReports[reportKey];
       const bucket=personal?EMPLOYEE_BUCKET:REPORT_BUCKET;
       notifySync(`Restoring saved reports ${i+1}/${registered.length}: ${reportTypes.find(t=>t.key===reportKey)?.title||reportKey}…`);
       try{
         const {data:blob,error}=await sb.storage.from(bucket).download(remote.storage_path);
         if(error)throw error;
         if(generation!==sessionGeneration||!staff)break;
         const file=new File([blob],remote.original_filename||reportKey+'.xlsx',{type:blob.type||'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
         const ok=await loadFile(file,kind,{silent:true});
         if(!ok)throw Error('Excel report could not be parsed.');
         if(generation!==sessionGeneration||!staff)break;
         reportSources[kind]={origin:personal?'personal':'official',uploaded_at:remote.uploaded_at||remote.updated_at};
         successful++;
       }catch(err){failed.push(reportKey);console.error('Saved report load failed',reportKey,err)}
     }
     if(generation===sessionGeneration&&staff){
       notifySync(registered.length?`${successful} saved reports ready${failed.length?' · Some reports could not be read':''}.`:'No saved reports yet. Upload a workbook once to save it to your account.',!!failed.length);
       if(page==='reportManagement'||page==='myReports'||page==='upload'||page==='grn'||page==='dataAnalysis')render();
       cleanupOldReports({quiet:true}).catch(console.error);
     }
   }else{
     notifySync('Saved report list refreshed.');
     if(page==='reportManagement'||page==='myReports')render();
   }
 }catch(ex){notifySync('Cannot check saved reports: '+(ex.message||ex),true);throw ex}
 finally{reportLoading=false}
}
// Personal uploads are secure Supabase objects, not browser-only data or official published reports.
function reportFileType(file){
 const x=file.name.toLowerCase();
 return x.endsWith('.xlsx')?'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
   x.endsWith('.xls')?'application/vnd.ms-excel':'text/csv';
}
async function savePersonalReport(file,reportKey){
 if(!sb||!staff)throw Error('Please sign in again before saving a report.');
 if(!Object.hasOwn(REPORT_KIND,reportKey))throw Error('Unknown report type.');
 if(!/\.(xlsx|xls|csv)$/i.test(file.name))throw Error('Choose an Excel or CSV report.');
 if(!file.size||file.size>50*1024*1024)throw Error('Your report must be under 50 MB.');
 const userId=staff.user_id;
 const hash=await sha256File(file);
 const active=personalReports[reportKey];
 if(hash&&active?.sha256===hash)return {...active,unchanged:true};
 const path=`${userId}/${reportKey}/${crypto.randomUUID()}-${file.name.replace(/[^a-z0-9._-]/gi,'_')}`;
 const {error:uploadError}=await sb.storage.from(EMPLOYEE_BUCKET).upload(path,file,{upsert:false,contentType:reportFileType(file)});
 if(uploadError)throw uploadError;
 let recorded=false;
 try{
   const {data:version,error:versionError}=await sb.from('employee_report_versions')
     .insert({user_id:userId,report_key:reportKey,storage_path:path,original_filename:file.name,bytes:file.size,sha256:hash})
     .select('id,report_key,storage_path,original_filename,bytes,uploaded_at,sha256').single();
   if(versionError)throw versionError;
   recorded=true;
   const {error:currentError}=await sb.from('employee_current_reports')
     .upsert({user_id:userId,report_key:reportKey,upload_id:version.id,updated_at:new Date().toISOString()},{onConflict:'user_id,report_key'});
   if(currentError)throw currentError;
   personalReports[reportKey]=version;
   cleanupOldReports({quiet:true}).catch(console.error);
   if(REPORT_KIND[reportKey])reportSources[REPORT_KIND[reportKey]]={origin:'personal',uploaded_at:version.uploaded_at};
   return version;
 }catch(err){if(!recorded)await sb.storage.from(EMPLOYEE_BUCKET).remove([path]).catch(()=>{});throw err;}
}
async function importAndSaveReport(file,type,reportKey=KEY_FOR_KIND[type]){
 if(!file||!staff)return;
 if(personalSaving){flash('Please wait for the previous workbook to finish saving.');return;}
 if(!reportKey){flash('Unsupported report type.');return;}
 personalSaving=true;
 const userId=staff.user_id;
 try{
   notifySync('Reading your workbook…');
   if(type){const ok=await loadFile(file,type);if(!ok)throw Error('Could not read this Excel report. Nothing was saved.');}
   else {if(typeof XLSX==='undefined')throw Error('Excel reader is unavailable.');const wb=XLSX.read(await file.arrayBuffer(),{type:'array',bookSheets:true});if(!wb.SheetNames?.length)throw Error('Workbook contains no sheets.');}
   if(!staff||staff.user_id!==userId)throw Error('Session changed while importing the file.');
   notifySync('Saving your report privately to Supabase…');
   const record=await savePersonalReport(file,reportKey);
   notifySync((record.unchanged?'Already up to date · ':'Cloud saved · ')+fmtTimestamp(record.uploaded_at));
   if(page==='myReports'||page==='upload'||page==='grn'||page==='dataAnalysis')render();
   flash('Report saved to your account. It will be available after your next login.');
 }catch(err){notifySync('Not saved to cloud · '+(err.message||err),true);flash('Report not saved: '+(err.message||err));}
 finally{personalSaving=false;}
}
async function downloadPersonalReport(key){
 if(!staff||!sb)return;
 const report=personalReports[key];if(!report)return flash('Report not found in your account.');
 try{
  notifySync('Downloading your private workbook…');
  const {data,error}=await sb.storage.from(EMPLOYEE_BUCKET).download(report.storage_path);
  if(error)throw error;
  const url=URL.createObjectURL(data),a=document.createElement('a');a.href=url;a.download=report.original_filename;a.style.display='none';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);
  notifySync('Your workbook download has started.');
 }catch(err){notifySync('Download failed: '+err.message,true);}
}
async function revertPersonalReport(key){
 if(!staff||!personalReports[key])return;
 if(!confirm('Stop using your private version of this report? The official version, if available, will be loaded instead. Previous workbook files are removed from cloud storage after replacement.'))return;
 try{
  const {error}=await sb.from('employee_current_reports').delete().eq('user_id',staff.user_id).eq('report_key',key);
  if(error)throw error;
  delete personalReports[key];
  cleanupOldReports({quiet:true}).catch(console.error);
  const kind=REPORT_KIND[key];
  if(kind){delete reportSources[kind];if(kind==='analytics'){analyticsRows=[];analyticsMeta=null;}else if(kind==='grn')grnData=null;else {delete uploads[kind];delete meta[kind];}review();calculate();}
  await refreshOfficialReports({autoImport:true});
  if(page==='myReports')render();flash('Private override removed. Official report restored if available.');
 }catch(err){flash('Could not change active report: '+(err.message||err));}
}
function myReports(){
 const types=reportTypes.length?reportTypes:[{key:'inventory',title:'Inventory Report'},{key:'total_sales',title:'Total Sales Report'},{key:'annual_sales',title:'One-Year Sales Report'},{key:'sales_30_days',title:'30-Day Summary Sales by Quantity'},{key:'store_stock',title:'Store Stock Balance'},{key:'warehouse_stock',title:'Warehouse Stock Balance'},{key:'pending_indents',title:'Pending Indents'}];
 const rows=types.map(r=>{
   const own=personalReports[r.key],official=latestReports[r.key],active=own||official;
   const note=r.key==='annual_sales'?'Stored for future annual analysis; charts are still under development.':own?'Private to your employee account.':official?'Using store-approved data.':'Upload a report to get started.';
   const time=active?.uploaded_at||active?.updated_at;
   return `<tr><td><strong>${esc(r.title)}</strong><small class="manage-description">${esc(active?.original_filename?active.original_filename+' · '+note:note)}</small></td><td><span class="report-state ${active?'ok':''}">${own?'My report':official?'Official':'Not loaded'}</span></td><td>${esc(fmtTimestamp(time))}</td><td><div class="my-report-actions"><label class="file-upload-action my-upload-button">${own?'Replace':'Upload'}<input type="file" accept=".xlsx,.xls,.csv" data-my-report="${esc(r.key)}" aria-label="Upload ${esc(r.title)}"></label>${own?`<button type="button" class="ghost" data-my-download="${esc(r.key)}">Download</button><button type="button" class="ghost" data-my-reset="${esc(r.key)}">${official?'Use official':'Remove override'}</button>`:''}</div></td></tr>`;
 }).join('');
 return heading('My Reports','Your private Excel reports stay with your employee account across sign-outs and devices.',`<button type="button" class="btn" id="refreshReports">↻ Refresh saved reports</button>`)+`
 <section class="panel"><div class="panel-head"><div><h3>My saved workbooks</h3><p class="muted small">Upload once and access them after your next login. Dates show when a report was uploaded, in India Standard Time (IST).</p></div><span class="pill">PRIVATE TO YOU</span></div>
 ${storagePanel()}<div class="manage-table-wrap"><table class="manage-table my-reports-table"><thead><tr><th>Report</th><th>Data source</th><th>Uploaded date &amp; time</th><th>Actions</th></tr></thead><tbody>${rows}</tbody></table></div>
 <p class="muted small">Your upload never replaces the store's official report for other employees. Use “Use official” to return to store-approved data when available. Older Excel files are automatically deleted after successful replacement; only audit metadata remains.</p></section>`;
}
async function publishOfficialReport(){
  if(!sb||staff?.role!=='admin')throw Error('Only an administrator can publish store reports.');
  if(officialUploadInProgress)throw Error('Wait for the current upload to finish.');
  const key=$('#officialReportType')?.value,file=$('#officialReportFile')?.files?.[0];
  if(!key||!Object.hasOwn(REPORT_KIND,key))throw Error('Choose a valid report type.');
  if(!file)throw Error('Choose the Excel workbook to upload.');
  if(!/\.(xlsx|xls)$/i.test(file.name))throw Error('Upload an Excel .xlsx or .xls workbook.');
  if(file.size<=0||file.size>50*1024*1024)throw Error('Report must be between 1 byte and 50 MB.');
  officialUploadInProgress=true;
  let remotePath=null,entryInserted=false;
  try{
    const status=$('#officialReportStatus');if(status)status.textContent='Checking report format…';
    // Verify report is readable and mapped to its assigned purpose before publication.
    if(REPORT_KIND[key])await loadFile(file,REPORT_KIND[key],{silent:true});
    else {if(typeof XLSX==='undefined')throw Error('Excel reader not loaded.');const wb=XLSX.read(await file.arrayBuffer(),{type:'array',bookSheets:true});if(!wb.SheetNames.length)throw Error('The workbook has no sheets.');}
    const hash=await sha256File(file);
    if(hash&&latestReports[key]?.sha256===hash){
      if(status)status.textContent='This report is already current. No new storage used.';
      flash('Identical report already published.');return;
    }
    if(status)status.textContent='Uploading securely to Supabase…';
    remotePath=`${key}/${new Date().toISOString().replace(/[:.]/g,'-')}-${crypto.randomUUID()}-${file.name.replace(/[^A-Za-z0-9._-]/g,'_')}`;
    const fileContentType=/\.xls$/i.test(file.name)?'application/vnd.ms-excel':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const {error:storageError}=await sb.storage.from(REPORT_BUCKET).upload(remotePath,file,{upsert:false,contentType:fileContentType});
    if(storageError)throw storageError;
    if(status)status.textContent='Registering the latest report version…';
    const {data:inserted,error:insertError}=await sb.from('report_uploads').insert({report_key:key,storage_path:remotePath,original_filename:file.name,bytes:file.size,sha256:hash,uploaded_by:staff.user_id}).select('id').single();
    if(insertError)throw insertError;entryInserted=true;
    const {error:currentError}=await sb.from('current_reports').upsert({report_key:key,upload_id:inserted.id,updated_by:staff.user_id,updated_at:new Date().toISOString()},{onConflict:'report_key'});
    if(currentError)throw currentError;
    await fetchReportRegistry();render();flash('Official report published. Previous file is queued for cleanup.');
    cleanupOldReports({quiet:true}).catch(console.error);
  }catch(err){if(remotePath&&!entryInserted)await sb.storage.from(REPORT_BUCKET).remove([remotePath]).catch(()=>{});throw err}
  finally{officialUploadInProgress=false;}
}

function showImport(message){const el=$('#importProgress');if(el){el.hidden=false;$('#importStatus').textContent=message||'Reading sales report…';}}
function updateImport(message){const el=$('#importStatus');if(el)el.textContent=message;}
function hideImport(){const el=$('#importProgress');if(el)el.hidden=true;}
function importAnalyticsInWorker(file){
  if(typeof Worker==='undefined'||typeof Blob==='undefined')return Promise.reject(new Error('Background import is unavailable in this browser.'));
  // Use a dedicated Worker to keep the interface responsive while XLSX parses very large reports.
  const src=[
    "importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');",
    "const key="+key.toString()+";", "const norm="+norm.toString()+";", "const n="+n.toString()+";", "const isNumeric="+isNumeric.toString()+";", "const dateOf="+dateOf.toString()+";",
    `function parseSheet(sheet){
      const grid=Array.isArray(sheet)?sheet:Array.isArray(sheet['!data'])?sheet['!data']:null;
      const val=(row,col)=>{if(col<0)return null;let cell=grid?grid[row]?.[col]:sheet[XLSX.utils.encode_cell({r:row,c:col})];return cell?.v??null;};
      let rowCount=grid?.length||0;
      if(!rowCount){const ref=sheet['!fullref']||sheet['!ref'];rowCount=ref?XLSX.utils.decode_range(ref).e.r+1:0;}
      let hi=-1,heads=[];const width=90;
      for(let i=0;i<Math.min(rowCount,35);i++){
        const h=Array.from({length:width},(_,j)=>norm(val(i,j)).replace(/[^a-z0-9]+/g,''));
        if(h.some(x=>['item','itemname','product','productname','particulars','sku'].includes(x))&&h.some(x=>['amount','totalamount','salesamount','netamount','saleamount'].includes(x))){hi=i;heads=h;break;}
      }
      if(hi<0)throw Error('Could not find ItemName and TotalAmount in the sales report.');
      const ix=names=>heads.findIndex(x=>names.includes(x));
      const item=ix(['item','itemname','product','productname','particulars','sku']);
      const amt=ix(['amount','totalamount','salesamount','netamount','saleamount']);
      const qty=ix(['quantity','totalquantity','qty','totalqty','salesqty']);
      const floor=ix(['profitcentre','profitcenter','floor','floorname','floorwise','profitcentrename']);
      const date=ix(['date','createddate','billdate','invoicedate','salesdate','transactiondate']);
      const cat=ix(['category','categoryname','department']);
      const subcat=ix(['subcategory','subcategoryname','subgroup']);
      const skuIx=ix(['itemcode','sku','skucode']);
   const brandIx=ix(['brand','brandname','itembrand','productbrand','manufacturerbrand','brands','branddescription','brandlabel']);
      if(floor<0&&date<0)throw Error('Floor or date field is missing in the sales report.');
      const out=[];
      for(let r=hi+1;r<rowCount;r++){
        const name=key(val(r,item)),amount=val(r,amt);
        if(!name||/^(total|grand total|subtotal|group)/i.test(name)||!isNumeric(amount))continue;
        const d=date>=0?dateOf(val(r,date)):null;
        const day=d?d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'):null;
        out.push({name,sku:skuIx>=0?key(val(r,skuIx)):'',amount:n(amount),qty:qty>=0?n(val(r,qty)):null,floor:floor>=0?key(val(r,floor)):'',day,cat:cat>=0?key(val(r,cat)):'',brand:brandIx>=0?(key(val(r,brandIx))||'Unspecified brand'):'Unspecified brand'});
      }
      if(!out.length)throw Error('No valid sales transactions were detected.');
      return {rows:out,hasFloor:floor>=0,hasDay:date>=0};
    }`,
    `self.onmessage=event=>{try{
       const wb=XLSX.read(event.data.buffer,{type:'array',cellDates:true,dense:true});
       if(!wb.SheetNames.length)throw Error('No sheet found in Excel report.');
       const result=parseSheet(wb.Sheets[wb.SheetNames[0]]);
       self.postMessage({ok:true,result});
     }catch(err){self.postMessage({ok:false,error:err?.message||String(err)})}}`
  ].join('\n');
  const url=URL.createObjectURL(new Blob([src],{type:'application/javascript'}));
  return new Promise((resolve,reject)=>{
    let worker,completed=false;
    const finish=(err,result)=>{if(completed)return;completed=true;try{worker?.terminate()}catch(e){}URL.revokeObjectURL(url);err?reject(err):resolve(result)};
    try{
      worker=new Worker(url);
      worker.onerror=event=>{event.preventDefault();finish(new Error('Excel import worker failed. Please check the Excel format or try again.'))};
      worker.onmessage=event=>{const msg=event.data;if(msg?.ok)finish(null,msg.result);else finish(new Error(msg?.error||'Unable to parse report.'))};
      file.arrayBuffer().then(buffer=>worker.postMessage({buffer},[buffer]),finish);
    }catch(err){finish(err)}
  });
}
async function handleAnalyticsUpload(file,opts={}){
  if(!opts.silent)showImport('Opening the report in the background…');
  await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
  let statusTimer=setTimeout(()=>updateImport('Reading transactions and identifying sales categories…'),2400);
  try {
    const result=await importAnalyticsInWorker(file);
    updateImport('Building charts and smart filters…');
    await new Promise(resolve=>requestAnimationFrame(resolve));
    analyticsRows=result.rows;analyticsImportError='';
    analyticsMeta={name:file.name,hasFloor:result.hasFloor,hasDay:result.hasDay};
    analysisMonth='all';analysisFloor='all';analysisCategory='all';analysisBrand='all';analysisSubcategory='all';analysisSearch='';analysisCategoryQuery='';selectedSalesDay=null;
    if(!opts.silent){go('dataAnalysis');flash('Sales report imported: '+fmt(analyticsRows.length)+' transactions')}else if(page==='dataAnalysis')render();
    return true;
  }catch(err){
    console.error('Sales report import failed',err);
    analyticsImportError=err?.message||String(err);if(page==='dataAnalysis')render();
    if(!opts.silent)flash('Sales report could not be imported: '+analyticsImportError);throw err;
  }finally{clearTimeout(statusTimer);if(!opts.silent)hideImport()}
}

async function loadFile(file,type,opts={}){if(type==='analytics'&&file){return await handleAnalyticsUpload(file,opts);}if(!file)return false;if(typeof XLSX==='undefined'){if(opts.silent)throw Error('Excel library could not load. Check your internet connection.');flash('Excel library could not load. Check your internet connection.');return false}if(type==='grn'){grnLoading=true;grnImportError='';if(!opts.silent)go('grn');await new Promise(resolve=>setTimeout(resolve,20));}try{let buf=await file.arrayBuffer(),wb=XLSX.read(buf,{type:'array',cellDates:true}),rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,raw:true,defval:null,blankrows:false});if(type==='grn'){grnData=parseInventoryGrn(rows);grnLevel='months';grnMonth='';grnDay='';grnNumber='';grnSearch='';grnImportError='';grnLoading=false;if(!opts.silent){go('grn');flash('Inventory GRN report prepared: '+fmt(grnData.allGrns.size)+' GRNs')}else if(page==='grn')render();return true;}if(type==='analytics'){const result=parseAnalytics(rows);analyticsRows=result.rows;analyticsMeta={name:file.name,hasFloor:result.hasFloor,hasDay:result.hasDay};analysisMonth='all';analysisFloor='all';analysisCategory='all';analysisBrand='all';analysisSubcategory='all';analysisSearch='';analysisCategoryQuery='';selectedSalesDay=null;go('dataAnalysis');flash('Additional sales report loaded: '+fmt(analyticsRows.length)+' rows');return}let parsed=type==='store'||type==='warehouse'?parseStock(rows):type==='sales'?parseSales(rows):parsePending(rows);if(!parsed.length)throw Error('No valid item rows were detected.');const signature=`${file.name}|${file.size}|${file.lastModified}`;if(!opts.silent&&Object.values(meta).some(x=>x.sig===signature)){if(!confirm('This exact file has already been uploaded. Replace anyway?'))return}uploads[type]=parsed;meta[type]={name:file.name,sig:signature};edits={};approved=false;review();calculate();render();if(!opts.silent)flash(`${type} loaded: ${fmt(parsed.length)} records`);return true;}catch(err){if(type==='grn'){grnImportError=err.message;grnLoading=false;go('grn');}if(!opts.silent)flash('Import error: '+err.message);console.error(err);if(opts.silent)throw err;}finally{if(type==='grn'&&grnLoading){grnLoading=false;if(page==='grn')render()}}}
function modalMap(sku){const item=matches.find(x=>x.sku===sku),opts=[...new Set(datasets().sales.map(r=>r.name))];if(!item)return;const current=mappings[sku]||'';const suggested=opts.filter(x=>norm(x).includes(norm(item.name).slice(0,11))).slice(0,12);let chosen=prompt(`Map sales product to store SKU ${sku}: ${item.name}\n\nEnter EXACT sales Item name.\nSuggestions:\n${suggested.join('\n')}`,current||suggested[0]||'');if(chosen===null)return;if(!opts.includes(chosen.trim())){flash('Name must exactly match an Item value in the sales report.');return}mappings[sku]=chosen.trim();review();calculate();render();flash('Product manually mapped; inspect updated matching results.')}
function categorySetting(){let cats=[...new Set(datasets().store.map(r=>r.cat))].sort();let list=cats.length>40?cats.slice(0,40):cats;let message=`Set category cover days. Default ${baseCover}.\n\n`+list.map(x=>`${x}: ${covers[x]??baseCover}`).join('\n');let chosen=prompt(message+'\n\nEnter EXACT category name to change:');if(!chosen)return;if(!cats.includes(chosen)){flash('Category not found. Enter exact category name.');return}let days=prompt('New stock-cover days for '+chosen,String(covers[chosen]??baseCover));if(!days)return;if(!(Number(days)>0&&Number(days)<=365)){flash('Enter 1–365 days.');return}covers[chosen]=Number(days);calculate();render();flash('Cover updated for '+chosen)}
function detail(sku){let r=records.find(x=>x.sku===sku);if(!r)return;let remark=prompt(`${r.name}\nItem Code: ${r.sku}\nStore ${r.qty}, Warehouse ${r.warehouse??'missing'}, Valid pending ${r.pending}\nRequired ${r.need??'unknown'}, Suggested ${r.suggest}, Final ${r.final}\nFlags: ${r.flags.join('; ')||'None'}\n\nEnter review remarks (optional):`,r.remark||'');if(remark===null)return;r.remark=remark;flash('SKU review note updated for this session.')}
function updateEdit(sku,value){let r=records.find(x=>x.sku===sku);if(!r||approved)return;if(value===''||!Number.isFinite(Number(value))||Number(value)<0){flash('Final quantity must be a non-negative number');return}let qty=Number(value);r.final=qty;edits[sku]=qty;if(qty>(r.warehouse??0)&&!r.flags.includes('Manager qty exceeds warehouse'))r.flags.push('Manager qty exceeds warehouse');if(qty<=(r.warehouse??0))r.flags=r.flags.filter(f=>f!=='Manager qty exceeds warehouse')}
function createDatabase(){return new Promise(resolve=>{try{if(typeof indexedDB==='undefined'){resolve(null);return}let q=indexedDB.open('todaysBasketHtmlV2',1);q.onupgradeneeded=()=>{if(!q.result.objectStoreNames.contains('history'))q.result.createObjectStore('history',{keyPath:'id'})};q.onsuccess=()=>resolve(q.result);q.onerror=()=>resolve(null)}catch(err){resolve(null)}})}
async function saveVersion(){if(!db){flash('Local history storage unavailable.');return}let last=history.filter(x=>x.date===today()).reduce((m,x)=>Math.max(m,x.version),0),row={id:crypto.randomUUID(),date:today(),version:last+1,created:new Date().toLocaleString('en-IN'),approved,records:JSON.parse(JSON.stringify(records)),settings:{salesDays,baseCover,covers,edits},actor:staff?.full_name||'Store manager',action:approved?'approved':'saved'};await new Promise((yes,no)=>{const t=db.transaction('history','readwrite');t.objectStore('history').put(row);t.oncomplete=yes;t.onerror=no});history.unshift(row);flash('Saved local version V'+row.version);if(page==='history')render()}
async function loadHistory(){if(!db)return;history=await new Promise(resolve=>{let t=db.transaction('history','readonly'),q=t.objectStore('history').getAll();q.onsuccess=()=>resolve(q.result.sort((a,b)=>b.created.localeCompare(a.created)));q.onerror=()=>resolve([])})}
function restore(id){let item=history.find(r=>r.id===id);if(!item||!confirm('Restore this saved analysis snapshot? Current unsaved edits may be lost.'))return;records=JSON.parse(JSON.stringify(item.records));salesDays=item.settings.salesDays;baseCover=item.settings.baseCover;covers=item.settings.covers;edits=item.settings.edits;approved=item.approved;go('indent');flash('Version restored locally.')}
function exportExcel(){if(typeof XLSX==='undefined'){flash('Excel library unavailable.');return}let header=['Item Code','Product','Category','UOM','Sales Period Days','Sales Total Qty','Sales Total Amount','Store Qty','Warehouse Qty','Valid Pending Qty','Expired Pending Qty','Unverified Pending Qty','Target Cover Days','30/60-day Target Qty','Required Qty','Warehouse-capped Suggested Qty','Manager Final Qty','Current Stock Cover Days','Flags','Review Remarks','Approval Status'];let rows=records.map(r=>[r.sku,r.name,r.cat,r.uom,salesDays,r.salesQty,r.amount,r.qty,r.warehouse,r.pending,r.expiredPending,r.unknownPending,r.coverDays,r.target,r.need,r.suggest,r.final,r.cover,r.flags.join('; '),r.remark||'',approved?'Approved':'Draft']);let wb=XLSX.utils.book_new(),ws=XLSX.utils.aoa_to_sheet([header,...rows]);ws['!cols']=[{wch:18},{wch:44},{wch:24},{wch:10},...Array(14).fill({wch:18}),{wch:48},{wch:35},{wch:18}];XLSX.utils.book_append_sheet(wb,ws,'Complete Indent Analysis');XLSX.writeFile(wb,`Todays_Basket_Indent_${today()}_${approved?'Approved':'Draft'}.xlsx`);flash('Complete analysis exported in one Excel worksheet')}
function approve(){if(!complete()){flash('Upload all four actual indent reports before approval.');return}let actor=prompt('Manager name / Employee ID for audit (local prototype only):');if(!actor)return;if(stats.duplicates.length&&!confirm('Duplicate Item Codes exist. Approve despite this data risk?'))return;if(records.some(r=>r.unknownPending>0)&&!confirm('Some PO expiry dates are unknown and their pending qty was NOT deducted. Approve anyway?'))return;if(!confirm(`Approve ${records.length} SKU rows as ${actor}? This indent approval is currently saved only in this browser; server-side approval is not implemented yet.`))return;approved=true;saveVersion();render();flash('Approved locally. Revisions are recorded only in this browser.')}
function unlock(){const who=prompt('Manager name / Employee ID unlocking approved indent:');if(!who)return;if(!confirm('Unlock this approved analysis for editing? Save a new version after changes.'))return;approved=false;flash('Unlocked by '+who+' (prototype audit only)');render()}
document.addEventListener('click',e=>{const t=e.target.closest('button');if(!t)return;if(t.id==='employeeReload'){loadEmployees();return;}if(t.id==='employeeEditCancel'){$('#employeeEditOverlay')?.remove();return;}if(t.dataset.employeeEdit){editEmployee(t.dataset.employeeEdit);return;}if(t.id==='logout'){if(personalSaving||officialUploadInProgress){flash('Please wait until your report finishes saving before logging out.');return;}logoutEmployee();return;}if(t.id==='runStorageCleanup'){if(staff?.role!=='admin')return;cleanupOldReports().catch(console.error);return;}if(t.id==='refreshReports'){refreshOfficialReports({autoImport:true}).catch(err=>flash(err.message));return;}if(t.dataset.myDownload){downloadPersonalReport(t.dataset.myDownload);return;}if(t.dataset.myReset){revertPersonalReport(t.dataset.myReset);return;}if(t.id==='overviewMenu'){const shouldOpen=$('#overviewMenu').getAttribute('aria-expanded')!=='true';go('overview');setOverviewMenu(shouldOpen);return;}if(t.id==='grnMenu'){setGrnMenu($('#grnMenu').getAttribute('aria-expanded')!=='true');return;}if(t.dataset.stockReqTab){stockReqTab=t.dataset.stockReqTab;render();return;}if(t.hasAttribute('data-grn-back')){grnNavigate(grnLevel==='products'?'grns':grnLevel==='grns'?'dates':'months');return;}if(t.hasAttribute('data-grn-nav')){grnNavigate(t.dataset.grnNav);return;}if(t.hasAttribute('data-grn-month')){grnMonth=t.dataset.grnMonth;grnNavigate('dates');return;}if(t.hasAttribute('data-grn-date')){grnDay=t.dataset.grnDate;grnNavigate('grns');return;}if(t.hasAttribute('data-grn-number')){grnNumber=t.dataset.grnNumber;grnNavigate('products');return;}if(t.id==='togglePassword'){const field=document.getElementById('employeePassword');field.type=field.type==='password'?'text':'password';t.textContent=field.type==='password'?'Show':'Hide';return;}if(t.dataset.page){go(t.dataset.page);if(t.dataset.page==='employeeManagement'&&isEmployeeAdministrator())loadEmployees();return;}if(t.hasAttribute('data-sales-day')){selectSalesDay(t.dataset.salesDay);return;}if(t.hasAttribute('data-day-move')){const d=currentAnalyticsView||dataGroup();const current=d.dates.includes(selectedSalesDay)?d.dates.indexOf(selectedSalesDay):d.dates.length-1;const next=d.dates[current+Number(t.dataset.dayMove)];if(next)selectSalesDay(next);return;}if(t.dataset.tab){tab=t.dataset.tab;render()}if(t.hasAttribute('data-indent-cat')){indentCategory=t.dataset.indentCat;render()}if(t.hasAttribute('data-analytics-cat')){analysisCategory=t.dataset.analyticsCat;analysisSubcategory='all';analysisBrand='all';analysisCategoryQuery='';selectedSalesDay=null;render()}if(t.hasAttribute('data-analytics-month')){analysisMonth=t.dataset.analyticsMonth;analysisFloor='all';analysisCategory='all';analysisBrand='all';analysisSubcategory='all';analysisCategoryQuery='';render()}if(t.hasAttribute('data-analytics-floor')){analysisFloor=t.dataset.analyticsFloor;analysisCategory='all';analysisBrand='all';analysisSubcategory='all';analysisCategoryQuery='';render()}if(t.id==='resetAnalytics'){analysisMonth='all';analysisFloor='all';analysisCategory='all';analysisBrand='all';analysisSubcategory='all';analysisSearch='';analysisCategoryQuery='';selectedSalesDay=null;render()}if(t.dataset.map)modalMap(t.dataset.map);if(t.dataset.detail)detail(t.dataset.detail);if(t.dataset.restore)restore(t.dataset.restore);switch(t.id){case'menu':$('.side').classList.toggle('open');break;case'reset':if(confirm('Clear uploaded indent reports? Locally saved history remains.')){uploads={};meta={};edits={};mappings={};approved=false;review();calculate();render()}break;case'calculateBtn':if(!complete())flash('Upload all four Excel reports first.');else{calculate();go('indent');flash('Calculated from uploaded reports. Review quantities before approval.')}break;case'categoryCover':categorySetting();break;case'recalc':calculate();render();flash('Analysis recalculated with current settings');break;case'approve':approve();break;case'unlock':unlock();break;case'exportXlsx':exportExcel();break;case'expandTags':window.open('./price-tags.html','_blank','noopener');break}});
document.addEventListener('change',e=>{const t=e.target;if(t.dataset.upload)importAndSaveReport(t.files?.[0],t.dataset.upload);if(t.dataset.myReport)importAndSaveReport(t.files?.[0],REPORT_KIND[t.dataset.myReport],t.dataset.myReport);if(t.id==='salesDays'){salesDays=+t.value;calculate();render()}if(t.id==='baseCover'){let k=+t.value;if(k>=1&&k<=365){baseCover=k;calculate();render()}else flash('Stock cover should be 1–365 days')}if(t.dataset.edit){updateEdit(t.dataset.edit,t.value);let focus=page;render();page=focus}});
document.addEventListener('change',e=>{
  const id=e.target.id;
  if(id==='indentCategorySelect'){indentCategory=e.target.value;render();}
  if(id==='smartMonth'){analysisMonth=e.target.value;analysisFloor='all';analysisCategory='all';analysisBrand='all';analysisSubcategory='all';analysisBrand='all';analysisCategoryQuery='';selectedSalesDay=null;render()}
  if(id==='smartFloor'){analysisFloor=e.target.value;analysisCategory='all';analysisBrand='all';analysisSubcategory='all';analysisBrand='all';analysisCategoryQuery='';selectedSalesDay=null;render()}
  if(id==='analysisCategorySelect'&&e.target.value){analysisCategory=e.target.value;analysisSubcategory='all';analysisBrand='all';analysisCategoryQuery='';selectedSalesDay=null;render()}
  if(id==='smartSubcategory'){analysisSubcategory=e.target.value;selectedSalesDay=null;render()}
   if(id==='smartBrand'){analysisBrand=e.target.value;selectedSalesDay=null;render()}
});
document.addEventListener('input',e=>{if(e.target.id==='analysisSearch'){analysisSearch=e.target.value;clearTimeout(window.analysisDebounce);window.analysisDebounce=setTimeout(()=>{const pos=$('#analysisSearch')?.selectionStart??analysisSearch.length;render();const el=$('#analysisSearch');if(el){el.focus();el.setSelectionRange(pos,pos)}},280)}if(e.target.id==='stockReqSearch'){stockReqSearch=e.target.value;const rows=stockRequirementsData()||[];const selected=rows.filter(x=>(stockReqTab==='all'||x.status===stockReqTab)&&(!stockReqSearch||norm(x.name+' '+x.sku+' '+x.cat).includes(norm(stockReqSearch))));const node=$('#stockReqRows');if(node)node.innerHTML=stockRequirementsRows(selected); }if(e.target.id==='indentSearch'){filter=e.target.value;const r=$('#indentRows');if(r)r.innerHTML=tableRows(filtered(),page==='indent')}if(e.target.id==='grnSearch'){grnSearch=e.target.value;grnUpdateSearch()}if(e.target.id==='matchSearch'||e.target.id==='matchFilter'){const r=$('#matchTable');if(r)r.innerHTML=matchTable()}});
document.getElementById('loginForm').addEventListener('submit',signInEmployee);
document.addEventListener('submit',async event=>{if(event.target.id!=='officialReportForm')return;event.preventDefault();try{await publishOfficialReport()}catch(err){flash('Report publish failed: '+err.message);const el=$('#officialReportStatus');if(el)el.textContent='Upload failed: '+err.message;}finally{if(page==='reportManagement'){$('#publishReport')?.removeAttribute('disabled');}}});
document.addEventListener('input',e=>{if(e.target.id==='analysisCategoryFilter'){analysisCategoryQuery=e.target.value;searchCategoryOptions(e.target.value)}});
document.addEventListener('keydown',e=>{
 if(e.target.id==='analysisCategoryFilter'&&e.key==='Enter'){
   const select=document.getElementById('analysisCategorySelect');
   const first=[...select.options].find(x=>!x.disabled&&x.value!=='all');
   if(first){e.preventDefault();analysisCategory=first.value;render();}
 }
});

// V4.9 - animate the corresponding SVG icon whenever a navigation item is pressed.
document.getElementById('nav')?.addEventListener('click',event=>{
  const button=event.target.closest('button');
  if(!button||!event.currentTarget.contains(button))return;
  button.classList.remove('nav-icon-tap');
  void button.offsetWidth;
  button.classList.add('nav-icon-tap');
  clearTimeout(button._iconTimer);
  button._iconTimer=setTimeout(()=>button.classList.remove('nav-icon-tap'),500);
});

// V4.8 - Password recovery & account security
let passwordMode='forgot';
const strongPassword=p=>p.length>=12&&/[A-Za-z]/.test(p)&&/\d/.test(p)&&/[^A-Za-z0-9]/.test(p);
function showPasswordDialog(mode){
 passwordMode=mode;
 const dialog=document.getElementById('passwordDialog');if(!dialog)return;
 dialog.hidden=false;
 const title=document.getElementById('passwordDialogTitle'),desc=document.getElementById('passwordDialogText');
 const old=document.getElementById('passwordOldField'),email=document.getElementById('passwordEmailField'),fresh=document.getElementById('passwordNewField');
 const message=document.getElementById('passwordActionMessage');message.textContent='';message.classList.remove('is-success');
 document.getElementById('passwordActionForm').reset();
 old.hidden=mode!=='change';email.hidden=mode!=='forgot';fresh.hidden=mode==='forgot';
 title.textContent=mode==='forgot'?'Reset your password':mode==='recovery'?'Choose a new password':'Change your password';
 desc.textContent=mode==='forgot'?'Enter the email registered with your employee account. We will send a secure reset link if the account is eligible. Check your inbox and spam folder.':mode==='recovery'?'Your reset link was verified. Set a strong new password.':'Enter your current password and choose a different, strong new password.';
 document.getElementById('passwordActionSubmit').textContent=mode==='forgot'?'Send reset link':'Update password';
 document.getElementById(mode==='forgot'?'passwordEmail':mode==='change'?'passwordOld':'passwordNew').focus();
}
function closePasswordDialog(){
 if(passwordRecoverySession)return;
 document.getElementById('passwordDialog').hidden=true;
}
async function completePasswordAction(event){
 event.preventDefault();
 const msg=document.getElementById('passwordActionMessage'),button=document.getElementById('passwordActionSubmit');
 if(!sb){msg.textContent='Secure login service is unavailable.';msg.classList.remove('is-success');return;}
 button.disabled=true;const label=button.textContent;button.textContent='Please wait…';msg.textContent='';msg.classList.remove('is-success');
 try{
  if(passwordMode==='forgot'){
    const email=document.getElementById('passwordEmail').value.trim();
    if(!/^\S+@\S+\.\S+$/.test(email))throw Error('Please enter your registered email address.');
    if(!/^https?:$/.test(location.protocol)){msg.textContent='To get a reset link, please use the store dashboard website.';msg.classList.add('is-success');flash('Open your store dashboard website to reset your password.');return;}
    // Dedicated callback, in the same GitHub Pages repository as index.html.
    const callback=new URL('reset-password.html',location.href).href.split('#')[0].split('?')[0];
    const {error}=await sb.auth.resetPasswordForEmail(email,{redirectTo:callback});
    if(error)throw Error('Unable to send the link. Check the Supabase redirect URL settings and try again.');
    msg.textContent='If this email is registered, you’ll receive a reset link shortly. Please check your inbox.';msg.classList.add('is-success');flash('Reset link requested. Please check your email.');
    return;
  }
  const pass=document.getElementById('passwordNew').value;
  const confirmation=document.getElementById('passwordConfirm').value;
  if(!strongPassword(pass))throw Error('Use 12+ characters, including letters, a number, and a symbol.');
  if(pass!==confirmation)throw Error('The new passwords do not match.');
  if(passwordMode==='change'){
    if(!staff||!serverSession)throw Error('Please sign in before changing your password.');
    const current=document.getElementById('passwordOld').value;
    const {data:{user},error:userError}=await sb.auth.getUser();
    if(userError||!user?.email)throw Error('Could not verify this employee account.');
    // Verify current password using an isolated, non-persistent client so the
    // active employee session is not replaced during password confirmation.
    const verificationClient=window.supabase.createClient(SB_URL,SB_PUBLISHABLE_KEY,{
      auth:{persistSession:false,autoRefreshToken:false,detectSessionInUrl:false}
    });
    const {error:verifyError}=await verificationClient.auth.signInWithPassword({email:user.email,password:current});
    if(verifyError)throw Error('Your current password is incorrect.');
    await verificationClient.auth.signOut().catch(()=>{});
    if(current===pass)throw Error('Your new password must be different from your current password.');
  } else if(passwordMode==='recovery'&&!passwordRecoverySession)throw Error('Open your verified password reset link first.');
  const {error}=await sb.auth.updateUser({password:pass});
  if(error)throw error;
  passwordRecoverySession=false;document.getElementById('passwordDialog').hidden=true;
  document.getElementById('passwordActionForm').reset();
  await logoutEmployee();
  const errorLabel=document.getElementById('loginError');if(errorLabel)errorLabel.textContent='Password successfully updated in Supabase. Sign in with your new password.';
  // Clean recovery tokens from the address bar without changing application location.
  if(location.hash)history.replaceState(null,'',location.pathname+location.search);
 }catch(err){msg.classList.remove('is-success');msg.textContent=err.message||'Password update failed.';}
 finally{button.disabled=false;button.textContent=label;}
}
document.getElementById('passwordActionForm').addEventListener('submit',completePasswordAction);
document.getElementById('forgotPassword').addEventListener('click',()=>showPasswordDialog('forgot'));
document.getElementById('changePassword').addEventListener('click',()=>showPasswordDialog('change'));
document.getElementById('passwordDialogClose').addEventListener('click',closePasswordDialog);

(async()=>{
  try{db=await createDatabase();await loadHistory();}
  catch(err){console.warn('Local history unavailable:',err);db=null;}
  review();calculate();render();applyLoginState();
  if(setupSupabase())await restoreSession();
})().catch(err=>{console.error('Application startup failed:',err);const label=document.getElementById('loginError');if(label)label.textContent='Dashboard could not start. Refresh the page or contact the store administrator.';});

// Employee Management forms: authenticated API calls; no employee passwords handled here.
document.addEventListener('submit',async event=>{
 if(!['employeeEditForm','employeeRegisterForm'].includes(event.target?.id))return;
 event.preventDefault();
 if(event.target.id==='employeeEditForm')await saveEmployeeForm(event.target);
 else await registerEmployee(event.target);
});

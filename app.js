/* Today's Basket HTML v2 — local prototype. No server authentication. */
const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const n=v=>{if(typeof v==='number')return Number.isFinite(v)?v:0;let x=Number(String(v??'').replace(/,/g,'').replace(/[₹ ]/g,''));return Number.isFinite(x)?x:0};
const fmt=v=>new Intl.NumberFormat('en-IN',{maximumFractionDigits:2}).format(v||0),money=v=>'₹'+fmt(v),key=v=>String(v??'').trim(),norm=v=>key(v).toLowerCase();
const round=v=>Math.ceil((v-1e-9)*100)/100;
let page='overview', tab='all', filter='',uploads={}, meta={}, records=[], matches=[], stats={}, edits={},covers={}, mappings={}, approved=false, revision=1, history=[], activeId=null;
let analysisSubcategory='all',analysisCategoryQuery='',grnRows=[],grnHeaders=[],grnLabel='',grnSearch='',grnShowAll=false;
let salesDays=60,baseCover=30,db=null,indentCategory='all',analysisCategory='all',analysisFloor='all',analysisSearch='',analyticsRows=[],analyticsMeta=null,analyticsMode='value',analysisMonth='all',selectedSalesDay=null,currentAnalyticsView=null;
const sample=[['TB-101','Amul Taaza Milk 500ml','DAIRY',12,140,184,5152,8],['TB-102','Pepsi 750ml','BEVERAGES',8,78,122,4880,0],['TB-103','Fortune Sunflower Oil 1L','GROCERY',70,80,60,8700,0],['TB-104','Britannia Bread 400g','BAKERY',0,42,103,4635,0],['TB-105','Aashirvaad Atta 5kg','GROCERY',15,90,81,23490,20],['TB-106','Cavins Curd 400g','DAIRY',21,50,143,6435,0],['TB-107','Tata Salt 1kg','GROCERY',48,200,47,1316,0],['TB-108','Lay’s Classic Salted','SNACKS',2,45,88,1760,0],['TB-109','Watermelon Striped','PRODUCE',9,32,101,7575,0],['TB-110','Amul Ice Cream Tub','FROZEN',5,19,36,7020,0]];
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
  if(floor<0 && date<0)throw Error('Floor/day report must also contain a Floor/Profit Centre column or a Date column.');
  let out=[];
  for(const r of rows.slice(hi+1)){
    const name=key(r[item]);if(!name||/^(total|grand total|subtotal|group)/i.test(name)||!isNumeric(r[amt]))continue;
    const d=date>=0?dateOf(r[date]):null;
    const day=d?`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`:null;
    out.push({name,sku:skuIx>=0?key(r[skuIx]):'',amount:n(r[amt]),qty:qty>=0?n(r[qty]):null,floor:floor>=0?key(r[floor]):'',day,cat:cat>=0?key(r[cat]):'',subcat:subcat>=0?key(r[subcat]):''});
  }
  if(!out.length)throw Error('No valid sales rows detected in the floor/day report.');
  return {rows:out,hasFloor:floor>=0,hasDay:date>=0};
}
function parsePending(rows){let hi=headerIndex(rows,['indent no']);if(hi<0)throw Error('Could not find Indent No in the previous-indent report.');const h=rows[hi];let out=[];for(const r of rows.slice(hi+1)){let sku=key(r[col(h,'itemcode',3)]||r[3]);if(!sku)continue;let exp=dateOf(r[col(h,'expirydate',2)]),pending=n(r[col(h,'pendingqty',7)]);if(!Number.isFinite(pending)||pending<=0)continue;out.push({sku,indent:key(r[col(h,'indent no',0)]),exp,days:compareDates(exp),pending,status:key(r[col(h,'indent status',11)]),isExpired:exp?compareDates(exp)<0:false,expiryUnknown:!exp})}return out}
function demoRows(){const store=sample.map(s=>({sku:s[0],name:s[1],cat:s[2],qty:s[3],uom:'NOS'})),warehouse=sample.map(s=>({sku:s[0],name:s[1],cat:s[2],qty:s[4],uom:'NOS'})),sales=sample.map(s=>({name:s[1],qty:s[5],amount:s[6]})),pending=sample.filter(s=>s[7]).map(s=>({sku:s[0],pending:s[7],isExpired:false,expiryUnknown:false,indent:'DEMO-PO',status:'Open'}));return {store,warehouse,sales,pending}}
function datasets(){if(['store','warehouse','sales','pending'].every(t=>uploads[t]))return uploads;return demoRows()}
function complete(){return ['store','warehouse','sales','pending'].every(t=>uploads[t])}
function review(){const d=datasets(),wm=new Map(),pm=new Map(),sm=new Map(),sn=new Map(),duplicates=[];for(const r of d.warehouse){if(wm.has(r.sku))duplicates.push('Warehouse duplicate Item Code '+r.sku);wm.set(r.sku,(wm.get(r.sku)||0)+r.qty)}for(const r of d.pending){const p=pm.get(r.sku)||{valid:0,expired:0,unknown:0};if(r.isExpired)p.expired+=r.pending;else if(r.expiryUnknown)p.unknown+=r.pending;else p.valid+=r.pending;pm.set(r.sku,p)}for(const r of d.sales){let name=key(r.name),prev=sm.get(name)||{qty:0,amount:0};prev.qty+=r.qty;prev.amount+=r.amount;sm.set(name,prev)}const nameCount=new Map();for(const r of d.store)nameCount.set(r.name,(nameCount.get(r.name)||0)+1);let rows=[];for(const s of d.store){if(sn.has(s.sku))duplicates.push('Store duplicate Item Code '+s.sku);sn.set(s.sku,true);let mapped=mappings[s.sku],salesName=mapped||s.name,sale=sm.get(salesName);let ambiguous=nameCount.get(s.name)>1&&!mapped;let pend=pm.get(s.sku)||{valid:0,expired:0,unknown:0};let warehouse=wm.has(s.sku)?wm.get(s.sku):null;rows.push({...s,warehouse,salesName,salesQty:sale?.qty??null,amount:sale?.amount??null,matched:!!sale&&!ambiguous,ambiguous,pending:pend.valid,expiredPending:pend.expired,unknownPending:pend.unknown})}const used=new Set(rows.filter(r=>r.matched).map(r=>r.salesName)),unmatchedSales=[...sm.keys()].filter(k=>!used.has(k));matches=rows;stats={duplicates,unmatchedSales,store:rows.length,warehouse:d.warehouse.length,sales:d.sales.length,pending:d.pending.length,matched:rows.filter(r=>r.matched).length,missingSales:rows.filter(r=>r.salesQty===null).length,missingWarehouse:rows.filter(r=>r.warehouse===null).length,expiredLines:d.pending.filter(p=>p.isExpired).length,unknownExpiry:d.pending.filter(p=>p.expiryUnknown).length};return rows}
function calculate(){if(!matches.length)review();records=matches.map(r=>{const cv=covers[r.cat]??baseCover,rate=r.matched?r.salesQty/salesDays:null,target=rate===null?null:round(rate*cv),need=target===null?null:Math.max(0,round(target-r.qty-r.pending)),suggest=need===null||r.warehouse===null?0:Math.min(need,Math.max(r.warehouse,0)),cover=rate===null?null:rate>0?r.qty/rate:null;let flags=[];if(r.salesQty===null)flags.push('No Sales Data');else if(r.salesQty===0)flags.push('Zero sales');if(r.ambiguous)flags.push('Ambiguous name');if(r.warehouse===null)flags.push('Warehouse unavailable');else if(need!==null&&need>r.warehouse)flags.push('Warehouse shortage');if(r.expiredPending>0)flags.push('Expired previous indent');if(r.unknownPending>0)flags.push('PO expiry unverified');if(cover!==null&&cover>cv*2)flags.push('High stock');let edit=Object.hasOwn(edits,r.sku)?edits[r.sku]:suggest;return {...r,rate,target,need,suggest,cover,coverDays:cv,final:edit,flags,remark:''}});approved=false;return records}
function summary(){let d=records.length?records:calculate(),revenue=d.reduce((a,r)=>a+(r.amount||0),0),req=d.reduce((a,r)=>a+(r.need||0),0),final=d.reduce((a,r)=>a+n(r.final),0);return {revenue,req,final,low:d.filter(r=>r.qty===0).length,short:d.filter(r=>r.flags.includes('Warehouse shortage')).length,issues:d.filter(r=>r.flags.length).length}}
function heading(title,sub='',buttons=''){return `<div class="page-head"><div><h1>${esc(title)}</h1>${sub?`<p>${esc(sub)}</p>`:''}</div><div class="links">${buttons}</div></div>`}
function cards(arr){return `<div class="cards">${arr.map(x=>`<div class="metric"><small>${esc(x[0])}</small><strong>${esc(x[1])}</strong><small>${esc(x[2]||'')}</small></div>`).join('')}</div>`}
function onlyDemo(){return !complete()?`<div class="notice compact-note"><b>Sample figures.</b> Upload all four indent reports for verified results.</div>`:''}
function categories(){let x=new Map();for(let r of records){let c=x.get(r.cat)||{cat:r.cat,amount:0,qty:0,need:0,count:0};c.amount+=r.amount||0;c.qty+=r.salesQty||0;c.need+=r.need||0;c.count++;x.set(r.cat,c)}return [...x.values()].sort((a,b)=>b.amount-a.amount)}
function overview(){const s=summary(),cats=categories(),max=Math.max(1,...cats.map(x=>x.amount));return heading('Executive Dashboard','Your daily store and warehouse decisions, in one intelligent workspace.',`<button class="btn" data-page="upload">Import reports ↗</button>`)+onlyDemo()+`<div class="hero"><div><h2>Know what to indent. Know why.</h2><p>Review SKU matches, calculate exact quantities, compare warehouse stock and approve daily indents—without losing the decisions made on the shop floor.</p></div><button class="btn orange" data-page="indent">Open Indent Workspace →</button></div>`+cards([['Sales revenue',money(s.revenue),'Actual sales Total Amount'],['Calculated requirement',fmt(s.req),'Before warehouse cap'],['Warehouse shortage SKUs',fmt(s.short),'Demand exceeds available stock'],['Review exceptions',fmt(s.issues),'Includes missing sales and expiry issues']])+`<div class="grid2"><section class="panel"><h3>Category-wise sales</h3><p class="muted small">Actual sales report amounts, not inferred prices</p>${cats.slice(0,9).map(c=>`<div class="row-item"><div style="width:100%"><div class="flexline"><strong>${esc(c.cat)}</strong><span>${money(c.amount)}</span></div><div class="bar" style="margin-top:9px"><span style="width:${c.amount/max*100}%"></span></div></div></div>`).join('')}</section><section class="panel"><div class="panel-head"><h3>Action center</h3><button class="ghost" data-page="exceptions">View alerts →</button></div>${records.filter(r=>r.flags.length).slice(0,9).map(r=>`<div class="row-item"><div><b>${esc(r.name)}</b><div class="muted small">${esc(r.sku)} · ${esc(r.flags.join(', '))}</div></div><span class="pill warn">Review</span></div>`).join('')||'<p>No current alerts</p>'}</section></div><div class="footer">${complete()?'Calculated using the four uploaded indent reports.':'Illustrative figures · Import reports to view actual results.'}</div>`}
function upload(){return heading('Import & Match','Upload reports one at a time. Check the matching preview before calculating.')+`<div class="upload-grid">${[['store','1 · Store stock','Column B Item Code, D Category, F Quantity'],['warehouse','2 · Warehouse stock','Column B Item Code, F Quantity'],['sales','3 · Sales report','Only Item, Total Quantity, Total Amount'],['pending','4 · Previous indents','ItemCode, PendingQty, ExpiryDate']].map(([k,title,desc])=>`<div class="upload-card ${uploads[k]?'loaded':''}"><div class="glyph">${uploads[k]?'✓':'⇧'}</div><h3>${title}</h3><p>${desc}</p><label class="file-upload-action">⇧ &nbsp; Choose report<input type="file" accept=".xlsx,.xls,.csv" data-upload="${k}"></label><p>${uploads[k]?fmt(uploads[k].length)+' records ready':'No data uploaded'}</p></div>`).join('')}</div><section class="panel"><div class="panel-head"><div><h3>Total Sales Report · Data Analysis only</h3><p class="muted small">Separate from the four indent reports. Never changes required or final indent quantities.</p></div><span class="pill">ANALYTICS</span></div><div class="analytics-upload"><div><strong>⇧ Upload Total Sales Report</strong><p class="muted small">Reads CreatedDate, ItemCode, ItemName, Category, Quantity, TotalAmount and ProfitCentre.</p><p class="muted small">${analyticsMeta?fmt(analyticsRows.length)+' sales records ready':'No sales data loaded'}</p></div><label class="file-upload-action">⇧ &nbsp; Choose sales report<input type="file" data-upload="analytics" accept=".xlsx,.xls,.csv"></label></div><button class="ghost" data-page="dataAnalysis">Open Data Analysis →</button></section><div class="panel"><div class="panel-head"><div><h3>Analysis settings</h3><p class="muted small">Editable before each calculation</p></div><span class="pill">30-day default stock cover</span></div><div class="filters"><label>Sales period <select id="salesDays"><option value="30" ${salesDays===30?'selected':''}>30 days</option><option value="60" ${salesDays===60?'selected':''}>60 days</option></select></label><label>Global stock cover <input id="baseCover" type="number" min="1" max="365" value="${baseCover}"> days</label><button class="ghost" id="categoryCover">Configure category targets</button></div><p class="muted small">Use category overrides whenever store conditions differ. Sales report filenames do not control the selected days.</p></div><div class="panel"><div class="panel-head"><h3>Match diagnostics</h3><span class="pill ${complete()?'':'warn'}">${complete()?'All four reports present':'Demo / incomplete reports'}</span></div><div class="statusline"><span>Store SKUs: <b>${fmt(stats.store)}</b></span><span>Exact-name matches: <b>${fmt(stats.matched)}</b></span><span>No sales match: <b>${fmt(stats.missingSales)}</b></span><span>Warehouse missing: <b>${fmt(stats.missingWarehouse)}</b></span><span>Expired PO lines: <b>${fmt(stats.expiredLines)}</b></span><span>Unknown PO expiry: <b>${fmt(stats.unknownExpiry)}</b></span></div>${stats.duplicates?.length?`<div class="notice">Duplicate Item Codes detected: ${esc(stats.duplicates.slice(0,5).join('; '))} — review before approval.</div>`:''}<div class="filters"><input id="matchSearch" placeholder="Search store product / Item Code"><select id="matchFilter"><option value="all">All matches</option><option value="unmatched">Unmatched or ambiguous</option><option value="warehouse">Missing warehouse stock</option></select><button class="btn" id="calculateBtn" ${!complete()?'disabled title="Upload all four real files first"':''}>Calculate indents →</button></div><div class="table-wrap" id="matchTable">${matchTable()}</div><p class="muted small">Select an unmatched store product to manually map its sales name. Exact product-name matching is otherwise required.</p></div>`}
function matchTable(){const q=norm($('#matchSearch')?.value),f=$('#matchFilter')?.value||'all',rows=matches.filter(r=>(!q||(r.name+' '+r.sku).toLowerCase().includes(q))&&(f==='all'||f==='unmatched'&&!r.matched||f==='warehouse'&&r.warehouse===null)).slice(0,300);return `<table><thead><tr><th>Item Code</th><th>Product (Store)</th><th>Category</th><th>Store Qty</th><th>Warehouse Qty</th><th>Sales Qty</th><th>Match</th><th>Action</th></tr></thead><tbody>${rows.map(r=>`<tr><td>${esc(r.sku)}</td><td><span class="product">${esc(r.name)}</span></td><td>${esc(r.cat)}</td><td>${fmt(r.qty)} ${esc(r.uom)}</td><td>${r.warehouse===null?'—':fmt(r.warehouse)}</td><td>${r.salesQty===null?'—':fmt(r.salesQty)}</td><td><span class="pill ${r.matched?'':'warn'}">${r.matched?'Exact match':r.ambiguous?'Ambiguous':'Unmatched'}</span></td><td><button class="ghost" data-map="${esc(r.sku)}">Map</button></td></tr>`).join('')}</tbody></table>${rows.length===300?'<p class="muted small">Showing first 300 matching records; refine search to inspect others.</p>':''}`}
function tableRows(list,editable){return `<div class="table-wrap"><table><thead><tr><th>Item Code</th><th>Product</th><th>Category</th><th>60/30d Sales</th><th>Sales ₹</th><th>Store Qty</th><th>Valid Pending</th><th>WH Qty</th><th>Cover Days</th><th>Requirement</th><th>Suggested</th><th>Final Qty</th><th>Flags</th><th>Review</th></tr></thead><tbody>${list.slice(0,800).map(r=>`<tr><td>${esc(r.sku)}</td><td><span class="product">${esc(r.name)}</span><span class="muted small">${esc(r.uom)}</span></td><td>${esc(r.cat)}</td><td>${r.salesQty===null?'No data':fmt(r.salesQty)}</td><td>${r.amount===null?'—':money(r.amount)}</td><td>${fmt(r.qty)}</td><td>${fmt(r.pending)}</td><td>${r.warehouse===null?'Unavailable':fmt(r.warehouse)}</td><td>${fmt(r.coverDays)}</td><td>${r.need===null?'Review':fmt(r.need)}</td><td>${fmt(r.suggest)}</td><td>${editable?`<input class="editqty" type="number" min="0" step="any" data-edit="${esc(r.sku)}" value="${r.final}" ${approved?'disabled':''}>`:fmt(r.final)}</td><td>${r.flags.length?r.flags.map(f=>`<span class="pill warn" style="display:inline-block;margin:2px">${esc(f)}</span>`).join(''):'<span class="pill">OK</span>'}</td><td><button class="ghost" data-detail="${esc(r.sku)}">Details</button></td></tr>`).join('')}</tbody></table>${list.length>800?'<div class="notice">Showing first 800 rows; use search and category filters to inspect remaining SKUs. All records are included in export.</div>':''}</div>`}
function filtered(){return records.filter(r=>(!filter||(r.name+' '+r.sku).toLowerCase().includes(norm(filter)))&&(indentCategory==='all'||r.cat===indentCategory)&&(tab==='all'||tab==='alerts'&&r.flags.length||tab==='positive'&&r.need>0||tab==='nosales'&&r.salesQty===null||tab==='zero'&&r.salesQty===0||tab==='shortage'&&r.flags.includes('Warehouse shortage')))}
function indent(){const s=summary();return heading('Indent Workspace','Review sales-based requirements and final replenishment quantities.',`<button class="btn" id="exportXlsx">Export detailed Excel ↓</button>`)+onlyDemo()+`<div class="dev-notice"><strong>Under active development.</strong> This workspace is usable for calculations and review. Please verify suggested quantities against your store process and share feedback for improvements.</div>`+cards([['Total required',fmt(s.req),'Calculated requirement'],['Warehouse-capped suggested',fmt(records.reduce((a,r)=>a+r.suggest,0)),'Before manager adjustments'],['Final indent quantity',fmt(s.final),'Editable by manager'],['Shortage SKU count',fmt(s.short),'Warehouse below requirement']])+`<section class="panel"><div class="panel-head"><h3>Indent controls</h3><span class="pill ${approved?'':'warn'}">${approved?'Approved · manager-unlock required':'Draft · editable'}</span></div><div class="filters"><label>Sales window <select id="salesDays"><option value="30" ${salesDays===30?'selected':''}>30 days</option><option value="60" ${salesDays===60?'selected':''}>60 days</option></select></label><label>Default cover <input type="number" id="baseCover" min="1" max="365" value="${baseCover}"> days</label><button class="ghost" id="categoryCover">Category cover settings</button><button class="ghost" id="recalc">Recalculate</button><button class="btn" id="approve" ${approved?'disabled':''}>Approve & Finalize</button>${approved?'<button class="ghost" id="unlock">Manager unlock</button>':''}</div><div class="filters"><input id="indentSearch" value="${esc(filter)}" placeholder="Search Item Code or product"><span class="muted small">${fmt(filtered().length)} products</span></div><div class="filter-label">FILTER INDENTS BY CATEGORY</div><div class="category-scroller">${[['all','All categories'],...[...new Set(records.map(r=>r.cat))].sort().map(c=>[c,c])].map(([id,name])=>`<button class="cat-chip ${indentCategory===id?'active':''}" data-indent-cat="${esc(id)}">${esc(name)} <span>${fmt(id==='all'?records.length:records.filter(r=>r.cat===id).length)}</span></button>`).join('')}</div><div class="tabs">${[['all','All SKUs'],['positive','Positive requirement'],['alerts','Exceptions'],['shortage','Warehouse shortages'],['nosales','No sales data'],['zero','Zero sales']].map(([id,name])=>`<button data-tab="${id}" class="${tab===id?'selected':''}">${name}</button>`).join('')}</div><div id="indentRows">${tableRows(filtered(),true)}</div></section>`}
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
  const passesFilter=r=>(analysisFloor==='all'||r.floor===analysisFloor)&&(analysisCategory==='all'||(r.cat||'Unmapped category')===analysisCategory)&&(analysisSubcategory==='all'||r.subcat===analysisSubcategory)&&(!analysisSearch||norm(r.name+' '+r.sku).includes(norm(analysisSearch)));
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
  return {fromReport,rows,visible,cats,skus,floors,months,categoryChoices,subcategories,daily,monthly,dates:daily.map(x=>x.day),total:visible.reduce((x,r)=>x+r.amount,0),qty:visible.reduce((x,r)=>x+(r.qty||0),0)};
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
function getGrnLabels(){
  const detect=variants=>grnHeaders.findIndex(h=>variants.includes(norm(h).replace(/[^a-z0-9]/g,'')));
  return {date:detect(['date','grndate','receiptdate','receiveddate','documentdate']),no:detect(['grnno','grnnumber','grn','grncode','receiptnumber','receiptno']),item:detect(['itemcode','sku','skucode','itemname','productname','particulars']),qty:detect(['receivedqty','receivedquantity','quantity','qty','grnqty','receiptqty'])};
}
function grn(){
  const count=grnRows.length,columns=getGrnLabels(),query=norm(grnSearch);
  const found=count?grnRows.filter(r=>!query||r.some(x=>norm(x).includes(query))):[];
  const selected=grnShowAll?found.slice(0,500):found.slice(0,80);
  const lines=grnHeaders.length;
  const sampleText='Upload your store GRN Excel report to preview receipts. The exact columns will be confirmed from your file.';
  return heading('Store GRN Reports','Store goods receipt records and verification.')+`
   <section class="panel grn-landing"><div class="panel-head"><div><h3>GRN Report</h3><p class="muted small">Goods receipt register · separate from stock indent calculations</p></div><span class="pill">STORE GRN</span></div>
   <div class="grn-upload"><div><strong>${count?'GRN report ready':'Upload GRN report'}</strong><p class="muted small">${count?fmt(count)+' rows · '+fmt(lines)+' columns detected':sampleText}</p></div><label class="file-upload-action">⇧ Choose Excel file<input type="file" data-upload="grn" accept=".xlsx,.xls,.csv"></label></div>
   ${count?`<div class="grn-kpis"><div><span>Receipt rows</span><strong>${fmt(count)}</strong></div><div><span>Detected columns</span><strong>${fmt(lines)}</strong></div><div><span>Search matches</span><strong id="grnMatchCount">${fmt(found.length)}</strong></div></div><div class="grn-detection">${columns.no>=0?'GRN number · ':''}${columns.date>=0?'Receipt date · ':''}${columns.item>=0?'Item · ':''}${columns.qty>=0?'Quantity':''}${Object.values(columns).every(v=>v<0)?'Preview mode · column mapping pending':''}</div><div class="panel-head grn-table-head"><div><h3>GRN records</h3><p class="muted small">Previewing up to ${grnShowAll?500:80} matching lines</p></div><input type="search" id="grnSearch" value="${esc(grnSearch)}" placeholder="Search receipt, item or date" aria-label="Search GRN report"></div><div id="grnTable">${grnTable(selected)}</div><button type="button" class="ghost" id="grnMore" ${found.length<=80?'disabled':''}>${grnShowAll?'Show fewer rows':'Show more rows'}</button>`:`<div class="grn-empty"><span aria-hidden="true">▣</span><strong>No GRN report uploaded</strong><p>Import an Excel report to review your store receipts.</p></div>`}
   </section>`;
}
function grnTable(rows){
  return `<div class="table-wrap grn-table-wrap"><table><thead><tr>${grnHeaders.slice(0,20).map(h=>`<th>${esc(h||'Unnamed column')}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${grnHeaders.slice(0,20).map((_,i)=>`<td>${esc(r[i] instanceof Date?r[i].toLocaleDateString('en-IN'):r[i])}</td>`).join('')}</tr>`).join('')||`<tr><td colspan="${Math.max(1,Math.min(20,grnHeaders.length))}">No matching records</td></tr>`}</tbody></table></div>`;
}
function grnUpdateSearch(){
  const query=norm(grnSearch),found=grnRows.filter(r=>!query||r.some(x=>norm(x).includes(query)));
  const counter=$('#grnMatchCount');if(counter)counter.textContent=fmt(found.length);
  const node=$('#grnTable');if(node)node.innerHTML=grnTable(found.slice(0,grnShowAll?500:80));
}
function dataAnalysis(){
  const d=dataGroup(),floorName=analysisFloor==='all'?'Entire store':analysisFloor,catName=analysisCategory==='all'?'All categories':analysisCategory,monthName=analysisMonth==='all'?'All months':monthLabel(analysisMonth);
  currentAnalyticsView=d;
  return heading('Data Analysis')+`${!d.fromReport?'<div class="notice compact-note">Upload the Total Sales Report to view actual floor, category and SKU performance.</div>':''}
  <section class="panel smart-filter-panel"><div class="panel-head"><div><h3>Smart Filters</h3><p class="muted small">A single set of filters for all charts and daily sales.</p></div><button class="ghost" id="resetAnalytics" type="button">↺ Reset all</button></div>
  <div class="smart-filter-grid">
    <label class="smart-filter-field"><span>Month</span><select id="smartMonth"><option value="all">All months</option>${d.months.map(m=>`<option value="${esc(m)}" ${analysisMonth===m?'selected':''}>${esc(monthLabel(m))}</option>`).join('')}</select></label>
    <label class="smart-filter-field"><span>Floor</span><select id="smartFloor"><option value="all">Entire store</option>${d.floors.map(f=>`<option value="${esc(f)}" ${analysisFloor===f?'selected':''}>${esc(f)}</option>`).join('')}</select></label>
    <div class="smart-filter-field smart-category-field"><label for="analysisCategorySelect">Category · A–Z</label><input type="search" id="analysisCategoryFilter" placeholder="Search categories" value="${esc(analysisCategoryQuery)}" aria-label="Search categories alphabetically"><select id="analysisCategorySelect" aria-label="Select category"><option value="all">All categories</option>${d.categoryChoices.map(cat=>`<option value="${esc(cat)}" ${analysisCategory===cat?'selected':''}>${esc(cat)}</option>`).join('')}</select></div>
    <label class="smart-filter-field"><span>Subcategory</span><select id="smartSubcategory"><option value="all">All subcategories</option>${d.subcategories.map(x=>`<option value="${esc(x)}" ${analysisSubcategory===x?'selected':''}>${esc(x)}</option>`).join('')}</select></label>
    <label class="smart-filter-field"><span>SKU / Product</span><input id="analysisSearch" type="search" value="${esc(analysisSearch)}" placeholder="Name or item code" aria-label="Search SKU or product"></label>
  </div>
  <div class="filter-active-line" aria-live="polite"><span class="filter-active-dot"></span><span>${fmt(d.visible.length)} transactions · ${fmt(d.skus.length)} SKUs · ${fmt(d.cats.length)} categories</span><span class="filter-context">${esc(monthName)} · ${esc(floorName)}</span></div>
  </section>
  ${cards([['Sales value',money(d.total),monthName+' · '+floorName+' · '+catName],['Quantity sold',fmt(d.qty),'Mixed UOM: interpret with care'],['SKUs',fmt(d.skus.length),'Visible distinct product lines'],['Categories',fmt(analysisCategory==='all'?d.cats.length:Math.min(1,d.cats.length)),'Selected month and floor']])}
  <section class="panel"><div class="panel-head"><div><h3>${analysisCategory==='all'?'Category overview · top 8 by revenue':esc(analysisCategory)+' · sales performance'}</h3><p class="muted small">${analysisCategory==='all'?'Select a category from the filter above to inspect its SKU results.':'Revenue and contribution for the selected category.'}</p></div><span class="pill">${analysisCategory==='all'?fmt(d.cats.length)+' available categories':'1 selected category'}</span></div>${analysisCategory==='all'?chartRows(d.cats,8):chartRows(d.cats.filter(x=>x.cat===analysisCategory),1)}${analysisCategory==='all'&&d.cats.length>8?'<p class="muted small">Only the top 8 are shown here to keep the dashboard clean. Use Filter by Category to access every category.</p>':''}</section>
  <section class="panel"><div class="panel-head"><div><h3>Monthly sales performance</h3><p class="muted small">Detected automatically from CreatedDate · Month names and revenue comparisons</p></div><span class="pill">${fmt(d.months.length)} months detected</span></div>${d.monthly.length?`<div class="monthly-grid">${d.monthly.map((r,i)=>{const previous=i?d.monthly[i-1]:null;const growth=previous&&previous.amount!==0?(r.amount-previous.amount)/Math.abs(previous.amount)*100:null;return `<button class="month-stat ${analysisMonth===r.month?'selected':''}" data-analytics-month="${r.month}"><span class="month-heading">${esc(monthLabel(r.month))}</span><strong>${money(r.amount)}</strong><small>${fmt(r.dates.size)} trading dates · ${fmt(r.qty)} units</small><span class="month-change ${growth===null?'neutral':growth>=0?'positive':'negative'}">${growth===null?'First available month':`${growth>=0?'+':''}${fmt(growth)}% vs previous month`}</span></button>`}).join('')}</div><p class="muted small">Comparisons are between available report periods; partial months are not directly comparable to full months.</p>`:'<div class="analytics-missing">Month-wise analysis requires CreatedDate in the Total Sales Report.</div>'}</section>
   <section class="panel"><div class="panel-head"><div><h3>Daily sales trend</h3><p class="muted small">Click any date bar to inspect that day's actual sales, categories and SKUs.</p></div><span class="pill ${d.dates.length?'':'warn'}">${d.dates.length?fmt(d.dates.length)+' dates':'Date data unavailable'}</span></div>${d.daily.length?`${dailyChart(d)}${dailyDetails(d)}<div class="table-wrap daily-table"><table style="min-width:520px"><thead><tr><th>Date</th><th>Sales Value</th><th>Reported Qty</th><th>View</th></tr></thead><tbody>${d.daily.map(r=>`<tr class="${((selectedSalesDay&&d.dates.includes(selectedSalesDay)?selectedSalesDay:d.dates[d.dates.length-1])===r.day)?'active':''}"><td>${esc(r.day)}</td><td>${money(r.amount)}</td><td>${fmt(r.qty)}</td><td><button type="button" class="daily-view-btn" data-sales-day="${esc(r.day)}">Inspect day →</button></td></tr>`).join('')}</tbody></table></div>`:'<div class="analytics-missing">Day-by-day figures cannot be generated from a 30/60-day total. Please upload dated transactions or daily summary data for a valid daily trend.</div>'}</section>
  <section class="panel"><div class="panel-head"><div><h3>SKU performance · ${esc(catName)}</h3><p class="muted small">Sales value is recorded Total Amount, not estimated inventory value.</p></div><span class="pill">${fmt(d.skus.length)} SKUs</span></div><div class="table-wrap"><table style="min-width:700px"><thead><tr><th>Product</th><th>Item Code</th><th>Category</th><th>Qty sold</th><th>Sales Amount</th><th>Sales mix</th></tr></thead><tbody id="analysisSkuRows">${analysisSkuRows(d.skus,d.total)}</tbody></table></div></section>
  <section class="panel"><div class="panel-head"><div><h3>Total Sales Report upload</h3><p class="muted small">Dedicated analytics report · never used for indents</p></div><span class="pill">DATA ANALYSIS ONLY</span></div><div class="analytics-upload"><div><strong>Upload Total Sales Report (.xlsx)</strong><p class="muted small">Reads ProfitCentre, Category, ItemCode, ItemName, Quantity, TotalAmount and CreatedDate.</p>${analyticsMeta?`<p class="muted small">${fmt(analyticsRows.length)} transactions ready</p>`:'<p class="muted small">No sales report loaded.</p>'}</div><label class="file-upload-action">⇧ &nbsp; Choose sales report<input type="file" data-upload="analytics" accept=".xlsx,.xls,.csv"></label></div>${analyticsMeta?'<button class="ghost" id="clearAnalytics">Remove Total Sales Report</button>':''}</section>`;
}
function analysisSkuRows(rows,total){return rows.slice(0,700).map(r=>`<tr><td><span class="product">${esc(r.name)}</span></td><td>${esc(r.sku||'Not mapped')}</td><td>${esc(r.cat)}</td><td>${fmt(r.qty)}</td><td><b>${money(r.amount)}</b></td><td>${total?fmt(r.amount/total*100):'0'}%</td></tr>`).join('')||'<tr><td colspan="6">No matching products</td></tr>'}
function exceptions(){let arr=records.filter(r=>r.flags.length),types=[...new Set(arr.flatMap(x=>x.flags))];return heading('Exceptions & Reviews','Missing data, shortages, expired POs, excessive stock and manual decisions.')+onlyDemo()+cards([['Flagged SKUs',fmt(arr.length),'Require attention'],['Missing sales',fmt(records.filter(x=>x.salesQty===null).length),'Not treated as zero sales'],['Warehouse missing',fmt(records.filter(x=>x.warehouse===null).length),'Not silently treated as available'],['Expired PO lines',fmt(stats.expiredLines),'Excluded from pending deduction']])+`<div class="panel"><h3>Exceptions by type</h3>${types.map(f=>`<div class="row-item"><b>${esc(f)}</b><span>${fmt(arr.filter(r=>r.flags.includes(f)).length)} SKUs</span></div>`).join('')||'<p>No exceptions.</p>'}</div><div class="panel"><h3>Flagged products</h3>${tableRows(arr,false)}</div>`}
function hist(){return heading('Daily History','Saved local versions and manager approvals. Restore a prior snapshot.')+`<div class="notice">This local application stores versions in your browser using IndexedDB. Clearing browser data or switching devices may remove them. This is not an off-device daily backup or secure role enforcement.</div><section class="panel"><div class="panel-head"><h3>Saved versions</h3></div>${history.length?history.map(r=>`<div class="row-item"><div><b>${esc(r.date)} · Version ${r.version}</b><div class="muted small">${esc(r.created)} · ${r.approved?'Approved':'Draft'} · ${fmt(r.records.length)} SKUs</div></div><button class="ghost" data-restore="${r.id}">Restore snapshot</button></div>`).join(''):'<div class="empty">No saved versions yet.</div>'}</section>`}
const AUTH_EMPLOYEE_ID='15907';
const AUTH_EMPLOYEE_NAME='Pranav Valluru';
const AUTH_PASSWORD_SHA256='f4845f23932b043ef3f5e3d6368640b4481f064fc41339ad94667b9b567e466c';
let temporaryEmployeeSession=false;
function hasEmployeeSession(){
  try { if(sessionStorage.getItem('tb-employee-session')===AUTH_EMPLOYEE_ID)return true; }
  catch(_err){}
  return temporaryEmployeeSession;
}
function applyLoginState(){
  const active=hasEmployeeSession();
  document.body.classList.toggle('login-locked',!active);
  const portal=document.getElementById('loginPortal');
  if(portal)portal.hidden=active;
  const id=document.getElementById('signedInId'),name=document.getElementById('signedInName');
  if(id)id.textContent=AUTH_EMPLOYEE_ID;
  if(name)name.textContent=AUTH_EMPLOYEE_NAME;
  if(!active)setTimeout(()=>document.getElementById('employeeId')?.focus(),40);
}
let loginInProgress=false;
const loginPause=ms=>new Promise(resolve=>setTimeout(resolve,ms));
async function signInEmployee(event){
  event.preventDefault();
  if(loginInProgress)return;
  const employeeId=document.getElementById('employeeId').value.trim();
  const password=document.getElementById('employeePassword').value;
  const error=document.getElementById('loginError');
  const form=document.getElementById('loginForm');
  const submit=form.querySelector('.login-submit');
  if(employeeId!==AUTH_EMPLOYEE_ID){error.textContent='Incorrect employee ID or password.';return;}
  loginInProgress=true;
  submit.disabled=true;
  submit.classList.add('signing-in');
  submit.querySelector('.login-submit-label').textContent='Signing in';
  error.textContent='';
  try {
    if(!crypto?.subtle)throw Error('Browser does not support the local credential check.');
    const digest=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(password));
    const encoded=Array.from(new Uint8Array(digest),x=>x.toString(16).padStart(2,'0')).join('');
    if(encoded!==AUTH_PASSWORD_SHA256){error.textContent='Incorrect employee ID or password.';return;}
    // Visual transition for the local HTML prototype, not a server authentication step.
    const transition=document.getElementById('loginLoading');
    transition.hidden=false;
    transition.classList.remove('loading-exit');
    const loadingMessage=document.getElementById('loadingStatus');
    const progress=transition.querySelector('.loading-progress');
    loadingMessage.textContent='Preparing your dashboard…';
    progress.setAttribute('aria-valuenow','0');
    // Deliberate, non-blocking 4.3-second visual transition after valid sign-in.
    // This is presentation feedback only, not a server-side authentication check.
    const started=performance.now();
    const duration=4300;
    let interval=setInterval(()=>{
      const elapsed=performance.now()-started;
      const percent=Math.min(100,Math.floor(elapsed/duration*100));
      progress.setAttribute('aria-valuenow',String(percent));
      loadingMessage.textContent=elapsed<1450?'Preparing your dashboard…':elapsed<3050?'Setting up your workspace…':'Almost ready…';
    },100);
    try { await loginPause(duration); } finally { clearInterval(interval); }
    progress.setAttribute('aria-valuenow','100');
    loadingMessage.textContent='Your dashboard is ready';
    transition.classList.add('loading-exit');
    await loginPause(260);
    temporaryEmployeeSession=true;
    try{sessionStorage.setItem('tb-employee-session',AUTH_EMPLOYEE_ID)}catch(_err){}
    form.reset();
    applyLoginState();
    go('overview');
    transition.hidden=true;
    transition.classList.remove('loading-exit');
  } catch(err) {
    document.getElementById('loginLoading').hidden=true;
    error.textContent='Sign-in unavailable in this browser. Try a recent Chrome browser.';
  } finally {
    loginInProgress=false;
    submit.disabled=false;
    submit.classList.remove('signing-in');
    submit.querySelector('.login-submit-label').textContent='Sign in';
  }
}
function logoutEmployee(){
  temporaryEmployeeSession=false;
  try{sessionStorage.removeItem('tb-employee-session');}catch(_err){}
  window.location.reload();
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
const views={overview,grn,upload,indent,dataAnalysis,exceptions},names={overview:'Executive Dashboard',grn:'GRN Report',upload:'Import & Match',indent:'Indent Workspace',dataAnalysis:'Data Analysis',exceptions:'Exceptions'};
function setGrnMenu(open){const menu=$('#grnMenu'),panel=$('#grnSubmenu');if(!menu||!panel)return;panel.hidden=!open;menu.setAttribute('aria-expanded',String(open));menu.classList.toggle('expanded',open)}
function render(){if(!records.length)calculate();if(page==='grn')setGrnMenu(true);$$('nav button').forEach(x=>x.classList.toggle('selected',x.dataset.page===page));$('#breadcrumb').textContent=names[page];const idNode=$('#signedInId');if(idNode)idNode.textContent=AUTH_EMPLOYEE_ID;$('#view').innerHTML=views[page]();if(page==='dataAnalysis'&&analysisCategoryQuery)searchCategoryOptions(analysisCategoryQuery);document.querySelector('.side').classList.remove('open')}
function go(p){if(views[p]){page=p;render();document.querySelector('.main-shell').scrollTop=0}}
async function loadFile(file,type){if(!file)return;if(typeof XLSX==='undefined'){flash('Excel library could not load. Check your internet connection.');return}try{let buf=await file.arrayBuffer(),wb=XLSX.read(buf,{type:'array',cellDates:true}),rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,raw:true,defval:null,blankrows:false});if(type==='grn'){const hdr=rows.findIndex(r=>r.filter(x=>x!==null&&x!==undefined&&key(x)!=='').length>=2);if(hdr<0)throw Error('The GRN file contains no readable table.');grnHeaders=rows[hdr].map((v,i)=>key(v)||'Column '+(i+1));grnRows=rows.slice(hdr+1).filter(r=>r.some(v=>v!==null&&v!==undefined&&key(v)!==''));grnLabel=file.name;grnSearch='';grnShowAll=false;go('grn');flash('GRN report loaded: '+fmt(grnRows.length)+' rows');return;}if(type==='analytics'){const result=parseAnalytics(rows);analyticsRows=result.rows;analyticsMeta={name:file.name,hasFloor:result.hasFloor,hasDay:result.hasDay};analysisMonth='all';analysisFloor='all';analysisCategory='all';analysisSubcategory='all';analysisSearch='';analysisCategoryQuery='';selectedSalesDay=null;go('dataAnalysis');flash('Additional sales report loaded: '+fmt(analyticsRows.length)+' rows');return}let parsed=type==='store'||type==='warehouse'?parseStock(rows):type==='sales'?parseSales(rows):parsePending(rows);if(!parsed.length)throw Error('No valid item rows were detected.');const signature=`${file.name}|${file.size}|${file.lastModified}`;if(Object.values(meta).some(x=>x.sig===signature)){if(!confirm('This exact file has already been uploaded. Replace anyway?'))return}uploads[type]=parsed;meta[type]={name:file.name,sig:signature};edits={};approved=false;review();calculate();render();flash(`${type} loaded: ${fmt(parsed.length)} records`)}catch(err){flash('Import error: '+err.message);console.error(err)}}
function modalMap(sku){const item=matches.find(x=>x.sku===sku),opts=[...new Set(datasets().sales.map(r=>r.name))];if(!item)return;const current=mappings[sku]||'';const suggested=opts.filter(x=>norm(x).includes(norm(item.name).slice(0,11))).slice(0,12);let chosen=prompt(`Map sales product to store SKU ${sku}: ${item.name}\n\nEnter EXACT sales Item name.\nSuggestions:\n${suggested.join('\n')}`,current||suggested[0]||'');if(chosen===null)return;if(!opts.includes(chosen.trim())){flash('Name must exactly match an Item value in the sales report.');return}mappings[sku]=chosen.trim();review();calculate();render();flash('Product manually mapped; inspect updated matching results.')}
function categorySetting(){let cats=[...new Set(datasets().store.map(r=>r.cat))].sort();let list=cats.length>40?cats.slice(0,40):cats;let message=`Set category cover days. Default ${baseCover}.\n\n`+list.map(x=>`${x}: ${covers[x]??baseCover}`).join('\n');let chosen=prompt(message+'\n\nEnter EXACT category name to change:');if(!chosen)return;if(!cats.includes(chosen)){flash('Category not found. Enter exact category name.');return}let days=prompt('New stock-cover days for '+chosen,String(covers[chosen]??baseCover));if(!days)return;if(!(Number(days)>0&&Number(days)<=365)){flash('Enter 1–365 days.');return}covers[chosen]=Number(days);calculate();render();flash('Cover updated for '+chosen)}
function detail(sku){let r=records.find(x=>x.sku===sku);if(!r)return;let remark=prompt(`${r.name}\nItem Code: ${r.sku}\nStore ${r.qty}, Warehouse ${r.warehouse??'missing'}, Valid pending ${r.pending}\nRequired ${r.need??'unknown'}, Suggested ${r.suggest}, Final ${r.final}\nFlags: ${r.flags.join('; ')||'None'}\n\nEnter review remarks (optional):`,r.remark||'');if(remark===null)return;r.remark=remark;flash('SKU review note updated for this session.')}
function updateEdit(sku,value){let r=records.find(x=>x.sku===sku);if(!r||approved)return;if(value===''||!Number.isFinite(Number(value))||Number(value)<0){flash('Final quantity must be a non-negative number');return}let qty=Number(value);r.final=qty;edits[sku]=qty;if(qty>(r.warehouse??0)&&!r.flags.includes('Manager qty exceeds warehouse'))r.flags.push('Manager qty exceeds warehouse');if(qty<=(r.warehouse??0))r.flags=r.flags.filter(f=>f!=='Manager qty exceeds warehouse')}
function createDatabase(){return new Promise(resolve=>{try{if(typeof indexedDB==='undefined'){resolve(null);return}let q=indexedDB.open('todaysBasketHtmlV2',1);q.onupgradeneeded=()=>{if(!q.result.objectStoreNames.contains('history'))q.result.createObjectStore('history',{keyPath:'id'})};q.onsuccess=()=>resolve(q.result);q.onerror=()=>resolve(null)}catch(err){resolve(null)}})}
async function saveVersion(){if(!db){flash('Local history storage unavailable.');return}let last=history.filter(x=>x.date===today()).reduce((m,x)=>Math.max(m,x.version),0),row={id:crypto.randomUUID(),date:today(),version:last+1,created:new Date().toLocaleString('en-IN'),approved,records:JSON.parse(JSON.stringify(records)),settings:{salesDays,baseCover,covers,edits},actor:AUTH_EMPLOYEE_NAME,action:approved?'approved':'saved'};await new Promise((yes,no)=>{const t=db.transaction('history','readwrite');t.objectStore('history').put(row);t.oncomplete=yes;t.onerror=no});history.unshift(row);flash('Saved local version V'+row.version);if(page==='history')render()}
async function loadHistory(){if(!db)return;history=await new Promise(resolve=>{let t=db.transaction('history','readonly'),q=t.objectStore('history').getAll();q.onsuccess=()=>resolve(q.result.sort((a,b)=>b.created.localeCompare(a.created)));q.onerror=()=>resolve([])})}
function restore(id){let item=history.find(r=>r.id===id);if(!item||!confirm('Restore this saved analysis snapshot? Current unsaved edits may be lost.'))return;records=JSON.parse(JSON.stringify(item.records));salesDays=item.settings.salesDays;baseCover=item.settings.baseCover;covers=item.settings.covers;edits=item.settings.edits;approved=item.approved;go('indent');flash('Version restored locally.')}
function exportExcel(){if(typeof XLSX==='undefined'){flash('Excel library unavailable.');return}let header=['Item Code','Product','Category','UOM','Sales Period Days','Sales Total Qty','Sales Total Amount','Store Qty','Warehouse Qty','Valid Pending Qty','Expired Pending Qty','Unverified Pending Qty','Target Cover Days','30/60-day Target Qty','Required Qty','Warehouse-capped Suggested Qty','Manager Final Qty','Current Stock Cover Days','Flags','Review Remarks','Approval Status'];let rows=records.map(r=>[r.sku,r.name,r.cat,r.uom,salesDays,r.salesQty,r.amount,r.qty,r.warehouse,r.pending,r.expiredPending,r.unknownPending,r.coverDays,r.target,r.need,r.suggest,r.final,r.cover,r.flags.join('; '),r.remark||'',approved?'Approved':'Draft']);let wb=XLSX.utils.book_new(),ws=XLSX.utils.aoa_to_sheet([header,...rows]);ws['!cols']=[{wch:18},{wch:44},{wch:24},{wch:10},...Array(14).fill({wch:18}),{wch:48},{wch:35},{wch:18}];XLSX.utils.book_append_sheet(wb,ws,'Complete Indent Analysis');XLSX.writeFile(wb,`Todays_Basket_Indent_${today()}_${approved?'Approved':'Draft'}.xlsx`);flash('Complete analysis exported in one Excel worksheet')}
function approve(){if(!complete()){flash('Upload all four actual indent reports before approval.');return}let actor=prompt('Manager name / Employee ID for audit (local prototype only):');if(!actor)return;if(stats.duplicates.length&&!confirm('Duplicate Item Codes exist. Approve despite this data risk?'))return;if(records.some(r=>r.unknownPending>0)&&!confirm('Some PO expiry dates are unknown and their pending qty was NOT deducted. Approve anyway?'))return;if(!confirm(`Approve ${records.length} SKU rows as ${actor}? This HTML-only approval does not replace server authorization.`))return;approved=true;saveVersion();render();flash('Approved locally. Revisions are recorded only in this browser.')}
function unlock(){const who=prompt('Manager name / Employee ID unlocking approved indent:');if(!who)return;if(!confirm('Unlock this approved analysis for editing? Save a new version after changes.'))return;approved=false;flash('Unlocked by '+who+' (prototype audit only)');render()}
document.addEventListener('click',e=>{const t=e.target.closest('button');if(!t)return;if(t.id==='logout'){logoutEmployee();return;}if(t.id==='grnMenu'){setGrnMenu($('#grnMenu').getAttribute('aria-expanded')!=='true');return;}if(t.id==='grnMore'){grnShowAll=!grnShowAll;render();return;}if(t.id==='togglePassword'){const field=document.getElementById('employeePassword');field.type=field.type==='password'?'text':'password';t.textContent=field.type==='password'?'Show':'Hide';return;}if(t.dataset.page)go(t.dataset.page);if(t.hasAttribute('data-sales-day')){selectSalesDay(t.dataset.salesDay);return;}if(t.hasAttribute('data-day-move')){const d=currentAnalyticsView||dataGroup();const current=d.dates.includes(selectedSalesDay)?d.dates.indexOf(selectedSalesDay):d.dates.length-1;const next=d.dates[current+Number(t.dataset.dayMove)];if(next)selectSalesDay(next);return;}if(t.dataset.tab){tab=t.dataset.tab;render()}if(t.hasAttribute('data-indent-cat')){indentCategory=t.dataset.indentCat;render()}if(t.hasAttribute('data-analytics-cat')){analysisCategory=t.dataset.analyticsCat;analysisSubcategory='all';analysisCategoryQuery='';selectedSalesDay=null;render()}if(t.hasAttribute('data-analytics-month')){analysisMonth=t.dataset.analyticsMonth;analysisFloor='all';analysisCategory='all';analysisSubcategory='all';analysisCategoryQuery='';render()}if(t.hasAttribute('data-analytics-floor')){analysisFloor=t.dataset.analyticsFloor;analysisCategory='all';analysisSubcategory='all';analysisCategoryQuery='';render()}if(t.id==='resetAnalytics'){analysisMonth='all';analysisFloor='all';analysisCategory='all';analysisSubcategory='all';analysisSearch='';analysisCategoryQuery='';selectedSalesDay=null;render()}if(t.dataset.map)modalMap(t.dataset.map);if(t.dataset.detail)detail(t.dataset.detail);if(t.dataset.restore)restore(t.dataset.restore);switch(t.id){case'menu':$('.side').classList.toggle('open');break;case'reset':if(confirm('Clear uploaded indent reports? Locally saved history remains.')){uploads={};meta={};edits={};mappings={};approved=false;review();calculate();render()}break;case'calculateBtn':if(!complete())flash('Upload all four Excel reports first.');else{calculate();go('indent');flash('Calculated from uploaded reports. Review quantities before approval.')}break;case'categoryCover':categorySetting();break;case'recalc':calculate();render();flash('Analysis recalculated with current settings');break;case'approve':approve();break;case'unlock':unlock();break;case'exportXlsx':exportExcel();break;case'expandTags':window.open('./price-tags.html','_blank','noopener');break}});
document.addEventListener('change',e=>{const t=e.target;if(t.dataset.upload)loadFile(t.files?.[0],t.dataset.upload);if(t.id==='salesDays'){salesDays=+t.value;calculate();render()}if(t.id==='baseCover'){let k=+t.value;if(k>=1&&k<=365){baseCover=k;calculate();render()}else flash('Stock cover should be 1–365 days')}if(t.dataset.edit){updateEdit(t.dataset.edit,t.value);let focus=page;render();page=focus}});
document.addEventListener('change',e=>{
  const id=e.target.id;
  if(id==='smartMonth'){analysisMonth=e.target.value;analysisFloor='all';analysisCategory='all';analysisSubcategory='all';analysisCategoryQuery='';selectedSalesDay=null;render()}
  if(id==='smartFloor'){analysisFloor=e.target.value;analysisCategory='all';analysisSubcategory='all';analysisCategoryQuery='';selectedSalesDay=null;render()}
  if(id==='analysisCategorySelect'&&e.target.value){analysisCategory=e.target.value;analysisSubcategory='all';analysisCategoryQuery='';selectedSalesDay=null;render()}
  if(id==='smartSubcategory'){analysisSubcategory=e.target.value;selectedSalesDay=null;render()}
});
document.addEventListener('input',e=>{if(e.target.id==='analysisSearch'){analysisSearch=e.target.value;clearTimeout(window.analysisDebounce);window.analysisDebounce=setTimeout(()=>{const pos=$('#analysisSearch')?.selectionStart??analysisSearch.length;render();const el=$('#analysisSearch');if(el){el.focus();el.setSelectionRange(pos,pos)}},280)}if(e.target.id==='indentSearch'){filter=e.target.value;const r=$('#indentRows');if(r)r.innerHTML=tableRows(filtered(),page==='indent')}if(e.target.id==='grnSearch'){grnSearch=e.target.value;grnUpdateSearch()}if(e.target.id==='matchSearch'||e.target.id==='matchFilter'){const r=$('#matchTable');if(r)r.innerHTML=matchTable()}});
document.getElementById('loginForm').addEventListener('submit',signInEmployee);
document.addEventListener('input',e=>{if(e.target.id==='analysisCategoryFilter'){analysisCategoryQuery=e.target.value;searchCategoryOptions(e.target.value)}});
document.addEventListener('keydown',e=>{
 if(e.target.id==='analysisCategoryFilter'&&e.key==='Enter'){
   const select=document.getElementById('analysisCategorySelect');
   const first=[...select.options].find(x=>!x.disabled&&x.value!=='all');
   if(first){e.preventDefault();analysisCategory=first.value;render();}
 }
});
(async()=>{db=await createDatabase();await loadHistory();review();calculate();render();applyLoginState()})();

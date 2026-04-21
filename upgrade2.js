const fs = require('fs');
let h = fs.readFileSync('C:/Users/El-Wattaneya/Desktop/basset/index.html','utf8');

// Find the dashboard stat cards grid and replace it entirely
// Find from "grid-cols-6" to the closing </div> of that grid
const startMarker = 'xl:grid-cols-6 gap-4 mb-8">';
const startIdx = h.indexOf(startMarker);
if (startIdx === -1) { console.log('Start not found'); process.exit(1); }

// Find the next section after stat cards (charts section)
const endMarker = '<div class="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-8">';
const endIdx = h.indexOf(endMarker, startIdx);
if (endIdx === -1) { console.log('End not found'); process.exit(1); }

const newStats = `xl:grid-cols-6 gap-5 mb-8">
<div class="stat-card"><div class="stat-icon stat-icon-blue"><span class="material-symbols-outlined">trending_up</span></div><p class="stat-label text-xs text-slate-400 font-semibold mb-1">إجمالي المبيعات</p><p id="statTotalSales" class="text-2xl font-black text-primary font-headline currency">0.00</p><p id="statSalesCount" class="text-xs text-slate-400 mt-1"></p></div>
<div class="stat-card"><div class="stat-icon stat-icon-amber"><span class="material-symbols-outlined">shopping_cart</span></div><p class="stat-label text-xs text-slate-400 font-semibold mb-1">إجمالي المشتريات</p><p id="statTotalPurchases" class="text-2xl font-black text-primary font-headline currency">0.00</p></div>
<div class="stat-card"><div class="stat-icon stat-icon-green"><span class="material-symbols-outlined">account_balance</span></div><p class="stat-label text-xs text-slate-400 font-semibold mb-1">صافي الربح</p><p id="statProfit" class="text-2xl font-black font-headline currency" style="color:#166534">0.00</p></div>
<div class="stat-card"><div class="stat-icon stat-icon-red"><span class="material-symbols-outlined">credit_card_off</span></div><p class="stat-label text-xs text-slate-400 font-semibold mb-1">المديونيات</p><p id="statTotalDue" class="text-2xl font-black font-headline currency" style="color:#ef4444">0.00</p><p id="statCustomersCount" class="text-xs text-slate-400 mt-1"></p></div>
<div class="stat-card stat-card-accent"><div class="stat-icon stat-icon-white"><span class="material-symbols-outlined">savings</span></div><p class="stat-label text-xs font-semibold mb-1" style="color:rgba(255,255,255,.7)">رصيد الخزينة</p><p id="statCashInBox" class="text-2xl font-black font-headline currency" style="color:#fff">0.00</p></div>
<div class="stat-card"><div class="stat-icon stat-icon-amber"><span class="material-symbols-outlined">person_pin</span></div><p class="stat-label text-xs text-slate-400 font-semibold mb-1">عهدة المناديب</p><p id="statRepCustody" class="text-2xl font-black font-headline currency" style="color:#854d0e">0.00</p></div>
</div>
`;

h = h.substring(0, startIdx) + newStats + h.substring(endIdx);

// Now add the premium CSS classes that were missing
const cssInsertPoint = '.stat-card:hover{';
const cssIdx = h.indexOf(cssInsertPoint);
if (cssIdx > -1) {
  // Find end of current stat-card style
  const endCssIdx = h.indexOf('}', h.indexOf('}', cssIdx) + 1) + 1;
  const additionalCss = `
.stat-card-accent{background:linear-gradient(135deg,#031636 0%,#1a2b4c 60%,#0058bb 100%);color:#fff;border:none;position:relative;overflow:hidden}.stat-card-accent::before{content:'';position:absolute;top:-50%;right:-50%;width:100%;height:100%;background:radial-gradient(circle,rgba(255,255,255,.08) 0%,transparent 70%);pointer-events:none}
.stat-card-accent p,.stat-card-accent span{color:#fff!important}.stat-card-accent .stat-label{color:rgba(255,255,255,.7)!important}
.stat-icon{width:48px;height:48px;border-radius:14px;display:flex;align-items:center;justify-content:center;font-size:1.25rem;margin-bottom:.75rem}
.stat-icon-blue{background:rgba(0,88,187,.08);color:#0058bb}.stat-icon-green{background:rgba(22,101,52,.08);color:#166534}.stat-icon-red{background:rgba(153,27,27,.08);color:#991b1b}.stat-icon-amber{background:rgba(133,77,14,.08);color:#854d0e}
.stat-icon-white{background:rgba(255,255,255,.15);color:#fff}
.page-hero{padding:2rem 2.5rem;margin-bottom:0;background:linear-gradient(135deg,#f8f9ff 0%,#e6eeff 100%);border-radius:1.5rem;position:relative;overflow:hidden}
.page-hero::after{content:'';position:absolute;top:-30%;left:-20%;width:50%;height:160%;background:radial-gradient(ellipse,rgba(0,88,187,.04) 0%,transparent 70%);pointer-events:none}
.page-hero h1,.page-hero h2,.page-hero p{position:relative;z-index:1}
.section-label{font-size:.65rem;font-weight:700;text-transform:uppercase;letter-spacing:2px;color:#94a3b8;margin-bottom:.25rem;display:block}`;
  h = h.substring(0, endCssIdx) + additionalCss + h.substring(endCssIdx);
}

// Add page-hero to dashboard header
h = h.replace(
  /(<div class="[^"]*mb-8[^"]*">\s*<h1[^>]*>نظرة عامة مالية)/,
  '<div class="page-hero mb-8"><p class="section-label">FINANCIAL OVERVIEW</p><h1 class="font-headline font-black text-3xl text-primary mb-1">نظرة عامة مالية'
);
// Close properly - find the old closing
h = h.replace('هذه هي حالة دفترك المالي</p></div>', 'هذه هي حالة دفترك المالية الحالية</p></div>');

fs.writeFileSync('C:/Users/El-Wattaneya/Desktop/basset/index.html', h);
console.log('Done! stat-icon:', h.includes('stat-icon'), 'stat-card-accent:', h.includes('stat-card-accent'), 'page-hero:', h.includes('page-hero'));

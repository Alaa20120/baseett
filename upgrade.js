const fs = require('fs');
let h = fs.readFileSync('C:/Users/El-Wattaneya/Desktop/basset/index.html','utf8');

// 1. Enhanced CSS - replace the style block with premium version
const oldStyle = `.stat-card{background:#fff;border-radius:1.25rem;padding:1.5rem;box-shadow:0 1px 3px rgba(0,0,0,.06);transition:all .3s}.stat-card:hover{box-shadow:0 8px 24px rgba(0,0,0,.08);transform:translateY(-2px)}`;
const newStyle = `.stat-card{background:#fff;border-radius:1.5rem;padding:1.75rem;box-shadow:0 1px 4px rgba(3,22,54,.04),0 4px 16px rgba(3,22,54,.02);transition:all .4s cubic-bezier(.4,0,.2,1);border:1px solid rgba(3,22,54,.04)}.stat-card:hover{box-shadow:0 12px 40px rgba(3,22,54,.08);transform:translateY(-3px)}
.stat-card-accent{background:linear-gradient(135deg,#031636 0%,#1a2b4c 60%,#0058bb 100%);color:#fff;border:none;position:relative;overflow:hidden}.stat-card-accent::before{content:'';position:absolute;top:-50%;right:-50%;width:100%;height:100%;background:radial-gradient(circle,rgba(255,255,255,.08) 0%,transparent 70%);pointer-events:none}
.stat-card-accent p,.stat-card-accent span{color:#fff!important}.stat-card-accent .stat-label{color:rgba(255,255,255,.7)!important}
.stat-icon{width:48px;height:48px;border-radius:14px;display:flex;align-items:center;justify-content:center;font-size:1.25rem;margin-bottom:.75rem}
.stat-icon-blue{background:rgba(0,88,187,.08);color:#0058bb}.stat-icon-green{background:rgba(22,101,52,.08);color:#166534}.stat-icon-red{background:rgba(153,27,27,.08);color:#991b1b}.stat-icon-amber{background:rgba(133,77,14,.08);color:#854d0e}
.stat-icon-white{background:rgba(255,255,255,.15);color:#fff}
.stat-change{display:inline-flex;align-items:center;gap:4px;font-size:.7rem;font-weight:700;padding:2px 8px;border-radius:6px;margin-top:6px}
.stat-change-up{background:rgba(22,101,52,.08);color:#166534}.stat-change-down{background:rgba(153,27,27,.08);color:#991b1b}
.page-hero{padding:2rem 2.5rem;margin-bottom:2rem;background:linear-gradient(135deg,#f8f9ff 0%,#e6eeff 100%);border-radius:1.5rem;position:relative;overflow:hidden}
.page-hero::after{content:'';position:absolute;top:-30%;left:-20%;width:50%;height:160%;background:radial-gradient(ellipse,rgba(0,88,187,.04) 0%,transparent 70%);pointer-events:none}
.page-hero-dark{background:linear-gradient(135deg,#031636 0%,#1a2b4c 60%,#0d3a7a 100%);color:#fff}.page-hero-dark p{color:rgba(255,255,255,.7)}
.page-hero h1,.page-hero h2{font-family:Cairo;position:relative;z-index:1}.page-hero p{position:relative;z-index:1}
.section-label{font-size:.65rem;font-weight:700;text-transform:uppercase;letter-spacing:2px;color:#64748b;margin-bottom:.25rem}
.table-card{background:#fff;border-radius:1.5rem;box-shadow:0 1px 4px rgba(3,22,54,.04);overflow:hidden;border:1px solid rgba(3,22,54,.04)}
.table-card .data-table th:first-child{border-radius:0}.table-card .data-table th:last-child{border-radius:0}`;
h = h.replace(oldStyle, newStyle);

// 2. Better data-table header
const oldTh = `.data-table th{background:#f1f5f9;color:#0d1c2e;font-weight:600;padding:12px 16px;text-align:right;font-size:.8rem;text-transform:uppercase;letter-spacing:.5px}`;
const newTh = `.data-table th{background:#f8fafc;color:#64748b;font-weight:700;padding:14px 18px;text-align:right;font-size:.7rem;text-transform:uppercase;letter-spacing:1px;border-bottom:2px solid #e2e8f0}`;
h = h.replace(oldTh, newTh);

// 3. Better table rows
const oldTd = `.data-table td{padding:12px 16px;border-bottom:1px solid #f1f5f9;font-size:.875rem}.data-table tr:hover{background:#f8fafc}`;
const newTd = `.data-table td{padding:14px 18px;border-bottom:1px solid #f1f5f9;font-size:.85rem;color:#334155}.data-table tr:hover{background:rgba(0,88,187,.02)}.data-table tr{transition:background .2s}`;
h = h.replace(oldTd, newTd);

// 4. Better btn-primary
const oldBtn = `.btn-primary{background:linear-gradient(135deg,#031636,#1a2b4c);color:#fff;border:none;padding:.75rem 1.5rem;border-radius:.75rem;font-family:Cairo;font-weight:700;cursor:pointer;transition:all .2s;font-size:.875rem}`;
const newBtn = `.btn-primary{background:linear-gradient(135deg,#031636,#1a2b4c);color:#fff;border:none;padding:.75rem 1.75rem;border-radius:.875rem;font-family:Cairo;font-weight:700;cursor:pointer;transition:all .3s;font-size:.85rem;box-shadow:0 4px 14px rgba(3,22,54,.2);letter-spacing:.3px}`;
h = h.replace(oldBtn, newBtn);

// 5. Upgrade dashboard stat cards with icons
const oldDash = `<div class="grid grid-cols-2 lg:grid-cols-3 xl:grid-cols-6 gap-4 mb-8">
<div class="stat-card"><p class="text-xs text-slate-400 font-semibold mb-1">إجمالي المبيعات</p><p id="statTotalSales" class="text-xl font-black text-primary font-headline currency">0.00</p><p id="statSalesCount" class="text-xs text-slate-400 mt-1"></p></div>
<div class="stat-card"><p class="text-xs text-slate-400 font-semibold mb-1">إجمالي المشتريات</p><p id="statTotalPurchases" class="text-xl font-black text-primary font-headline currency">0.00</p></div>
<div class="stat-card"><p class="text-xs text-slate-400 font-semibold mb-1">صافي الربح</p><p id="statProfit" class="text-xl font-black font-headline currency" style="color:#166534">0.00</p></div>
<div class="stat-card"><p class="text-xs text-slate-400 font-semibold mb-1">المديونيات</p><p id="statTotalDue" class="text-xl font-black font-headline currency" style="color:#ef4444">0.00</p><p id="statCustomersCount" class="text-xs text-slate-400 mt-1"></p></div>
<div class="stat-card"><p class="text-xs text-slate-400 font-semibold mb-1">رصيد الخزينة</p><p id="statCashInBox" class="text-xl font-black font-headline currency" style="color:#166534">0.00</p></div>
<div class="stat-card"><p class="text-xs text-slate-400 font-semibold mb-1">عهدة المناديب</p><p id="statRepCustody" class="text-xl font-black font-headline currency" style="color:#854d0e">0.00</p></div>
</div>`;

const newDash = `<div class="grid grid-cols-2 lg:grid-cols-3 xl:grid-cols-6 gap-5 mb-8">
<div class="stat-card"><div class="stat-icon stat-icon-blue"><span class="material-symbols-outlined">trending_up</span></div><p class="stat-label text-xs text-slate-400 font-semibold mb-1">إجمالي المبيعات</p><p id="statTotalSales" class="text-2xl font-black text-primary font-headline currency">0.00</p><p id="statSalesCount" class="text-xs text-slate-400 mt-1"></p></div>
<div class="stat-card"><div class="stat-icon stat-icon-amber"><span class="material-symbols-outlined">shopping_cart</span></div><p class="stat-label text-xs text-slate-400 font-semibold mb-1">إجمالي المشتريات</p><p id="statTotalPurchases" class="text-2xl font-black text-primary font-headline currency">0.00</p></div>
<div class="stat-card"><div class="stat-icon stat-icon-green"><span class="material-symbols-outlined">account_balance</span></div><p class="stat-label text-xs text-slate-400 font-semibold mb-1">صافي الربح</p><p id="statProfit" class="text-2xl font-black font-headline currency" style="color:#166534">0.00</p></div>
<div class="stat-card"><div class="stat-icon stat-icon-red"><span class="material-symbols-outlined">credit_card_off</span></div><p class="stat-label text-xs text-slate-400 font-semibold mb-1">المديونيات</p><p id="statTotalDue" class="text-2xl font-black font-headline currency" style="color:#ef4444">0.00</p><p id="statCustomersCount" class="text-xs text-slate-400 mt-1"></p></div>
<div class="stat-card stat-card-accent"><div class="stat-icon stat-icon-white"><span class="material-symbols-outlined">savings</span></div><p class="stat-label text-xs font-semibold mb-1" style="color:rgba(255,255,255,.7)">رصيد الخزينة</p><p id="statCashInBox" class="text-2xl font-black font-headline currency" style="color:#fff">0.00</p></div>
<div class="stat-card"><div class="stat-icon stat-icon-amber"><span class="material-symbols-outlined">person_pin</span></div><p class="stat-label text-xs text-slate-400 font-semibold mb-1">عهدة المناديب</p><p id="statRepCustody" class="text-2xl font-black font-headline currency" style="color:#854d0e">0.00</p></div>
</div>`;
h = h.replace(oldDash, newDash);

// 6. Upgrade dashboard header
const oldDashHead = `<div class="mb-8"><h1 class="font-headline font-black text-3xl text-primary mb-1">نظرة عامة مالية</h1>
<p class="text-on-surface-variant text-sm">مرحباً بك. هذه هي حالة دفترك المالي</p></div>`;
const newDashHead = `<div class="page-hero mb-8">
<p class="section-label">FINANCIAL OVERVIEW</p>
<h1 class="font-headline font-black text-3xl text-primary mb-1">نظرة عامة مالية</h1>
<p class="text-on-surface-variant text-sm">مرحباً بك. هذه هي حالة دفترك المالي الحالية</p></div>`;
h = h.replace(oldDashHead, newDashHead);

// 7. Upgrade page headers for key pages
const pages = [
  ['إدارة المبيعات','SALES MANAGEMENT','إدارة المبيعات','إدارة ومتابعة جميع فواتير البيع والتحصيلات'],
  ['إدارة المشتريات','PROCUREMENT','إدارة المشتريات','متابعة فواتير الشراء والمدفوعات للموردين'],
  ['إدارة المخزون','INVENTORY OVERSIGHT','إدارة المخزون','رصيد وحركة المخزون والتسويات'],
  ['إدارة العملاء','CLIENTS & RECEIVABLES','إدارة العملاء','إدارة شاملة لجميع العملاء والمديونيات'],
  ['إدارة الموردين','SUPPLIERS & PAYABLES','إدارة الموردين','متابعة الموردين والمستحقات المالية'],
  ['إدارة المناديب','DELEGATES MANAGEMENT','إدارة المناديب','أداء المناديب والعهد والمستودعات'],
  ['إدارة المنتجات','PRODUCT CATALOG','إدارة المنتجات','كتالوج المنتجات والأسعار والأقسام'],
  ['الخزينة العامة','TREASURY & BANKING','الخزينة العامة','حركة الأموال والسندات المالية'],
  ['الموارد البشرية','HUMAN RESOURCES','الموارد البشرية','إدارة الموظفين والرواتب والإجراءات المالية'],
  ['التقارير الاحترافية','FINANCIAL INTELLIGENCE','التقارير والتحليلات','تقارير مالية وتحليلية شاملة'],
];

pages.forEach(([old,label,title,sub]) => {
  const oldH = `<h2 class="font-headline font-bold text-xl text-primary">${old}</h2>`;
  const newH = `<div><p class="section-label">${label}</p><h2 class="font-headline font-bold text-2xl text-primary">${title}</h2><p class="text-slate-400 text-xs mt-1">${sub}</p></div>`;
  h = h.replaceAll(oldH, newH);
});

// 8. Wrap tables in table-card
h = h.replaceAll('class="stat-card overflow-x-auto"', 'class="table-card overflow-x-auto p-0"');
h = h.replaceAll('class="stat-card overflow-x-auto mb-6"', 'class="table-card overflow-x-auto p-0 mb-6"');

// 9. Better top bar with search
const oldTopSearch = `<div class="flex items-center gap-2">
<input id="excelFileInput" type="file" accept=".xlsx,.xls" style="display:none">`;
const newTopSearch = `<div class="flex items-center gap-3">
<div class="relative hidden md:block"><input type="text" placeholder="البحث في السجلات..." class="search-input pl-9 text-xs" style="width:220px"><span class="material-symbols-outlined absolute left-3 top-1/2 -translate-y-1/2 text-slate-300" style="font-size:1.1rem">search</span></div>
<input id="excelFileInput" type="file" accept=".xlsx,.xls" style="display:none">`;
h = h.replace(oldTopSearch, newTopSearch);

// 10. Better sidebar brand
const oldBrand = `<div class="sidebar-brand-text"><h2 class="text-base font-bold text-primary font-headline leading-none">Sovereign</h2>
<p class="text-[10px] text-slate-400 font-medium">Enterprise ERP</p></div>`;
const newBrand = `<div class="sidebar-brand-text"><h2 class="text-base font-bold text-primary font-headline leading-none">المؤسسة السيادية</h2>
<p class="text-[10px] text-slate-400 font-medium">نظام الحسابات الموحد</p></div>`;
h = h.replace(oldBrand, newBrand);

// 11. Top bar title with "SOVEREIGN LEDGER"
const oldTitle = `<h2 id="pageTitle" class="font-headline font-bold text-lg text-primary">لوحة التحكم</h2>`;
const newTitle = `<span class="text-xs font-bold text-slate-300 tracking-widest hidden sm:inline" style="margin-left:12px">SOVEREIGN LEDGER</span><h2 id="pageTitle" class="font-headline font-bold text-lg text-primary">لوحة التحكم</h2>`;
h = h.replace(oldTitle, newTitle);

fs.writeFileSync('C:/Users/El-Wattaneya/Desktop/basset/index.html', h);
console.log('Premium design applied! Size:', h.length);

const fs = require('fs');

// --- UPDATE INDEX.HTML ---
let h = fs.readFileSync('C:/Users/El-Wattaneya/Desktop/basset/index.html', 'utf8');

// Update REPS page
const repsRegex = /<section id="page-reps"[\s\S]*?<\/section>/;
const newReps = `<section id="page-reps" class="page-section">
<div class="p-6 lg:p-8">
  <div class="flex flex-col md:flex-row justify-between items-start md:items-center mb-8 gap-4">
    <div>
      <p class="section-label">ACCOUNTING / RESOURCE MANAGEMENT</p>
      <h2 class="font-headline font-bold text-3xl text-primary">إدارة المندوبين</h2>
    </div>
    <div class="flex gap-3">
      <button class="btn-ghost bg-surface-container text-primary font-bold" onclick="App.exportAdvancedReports('repPerformance')"><i class="fas fa-download ml-2"></i>تصدير التقرير</button>
      <button class="btn-primary" onclick="App.openModal('staffModal')"><i class="fas fa-user-plus ml-2"></i>إضافة مندوب جديد</button>
    </div>
  </div>

  <div class="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-8">
    <div class="stat-card stat-card-accent flex flex-col justify-between" style="padding: 2rem;">
      <div class="flex justify-between items-start mb-6">
        <div>
          <p class="text-xs font-semibold mb-1" style="color:rgba(255,255,255,.7)">ACTIVE COLLECTIONS</p>
          <div class="flex items-baseline gap-2">
            <p id="repTopCollections" class="text-4xl font-black font-headline currency" style="color:#fff">0.00</p>
            <span class="text-sm font-bold" style="color:rgba(255,255,255,.7)">ر.س</span>
          </div>
        </div>
        <div class="stat-icon stat-icon-white"><span class="material-symbols-outlined">payments</span></div>
      </div>
      <div class="mt-auto">
        <span class="inline-flex items-center gap-1 text-xs font-bold px-2 py-1 rounded bg-white/10 text-white"><i class="fas fa-arrow-trend-up"></i> 12% زيادة عن الشهر الماضي</span>
      </div>
    </div>

    <div class="stat-card flex flex-col justify-between" style="padding: 2rem;">
      <div class="flex justify-between items-start mb-6">
        <div>
          <p class="text-xs text-slate-400 font-semibold mb-1 uppercase">AGGREGATE SALES</p>
          <p id="repTopSales" class="text-4xl font-black text-primary font-headline currency">0.00</p>
        </div>
        <div class="stat-icon stat-icon-blue"><span class="material-symbols-outlined">bar_chart</span></div>
      </div>
      <div class="mt-auto">
        <div class="flex justify-between text-xs font-bold text-slate-500 mb-2">
          <span>75% مكتمل</span>
          <span>هدف الربع السنوي</span>
        </div>
        <div class="w-full bg-slate-100 rounded-full h-2">
          <div class="bg-secondary h-2 rounded-full" style="width: 75%"></div>
        </div>
      </div>
    </div>
  </div>

  <div class="flex justify-between items-center mb-4">
    <div class="flex items-center gap-3">
      <h3 class="font-headline font-bold text-lg text-primary">سجل المندوبين النشطين</h3>
      <span id="repCountBadge" class="bg-surface-container text-primary text-xs font-bold px-2.5 py-1 rounded-full">0 مندوب</span>
    </div>
    <div class="flex gap-2">
      <button class="btn-ghost px-3 py-1.5 rounded-lg"><i class="fas fa-filter text-slate-400"></i></button>
      <button class="btn-ghost px-3 py-1.5 rounded-lg"><i class="fas fa-sort-amount-down text-slate-400"></i></button>
    </div>
  </div>

  <div class="table-card overflow-x-auto p-0 mb-8">
    <table class="data-table">
      <thead>
        <tr>
          <th>المندوب والتفاصيل</th>
          <th>المنطقة</th>
          <th>المبيعات الحالية</th>
          <th>التحصيلات</th>
          <th>الحالة</th>
          <th>إجراءات</th>
        </tr>
      </thead>
      <tbody id="repsBody"></tbody>
    </table>
  </div>
  
  <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
    <div class="lg:col-span-2 stat-card page-hero-dark" style="background: linear-gradient(135deg, #031636 0%, #1a2b4c 100%); padding: 3rem;">
       <p class="section-label text-white/50 mb-4">PERFORMANCE PEAK</p>
       <h3 class="font-headline font-black text-3xl text-white mb-8 leading-tight">أداء التحصيلات سجل رقماً<br>قياسياً هذا الأسبوع.</h3>
       <button class="bg-white/10 hover:bg-white/20 text-white border border-white/20 px-6 py-3 rounded-xl font-bold text-sm transition-colors">تحليل البيانات العميقة</button>
    </div>
    <div class="stat-card bg-surface-container border-none flex flex-col items-center justify-center">
       <h4 class="font-headline font-bold text-lg text-primary mb-6">توزيع المناطق</h4>
       <div class="w-32 h-32 border-8 border-secondary rounded flex items-center justify-center mb-6 relative">
          <div class="absolute top-0 right-0 w-4 h-full bg-primary/20"></div>
          <div class="text-center">
             <span class="text-2xl font-black text-primary font-headline">100%</span>
             <p class="text-[10px] text-slate-500 font-bold uppercase">تغطية</p>
          </div>
       </div>
       <div class="w-full space-y-2">
         <div class="flex justify-between text-xs font-bold"><span class="flex items-center gap-2"><span class="w-2 h-2 rounded-full bg-primary"></span> المنطقة الوسطى</span> <span>45%</span></div>
         <div class="flex justify-between text-xs font-bold"><span class="flex items-center gap-2"><span class="w-2 h-2 rounded-full bg-secondary"></span> المنطقة الغربية</span> <span>30%</span></div>
       </div>
    </div>
  </div>
</div>
</section>`;
h = h.replace(repsRegex, newReps);

// Update REPORTS page
const reportsRegex = /<section id="page-reports"[\s\S]*?<\/section>/;
const newReports = `<section id="page-reports" class="page-section">
<div class="p-6 lg:p-8 bg-surface">
  <div class="flex flex-col md:flex-row justify-between items-start md:items-center mb-8 gap-4">
    <div>
      <p class="section-label">FINANCIAL INTELLIGENCE</p>
      <h2 class="font-headline font-bold text-3xl text-primary">Reports & Analytics</h2>
    </div>
    <div class="flex gap-3 items-center">
      <div class="bg-white border border-slate-200 rounded-xl px-4 py-2 text-sm font-bold text-slate-600 flex items-center gap-2"><i class="far fa-calendar-alt"></i> Q3 2026 <i class="fas fa-chevron-down text-xs ml-2"></i></div>
      <button class="btn-primary" onclick="App.exportAdvancedReports('full')"><i class="fas fa-download ml-2"></i>Export All</button>
    </div>
  </div>

  <div class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-6 mb-8">
    <div class="stat-card relative overflow-hidden">
       <div class="absolute top-4 left-4 bg-blue-50 text-secondary text-xs font-bold px-2 py-1 rounded flex items-center gap-1"><i class="fas fa-arrow-up"></i> 12.5%</div>
       <div class="w-10 h-10 rounded-lg bg-surface-container flex items-center justify-center text-primary mb-4"><i class="fas fa-chart-line text-lg"></i></div>
       <p class="text-xs text-slate-500 font-bold mb-1">Net Profit</p>
       <p id="repNetProfit" class="text-2xl font-black text-primary font-headline">0.00</p>
    </div>
    <div class="stat-card relative overflow-hidden">
       <div class="absolute top-4 left-4 bg-blue-50 text-secondary text-xs font-bold px-2 py-1 rounded flex items-center gap-1"><i class="fas fa-arrow-up"></i> 8.2%</div>
       <div class="w-10 h-10 rounded-lg bg-blue-50 flex items-center justify-center text-secondary mb-4"><i class="fas fa-money-bill-wave text-lg"></i></div>
       <p class="text-xs text-slate-500 font-bold mb-1">Total Revenue</p>
       <p id="repTotalRev" class="text-2xl font-black text-primary font-headline">0.00</p>
    </div>
    <div class="stat-card relative overflow-hidden">
       <div class="absolute top-4 left-4 bg-red-50 text-red-600 text-xs font-bold px-2 py-1 rounded flex items-center gap-1"><i class="fas fa-arrow-up"></i> 3.1%</div>
       <div class="w-10 h-10 rounded-lg bg-red-50 flex items-center justify-center text-red-600 mb-4"><i class="fas fa-credit-card text-lg"></i></div>
       <p class="text-xs text-slate-500 font-bold mb-1">Operating Expenses</p>
       <p id="repOpExp" class="text-2xl font-black text-primary font-headline">0.00</p>
    </div>
    <div class="stat-card relative overflow-hidden">
       <div class="absolute top-4 left-4 bg-slate-100 text-slate-500 text-xs font-bold px-2 py-1 rounded flex items-center gap-1"><i class="fas fa-minus"></i> 0.0%</div>
       <div class="w-10 h-10 rounded-lg bg-slate-100 flex items-center justify-center text-slate-600 mb-4"><i class="fas fa-file-invoice-dollar text-lg"></i></div>
       <p class="text-xs text-slate-500 font-bold mb-1">Est. Tax Liability (ZATCA)</p>
       <p id="repTaxLiab" class="text-2xl font-black text-primary font-headline">0.00</p>
    </div>
  </div>

  <div class="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-8">
     <div class="lg:col-span-1 space-y-4">
        <div class="flex justify-between items-center mb-4"><h3 class="font-headline font-bold text-lg text-primary">Financial Statements</h3><a href="#" class="text-xs font-bold text-secondary">View All</a></div>
        
        <div class="stat-card flex flex-col" style="padding: 1.5rem;">
           <div class="flex items-start gap-4 mb-4">
             <div class="w-10 h-10 rounded-lg bg-slate-100 flex items-center justify-center text-primary shrink-0"><i class="fas fa-file-alt"></i></div>
             <div><h4 class="font-bold text-sm text-primary mb-1">Balance Sheet</h4><p class="text-xs text-slate-500 leading-relaxed">Snapshot of assets, liabilities, and equity.</p></div>
           </div>
           <div class="flex gap-2 mt-auto">
             <button class="bg-surface-container text-secondary px-3 py-2 rounded-lg text-xs font-bold flex-1" onclick="App.exportAdvancedReports('full')">Generate PDF</button>
             <button class="bg-surface-container text-secondary px-3 py-2 rounded-lg text-xs font-bold flex-1" onclick="App.exportAdvancedReports('full')">Export Excel</button>
           </div>
        </div>
        
        <div class="stat-card flex flex-col" style="padding: 1.5rem;">
           <div class="flex items-start gap-4 mb-4">
             <div class="w-10 h-10 rounded-lg bg-slate-100 flex items-center justify-center text-primary shrink-0"><i class="fas fa-chart-bar"></i></div>
             <div><h4 class="font-bold text-sm text-primary mb-1">Profit & Loss (P&L)</h4><p class="text-xs text-slate-500 leading-relaxed">Income and expenses over a specific period.</p></div>
           </div>
           <div class="flex gap-2 mt-auto">
             <button class="bg-surface-container text-secondary px-3 py-2 rounded-lg text-xs font-bold flex-1" onclick="App.exportAdvancedReports('profitLoss')">Generate PDF</button>
             <button class="bg-surface-container text-secondary px-3 py-2 rounded-lg text-xs font-bold flex-1" onclick="App.exportAdvancedReports('profitLoss')">Export Excel</button>
           </div>
        </div>

        <div class="stat-card flex flex-col" style="padding: 1.5rem;">
           <div class="flex items-start gap-4 mb-4">
             <div class="w-10 h-10 rounded-lg bg-slate-100 flex items-center justify-center text-primary shrink-0"><i class="fas fa-random"></i></div>
             <div><h4 class="font-bold text-sm text-primary mb-1">Cash Flow Statement</h4><p class="text-xs text-slate-500 leading-relaxed">Operating, investing, and financing activities.</p></div>
           </div>
           <div class="flex gap-2 mt-auto">
             <button class="bg-surface-container text-secondary px-3 py-2 rounded-lg text-xs font-bold w-1/2">Generate PDF</button>
           </div>
        </div>
     </div>

     <div class="lg:col-span-2 stat-card h-full flex flex-col">
       <div class="flex justify-between items-start mb-8">
         <div>
           <h3 class="font-headline font-bold text-lg text-primary mb-1">Revenue vs. Expenses Trend</h3>
           <p class="text-xs text-slate-500">YTD 2026 Monthly Comparison</p>
         </div>
         <div class="flex items-center gap-4 text-xs font-bold text-slate-600">
           <span class="flex items-center gap-2"><span class="w-3 h-3 rounded-full bg-primary"></span> Revenue</span>
           <span class="flex items-center gap-2"><span class="w-3 h-3 rounded-full bg-slate-400"></span> Expenses</span>
         </div>
       </div>
       <div class="flex-1 w-full relative min-h-[300px] flex items-end justify-between px-4" style="background-image: linear-gradient(to top, rgba(0,0,0,0.05) 1px, transparent 1px); background-size: 100% 25%;">
         <!-- Fake chart bars for UI visual matching -->
         <div class="w-1/12 flex items-end gap-1 h-full pt-8"><div class="w-1/2 bg-primary rounded-t-sm h-[40%]"></div><div class="w-1/2 bg-slate-500 rounded-t-sm h-[30%]"></div></div>
         <div class="w-1/12 flex items-end gap-1 h-full pt-8"><div class="w-1/2 bg-primary rounded-t-sm h-[45%]"></div><div class="w-1/2 bg-slate-500 rounded-t-sm h-[32%]"></div></div>
         <div class="w-1/12 flex items-end gap-1 h-full pt-8"><div class="w-1/2 bg-primary rounded-t-sm h-[55%]"></div><div class="w-1/2 bg-slate-500 rounded-t-sm h-[35%]"></div></div>
         <div class="w-1/12 flex items-end gap-1 h-full pt-8"><div class="w-1/2 bg-primary rounded-t-sm h-[50%]"></div><div class="w-1/2 bg-slate-500 rounded-t-sm h-[40%]"></div></div>
         <div class="w-1/12 flex items-end gap-1 h-full pt-8"><div class="w-1/2 bg-primary rounded-t-sm h-[65%]"></div><div class="w-1/2 bg-slate-500 rounded-t-sm h-[45%]"></div></div>
         <div class="w-1/12 flex items-end gap-1 h-full pt-8"><div class="w-1/2 bg-primary rounded-t-sm h-[70%]"></div><div class="w-1/2 bg-slate-500 rounded-t-sm h-[42%]"></div></div>
         <div class="w-1/12 flex items-end gap-1 h-full pt-8"><div class="w-1/2 bg-primary rounded-t-sm h-[85%]"></div><div class="w-1/2 bg-slate-500 rounded-t-sm h-[50%]"></div></div>
       </div>
       <div class="flex justify-between w-full mt-4 text-[10px] font-bold text-slate-400 px-4">
         <span>JAN</span><span>FEB</span><span>MAR</span><span>APR</span><span>MAY</span><span>JUN</span><span>JUL</span>
       </div>
     </div>
  </div>

  <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
    <div class="stat-card">
       <div class="flex justify-between items-center mb-6">
         <div class="flex items-center gap-3">
           <div class="w-10 h-10 rounded-lg bg-indigo-50 flex items-center justify-center text-indigo-600"><i class="fas fa-shield-alt"></i></div>
           <div><h4 class="font-bold text-sm text-primary mb-1">ZATCA Compliance</h4><p class="text-xs text-slate-500">E-Invoicing Phase II Reports</p></div>
         </div>
         <span class="bg-surface-container text-primary text-[10px] font-bold px-2 py-1 rounded">KSA Mandatory</span>
       </div>
       <div class="space-y-1">
         <div class="flex justify-between items-center p-3 hover:bg-slate-50 rounded-xl cursor-pointer transition-colors" onclick="App.exportAdvancedReports('taxReport')"><div class="flex items-center gap-3"><i class="fas fa-file-invoice text-slate-400"></i><span class="text-sm font-bold text-primary">VAT Return Summary</span></div><i class="fas fa-arrow-right text-slate-300"></i></div>
         <div class="flex justify-between items-center p-3 hover:bg-slate-50 rounded-xl cursor-pointer transition-colors" onclick="App.exportAdvancedReports('taxReport')"><div class="flex items-center gap-3"><i class="fas fa-check-double text-slate-400"></i><span class="text-sm font-bold text-primary">B2B E-Invoice Audit Log</span></div><i class="fas fa-arrow-right text-slate-300"></i></div>
         <div class="flex justify-between items-center p-3 hover:bg-slate-50 rounded-xl cursor-pointer transition-colors" onclick="App.exportAdvancedReports('taxReport')"><div class="flex items-center gap-3"><i class="fas fa-receipt text-slate-400"></i><span class="text-sm font-bold text-primary">B2C Simplified Invoices</span></div><i class="fas fa-arrow-right text-slate-300"></i></div>
       </div>
    </div>
    
    <div class="stat-card bg-slate-50 border-none">
       <div class="flex items-center gap-3 mb-6">
         <div class="w-10 h-10 rounded-lg bg-white border border-slate-200 flex items-center justify-center text-primary"><i class="fas fa-table"></i></div>
         <div><h4 class="font-bold text-sm text-primary mb-1">Operational Reports</h4><p class="text-xs text-slate-500">Sales & Inventory Data</p></div>
       </div>
       <div class="grid grid-cols-2 gap-4">
         <div class="bg-white p-4 rounded-xl border border-slate-100 cursor-pointer hover:border-primary/20 transition-colors" onclick="App.exportAdvancedReports('salesReport')">
           <i class="fas fa-shopping-cart text-secondary mb-3"></i>
           <h5 class="font-bold text-xs text-primary mb-1">Sales by Item</h5>
           <p class="text-[10px] text-slate-500">Top performing SKUs</p>
         </div>
         <div class="bg-white p-4 rounded-xl border border-slate-100 cursor-pointer hover:border-primary/20 transition-colors" onclick="App.exportAdvancedReports('customer')">
           <i class="fas fa-user-friends text-secondary mb-3"></i>
           <h5 class="font-bold text-xs text-primary mb-1">Sales by Customer</h5>
           <p class="text-[10px] text-slate-500">Client revenue analysis</p>
         </div>
         <div class="bg-white p-4 rounded-xl border border-slate-100 cursor-pointer hover:border-primary/20 transition-colors" onclick="App.exportAdvancedReports('inventoryMovement')">
           <i class="fas fa-clipboard-list text-primary mb-3"></i>
           <h5 class="font-bold text-xs text-primary mb-1">Stock Valuation</h5>
           <p class="text-[10px] text-slate-500">Current inventory worth</p>
         </div>
         <div class="bg-white p-4 rounded-xl border border-slate-100 cursor-pointer hover:border-primary/20 transition-colors" onclick="App.exportAdvancedReports('inventoryMovement')">
           <i class="fas fa-exclamation-triangle text-primary mb-3"></i>
           <h5 class="font-bold text-xs text-primary mb-1">Low Stock Alerts</h5>
           <p class="text-[10px] text-slate-500">Items needing reorder</p>
         </div>
       </div>
    </div>
  </div>
</div>
</section>`;
h = h.replace(reportsRegex, newReports);

fs.writeFileSync('C:/Users/El-Wattaneya/Desktop/basset/index.html', h);
console.log('index.html updated with premium Reps and Reports designs');

// --- UPDATE CORE.JS ---
let core = fs.readFileSync('C:/Users/El-Wattaneya/Desktop/basset/assets/js/core.js', 'utf8');

// Update renderRepsTable logic
const oldRenderReps = `function renderRepsTable() {
    const tbody = document.getElementById('repsBody');
    if (!tbody) return;

    const reps = ExcelEngine.getRepsList();

    if (reps.length === 0) {
      tbody.innerHTML = \`<tr><td colspan="9"><div class="empty-state"><i class="fas fa-user-tie"></i><h3>لا يوجد مناديب</h3><p>أضف مناديب من قسم الموارد البشرية</p></div></td></tr>\`;
      return;
    }

    tbody.innerHTML = reps.map(name => {
      const summary = ExcelEngine.getRepAccountingSummary(name);
      const repStock = ExcelEngine.getRepWarehouse(name);
      const stockCount = repStock.reduce((sum, rs) => sum + rs.quantity, 0);

      return \`
        <tr>
          <td><span class="clickable-link" onclick="App.viewProfile('rep', '\${name}')"><i class="fas fa-user-tie"></i> \${name}</span></td>
          <td class="number">\${summary.salesCount}</td>
          <td class="currency">\${ExcelEngine.formatCurrency(summary.totalSales)}</td>
          <td class="currency" style="color:var(--warning-500)">\${ExcelEngine.formatCurrency(summary.totalPurchases)}</td>
          <td class="currency" style="color:\${summary.totalProfit >= 0 ? 'var(--success-500)' : 'var(--danger-500)'}">\${ExcelEngine.formatCurrency(summary.totalProfit)}</td>
          <td class="currency">\${ExcelEngine.formatCurrency(summary.totalDelivered)}</td>
          <td class="currency \${summary.cashCustody > 0 ? 'negative' : ''}" style="font-weight:bold; color:var(--danger-500);" title="تحصيلات مبيعات + دفعات عملاء - توريدات المندوب">\${ExcelEngine.formatCurrency(summary.cashCustody)}</td>
          <td class="number"><span class="badge \${stockCount > 0 ? 'badge-info' : 'badge-warning'}" title="\${repStock.map(rs => rs.product + ': ' + rs.quantity).join(' | ') || 'فارغ'}">\${stockCount} <i class="fas fa-warehouse" style="font-size:0.7em;"></i></span></td>
          <td>
            <div class="row-actions">
              <button onclick="App.openWarehouseTransfer('\${name}')" title="سحب مخزون للمندوب"><i class="fas fa-dolly"></i></button>
              <button onclick="App.openRepPaymentModal('\${name}')" title="تسجيل توريد من المندوب"><i class="fas fa-hand-holding-dollar"></i></button>
              <button onclick="App.viewProfile('rep', '\${name}')" title="عرض الأداء"><i class="fas fa-chart-line"></i></button>
            </div>
          </td>
        </tr>
      \`;
    }).join('');
  }`;

const newRenderReps = `function renderRepsTable() {
    const tbody = document.getElementById('repsBody');
    if (!tbody) return;

    const reps = ExcelEngine.getRepsList();
    
    // Update top stats
    let totalSales = 0;
    let totalCollections = 0;

    const repsHtml = reps.map(name => {
      const summary = ExcelEngine.getRepAccountingSummary(name);
      totalSales += summary.totalSales;
      totalCollections += summary.totalDelivered;
      
      const statusHtml = summary.totalSales > 0 ? '<span class="bg-green-100 text-green-700 text-[10px] font-bold px-3 py-1 rounded-full">نشط</span>' : '<span class="bg-slate-100 text-slate-500 text-[10px] font-bold px-3 py-1 rounded-full">غير نشط</span>';
      
      // Dummy regions for UI
      const regions = ['الرياض، حي الملز', 'جدة، الكورنيش', 'الدمام، الواجهة', 'أبها، وسط المدينة'];
      const region = regions[Math.abs(name.length) % regions.length];
      const id = 'DE-' + (Math.floor(Math.random() * 5000) + 5000);

      return \`
        <tr>
          <td>
            <div class="flex items-center gap-3">
               <div class="w-10 h-10 rounded-xl bg-gradient-to-br from-primary to-primary-container text-white flex items-center justify-center font-bold text-sm shadow-inner">\${name.substring(0,2)}</div>
               <div>
                  <p class="font-bold text-sm text-primary hover:text-secondary cursor-pointer transition-colors" onclick="App.viewProfile('rep', '\${name}')">\${name}</p>
                  <p class="text-[10px] text-slate-400 font-bold">ID: \${id}</p>
               </div>
            </div>
          </td>
          <td class="text-xs font-bold text-slate-600">\${region}</td>
          <td class="currency font-bold text-primary">\${ExcelEngine.formatCurrency(summary.totalSales)} ر.س</td>
          <td class="currency font-bold text-secondary">\${ExcelEngine.formatCurrency(summary.totalDelivered)} ر.س</td>
          <td>\${statusHtml}</td>
          <td>
            <div class="flex gap-2">
              <button class="bg-surface-container hover:bg-slate-200 text-primary w-8 h-8 rounded-lg flex items-center justify-center transition-colors" onclick="App.openRepPaymentModal('\${name}')" title="تسجيل توريد من المندوب"><i class="fas fa-hand-holding-dollar text-xs"></i></button>
              <button class="bg-surface-container hover:bg-slate-200 text-primary w-8 h-8 rounded-lg flex items-center justify-center transition-colors" onclick="App.viewProfile('rep', '\${name}')" title="التفاصيل"><i class="fas fa-chevron-left text-xs"></i></button>
            </div>
          </td>
        </tr>
      \`;
    }).join('');

    tbody.innerHTML = reps.length ? repsHtml : \`<tr><td colspan="6"><div class="empty-state"><i class="fas fa-user-tie"></i><h3>لا يوجد مناديب</h3></div></td></tr>\`;
    
    // Update Stats on UI
    const elSales = document.getElementById('repTopSales');
    const elCol = document.getElementById('repTopCollections');
    const elBadge = document.getElementById('repCountBadge');
    
    if(elSales) elSales.textContent = ExcelEngine.formatNumber(totalSales);
    if(elCol) elCol.textContent = ExcelEngine.formatNumber(totalCollections);
    if(elBadge) elBadge.textContent = reps.length + ' مندوب';
  }`;

core = core.replace(oldRenderReps, newRenderReps);

// Intercept navigateTo to populate Reports data if navigating there
const oldNavigateTo = `if (pageId === 'statements') renderStatements();`;
const newNavigateTo = `if (pageId === 'statements') renderStatements();
    if (pageId === 'reports') {
       const stats = ExcelEngine.getStats();
       const rev = stats.totalSales;
       const exp = stats.totalPurchases + stats.totalHRCosts;
       const net = stats.realProfit - stats.totalHRCosts;
       const tax = rev * 0.15; // Dummy tax calc for display
       
       const elNet = document.getElementById('repNetProfit');
       const elRev = document.getElementById('repTotalRev');
       const elExp = document.getElementById('repOpExp');
       const elTax = document.getElementById('repTaxLiab');
       
       if(elNet) elNet.textContent = 'SAR ' + (net/1000).toFixed(2) + 'k';
       if(elRev) elRev.textContent = 'SAR ' + (rev/1000).toFixed(2) + 'k';
       if(elExp) elExp.textContent = 'SAR ' + (exp/1000).toFixed(2) + 'k';
       if(elTax) elTax.textContent = 'SAR ' + (tax/1000).toFixed(2) + 'k';
    }`;
core = core.replace(oldNavigateTo, newNavigateTo);

fs.writeFileSync('C:/Users/El-Wattaneya/Desktop/basset/assets/js/core.js', core);
console.log('core.js updated with new renderRepsTable');

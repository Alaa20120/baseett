const fs = require('fs');

// --- UPDATE INDEX.HTML ---
let h = fs.readFileSync('C:/Users/El-Wattaneya/Desktop/basset/index.html', 'utf8');

const invRegex = /<section id="page-inventory"[\s\S]*?<\/section>/;
const newInv = `<section id="page-inventory" class="page-section">
<div class="p-6 lg:p-8">
  <div class="flex flex-col md:flex-row justify-between items-start md:items-center mb-8 gap-4">
    <div>
      <p class="section-label">INVENTORY OVERSIGHT & ASSET MANAGEMENT</p>
      <h2 class="font-headline font-bold text-3xl text-primary">إدارة المخزون</h2>
    </div>
    <div class="flex gap-2">
      <button class="btn-ghost bg-surface-container text-primary font-bold" onclick="App.openModal('warehouseTransferModal')"><i class="fas fa-dolly ml-2"></i>نقل مخزون</button>
      <button class="btn-primary" onclick="App.openModal('stockAdjModal')"><i class="fas fa-plus ml-2"></i>تسوية مخزون</button>
    </div>
  </div>

  <div class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-6 mb-8">
    <div class="stat-card">
       <div class="flex justify-between items-start mb-4">
          <div class="w-10 h-10 rounded-lg bg-blue-50 flex items-center justify-center text-blue-600"><i class="fas fa-box-open text-lg"></i></div>
          <span class="bg-blue-50 text-blue-600 text-[10px] font-bold px-2 py-1 rounded">Daily</span>
       </div>
       <p class="text-xs text-slate-500 font-bold mb-1">PENDING SHIPMENTS</p>
       <p id="invPending" class="text-3xl font-black text-primary font-headline">0</p>
    </div>
    <div class="stat-card border border-red-100">
       <div class="flex justify-between items-start mb-4">
          <div class="w-10 h-10 rounded-lg bg-red-50 flex items-center justify-center text-red-600"><i class="fas fa-exclamation-triangle text-lg"></i></div>
          <span class="bg-red-50 text-red-600 text-[10px] font-bold px-2 py-1 rounded">Action Required</span>
       </div>
       <p class="text-xs text-slate-500 font-bold mb-1">LOW STOCK ITEMS</p>
       <p id="invLowStock" class="text-3xl font-black text-red-600 font-headline">0</p>
    </div>
    <div class="stat-card">
       <div class="flex justify-between items-start mb-4">
          <div class="w-10 h-10 rounded-lg bg-emerald-50 flex items-center justify-center text-emerald-600"><i class="fas fa-dollar-sign text-lg"></i></div>
          <span class="bg-emerald-50 text-emerald-600 text-[10px] font-bold px-2 py-1 rounded">YTD</span>
       </div>
       <p class="text-xs text-slate-500 font-bold mb-1">INVENTORY VALUE</p>
       <p id="invValue" class="text-3xl font-black text-primary font-headline">0.00</p>
    </div>
    <div class="stat-card">
       <div class="flex justify-between items-start mb-4">
          <div class="w-10 h-10 rounded-lg bg-slate-100 flex items-center justify-center text-slate-600"><i class="fas fa-tags text-lg"></i></div>
       </div>
       <p class="text-xs text-slate-500 font-bold mb-1">TOTAL SKU'S</p>
       <p id="invSkus" class="text-3xl font-black text-primary font-headline">0</p>
    </div>
  </div>

  <div class="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-6">
    <div class="flex items-center gap-2">
      <h3 class="font-headline font-bold text-lg text-primary ml-4">Product Ledger</h3>
      <div class="flex gap-2">
         <span class="bg-primary text-white text-xs font-bold px-4 py-2 rounded-full cursor-pointer shadow-sm">All Categories</span>
         <span class="bg-surface-container text-slate-500 text-xs font-bold px-4 py-2 rounded-full cursor-pointer hover:bg-slate-200 transition-colors">Apparel</span>
         <span class="bg-surface-container text-slate-500 text-xs font-bold px-4 py-2 rounded-full cursor-pointer hover:bg-slate-200 transition-colors">Electronics</span>
      </div>
    </div>
    <div class="relative">
      <input type="text" placeholder="Search inventory..." class="search-input pl-10 text-xs py-2 w-64">
      <i class="fas fa-search absolute left-3 top-1/2 -translate-y-1/2 text-slate-300"></i>
    </div>
  </div>

  <div class="table-card overflow-x-auto p-0 mb-8">
    <table class="data-table" style="min-width: 800px;">
      <thead>
        <tr>
          <th>PRODUCT DETAILS</th>
          <th>CATEGORY</th>
          <th>STOCK LEVEL</th>
          <th>UNIT PRICE</th>
          <th>STATUS</th>
          <th>ACTIONS</th>
        </tr>
      </thead>
      <tbody id="inventoryBody"></tbody>
    </table>
  </div>
  
  <div class="stat-card p-0 overflow-hidden">
     <div class="p-6 border-b border-slate-100 flex justify-between items-center">
        <h3 class="font-headline font-bold text-lg text-primary">حركة المخزون (Ledger)</h3>
        <button class="btn-ghost text-xs" onclick="App.exportAdvancedReports('inventoryMovement')">تصدير التقرير <i class="fas fa-download mr-1"></i></button>
     </div>
     <div class="overflow-x-auto">
        <table class="data-table w-full">
           <thead><tr><th>التاريخ</th><th>المنتج</th><th>النوع</th><th>الكمية</th><th>الموقع</th><th>الطرف</th></tr></thead>
           <tbody id="inventoryLedgerBody"></tbody>
        </table>
     </div>
  </div>

</div>
</section>`;

h = h.replace(invRegex, newInv);

fs.writeFileSync('C:/Users/El-Wattaneya/Desktop/basset/index.html', h);
console.log('index.html updated for Inventory');

// --- UPDATE CORE.JS ---
let core = fs.readFileSync('C:/Users/El-Wattaneya/Desktop/basset/assets/js/core.js', 'utf8');

const oldRenderInv = `function renderInventoryTable() {
    const products = ExcelEngine.getProducts();
    const tbody = document.getElementById('inventoryBody');
    if (!tbody) return;

    if (products.length === 0) {
      tbody.innerHTML = \`<tr><td colspan="5"><div class="empty-state"><i class="fas fa-boxes"></i><h3>لا توجد حركة مخزون</h3></div></td></tr>\`;
      return;
    }

    tbody.innerHTML = products.map(p => {
      const summary = ExcelEngine.getProductStockSummary(p.name);
      return \`
        <tr>
          <td><span class="clickable-link" onclick="App.viewProfile('product', '\${p.name}')"><i class="fas fa-box"></i> \${p.name}</span></td>
          <td class="number">\${summary.prevStock}</td>
          <td class="number text-success">\${summary.purchased}</td>
          <td class="number text-danger">\${summary.sold}</td>
          <td class="number" style="font-weight:bold; color:\${summary.currentStock > 0 ? 'var(--primary-600)' : 'var(--danger-500)'}">\${summary.currentStock}</td>
        </tr>
      \`;
    }).join('');
  }`;

const newRenderInv = `function renderInventoryTable() {
    const products = ExcelEngine.getProducts();
    const tbody = document.getElementById('inventoryBody');
    if (!tbody) return;

    if (products.length === 0) {
      tbody.innerHTML = \`<tr><td colspan="6"><div class="empty-state"><i class="fas fa-boxes"></i><h3>لا توجد منتجات</h3></div></td></tr>\`;
      return;
    }
    
    let totalValue = 0;
    let lowStockCount = 0;

    tbody.innerHTML = products.map(p => {
      const summary = ExcelEngine.getProductStockSummary(p.name);
      const stock = summary.currentStock;
      totalValue += (stock * p.sellPrice);
      if(stock <= 10) lowStockCount++;
      
      const maxStock = 500; // arbitrary max for progress bar
      const pct = Math.min(100, Math.max(0, (stock / maxStock) * 100));
      let statusHtml = '';
      let barColor = 'bg-primary';
      
      if(stock <= 0) {
         statusHtml = '<span class="bg-red-50 text-red-600 text-[10px] font-bold px-2 py-1 rounded-md">Out of Stock</span>';
         barColor = 'bg-red-500';
      } else if (stock <= 10) {
         statusHtml = '<span class="bg-amber-50 text-amber-600 text-[10px] font-bold px-2 py-1 rounded-md">Low Stock</span>';
         barColor = 'bg-amber-500';
      } else {
         statusHtml = '<span class="bg-emerald-50 text-emerald-600 text-[10px] font-bold px-2 py-1 rounded-md">In Stock</span>';
         barColor = 'bg-emerald-500';
      }

      const id = 'SKU-' + (Math.floor(Math.random() * 90000) + 10000);

      return \`
        <tr>
          <td>
            <div class="flex items-center gap-3">
               <div class="w-10 h-10 rounded-lg bg-surface-container flex items-center justify-center text-slate-400 shrink-0"><i class="fas fa-image"></i></div>
               <div>
                  <p class="font-bold text-sm text-primary hover:text-secondary cursor-pointer" onclick="App.viewProfile('product', '\${p.name}')">\${p.name}</p>
                  <p class="text-[10px] text-slate-400 font-bold">\${id}</p>
               </div>
            </div>
          </td>
          <td><span class="bg-slate-100 text-slate-600 text-[10px] font-bold px-3 py-1 rounded-full">\${p.category}</span></td>
          <td style="min-width: 150px;">
             <div class="flex items-center gap-2">
                <div class="w-full bg-slate-100 rounded-full h-1.5 flex-1"><div class="\${barColor} h-1.5 rounded-full" style="width: \${pct}%"></div></div>
                <span class="text-xs font-bold text-slate-600 w-8 text-left">\${stock}</span>
             </div>
          </td>
          <td class="currency font-bold text-slate-700">\${ExcelEngine.formatCurrency(p.sellPrice)}</td>
          <td>\${statusHtml}</td>
          <td>
            <div class="flex gap-2">
              <button class="bg-surface-container hover:bg-slate-200 text-primary w-8 h-8 rounded-lg flex items-center justify-center transition-colors" onclick="App.editProduct(\${p.id})" title="تعديل"><i class="fas fa-edit text-xs"></i></button>
            </div>
          </td>
        </tr>
      \`;
    }).join('');
    
    // Update Stats
    const elSkus = document.getElementById('invSkus');
    const elVal = document.getElementById('invValue');
    const elLow = document.getElementById('invLowStock');
    const elPend = document.getElementById('invPending'); // Dummy pending
    
    if(elSkus) elSkus.textContent = ExcelEngine.formatNumber(products.length);
    if(elVal) elVal.textContent = ExcelEngine.formatCurrency(totalValue) + ' ر.س';
    if(elLow) elLow.textContent = lowStockCount;
    if(elPend) elPend.textContent = '12';
  }`;

core = core.replace(oldRenderInv, newRenderInv);
fs.writeFileSync('C:/Users/El-Wattaneya/Desktop/basset/assets/js/core.js', core);
console.log('core.js updated for Inventory');

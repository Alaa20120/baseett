const fs = require('fs');

// --- UPDATE INDEX.HTML ---
let h = fs.readFileSync('C:/Users/El-Wattaneya/Desktop/basset/index.html', 'utf8');

const custRegex = /<section id="page-customers"[\s\S]*?<\/section>/;
const newCust = `<section id="page-customers" class="page-section">
<div class="p-6 lg:p-8">
  <div class="flex flex-col md:flex-row justify-between items-start md:items-center mb-8 gap-4">
    <div>
      <p class="section-label">CLIENTS & RECEIVABLES</p>
      <h2 class="font-headline font-bold text-3xl text-primary">إدارة العملاء</h2>
    </div>
    <div class="flex gap-2">
      <button class="btn-ghost bg-surface-container text-primary font-bold" onclick="App.exportAdvancedReports('customer')"><i class="fas fa-download ml-2"></i>تصدير</button>
      <button class="btn-primary" onclick="App.openModal('customerModal')"><i class="fas fa-plus ml-2"></i>إضافة عميل جديد</button>
    </div>
  </div>

  <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
    <div class="stat-card">
       <div class="flex items-center gap-4 mb-4">
          <div class="w-12 h-12 rounded-xl bg-blue-50 flex items-center justify-center text-secondary"><i class="fas fa-users text-xl"></i></div>
          <div><p class="text-xs text-slate-500 font-bold mb-1">TOTAL CLIENTS</p><p id="custTotal" class="text-2xl font-black text-primary font-headline">0</p></div>
       </div>
    </div>
    <div class="stat-card">
       <div class="flex items-center gap-4 mb-4">
          <div class="w-12 h-12 rounded-xl bg-emerald-50 flex items-center justify-center text-emerald-600"><i class="fas fa-hand-holding-dollar text-xl"></i></div>
          <div><p class="text-xs text-slate-500 font-bold mb-1">TOTAL RECEIVABLES</p><p id="custDue" class="text-2xl font-black text-primary font-headline currency">0.00</p></div>
       </div>
    </div>
    <div class="stat-card">
       <div class="flex items-center gap-4 mb-4">
          <div class="w-12 h-12 rounded-xl bg-red-50 flex items-center justify-center text-red-600"><i class="fas fa-exclamation-circle text-xl"></i></div>
          <div><p class="text-xs text-slate-500 font-bold mb-1">OVERDUE CLIENTS</p><p id="custOverdue" class="text-2xl font-black text-red-600 font-headline">0</p></div>
       </div>
    </div>
  </div>

  <div class="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-6">
    <div class="flex items-center gap-2">
      <h3 class="font-headline font-bold text-lg text-primary ml-4">سجل العملاء</h3>
      <div class="flex gap-2">
         <span class="bg-primary text-white text-xs font-bold px-4 py-2 rounded-full cursor-pointer shadow-sm">الكل</span>
         <span class="bg-surface-container text-slate-500 text-xs font-bold px-4 py-2 rounded-full cursor-pointer hover:bg-slate-200 transition-colors">نشط</span>
         <span class="bg-surface-container text-slate-500 text-xs font-bold px-4 py-2 rounded-full cursor-pointer hover:bg-slate-200 transition-colors">متأخر</span>
      </div>
    </div>
    <div class="relative">
      <input type="text" placeholder="البحث عن عميل..." class="search-input pl-10 text-xs py-2 w-64">
      <i class="fas fa-search absolute left-3 top-1/2 -translate-y-1/2 text-slate-300"></i>
    </div>
  </div>

  <div class="table-card overflow-x-auto p-0">
    <table class="data-table w-full" style="min-width: 800px;">
      <thead>
        <tr>
          <th>العميل</th>
          <th>رقم التواصل</th>
          <th>إجمالي التعاملات</th>
          <th>المدفوع</th>
          <th>المتبقي (عليه)</th>
          <th>الحالة</th>
          <th>إجراءات</th>
        </tr>
      </thead>
      <tbody id="customersBody"></tbody>
    </table>
  </div>
</div>
</section>`;
h = h.replace(custRegex, newCust);

const suppRegex = /<section id="page-suppliers"[\s\S]*?<\/section>/;
const newSupp = `<section id="page-suppliers" class="page-section">
<div class="p-6 lg:p-8">
  <div class="flex flex-col md:flex-row justify-between items-start md:items-center mb-8 gap-4">
    <div>
      <p class="section-label">SUPPLIERS & PAYABLES</p>
      <h2 class="font-headline font-bold text-3xl text-primary">إدارة الموردين</h2>
    </div>
    <div class="flex gap-2">
      <button class="btn-ghost bg-surface-container text-primary font-bold" onclick="App.exportAdvancedReports('debts')"><i class="fas fa-download ml-2"></i>تصدير</button>
      <button class="btn-primary" onclick="App.openModal('supplierModal')"><i class="fas fa-plus ml-2"></i>إضافة مورد جديد</button>
    </div>
  </div>

  <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
    <div class="stat-card">
       <div class="flex items-center gap-4 mb-4">
          <div class="w-12 h-12 rounded-xl bg-blue-50 flex items-center justify-center text-secondary"><i class="fas fa-truck text-xl"></i></div>
          <div><p class="text-xs text-slate-500 font-bold mb-1">TOTAL SUPPLIERS</p><p id="suppTotal" class="text-2xl font-black text-primary font-headline">0</p></div>
       </div>
    </div>
    <div class="stat-card">
       <div class="flex items-center gap-4 mb-4">
          <div class="w-12 h-12 rounded-xl bg-emerald-50 flex items-center justify-center text-emerald-600"><i class="fas fa-file-invoice-dollar text-xl"></i></div>
          <div><p class="text-xs text-slate-500 font-bold mb-1">TOTAL PAYABLES</p><p id="suppDue" class="text-2xl font-black text-primary font-headline currency">0.00</p></div>
       </div>
    </div>
    <div class="stat-card">
       <div class="flex items-center gap-4 mb-4">
          <div class="w-12 h-12 rounded-xl bg-red-50 flex items-center justify-center text-red-600"><i class="fas fa-exclamation-triangle text-xl"></i></div>
          <div><p class="text-xs text-slate-500 font-bold mb-1">OVERDUE PAYMENTS</p><p id="suppOverdue" class="text-2xl font-black text-red-600 font-headline">0</p></div>
       </div>
    </div>
  </div>

  <div class="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-6">
    <div class="flex items-center gap-2">
      <h3 class="font-headline font-bold text-lg text-primary ml-4">سجل الموردين</h3>
      <div class="flex gap-2">
         <span class="bg-primary text-white text-xs font-bold px-4 py-2 rounded-full cursor-pointer shadow-sm">الكل</span>
         <span class="bg-surface-container text-slate-500 text-xs font-bold px-4 py-2 rounded-full cursor-pointer hover:bg-slate-200 transition-colors">نشط</span>
      </div>
    </div>
    <div class="relative">
      <input type="text" placeholder="البحث عن مورد..." class="search-input pl-10 text-xs py-2 w-64">
      <i class="fas fa-search absolute left-3 top-1/2 -translate-y-1/2 text-slate-300"></i>
    </div>
  </div>

  <div class="table-card overflow-x-auto p-0">
    <table class="data-table w-full" style="min-width: 800px;">
      <thead>
        <tr>
          <th>المورد</th>
          <th>رقم التواصل</th>
          <th>إجمالي التعاملات</th>
          <th>المدفوع</th>
          <th>المتبقي (له)</th>
          <th>الحالة</th>
          <th>إجراءات</th>
        </tr>
      </thead>
      <tbody id="suppliersBody"></tbody>
    </table>
  </div>
</div>
</section>`;
h = h.replace(suppRegex, newSupp);

fs.writeFileSync('C:/Users/El-Wattaneya/Desktop/basset/index.html', h);
console.log('index.html updated for Customers and Suppliers');

// --- UPDATE CORE.JS ---
let core = fs.readFileSync('C:/Users/El-Wattaneya/Desktop/basset/assets/js/core.js', 'utf8');

const oldRenderCust = `function renderCustomersTable() {
    const customers = ExcelEngine.getCustomers();
    const tbody = document.getElementById('customersBody');
    if (!tbody) return;

    if (customers.length === 0) {
      tbody.innerHTML = \`<tr><td colspan="6"><div class="empty-state"><i class="fas fa-users"></i><h3>لا يوجد عملاء</h3><p>أضف عميل جديد لتبدأ</p></div></td></tr>\`;
      return;
    }

    tbody.innerHTML = customers.map(c => {
      const isDue = c.due > 0;
      return \`
        <tr>
          <td><span class="clickable-link" onclick="App.viewProfile('customer', '\${c.name}')"><i class="fas fa-user"></i> \${c.name}</span></td>
          <td>\${c.phone}</td>
          <td class="currency">\${ExcelEngine.formatCurrency(c.totalTransactions)}</td>
          <td class="currency text-success">\${ExcelEngine.formatCurrency(c.totalPaid)}</td>
          <td class="currency" style="font-weight:bold; color:\${isDue ? 'var(--danger-500)' : 'inherit'}">\${ExcelEngine.formatCurrency(c.due)}</td>
          <td>
            <div class="row-actions">
              <button class="primary" onclick="App.openCustomerPaymentModal('\${c.name}')" title="تسجيل دفعة"><i class="fas fa-money-bill-wave"></i></button>
              <button onclick="App.editCustomer(\${c.id})" title="تعديل"><i class="fas fa-edit"></i></button>
            </div>
          </td>
        </tr>
      \`;
    }).join('');
  }`;

const newRenderCust = `function renderCustomersTable() {
    const customers = ExcelEngine.getCustomers();
    const tbody = document.getElementById('customersBody');
    if (!tbody) return;

    if (customers.length === 0) {
      tbody.innerHTML = \`<tr><td colspan="7"><div class="empty-state"><i class="fas fa-users"></i><h3>لا يوجد عملاء</h3></div></td></tr>\`;
      return;
    }

    let totalDue = 0;
    let overdueCount = 0;

    tbody.innerHTML = customers.map(c => {
      const isDue = c.due > 0;
      totalDue += c.due;
      if (isDue) overdueCount++;
      
      const statusHtml = isDue ? '<span class="bg-red-50 text-red-600 text-[10px] font-bold px-3 py-1 rounded-full">عليه مديونية</span>' : '<span class="bg-green-50 text-green-600 text-[10px] font-bold px-3 py-1 rounded-full">خالص</span>';

      return \`
        <tr>
          <td>
            <div class="flex items-center gap-3">
               <div class="w-10 h-10 rounded-full bg-surface-container text-primary flex items-center justify-center font-bold text-sm border border-slate-200"><i class="fas fa-user text-xs"></i></div>
               <div>
                  <p class="font-bold text-sm text-primary hover:text-secondary cursor-pointer" onclick="App.viewProfile('customer', '\${c.name}')">\${c.name}</p>
               </div>
            </div>
          </td>
          <td class="text-xs text-slate-500 font-bold">\${c.phone || '-'}</td>
          <td class="currency font-bold text-slate-700">\${ExcelEngine.formatCurrency(c.totalTransactions)}</td>
          <td class="currency font-bold text-emerald-600">\${ExcelEngine.formatCurrency(c.totalPaid)}</td>
          <td class="currency font-bold \${isDue ? 'text-red-600' : 'text-slate-400'}">\${ExcelEngine.formatCurrency(c.due)}</td>
          <td>\${statusHtml}</td>
          <td>
            <div class="flex gap-2">
              <button class="bg-emerald-50 hover:bg-emerald-100 text-emerald-600 w-8 h-8 rounded-lg flex items-center justify-center transition-colors" onclick="App.openCustomerPaymentModal('\${c.name}')" title="تسجيل دفعة"><i class="fas fa-money-bill-wave text-xs"></i></button>
              <button class="bg-surface-container hover:bg-slate-200 text-primary w-8 h-8 rounded-lg flex items-center justify-center transition-colors" onclick="App.editCustomer(\${c.id})" title="تعديل"><i class="fas fa-edit text-xs"></i></button>
            </div>
          </td>
        </tr>
      \`;
    }).join('');
    
    const elTot = document.getElementById('custTotal');
    const elDue = document.getElementById('custDue');
    const elOv = document.getElementById('custOverdue');
    if(elTot) elTot.textContent = customers.length;
    if(elDue) elDue.textContent = ExcelEngine.formatCurrency(totalDue);
    if(elOv) elOv.textContent = overdueCount;
  }`;
core = core.replace(oldRenderCust, newRenderCust);

const oldRenderSupp = `function renderSuppliersTable() {
    const suppliers = ExcelEngine.getSuppliers();
    const tbody = document.getElementById('suppliersBody');
    if (!tbody) return;

    if (suppliers.length === 0) {
      tbody.innerHTML = \`<tr><td colspan="6"><div class="empty-state"><i class="fas fa-truck"></i><h3>لا يوجد موردين</h3><p>أضف مورد جديد لتبدأ</p></div></td></tr>\`;
      return;
    }

    tbody.innerHTML = suppliers.map(s => {
      const isDue = s.due > 0;
      return \`
        <tr>
          <td><span class="clickable-link" onclick="App.viewProfile('supplier', '\${s.name}')"><i class="fas fa-building"></i> \${s.name}</span></td>
          <td>\${s.phone}</td>
          <td class="currency">\${ExcelEngine.formatCurrency(s.totalTransactions)}</td>
          <td class="currency text-success">\${ExcelEngine.formatCurrency(s.totalPaid)}</td>
          <td class="currency" style="font-weight:bold; color:\${isDue ? 'var(--danger-500)' : 'inherit'}">\${ExcelEngine.formatCurrency(s.due)}</td>
          <td>
            <div class="row-actions">
              <button class="primary" onclick="App.openSupplierPaymentModal('\${s.name}')" title="تسجيل دفعة"><i class="fas fa-money-bill-wave"></i></button>
              <button onclick="App.editSupplier(\${s.id})" title="تعديل"><i class="fas fa-edit"></i></button>
            </div>
          </td>
        </tr>
      \`;
    }).join('');
  }`;

const newRenderSupp = `function renderSuppliersTable() {
    const suppliers = ExcelEngine.getSuppliers();
    const tbody = document.getElementById('suppliersBody');
    if (!tbody) return;

    if (suppliers.length === 0) {
      tbody.innerHTML = \`<tr><td colspan="7"><div class="empty-state"><i class="fas fa-truck"></i><h3>لا يوجد موردين</h3></div></td></tr>\`;
      return;
    }
    
    let totalDue = 0;
    let overdueCount = 0;

    tbody.innerHTML = suppliers.map(s => {
      const isDue = s.due > 0;
      totalDue += s.due;
      if (isDue) overdueCount++;
      
      const statusHtml = isDue ? '<span class="bg-red-50 text-red-600 text-[10px] font-bold px-3 py-1 rounded-full">مستحق الدفع</span>' : '<span class="bg-green-50 text-green-600 text-[10px] font-bold px-3 py-1 rounded-full">خالص</span>';

      return \`
        <tr>
          <td>
            <div class="flex items-center gap-3">
               <div class="w-10 h-10 rounded-full bg-slate-100 text-slate-500 flex items-center justify-center font-bold text-sm"><i class="fas fa-building text-xs"></i></div>
               <div>
                  <p class="font-bold text-sm text-primary hover:text-secondary cursor-pointer" onclick="App.viewProfile('supplier', '\${s.name}')">\${s.name}</p>
               </div>
            </div>
          </td>
          <td class="text-xs text-slate-500 font-bold">\${s.phone || '-'}</td>
          <td class="currency font-bold text-slate-700">\${ExcelEngine.formatCurrency(s.totalTransactions)}</td>
          <td class="currency font-bold text-emerald-600">\${ExcelEngine.formatCurrency(s.totalPaid)}</td>
          <td class="currency font-bold \${isDue ? 'text-red-600' : 'text-slate-400'}">\${ExcelEngine.formatCurrency(s.due)}</td>
          <td>\${statusHtml}</td>
          <td>
            <div class="flex gap-2">
              <button class="bg-emerald-50 hover:bg-emerald-100 text-emerald-600 w-8 h-8 rounded-lg flex items-center justify-center transition-colors" onclick="App.openSupplierPaymentModal('\${s.name}')" title="تسجيل دفعة"><i class="fas fa-money-bill-wave text-xs"></i></button>
              <button class="bg-surface-container hover:bg-slate-200 text-primary w-8 h-8 rounded-lg flex items-center justify-center transition-colors" onclick="App.editSupplier(\${s.id})" title="تعديل"><i class="fas fa-edit text-xs"></i></button>
            </div>
          </td>
        </tr>
      \`;
    }).join('');
    
    const elTot = document.getElementById('suppTotal');
    const elDue = document.getElementById('suppDue');
    const elOv = document.getElementById('suppOverdue');
    if(elTot) elTot.textContent = suppliers.length;
    if(elDue) elDue.textContent = ExcelEngine.formatCurrency(totalDue);
    if(elOv) elOv.textContent = overdueCount;
  }`;

core = core.replace(oldRenderSupp, newRenderSupp);
fs.writeFileSync('C:/Users/El-Wattaneya/Desktop/basset/assets/js/core.js', core);
console.log('core.js updated for Customers and Suppliers');

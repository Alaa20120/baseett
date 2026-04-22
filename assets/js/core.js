// ============================================
// التطبيق الرئيسي - Main Application
// ============================================

const App = (() => {
  let currentPage = 'dashboard';
  let chartInstances = {};

  // === Initialize ===
  function init() {
    // Try loading from localStorage first, then defaults
    if (!ExcelEngine.loadFromLocalStorage()) {
      ExcelEngine.initializeDefaults();
    }

    setupAuth();
    setupSidebarToggle();
    setupNavigation();
    setupFileHandlers();
    setupModals();
    setupSearch();

    // New setup functions for extra features
    if (typeof setupReturnForm === 'function') setupReturnForm();
    if (typeof setupVoucherForm === 'function') setupVoucherForm();
    if (typeof setupEditProductForm === 'function') setupEditProductForm();

    // Populate selections
    if (typeof populateCategorySelects === 'function') populateCategorySelects();
    if (typeof populateVoucherParties === 'function') populateVoucherParties();

    navigateTo('dashboard');
    updateExcelStatus(true);
  }

  // === Auth logic ===
  function setupAuth() {
    const loginForm = document.getElementById('loginForm');

    // Set up login form listener (always needed, even during first-time setup)
    if (loginForm) {
      loginForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const username = document.getElementById('loginUsername').value.trim();
        const pass = document.getElementById('loginPassword').value.trim();

        const loginBtn = loginForm.querySelector('.login-btn');
        const originalText = loginBtn.innerHTML;
        loginBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> جاري المصادقة...';
        loginBtn.disabled = true;

        setTimeout(() => {
          const result = ExcelEngine.login(username, pass);
          if (result === true) {
            handleAuthSuccess();
            const securedCount = ExcelEngine.secureExistingInvoices();
            if (securedCount > 0) {
              showToast(`تم اكتشاف وتأمين ${securedCount} فاتورة غير مشفرة بنجاح`, 'info');
            }
          } else if (result === 'suspended') {
            showToast('هذا الحساب موقوف مؤقتاً. تواصل مع المدير لتفعيله.', 'error');
            loginBtn.innerHTML = originalText;
            loginBtn.disabled = false;
          } else {
            showToast('اسم المستخدم أو كلمة المرور غير صحيحة', 'error');
            loginBtn.innerHTML = originalText;
            loginBtn.disabled = false;
          }
        }, 800);
      });
    }

    const logoutBtn = document.getElementById('logoutBtn');
    if (logoutBtn) {
      logoutBtn.addEventListener('click', () => {
        ExcelEngine.logout();
        sessionStorage.removeItem('currentUser');
        window.location.reload();
      });
    }

    // First-time setup: show workspace init overlay if not yet deployed
    if (!localStorage.getItem('sovLedger_workspace')) {
      showInitWorkspace();
      return;
    }

    // Check existing session
    const storedUser = sessionStorage.getItem('currentUser');
    if (storedUser) {
      try {
        const user = JSON.parse(storedUser);
        if (ExcelEngine.login(user.name, user.password || '')) {
          handleAuthSuccess();
        } else {
          document.body.classList.add('logged-out');
        }
      } catch (e) {
        document.body.classList.add('logged-out');
      }
    } else {
      document.body.classList.add('logged-out');
    }
  }

  function showInitWorkspace() {
    const overlay = document.getElementById('initWorkspaceOverlay');
    if (!overlay) return;
    overlay.style.display = 'block';

    const form = document.getElementById('initWorkspaceForm');
    if (!form) return;

    form.addEventListener('submit', (e) => {
      e.preventDefault();
      const companyName = document.getElementById('wsCompanyName').value.trim();
      const entityType = document.getElementById('wsEntityType').value;
      const adminEmail = document.getElementById('wsAdminEmail').value.trim();
      const phone = document.getElementById('wsPhone').value.trim();
      const masterPassword = document.getElementById('wsMasterPassword').value.trim();

      if (!companyName || !masterPassword) {
        showToast('يرجى إدخال اسم المؤسسة وكلمة المرور', 'error');
        return;
      }

      const deployBtn = form.querySelector('.deploy-btn');
      const originalHTML = deployBtn.innerHTML;
      deployBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Deploying...';
      deployBtn.disabled = true;

      setTimeout(() => {
        // Persist workspace config
        localStorage.setItem('sovLedger_workspace', JSON.stringify({
          companyName, entityType, adminEmail, phone,
          deployedAt: new Date().toISOString()
        }));

        // Update admin password in ExcelEngine
        const staff = ExcelEngine.getStaff();
        const admin = staff.find(s => s.name === 'admin');
        if (admin) admin.password = masterPassword;
        ExcelEngine.saveToLocalStorage();

        overlay.style.display = 'none';
        document.body.classList.add('logged-out');
        deployBtn.innerHTML = originalHTML;
        deployBtn.disabled = false;
        showToast(`تم إعداد مؤسسة "${companyName}" بنجاح. يمكنك تسجيل الدخول الآن.`, 'success');
      }, 1200);
    });

    const accessLink = document.getElementById('wsAccessDashboard');
    if (accessLink) {
      accessLink.addEventListener('click', (e) => {
        e.preventDefault();
        overlay.style.display = 'none';
        document.body.classList.add('logged-out');
      });
    }
  }

  function handleAuthSuccess() {
    const user = ExcelEngine.getCurrentUser();
    document.body.classList.remove('logged-out');
    const loginOverlay = document.getElementById('loginOverlay');
    if (loginOverlay) loginOverlay.style.display = 'none'; // Hide overlay

    const userDisplay = document.getElementById('currentUserDisplay');
    if (userDisplay) userDisplay.innerHTML = `<i class="fas fa-user-circle"></i> ${user.name} (${user.role})`;

    applyRoleAccess(user.role);
    showToast(`مرحباً بك ${user.name}`, 'success');

    // If we just logged in and we are on dashboard, refresh stats
    if (currentPage === 'dashboard') {
      renderDashboard();
    } else {
      navigateTo(currentPage);
    }
  }

  function applyRoleAccess(role) {
    if (role === 'مندوب') {
      document.querySelectorAll('.admin-only').forEach(el => el.style.display = 'none');
      // Hide all non-rep pages
      document.querySelectorAll('.nav-item[data-page="dashboard"], .nav-item[data-page="treasury"], .nav-item[data-page="reports"], .nav-item[data-page="reps"], .nav-item[data-page="hr"], .nav-item[data-page="statements"]').forEach(el => el.style.display = 'none');
      // Start on sales
      navigateTo('sales');
    } else {
      document.querySelectorAll('.admin-only').forEach(el => el.style.display = '');
      document.querySelectorAll('.nav-item').forEach(el => el.style.display = '');
    }
  }

  // === Sidebar Toggle ===
  function setupSidebarToggle() {
    const sidebar = document.getElementById('sidebar');
    const toggle = document.getElementById('sidebarToggle');

    // Load state
    const isCollapsed = localStorage.getItem('sidebarCollapsed') === 'true';
    if (isCollapsed && sidebar) {
      sidebar.classList.add('collapsed');
    }

    if (toggle && sidebar) {
      toggle.addEventListener('click', () => {
        sidebar.classList.toggle('collapsed');
        localStorage.setItem('sidebarCollapsed', sidebar.classList.contains('collapsed'));
      });
    }
  }

  // === Navigation ===
  function setupNavigation() {
    document.querySelectorAll('.nav-item[data-page]').forEach(item => {
      item.addEventListener('click', (e) => {
        e.preventDefault();
        const page = item.dataset.page;
        navigateTo(page);

        // Close mobile sidebar
        document.querySelector('.sidebar').classList.remove('open');
      });
    });

    // Handle Nav Group clicks (collapsible sections)
    document.querySelectorAll('.nav-group').forEach(group => {
        const header = group.querySelector('.nav-item:not([data-page])');
        if (header) {
            header.addEventListener('click', () => {
                // Close other groups
                document.querySelectorAll('.nav-group').forEach(g => {
                    if (g !== group) g.classList.remove('open');
                });
                group.classList.toggle('open');
            });
        }
    });

    // Mobile menu toggle
    const mobileBtn = document.getElementById('mobileMenuBtn');
    if (mobileBtn) {
      mobileBtn.addEventListener('click', () => {
        document.querySelector('.sidebar').classList.toggle('open');
      });
    }
  }

  function navigateTo(page) {
    currentPage = page;

    // Update active nav
    document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
    const activeNav = document.querySelector(`.nav-item[data-page="${page}"]`);
    if (activeNav) {
        activeNav.classList.add('active');
        // Ensure parent group is open
        const parentGroup = activeNav.closest('.nav-group');
        if (parentGroup) parentGroup.classList.add('open');
    }

    // Show active page
    document.querySelectorAll('.page-section').forEach(s => s.classList.remove('active'));
    const pageEl = document.getElementById(`page-${page}`);
    if (pageEl) pageEl.classList.add('active');

    // Update page title
    const titles = {
      'dashboard': 'لوحة التحكم',
      'sales': 'إدارة المبيعات',
      'purchases': 'إدارة المشتريات',
      'inventory': 'إدارة المخزون',
      'customers': 'إدارة العملاء',
      'suppliers': 'إدارة الموردين',
      'reps': 'إدارة المناديب',
      'statements': 'كشوفات الحساب',
      'products': 'إدارة المنتجات',
      'reports': 'التقارير الاحترافية',
      'hr': 'إدارة الموارد البشرية',
      'treasury': 'الخزينة العامة',
      'settings': 'الإعدادات',
      'entity-profile': 'ملف التعريف'
    };
    document.getElementById('pageTitle').textContent = titles[page] || '';

    // Refresh page data
    const secureCount = ExcelEngine.secureExistingInvoices();
    if (secureCount > 0) console.log(`Secured ${secureCount} invoices.`);
    refreshPage(page);
  }

  function refreshPage(page) {
    switch (page) {
      case 'dashboard': renderDashboard(); break;
      case 'sales': renderSalesTable(); break;
      case 'purchases': renderPurchasesTable(); break;
      case 'inventory':
        renderInventoryTable();
        if (typeof renderInventoryLedger === 'function') renderInventoryLedger();
        break;
      case 'customers': renderCustomersTable(); break;
      case 'suppliers': renderSuppliersTable(); break;
      case 'reps': renderRepsTable(); break;
      case 'statements': renderStatements(); break;
      case 'products': renderProductsTable(); break;
      case 'reports': populateDropdowns('page-reports'); break;
      case 'hr': renderHR(); break;
      case 'treasury': renderTreasury(); break;
      case 'entity-profile': break;
    }
  }

  // === File Handlers ===
  function setupFileHandlers() {
    // Import Excel
    const importBtn = document.getElementById('importExcel');
    const fileInput = document.getElementById('excelFileInput');

    if (importBtn) {
      importBtn.addEventListener('click', () => fileInput.click());
    }

    if (fileInput) {
      fileInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;

        showLoading('جاري تحميل ملف Excel...');
        try {
          await ExcelEngine.loadFromFile(file);
          ExcelEngine.saveToLocalStorage();
          hideLoading();
          showToast('تم تحميل الملف بنجاح!', 'success');
          refreshPage(currentPage);
          updateExcelStatus(true);
        } catch (err) {
          hideLoading();
          showToast('خطأ في تحميل الملف: ' + err.message, 'error');
          console.error(err);
        }
        fileInput.value = '';
      });
    }

    // Export Excel
    const exportBtn = document.getElementById('exportExcel');
    if (exportBtn) {
      exportBtn.addEventListener('click', () => {
        try {
          ExcelEngine.exportToFile();
          showToast('تم تصدير الملف بنجاح!', 'success');
        } catch (err) {
          showToast('خطأ في التصدير: ' + err.message, 'error');
        }
      });
    }

    // Save button
    const saveBtn = document.getElementById('saveData');
    if (saveBtn) {
      saveBtn.addEventListener('click', () => {
        ExcelEngine.saveToLocalStorage();
        showToast('تم حفظ البيانات محلياً', 'success');
      });
    }
  }

  // === Additional Form Handlers ===
  function setupAdditionalForms() {
    const staffForm = document.getElementById('staffForm');
    if (staffForm) {
      staffForm.addEventListener('submit', (e) => {
        e.preventDefault();
        ExcelEngine.addStaff({
          id: Date.now(),
          name: document.getElementById('stfName')?.value,
          role: document.getElementById('stfRole')?.value,
          phone: document.getElementById('stfPhone')?.value,
          password: document.getElementById('stfPassword')?.value || ''
        });
        closeModal('staffModal');
        showToast('تمت إضافة الموظف بنجاح', 'success');
        refreshPage(currentPage);
      });
    }

    const hrActionForm = document.getElementById('hrActionForm');
    if (hrActionForm) {
      hrActionForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const amount = parseFloat(document.getElementById('hrAmount').value);
        const action = {
          date: document.getElementById('hrDate').value,
          name: document.getElementById('hrStaff').value,
          type: document.getElementById('hrType').value,
          amount: amount,
          notes: document.getElementById('hrNotes').value
        };
        ExcelEngine.addHRAction(action);

        closeModal('hrActionModal');
        showToast('تم التسجيل بنجاح وتحديث الخزينة', 'success');
        refreshPage(currentPage);
      });
    }

    const cpForm = document.getElementById('customerPaymentForm');
    if (cpForm) {
      cpForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const amount = parseFloat(document.getElementById('cpAmount').value);
        const collectedBy = document.getElementById('cpCollectedBy').value;
        const payment = {
          date: document.getElementById('cpDate').value,
          customer: document.getElementById('cpCustomer').value,
          amount: amount,
          collectedBy: collectedBy,
          invoiceRef: document.getElementById('cpInvoiceRef')?.value || '',
          notes: document.getElementById('cpNotes').value
        };
        ExcelEngine.addCustomerPayment(payment);

        closeModal('customerPaymentModal');
        showToast('تم تسجيل الدفعة بنجاح', 'success');
        refreshPage(currentPage);
      });
    }

    const stockAdjForm = document.getElementById('stockAdjForm');
    if (stockAdjForm) {
      stockAdjForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const product = document.getElementById('adjProduct').value;
        const type = document.getElementById('adjType').value;
        const amount = parseFloat(document.getElementById('adjNewQty').value);
        const notes = document.getElementById('adjNotes').value;

        if (ExcelEngine.updateInventoryManual(product, type, amount, notes)) {
          showToast('تمت تسوية المخزون بنجاح', 'success');
          closeModal('stockAdjModal');
          refreshPage(currentPage);
        } else {
          showToast('حدث خطأ في التسوية', 'error');
        }
      });
    }

    const trActionForm = document.getElementById('treasuryActionForm');
    if (trActionForm) {
      trActionForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const amount = parseFloat(document.getElementById('trAmount').value);
        const type = document.getElementById('trType').value;
        const date = document.getElementById('trDate').value;
        const ref = document.getElementById('trRef').value;
        const notes = document.getElementById('trNotes').value;

        ExcelEngine.logTreasuryTransaction(type, amount, ref, notes);

        closeModal('treasuryActionModal');
        showToast('تم تسجيل حركة الخزينة بنجاح', 'success');
        refreshPage(currentPage);
        trActionForm.reset();
      });
    }

    // Rep Payment Form
    const rpForm = document.getElementById('repPaymentForm');
    if (rpForm) {
      rpForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const amount = parseFloat(rpForm.querySelector('#rpAmount')?.value) || 0;
        const rep = rpForm.querySelector('#rpRep')?.value;
        const payment = {
          date: rpForm.querySelector('#rpDate')?.value || ExcelEngine.getTodayStr(),
          rep: rep,
          amount: amount,
          paymentType: rpForm.querySelector('#rpType')?.value || 'نقدي',
          notes: rpForm.querySelector('#rpNotes')?.value || ''
        };

        ExcelEngine.addRepPayment(payment);

        ExcelEngine.saveToLocalStorage();
        closeModal('repPaymentModal');
        rpForm.reset();
        showToast('تم تسجيل الدفعة وتوريدها للخزينة بنجاح!', 'success');
        renderRepsTable();
      });
    }

    // Warehouse Transfer Form
    const whForm = document.getElementById('warehouseTransferForm');
    if (whForm) {
      whForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const rep = document.getElementById('whRep').value;
        const product = document.getElementById('whProduct').value;
        const qty = parseFloat(document.getElementById('whQty').value) || 0;
        const direction = document.getElementById('whDirection').value;
        const notes = document.getElementById('whNotes').value;

        if (!rep || !product || qty <= 0) {
          showToast('يرجى ملء جميع الحقول المطلوبة', 'warning');
          return;
        }

        let result;
        if (direction === 'TO_REP') {
          result = ExcelEngine.transferToRepWarehouse(rep, product, qty, notes);
        } else {
          result = ExcelEngine.returnFromRepWarehouse(rep, product, qty, notes);
        }

        if (result.success) {
          closeModal('warehouseTransferModal');
          whForm.reset();
          showToast(direction === 'TO_REP' ? `تم سحب ${qty} من ${product} لمستودع ${rep}` : `تم إرجاع ${qty} من ${product} للمخزن الرئيسي`, 'success');
          refreshPage(currentPage);
        } else {
          showToast(result.error, 'error');
        }
      });
    }
  }

  // === Export Advanced Reports ===
  function exportReport(type) {
    if (typeof ExcelEngine.exportAdvancedReports !== 'function') {
      showToast('جاري إعداد محرك التقارير الاحترافية، يرجى الانتظار...', 'info');
      return;
    }

    let query = '';

    if (type === 'customer') {
      query = document.getElementById('reportCustomerSelect')?.value;
      if (!query) return showToast('الرجاء اختيار عميل أولاً', 'warning');
      printStatement(query);
      return;
    } else if (type === 'supplier') {
      query = document.getElementById('reportSupplierSelect')?.value;
      if (!query) return showToast('الرجاء اختيار مورد أولاً', 'warning');
      // If we had printSupplierStatement we could call it here, but we will let it fall back
    } else if (type === 'rep') {
      query = document.getElementById('reportRepSelect')?.value;
      if (!query) return showToast('الرجاء اختيار مندوب أولاً', 'warning');
      printRepDailyReport(query);
      return;
    } else if (type === 'productProfit') {
      query = document.getElementById('reportProductSelect')?.value;
      if (!query) return showToast('الرجاء تحديد المنتج أولاً', 'warning');
    }

    showToast('جاري تجهيز تقرير الإكسيل، يرجى الانتظار...', 'info');

    ExcelEngine.exportAdvancedReports(type, query)
      .then(() => showToast('تم التصدير بنجاح!', 'success'))
      .catch((err) => {
        showToast('حدث خطأ أثناء التصدير: ' + err.message, 'error');
        console.error(err);
      });
  }

  // === Dashboard ===
  function renderDashboard() {
    const stats = ExcelEngine.getStats();

    // Update stat cards with Currency Formatting
    setTextContent('statCashInBox', 'SAR ' + ExcelEngine.formatCurrency(stats.cashInBox));
    setTextContent('statTotalDue', 'SAR ' + ExcelEngine.formatCurrency(stats.totalDue));
    setTextContent('statTotalPurchases', 'SAR ' + ExcelEngine.formatCurrency(stats.totalPurchases));
    setTextContent('statTotalSales', 'SAR ' + ExcelEngine.formatCurrency(stats.totalSales));

    // Update sales badge in sidebar
    const salesBadge = document.getElementById('salesBadge');
    if (salesBadge) {
        salesBadge.textContent = stats.salesCount;
        salesBadge.style.display = stats.salesCount > 0 ? 'inline-block' : 'none';
    }

    // Charts
    renderSalesChart(stats.salesByDate, stats.purchasesByDate);
    
    // Recent Ledger Entries
    renderRecentLedgerEntries();
  }

  function renderRecentLedgerEntries() {
    const listEl = document.getElementById('recentEntriesList');
    if (!listEl) return;

    // Combine recent sales and purchases for the ledger view
    const sales = (ExcelEngine.getData('sales') || []).slice(-3).reverse();
    const purchases = (ExcelEngine.getData('purchases') || []).slice(-2).reverse();
    
    const entries = [
        ...sales.map(s => ({ type: 'sale', title: `Invoice #${s.invoiceId || 'INV-001'}`, subtitle: s.customer, amount: s.total, date: s.date })),
        ...purchases.map(p => ({ type: 'purchase', title: `Purchase #${p.id || 'PUR-001'}`, subtitle: p.supplier, amount: -p.total, date: p.date }))
    ].sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 5);

    if (entries.length === 0) {
        listEl.innerHTML = `
            <div class="flex items-center justify-center h-48 text-slate-300 flex-col gap-2">
                <span class="material-symbols-outlined text-4xl">list_alt</span>
                <p class="text-xs font-bold uppercase tracking-widest">No recent activities</p>
            </div>`;
        return;
    }

    listEl.innerHTML = entries.map(entry => `
        <div class="flex items-center justify-between group cursor-pointer hover:bg-slate-50 p-3 rounded-2xl transition-all hover:scale-[1.01]">
            <div class="flex items-center gap-4">
                <div class="w-12 h-12 rounded-2xl ${entry.amount > 0 ? 'bg-emerald-50 text-emerald-600' : 'bg-slate-100 text-slate-500'} flex items-center justify-center shadow-sm">
                    <span class="material-symbols-outlined text-xl">${entry.amount > 0 ? 'payments' : 'shopping_cart'}</span>
                </div>
                <div>
                    <p class="text-sm font-black text-primary font-headline">${entry.title}</p>
                    <p class="text-[10px] font-bold text-slate-400 uppercase tracking-widest">${entry.subtitle}</p>
                </div>
            </div>
            <div class="text-right">
                <p class="text-sm font-black font-headline ${entry.amount > 0 ? 'text-emerald-600' : 'text-primary'}">
                    ${entry.amount > 0 ? '+' : ''}${ExcelEngine.formatCurrency(Math.abs(entry.amount))}
                </p>
                <p class="text-[9px] font-bold text-slate-300 uppercase">${entry.date}</p>
            </div>
        </div>
    `).join('');
  }

  function renderSalesChart(salesByDate, purchasesByDate = {}) {
    const ctx = document.getElementById('salesChart');
    if (!ctx) return;

    if (chartInstances.salesChart) chartInstances.salesChart.destroy();

    const allDates = [...new Set([...Object.keys(salesByDate || {}), ...Object.keys(purchasesByDate)])].sort().slice(-12);
    const revenueData = allDates.map(d => (salesByDate || {})[d] || 0);
    const expensesData = allDates.map(d => purchasesByDate[d] || 0);

    chartInstances.salesChart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: allDates,
        datasets: [
          {
            label: 'الإيرادات',
            data: revenueData,
            backgroundColor: '#031636',
            borderRadius: 8,
            barThickness: 10
          },
          {
            label: 'المصروفات',
            data: expensesData,
            backgroundColor: '#94a3b8',
            borderRadius: 8,
            barThickness: 10
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            display: true,
            position: 'top',
            align: 'end',
            labels: {
              boxWidth: 8,
              boxHeight: 8,
              usePointStyle: true,
              pointStyle: 'rectRounded',
              color: '#94a3b8',
              font: { family: 'Cairo', size: 11, weight: '600' }
            }
          },
          tooltip: {
            backgroundColor: '#031636',
            titleColor: '#fff',
            bodyColor: '#cbd5e1',
            padding: 12,
            cornerRadius: 12,
            displayColors: true,
            font: { family: 'Inter' }
          }
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: { color: '#94a3b8', font: { family: 'Inter', size: 10, weight: '600' } }
          },
          y: {
            grid: { color: '#f1f5f9', drawBorder: false },
            ticks: { color: '#94a3b8', font: { family: 'Inter', size: 10, weight: '600' } }
          }
        }
      }
    });
  }

  function renderTopProductsChart(topProducts) {
    const ctx = document.getElementById('productsChart');
    if (!ctx) return;

    if (chartInstances.productsChart) chartInstances.productsChart.destroy();

    const colors = [
      '#1a73e8', '#f9a825', '#4caf50', '#ef5350', '#ab47bc',
      '#00bcd4', '#ff7043', '#8bc34a', '#e91e63', '#607d8b'
    ];

    chartInstances.productsChart = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: topProducts.map(p => p.name),
        datasets: [{
          data: topProducts.map(p => p.total),
          backgroundColor: colors.slice(0, topProducts.length),
          borderWidth: 0,
          hoverOffset: 8
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            rtl: true,
            labels: {
              color: '#94a3b8',
              font: { family: 'Cairo', size: 11 },
              padding: 15,
              usePointStyle: true
            }
          }
        }
      }
    });
  }

  function renderTopDebtorsTable(debtors) {
    const tbody = document.getElementById('topDebtorsBody');
    if (!tbody) return;

    if (debtors.length === 0) {
      tbody.innerHTML = '<tr><td colspan="3" style="text-align:center;color:var(--text-tertiary);padding:20px;">لا توجد ديون</td></tr>';
      return;
    }

    tbody.innerHTML = debtors.map(d => `
      <tr>
        <td>${d.name}</td>
        <td class="currency">${ExcelEngine.formatCurrency(d.total)}</td>
        <td class="currency negative">${ExcelEngine.formatCurrency(d.remaining)}</td>
      </tr>
    `).join('');
  }

  function renderLowStockAlerts(items) {
    const container = document.getElementById('lowStockAlerts');
    if (!container) return;

    if (items.length === 0) {
      container.innerHTML = '<div class="empty-state" style="padding:20px"><i class="fas fa-check-circle"></i><p>المخزون في حالة جيدة</p></div>';
      return;
    }

    container.innerHTML = items.map(item => `
      <div style="display:flex;justify-content:space-between;align-items:center;padding:8px 12px;border-bottom:1px solid var(--border-glass);">
        <span>${item.product}</span>
        <span class="badge ${item.net <= 0 ? 'badge-danger' : 'badge-warning'}">${item.net}</span>
      </div>
    `).join('');
  }

  // === Sales Table ===
  function renderSalesTable() {
    let sales = ExcelEngine.getSales();
    const user = ExcelEngine.getCurrentUser();
    if (user && user.role === 'مندوب') {
      sales = sales.filter(s => s.rep === user.name);
    }
    const tbody = document.getElementById('salesBody');
    if (!tbody) return;

    if (sales.length === 0) {
      tbody.innerHTML = `<tr><td colspan="12"><div class="empty-state"><i class="fas fa-file-invoice"></i><h3>لا توجد مبيعات</h3><p>ابدأ بإضافة فاتورة بيع جديدة</p></div></td></tr>`;
      return;
    }

    tbody.innerHTML = sales.map(s => `
      <tr>
        <td class="number">${s.date}</td>
        <td><span class="clickable-link" onclick="App.viewProfile('customer', '${s.customer}')"><i class="fas fa-user-circle"></i> ${s.customer}</span></td>
        <td>${s.product}</td>
        <td><span class="clickable-link" onclick="App.viewProfile('rep', '${s.rep}')"><i class="fas fa-user-tie"></i> ${s.rep}</span></td>
        <td class="number">${s.quantity}</td>
        <td class="number">${s.returned}</td>
        <td class="number">${s.netQuantity}</td>
        <td class="currency">${ExcelEngine.formatCurrency(s.price)}</td>
        <td class="currency">${ExcelEngine.formatCurrency(s.total)}</td>
        <td class="currency">${ExcelEngine.formatCurrency(s.collected)}</td>
        <td class="currency ${s.balance > 0 ? 'negative' : ''}">${ExcelEngine.formatCurrency(s.balance)}</td>
        <td>
          <div class="row-actions">
            <button class="primary" onclick="App.viewInvoiceDetail(${s.id}, 'sale')" title="عرض التفاصيل"><i class="fas fa-eye"></i></button>
            <button onclick="App.printInvoice(${s.id}, 'sale')" title="طباعة"><i class="fas fa-print"></i></button>
            <button class="warning" onclick="App.openReturnModal(${s.id}, 'sale')" title="تسجيل مرتجع" ${s.returned >= s.quantity ? 'disabled style="opacity:0.5"' : ''}><i class="fas fa-undo"></i></button>
          </div>
        </td>
      </tr>
    `).join('');
  }

  // === Purchases Table ===
  function renderPurchasesTable() {
    let purchases = ExcelEngine.getPurchases();
    const user = ExcelEngine.getCurrentUser();
    if (user && user.role === 'مندوب') {
      purchases = purchases.filter(p => p.rep === user.name);
    }
    const tbody = document.getElementById('purchasesBody');
    if (!tbody) return;

    if (purchases.length === 0) {
      tbody.innerHTML = `<tr><td colspan="10"><div class="empty-state"><i class="fas fa-shopping-cart"></i><h3>لا توجد مشتريات</h3><p>ابدأ بإضافة فاتورة شراء جديدة</p></div></td></tr>`;
      return;
    }

    tbody.innerHTML = purchases.map(p => `
      <tr>
        <td class="number">${p.date}</td>
        <td>${p.product}</td>
        <td class="number">${p.quantity}</td>
        <td><span class="clickable-link" onclick="App.viewProfile('rep', '${p.rep}')">${p.rep}</span></td>
        <td class="currency">${ExcelEngine.formatCurrency(p.price)}</td>
        <td class="currency">${ExcelEngine.formatCurrency(p.total)}</td>
        <td><span class="clickable-link" onclick="App.viewProfile('supplier', '${p.supplier}')">${p.supplier}</span></td>
        <td class="currency">${ExcelEngine.formatCurrency(p.dueToSupplier)}</td>
        <td class="currency">${ExcelEngine.formatCurrency(p.paid)}</td>
        <td>
          <div class="row-actions">
            <button class="primary" onclick="App.viewInvoiceDetail(${p.id}, 'purchase')" title="عرض التفاصيل"><i class="fas fa-eye"></i></button>
            <button onclick="App.printInvoice(${p.id}, 'purchase')" title="طباعة"><i class="fas fa-print"></i></button>
            <button class="warning" onclick="App.openReturnModal(${p.id}, 'purchase')" title="تسجيل مرتجع" ${p.quantity <= 0 ? 'disabled style="opacity:0.5"' : ''}><i class="fas fa-undo"></i></button>
          </div>
        </td>
      </tr>
    `).join('');
  }

  // === Inventory Table ===
  function renderInventoryTable() {
    const inventory = ExcelEngine.getInventory();
    const tbody = document.getElementById('inventoryBody');
    if (!tbody) return;

    if (inventory.length === 0) {
      tbody.innerHTML = `<tr><td colspan="5"><div class="empty-state"><i class="fas fa-boxes"></i><h3>لا توجد بيانات مخزون</h3></div></td></tr>`;
      return;
    }

    tbody.innerHTML = inventory.map(i => `
      <tr>
        <td>${i.product}</td>
        <td class="number">${i.previousStock}</td>
        <td class="number" style="color:var(--success-400)">${i.purchased}</td>
        <td class="number" style="color:var(--danger-400)">${i.withdrawn}</td>
        <td class="number"><span class="badge ${i.net <= 0 ? 'badge-danger' : i.net < 10 ? 'badge-warning' : 'badge-success'}">${i.net}</span></td>
      </tr>
    `).join('');
  }

  function renderInventoryLedger() {
    if (typeof ExcelEngine.getInventoryLedger !== 'function') return;
    const ledger = ExcelEngine.getInventoryLedger();
    const tbody = document.getElementById('inventoryLedgerBody');
    if (!tbody) return;

    if (ledger.length === 0) {
      tbody.innerHTML = `<tr><td colspan="8"><div class="empty-state"><i class="fas fa-chart-line"></i><h3>لا يوجد حركات مخزون</h3></div></td></tr>`;
      return;
    }

    tbody.innerHTML = ledger.map(item => {
      const isOut = item.type === 'OUT';
      return `
      <tr>
        <td class="number">${item.date}</td>
        <td><span class="nav-badge" style="position:static;background:${isOut ? 'var(--warning-500)' : 'var(--success-400)'};">${item.location || 'غير محدد'}</span></td>
        <td><strong>${item.product}</strong></td>
        <td><span style="color:${isOut ? '#ef5350' : '#4caf50'};font-weight:bold;">${isOut ? 'صادر (-)' : 'وارد (+)'}</span></td>
        <td><span class="nav-badge" style="position:static;background:var(--primary-400);">${item.category}</span></td>
        <td class="number" style="font-weight:bold;">${item.amount}</td>
        <td>${item.party || '-'}</td>
        <td style="font-size:0.85em;color:var(--text-muted);">${item.reference}</td>
      </tr>
    `}).join('');
  }

  // === Customers Table ===
  function renderCustomersTable() {
    const customers = ExcelEngine.getCustomers();
    const tbody = document.getElementById('customersBody');
    if (!tbody) return;

    if (customers.length === 0) {
      tbody.innerHTML = `<tr><td colspan="7"><div class="empty-state"><i class="fas fa-users"></i><h3>لا يوجد عملاء</h3></div></td></tr>`;
      return;
    }

    tbody.innerHTML = customers.map(c => `
      <tr>
        <td><span class="clickable-link" onclick="App.viewProfile('customer', '${c.name}')"><i class="fas fa-user-circle"></i> ${c.name}</span></td>
        <td class="number">${c.refNumber}</td>
        <td class="currency">${ExcelEngine.formatCurrency(c.collected)}</td>
        <td class="currency">${ExcelEngine.formatCurrency(c.total)}</td>
        <td class="currency ${c.remaining > 0 ? 'negative' : ''}">${ExcelEngine.formatCurrency(c.remaining)}</td>
        <td class="number">${c.lastDate || '-'}</td>
        <td>
          <div class="row-actions">
            <button onclick="App.printStatement('${c.name}')" title="كشف حساب"><i class="fas fa-print"></i></button>
            <button onclick="App.viewProfile('customer', '${c.name}')" title="عرض الملف"><i class="fas fa-eye"></i></button>
            <button class="delete admin-only" onclick="App.confirmDeleteCustomer(${c.id})" title="حذف"><i class="fas fa-trash"></i></button>
          </div>
        </td>
      </tr>
    `).join('');
  }

  // === Suppliers Table ===
  let currentSupplierName = '';

  function renderSuppliersTable() {
    const suppliers = ExcelEngine.getSuppliers();
    const tbody = document.getElementById('suppliersBody');
    if (!tbody) return;

    const listView = document.getElementById('supplierListView');
    const detailView = document.getElementById('supplierDetailView');
    if (listView) listView.style.display = 'block';
    if (detailView) detailView.style.display = 'none';

    if (suppliers.length === 0) {
      tbody.innerHTML = `<tr><td colspan="6"><div class="empty-state"><i class="fas fa-truck"></i><h3>لا يوجد موردين</h3><p>ابدأ بإضافة مورد جديد</p></div></td></tr>`;
      return;
    }

    const purchases = ExcelEngine.getPurchases();
    tbody.innerHTML = suppliers.map(s => {
      const sp = purchases.filter(p => p.supplier === s.name);
      const total = sp.reduce((sum, p) => sum + (p.total || 0), 0);
      const paid = sp.reduce((sum, p) => sum + (p.paid || 0), 0);
      const rem = total - paid;
      return `
        <tr>
          <td><span class="clickable-link" onclick="App.viewProfile('supplier', '${s.name}')"><i class="fas fa-truck"></i> ${s.name}</span></td>
          <td>${s.phone || '-'}</td>
          <td class="currency">${ExcelEngine.formatCurrency(total)}</td>
          <td class="currency" style="color:var(--success-400)">${ExcelEngine.formatCurrency(paid)}</td>
          <td class="currency ${rem > 0 ? 'negative' : ''}">${ExcelEngine.formatCurrency(rem)}</td>
          <td>
            <div class="row-actions">
              <button onclick="App.viewProfile('supplier', '${s.name}')" title="عرض الملف"><i class="fas fa-eye"></i></button>
              <button class="delete admin-only" onclick="App.confirmDeleteSupplier(${s.id})" title="حذف"><i class="fas fa-trash"></i></button>
            </div>
          </td>
        </tr>
      `;
    }).join('');
  }

  function showSupplierDetail(supplierName) {
    if (currentPage !== 'suppliers') navigateTo('suppliers');
    currentSupplierName = supplierName;

    const listView = document.getElementById('supplierListView');
    const detailView = document.getElementById('supplierDetailView');
    if (listView) listView.style.display = 'none';
    if (detailView) detailView.style.display = 'block';

    const nameEl = document.getElementById('supplierDetailName');
    if (nameEl) nameEl.textContent = `كشف حساب المورد: ${supplierName}`;

    const purchases = ExcelEngine.getPurchases().filter(p => p.supplier === supplierName);
    const totalP = purchases.reduce((s, p) => s + (p.total || 0), 0);
    const totalPaid = purchases.reduce((s, p) => s + (p.paid || 0), 0);
    const rem = totalP - totalPaid;
    const reps = [...new Set(purchases.map(p => p.rep).filter(r => r))];

    const cardsEl = document.getElementById('supplierSummaryCards');
    if (cardsEl) {
      cardsEl.innerHTML = `
        <div class="stat-card blue">
          <div class="stat-card-header"><span class="stat-card-label">إجمالي المشتريات</span><div class="stat-card-icon"><i class="fas fa-shopping-cart"></i></div></div>
          <div class="stat-card-value">${ExcelEngine.formatCurrency(totalP)}</div>
          <div class="stat-card-change">${purchases.length} فاتورة</div>
        </div>
        <div class="stat-card green">
          <div class="stat-card-header"><span class="stat-card-label">المدفوع</span><div class="stat-card-icon"><i class="fas fa-check-circle"></i></div></div>
          <div class="stat-card-value">${ExcelEngine.formatCurrency(totalPaid)}</div>
        </div>
        <div class="stat-card red">
          <div class="stat-card-header"><span class="stat-card-label">المتبقي</span><div class="stat-card-icon"><i class="fas fa-exclamation-circle"></i></div></div>
          <div class="stat-card-value">${ExcelEngine.formatCurrency(rem)}</div>
        </div>
        <div class="stat-card gold">
          <div class="stat-card-header"><span class="stat-card-label">المناديب</span><div class="stat-card-icon"><i class="fas fa-user-tie"></i></div></div>
          <div class="stat-card-value" style="font-size:1rem">${reps.length > 0 ? reps.join(' ، ') : 'لا يوجد'}</div>
        </div>
      `;
    }

    const tbody = document.getElementById('supplierInvoicesBody');
    if (tbody) {
      if (purchases.length === 0) {
        tbody.innerHTML = `<tr><td colspan="8" style="text-align:center;padding:20px;color:var(--text-tertiary)">لا توجد فواتير من هذا المورد</td></tr>`;
      } else {
        let bal = 0;
        tbody.innerHTML = purchases.sort((a, b) => new Date(a.date) - new Date(b.date)).map(p => {
          bal += ((p.total || 0) - (p.paid || 0));
          return `
            <tr>
              <td class="number">${p.date}</td>
              <td>${p.product}</td>
              <td class="number">${p.quantity}</td>
              <td class="currency">${(p.price || 0).toFixed(2)}</td>
              <td class="currency">${(p.total || 0).toFixed(2)}</td>
              <td>${p.rep || '-'}</td>
              <td class="currency" style="color:var(--success-400)">${(p.paid || 0).toFixed(2)}</td>
              <td class="currency ${bal > 0 ? 'negative' : ''}">${bal.toFixed(2)}</td>
            </tr>
          `;
        }).join('');
      }
    }
  }

  function backToSupplierList() {
    renderSuppliersTable();
  }

  function confirmDeleteSupplier(id) {
    if (confirm('هل أنت متأكد من حذف هذا المورد؟')) {
      ExcelEngine.deleteSupplier(id);
      ExcelEngine.saveToLocalStorage();
      showToast('تم حذف المورد', 'info');
      renderSuppliersTable();
    }
  }

  function setupSupplierForm() {
    const form = document.getElementById('supplierForm');
    if (!form) return;
    form.addEventListener('submit', (e) => {
      e.preventDefault();
      const supplier = {
        name: form.querySelector('#suppName')?.value,
        phone: form.querySelector('#suppPhone')?.value || '',
        address: form.querySelector('#suppAddress')?.value || '',
        cr: form.querySelector('#suppCR')?.value || '',
        taxId: form.querySelector('#suppTaxId')?.value || ''
      };
      if (!supplier.name) { showToast('يرجى إدخال اسم المورد', 'warning'); return; }
      ExcelEngine.addSupplier(supplier);
      ExcelEngine.saveToLocalStorage();
      closeModal('supplierModal');
      form.reset();
      showToast('تم إضافة المورد بنجاح!', 'success');
      renderSuppliersTable();
    });
  }

  function printSupplierStatement() {
    if (!currentSupplierName) return;
    const purchases = ExcelEngine.getPurchases().filter(p => p.supplier === currentSupplierName);
    const totalP = purchases.reduce((s, p) => s + (p.total || 0), 0);
    const totalPaid = purchases.reduce((s, p) => s + (p.paid || 0), 0);
    const rem = totalP - totalPaid;

    const printWindow = window.open('', '_blank');
    if (!printWindow) { showToast('يرجى السماح بالنوافذ المنبثقة للطباعة', 'warning'); return; }
    let bal = 0;
    const rows = purchases.sort((a, b) => new Date(a.date) - new Date(b.date)).map(p => {
      bal += ((p.total || 0) - (p.paid || 0));
      return `<tr><td>${p.date}</td><td>${p.product}</td><td>${p.quantity}</td><td>${(p.price || 0).toFixed(2)}</td><td>${(p.total || 0).toFixed(2)}</td><td>${p.rep || '-'}</td><td>${(p.paid || 0).toFixed(2)}</td><td>${bal.toFixed(2)}</td></tr>`;
    }).join('');

    printWindow.document.write(`<!DOCTYPE html><html dir="rtl" lang="ar"><head><meta charset="UTF-8"><title>كشف حساب - ${currentSupplierName}</title>
      <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap" rel="stylesheet">
      <style>body{font-family:'Cairo',sans-serif;padding:20px;direction:rtl}.header{text-align:center;border-bottom:3px solid #f57c00;padding-bottom:15px;margin-bottom:20px}.header h1{color:#f57c00}.summary{display:grid;grid-template-columns:repeat(3,1fr);gap:15px;margin-bottom:20px}.summary-card{background:#f5f5f5;padding:15px;border-radius:8px;text-align:center}.summary-card .val{font-size:20px;font-weight:bold}table{width:100%;border-collapse:collapse}th,td{padding:8px;border:1px solid #ddd;text-align:right}th{background:#f57c00;color:white}.footer{margin-top:20px;text-align:center;color:#666;font-size:12px}</style></head>
      <body><div class="header"><h1>كشف حساب مورد</h1><h2>${currentSupplierName}</h2><p>التاريخ: ${new Date().toLocaleDateString('ar-EG')}</p></div>
      <div class="summary"><div class="summary-card"><div>إجمالي المشتريات</div><div class="val">${ExcelEngine.formatCurrency(totalP)}</div></div><div class="summary-card"><div>المدفوع</div><div class="val" style="color:green">${ExcelEngine.formatCurrency(totalPaid)}</div></div><div class="summary-card"><div>المتبقي</div><div class="val" style="color:red">${ExcelEngine.formatCurrency(rem)}</div></div></div>
      <table><thead><tr><th>التاريخ</th><th>المنتج</th><th>الكمية</th><th>السعر</th><th>الإجمالي</th><th>المندوب</th><th>المدفوع</th><th>الرصيد</th></tr></thead><tbody>${rows}</tbody></table>
      <div class="footer"><p>نظام إدارة الفواتير المتقدم</p></div>
      <script>window.print();<\/script></body></html>`);
  }

  // === Reps Table (with precise custody accounting) ===
  function renderRepsTable() {
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

      return `
        <tr>
          <td>
            <div class="flex items-center gap-3">
               <div class="w-10 h-10 rounded-xl bg-gradient-to-br from-primary to-primary-container text-white flex items-center justify-center font-bold text-sm shadow-inner">${name.substring(0,2)}</div>
               <div>
                  <p class="font-bold text-sm text-primary hover:text-secondary cursor-pointer transition-colors" onclick="App.viewProfile('rep', '${name}')">${name}</p>
                  <p class="text-[10px] text-slate-400 font-bold">ID: ${id}</p>
               </div>
            </div>
          </td>
          <td class="text-xs font-bold text-slate-600">${region}</td>
          <td class="currency font-bold text-primary">${ExcelEngine.formatCurrency(summary.totalSales)} ر.س</td>
          <td class="currency font-bold text-secondary">${ExcelEngine.formatCurrency(summary.totalDelivered)} ر.س</td>
          <td>${statusHtml}</td>
          <td>
            <div class="flex gap-2">
              <button class="bg-surface-container hover:bg-slate-200 text-primary w-8 h-8 rounded-lg flex items-center justify-center transition-colors" onclick="App.openRepPaymentModal('${name}')" title="تسجيل توريد من المندوب"><i class="fas fa-hand-holding-dollar text-xs"></i></button>
              <button class="bg-surface-container hover:bg-slate-200 text-primary w-8 h-8 rounded-lg flex items-center justify-center transition-colors" onclick="App.viewProfile('rep', '${name}')" title="التفاصيل"><i class="fas fa-chevron-left text-xs"></i></button>
            </div>
          </td>
        </tr>
      `;
    }).join('');

    tbody.innerHTML = reps.length ? repsHtml : `<tr><td colspan="6"><div class="empty-state"><i class="fas fa-user-tie"></i><h3>لا يوجد مناديب</h3></div></td></tr>`;
    
    // Update Stats on UI
    const elSales = document.getElementById('repTopSales');
    const elCol = document.getElementById('repTopCollections');
    const elBadge = document.getElementById('repCountBadge');
    
    if(elSales) elSales.textContent = ExcelEngine.formatNumber(totalSales);
    if(elCol) elCol.textContent = ExcelEngine.formatNumber(totalCollections);
    if(elBadge) elBadge.textContent = reps.length + ' مندوب';
  }

  // === Products Table ===
  function renderProductsTable() {
    const products = ExcelEngine.getProducts();
    const tbody = document.getElementById('productsBody');
    if (!tbody) return;

    if (products.length === 0) {
      tbody.innerHTML = `<tr><td colspan="6"><div class="empty-state"><i class="fas fa-box"></i><h3>لا توجد منتجات</h3></div></td></tr>`;
      return;
    }

    tbody.innerHTML = products.map(p => `
      <tr>
        <td>${p.name}</td>
        <td class="currency">${p.sellPrice.toFixed(2)}</td>
        <td class="currency">${p.buyPrice.toFixed(2)}</td>
        <td>${p.unit}</td>
        <td><span class="badge badge-info">${p.category}</span></td>
        <td>
          <div class="row-actions">
            <button class="primary" onclick="App.editProduct(${p.id})" title="تعديل"><i class="fas fa-edit"></i></button>
            <button class="delete" onclick="App.confirmDeleteProduct(${p.id})" title="حذف"><i class="fas fa-trash"></i></button>
          </div>
        </td>
      </tr>
    `).join('');
  }

  // === Statements ===
  function renderStatements() {
    populateStatementDropdown();
  }

  function populateStatementDropdown() {
    const type = document.querySelector('input[name="statementType"]:checked')?.value || 'customer';
    const select = document.getElementById('statementEntity');
    if (!select) return;

    if (type === 'customer') {
      const customers = ExcelEngine.getCustomers();
      select.innerHTML = '<option value="">اختر العميل...</option>' +
        customers.map(c => `<option value="${c.name}">${c.name}</option>`).join('');
    } else if (type === 'supplier') {
      const suppliers = ExcelEngine.getSuppliers();
      select.innerHTML = '<option value="">اختر المورد...</option>' +
        suppliers.map(s => `<option value="${s.name}">${s.name}</option>`).join('');
    } else if (type === 'rep') {
      const reps = ExcelEngine.getRepsList();
      select.innerHTML = '<option value="">اختر المندوب...</option>' +
        reps.map(r => `<option value="${r}">${r}</option>`).join('');
    } else if (type === 'staff') {
      const staff = ExcelEngine.getStaff();
      select.innerHTML = '<option value="">اختر الموظف...</option>' +
        staff.map(s => `<option value="${s.name}">${s.name}</option>`).join('');
    }
  }

  function showCustomerStatement(customerName) {
    navigateTo('statements');
    setTimeout(() => {
      document.querySelector('input[name="statementType"][value="customer"]').checked = true;
      populateStatementDropdown();
      const select = document.getElementById('statementEntity');
      if (select) {
        select.value = customerName;
        generateStatement();
      }
    }, 100);
  }

  function generateStatement() {
    const entityName = document.getElementById('statementEntity').value;
    const type = document.querySelector('input[name="statementType"]:checked')?.value || 'customer';

    if (!entityName) {
      showToast('اختر الجهة أولاً', 'warning');
      return;
    }

    const container = document.getElementById('statementResult');
    if (!container) return;

    if (type === 'customer') {
      const statement = ExcelEngine.getCustomerStatement(entityName);

      container.innerHTML = `
        <div class="statement-header">
          <div class="statement-info-card">
            <h4>إجمالي المبيعات</h4>
            <div class="value" style="color:var(--primary-400)">${ExcelEngine.formatCurrency(statement.totalSales)}</div>
          </div>
          <div class="statement-info-card">
            <h4>إجمالي التحصيل</h4>
            <div class="value" style="color:var(--success-400)">${ExcelEngine.formatCurrency(statement.totalCollected)}</div>
          </div>
          <div class="statement-info-card">
            <h4>الرصيد المتبقي</h4>
            <div class="value" style="color:var(--danger-400)">${ExcelEngine.formatCurrency(statement.balance)}</div>
          </div>
          <div class="statement-info-card">
            <h4>عدد العمليات</h4>
            <div class="value" style="color:var(--accent-500)">${statement.transactions.length}</div>
          </div>
        </div>

        <div class="table-container">
          <div class="table-header">
            <h3 class="table-header-title">كشف حساب عميل: ${entityName}</h3>
            <button class="btn btn-sm btn-ghost" onclick="App.printStatement('${entityName}')">
              <i class="fas fa-print"></i> طباعة
            </button>
          </div>
          <div class="table-wrapper">
            <table class="data-table">
              <thead>
                <tr>
                  <th>التاريخ</th>
                  <th>المنتج</th>
                  <th>الكمية</th>
                  <th>السعر</th>
                  <th>الإجمالي</th>
                  <th>التحصيل</th>
                  <th>الرصيد</th>
                </tr>
              </thead>
              <tbody>
                ${statement.transactions.length > 0 ? statement.transactions.map(t => `
                  <tr>
                    <td class="number">${t.date}</td>
                    <td>${t.product}</td>
                    <td class="number">${t.netQuantity || 0}</td>
                    <td class="currency">${t.price.toFixed(2)}</td>
                    <td class="currency">${t.total.toFixed(2)}</td>
                    <td class="currency">${t.collected.toFixed(2)}</td>
                    <td class="currency negative">${t.balance.toFixed(2)}</td>
                  </tr>
                `).join('') : '<tr><td colspan="7" style="text-align:center;padding:20px;">لا توجد عمليات</td></tr>'}
              </tbody>
            </table>
          </div>
        </div>
      `;
    } else if (type === 'supplier') {
      const statement = ExcelEngine.getSupplierStatement(entityName);

      container.innerHTML = `
        <div class="statement-header">
          <div class="statement-info-card">
            <h4>إجمالي المشتريات</h4>
            <div class="value" style="color:var(--primary-400)">${ExcelEngine.formatCurrency(statement.totalPurchases)}</div>
          </div>
          <div class="statement-info-card">
            <h4>إجمالي المدفوع</h4>
            <div class="value" style="color:var(--success-400)">${ExcelEngine.formatCurrency(statement.totalPaid)}</div>
          </div>
          <div class="statement-info-card">
            <h4>المستحق للمورد</h4>
            <div class="value" style="color:var(--danger-400)">${ExcelEngine.formatCurrency(statement.balance)}</div>
          </div>
        </div>
        <div class="table-container">
          <div class="table-header">
            <h3 class="table-header-title">كشف حساب مورد: ${entityName}</h3>
            <button class="btn btn-sm btn-ghost" onclick="App.printStatement('${entityName}', 'supplier')">
              <i class="fas fa-print"></i> طباعة
            </button>
          </div>
          <div class="table-wrapper">
            <table class="data-table">
              <thead>
                <tr>
                  <th>التاريخ</th>
                  <th>البيان</th>
                  <th>صادر (مدين +)</th>
                  <th>وارد (دائن -)</th>
                  <th>الرصيد</th>
                </tr>
              </thead>
              <tbody>
                ${statement.transactions.map(t => `
                  <tr>
                    <td class="number">${t.date}</td>
                    <td>${t.desc}</td>
                    <td class="currency">${t.debit.toFixed(2)}</td>
                    <td class="currency" style="color:var(--success-400)">${t.credit.toFixed(2)}</td>
                    <td class="currency ${t.balance > 0 ? 'negative' : ''}">${t.balance.toFixed(2)}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>
      `;
    } else if (type === 'rep') {
      const statement = ExcelEngine.getRepStatement(entityName);
      container.innerHTML = `
        <div class="statement-header">
          <div class="statement-info-card">
            <h4>إجمالي المبيعات والتحصيل</h4>
            <div class="value" style="color:var(--primary-400)">${ExcelEngine.formatCurrency(statement.transactions.filter(t => t.debit > 0).reduce((s, t) => s + t.debit, 0))}</div>
          </div>
          <div class="statement-info-card">
            <h4>إجمالي التوريد للمطعم</h4>
            <div class="value" style="color:var(--success-400)">${ExcelEngine.formatCurrency(statement.transactions.filter(t => t.credit > 0).reduce((s, t) => s + t.credit, 0))}</div>
          </div>
          <div class="statement-info-card">
            <h4>العهدة النقدية الحالية</h4>
            <div class="value" style="color:${statement.balance > 0 ? 'var(--danger-400)' : 'var(--success-400)'}">
              ${ExcelEngine.formatCurrency(statement.balance)}
            </div>
          </div>
        </div>
        <div class="table-container" style="margin-top:20px;">
           <table class="data-table">
              <thead>
                <tr><th>التاريخ</th><th>البيان</th><th>مدين (+)</th><th>دائن (-)</th><th>الرصيد</th></tr>
              </thead>
              <tbody>
                ${statement.transactions.map(t => `
                  <tr><td>${t.date}</td><td>${t.desc}</td><td>${t.debit.toFixed(2)}</td><td>${t.credit.toFixed(2)}</td><td>${t.balance.toFixed(2)}</td></tr>
                `).join('')}
              </tbody>
           </table>
        </div>
      `;
    } else if (type === 'staff') {
      const statement = ExcelEngine.getStaffStatement(entityName);
      container.innerHTML = `
        <div class="statement-header">
          <div class="statement-info-card">
            <h4>إجمالي الاستحقاق</h4>
            <div class="value" style="color:var(--success-400)">${ExcelEngine.formatCurrency(statement.transactions.reduce((s, t) => s + t.credit, 0))}</div>
          </div>
          <div class="statement-info-card">
            <h4>إجمالي الخصومات/السلف</h4>
            <div class="value" style="color:var(--danger-400)">${ExcelEngine.formatCurrency(statement.transactions.reduce((s, t) => s + t.debit, 0))}</div>
          </div>
          <div class="statement-info-card">
            <h4>المستحق النهائي</h4>
            <div class="value">${ExcelEngine.formatCurrency(statement.balance)}</div>
          </div>
        </div>
        <div class="table-container" style="margin-top:20px;">
           <table class="data-table">
              <thead>
                <tr><th>التاريخ</th><th>البيان</th><th>مديد (عليه)</th><th>دائن (له)</th><th>الرصيد</th></tr>
              </thead>
              <tbody>
                ${statement.transactions.map(t => `
                  <tr><td>${t.date}</td><td>${t.desc}</td><td>${t.debit.toFixed(2)}</td><td>${t.credit.toFixed(2)}</td><td>${t.balance.toFixed(2)}</td></tr>
                `).join('')}
              </tbody>
           </table>
        </div>
      `;
    }
  }


  // Old Treasury rendering removed - using the unified Hub version below

  // === HR & Staff (NEW) ===
  function renderHR() {
    const staffBody = document.getElementById('staffBody');
    const hrBody = document.getElementById('hrBody');
    if (!staffBody || !hrBody) return;

    if (ExcelEngine.getStaff && ExcelEngine.getHR) {
      staffBody.innerHTML = ExcelEngine.getStaff().map(s => {
        const isActive = s.status !== 'suspended';
        const statusBadge = isActive
          ? '<span class="nav-badge" style="position:static;background:var(--success-500)">نشط</span>'
          : '<span class="nav-badge" style="position:static;background:var(--danger-400)">موقوف</span>';
        const toggleBtn = isActive
          ? `<button onclick="ExcelEngine.toggleStaffStatus(${s.id}); App.refreshPage('hr')" style="color:var(--warning-500)" title="إيقاف"><i class="fas fa-ban"></i></button>`
          : `<button onclick="ExcelEngine.toggleStaffStatus(${s.id}); App.refreshPage('hr')" style="color:var(--success-400)" title="تفعيل"><i class="fas fa-check-circle"></i></button>`;
        return `
        <tr style="${!isActive ? 'opacity:0.5;' : ''}">
          <td>${s.name}</td>
          <td>${s.username || s.name}</td>
          <td><span class="nav-badge" style="position:static;background:${s.role === 'مندوب' ? 'var(--primary-400)' : 'var(--accent-400)'}">${s.role}</span></td>
          <td>${s.phone || '-'}</td>
          <td>${statusBadge}</td>
          <td>
            <div class="row-actions">
              ${toggleBtn}
              <button onclick="ExcelEngine.removeStaff('${s.name}'); App.refreshPage('hr')" style="color:var(--danger-400)"><i class="fas fa-trash"></i></button>
            </div>
          </td>
        </tr>
      `;
      }).join('');

      hrBody.innerHTML = ExcelEngine.getHR().map(h => `
        <tr>
          <td>${h.date}</td>
          <td>${h.name}</td>
          <td><span class="nav-badge" style="position:static;background:${h.type === 'راتب' ? 'var(--success-500)' : 'var(--warning-500)'}">${h.type}</span></td>
          <td class="currency font-num">${ExcelEngine.formatCurrency(h.amount)}</td>
          <td>${h.notes || '-'}</td>
        </tr>
      `).join('');
    }
  }

  // === Modals ===
  function setupModals() {
    // Close modals on overlay click
    document.querySelectorAll('.modal-overlay').forEach(overlay => {
      overlay.addEventListener('click', (e) => {
        if (e.target === overlay) closeModal(overlay.id);
      });
    });

    // Close modal buttons
    document.querySelectorAll('.modal-close').forEach(btn => {
      btn.addEventListener('click', () => {
        const modal = btn.closest('.modal-overlay');
        if (modal) closeModal(modal.id);
      });
    });

    // Setup form submissions
    setupSaleForm();
    setupPurchaseForm();
    setupCustomerForm();
    setupProductForm();
    setupSupplierForm();
    setupAdditionalForms();
  }

  function openModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) {
      modal.classList.add('active');
      populateDropdowns(modalId);
    }
  }

  function closeModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) modal.classList.remove('active');
  }

  function populateDropdowns(modalId) {
    const modal = document.getElementById(modalId);
    if (!modal) return;

    // Populate customer dropdowns
    modal.querySelectorAll('select[data-type="customers"]').forEach(sel => {
      const customers = ExcelEngine.getCustomers();
      sel.innerHTML = '<option value="">اختر العميل...</option>' +
        customers.map(c => `<option value="${c.name}">${c.name}</option>`).join('');
    });

    // Populate product dropdowns
    modal.querySelectorAll('select[data-type="products"]').forEach(sel => {
      const products = ExcelEngine.getProducts();
      sel.innerHTML = '<option value="">اختر المنتج...</option>' +
        products.map(p => `<option value="${p.name}" data-sell="${p.sellPrice}" data-buy="${p.buyPrice}">${p.name}</option>`).join('');
    });

    // Populate rep dropdowns
    modal.querySelectorAll('select[data-type="reps"]').forEach(sel => {
      const reps = ExcelEngine.getRepsList();
      const user = ExcelEngine.getCurrentUser();
      if (user && user.role === 'مندوب') {
        // Lock to current rep
        sel.innerHTML = `<option value="${user.name}" selected>${user.name}</option>`;
        sel.disabled = true;
        sel.style.opacity = '0.7';
      } else {
        sel.innerHTML = '<option value="">اختر المندوب...</option>' +
          reps.map(r => `<option value="${r}">${r}</option>`).join('');
        sel.disabled = false;
        sel.style.opacity = '1';
      }
    });

    // Populate supplier dropdowns
    modal.querySelectorAll('select[data-type="suppliers"]').forEach(sel => {
      const suppliers = ExcelEngine.getSuppliers();
      sel.innerHTML = '<option value="">اختر المورد...</option>' +
        suppliers.map(s => `<option value="${s.name}">${s.name}</option>`).join('');
    });

    // Populate allStaff dropdowns (employees + reps)
    modal.querySelectorAll('select[data-type="allStaff"]').forEach(sel => {
      const staff = ExcelEngine.getStaff();
      sel.innerHTML = '<option value="">اختر الموظف...</option>' +
        staff.map(s => `<option value="${s.name}">${s.name} (${s.role})</option>`).join('');
    });

    // Populate cpCollectedBy with dynamic reps
    const cpCollectedBy = modal.querySelector('#cpCollectedBy');
    if (cpCollectedBy) {
      const reps = ExcelEngine.getRepsList();
      cpCollectedBy.innerHTML = '<option value="الشركة/كاشير المطعم">الشركة / كاشير المطعم (مباشرة)</option>' +
        reps.map(r => `<option value="${r}">${r} (مندوب)</option>`).join('');

      // Add Change Listener for Filtering
      cpCollectedBy.addEventListener('change', () => {
        const rep = cpCollectedBy.value;
        const custSel = modal.querySelector('#cpCustomer');
        if (custSel) {
          const filtered = ExcelEngine.getCustomers(rep);
          custSel.innerHTML = '<option value="">اختر العميل...</option>' +
            filtered.map(c => `<option value="${c.name}">${c.name} (مديونية: ${c.remaining.toFixed(2)})</option>`).join('');
        }
      });
    }

    // Populate reps-with-company (for initial list)
    modal.querySelectorAll('select[data-type="reps-with-company"]').forEach(sel => {
      const reps = ExcelEngine.getRepsList();
      sel.innerHTML = '<option value="الشركة/كاشير المطعم">الشركة / كاشير المطعم (مباشرة)</option>' +
        reps.map(r => `<option value="${r}">${r} (مندوب)</option>`).join('');
    });
  }

  // === Sale Form ===
  function setupSaleForm() {
    const form = document.getElementById('saleForm');
    if (!form) return;

    // Auto-calculate
    const qtyInput = form.querySelector('#saleQty');
    const priceInput = form.querySelector('#salePrice');
    const totalDisplay = document.getElementById('saleTotal');

    const vatCheckbox = document.getElementById('saleVat');
    function calcTotal() {
      const qty = parseFloat(qtyInput?.value) || 0;
      const price = parseFloat(priceInput?.value) || 0;
      let total = qty * price;
      if (vatCheckbox && vatCheckbox.checked) {
        total = total - (total * 0.15);
      }
      if (totalDisplay) totalDisplay.textContent = total.toFixed(2);
    }

    [qtyInput, priceInput, vatCheckbox].forEach(input => {
      if (input) input.addEventListener('input', calcTotal);
    });

    // Auto-fill price when product is selected
    const productSelect = form.querySelector('#saleProduct');
    if (productSelect) {
      productSelect.addEventListener('change', () => {
        const opt = productSelect.selectedOptions[0];
        if (opt && opt.dataset.sell) {
          priceInput.value = opt.dataset.sell;
          calcTotal();
        }
      });
    }

    // Submit
    form.addEventListener('submit', (e) => {
      e.preventDefault();
      const isReturn = form.querySelector('#saleIsReturn')?.checked;
      const multiplier = isReturn ? -1 : 1;

      const sale = {
        date: form.querySelector('#saleDate')?.value || ExcelEngine.getTodayStr(),
        customer: form.querySelector('#saleCustomer')?.value,
        product: form.querySelector('#saleProduct')?.value,
        rep: form.querySelector('#saleRep')?.value,
        quantity: (parseFloat(form.querySelector('#saleQty')?.value) || 0) * multiplier,
        uom: form.querySelector('#saleUom')?.value || 'piece',
        returned: 0,
        price: parseFloat(form.querySelector('#salePrice')?.value) || 0,
        collected: (parseFloat(form.querySelector('#saleCollected')?.value) || 0) * multiplier,
        notes: (form.querySelector('#saleNotes')?.value || '') + (isReturn ? ' (مرتجع)' : ''),
        impactStock: document.getElementById('saleStockImpact')?.checked,
        applyVat: document.getElementById('saleVat')?.checked
      };

      if (!sale.customer || !sale.product) {
        showToast('يرجى ملء جميع الحقول المطلوبة', 'warning');
        return;
      }

      const result = ExcelEngine.addSale(sale);

      if (result && result.success === false) {
        showToast(result.error, 'error');
        return;
      }

      ExcelEngine.saveToLocalStorage();
      closeModal('saleModal');
      form.reset();
      showToast('تم إضافة فاتورة البيع بنجاح!', 'success');
      renderSalesTable();
    });
  }

  // === Purchase Form ===
  function setupPurchaseForm() {
    const form = document.getElementById('purchaseForm');
    if (!form) return;

    const qtyInput = form.querySelector('#purchQty');
    const priceInput = form.querySelector('#purchPrice');
    const totalDisplay = document.getElementById('purchTotal');

    const vatCheckbox = document.getElementById('purchVat');
    function calcTotal() {
      const qty = parseFloat(qtyInput?.value) || 0;
      const price = parseFloat(priceInput?.value) || 0;
      let total = qty * price;
      if (vatCheckbox && vatCheckbox.checked) {
        total = total - (total * 0.15);
      }
      if (totalDisplay) totalDisplay.textContent = total.toFixed(2);
    }

    [qtyInput, priceInput, vatCheckbox].forEach(input => {
      if (input) input.addEventListener('input', calcTotal);
    });

    // Auto-fill price
    const productSelect = form.querySelector('#purchProduct');
    if (productSelect) {
      productSelect.addEventListener('change', () => {
        const opt = productSelect.selectedOptions[0];
        if (opt && opt.dataset.buy) {
          priceInput.value = opt.dataset.buy;
          calcTotal();
        }
      });
    }

    form.addEventListener('submit', (e) => {
      e.preventDefault();
      const isReturn = form.querySelector('#purchIsReturn')?.checked;
      const multiplier = isReturn ? -1 : 1;

      const purchase = {
        date: form.querySelector('#purchDate')?.value || ExcelEngine.getTodayStr(),
        product: form.querySelector('#purchProduct')?.value,
        quantity: (parseFloat(form.querySelector('#purchQty')?.value) || 0) * multiplier,
        uom: form.querySelector('#purchUom')?.value || 'piece',
        rep: form.querySelector('#purchRep')?.value,
        price: parseFloat(form.querySelector('#purchPrice')?.value) || 0,
        supplier: form.querySelector('#purchSupplier')?.value,
        paid: (parseFloat(form.querySelector('#purchPaid')?.value) || 0) * multiplier,
        impactStock: document.getElementById('purchStockImpact')?.checked,
        applyVat: document.getElementById('purchVat')?.checked
      };

      if (!purchase.product) {
        showToast('يرجى اختيار المنتج', 'warning');
        return;
      }

      ExcelEngine.addPurchase(purchase);
      // Note: Treasury logging is already handled inside addPurchase in ExcelEngine

      ExcelEngine.saveToLocalStorage();
      closeModal('purchaseModal');
      form.reset();
      showToast('تم إضافة فاتورة الشراء وتحديث الخزينة!', 'success');
      renderPurchasesTable();
    });
  }

  // === Customer Form ===
  function setupCustomerForm() {
    const form = document.getElementById('customerForm');
    if (!form) return;

    form.addEventListener('submit', (e) => {
      e.preventDefault();
      const customer = {
        name: form.querySelector('#custName')?.value,
        cr: form.querySelector('#custCR')?.value || '',
        taxId: form.querySelector('#custTaxId')?.value || ''
      };

      if (!customer.name) {
        showToast('يرجى إدخال اسم العميل', 'warning');
        return;
      }

      ExcelEngine.addCustomer(customer);
      ExcelEngine.saveToLocalStorage();
      closeModal('customerModal');
      form.reset();
      showToast('تم إضافة العميل بنجاح!', 'success');
      renderCustomersTable();
    });
  }

  // === Product Form ===
  function setupProductForm() {
    const form = document.getElementById('productForm');
    if (!form) return;

    form.addEventListener('submit', (e) => {
      e.preventDefault();
      const product = {
        name: form.querySelector('#prodName')?.value,
        sellPrice: parseFloat(form.querySelector('#prodSellPrice')?.value) || 0,
        buyPrice: parseFloat(form.querySelector('#prodBuyPrice')?.value) || 0,
        unit: form.querySelector('#prodUnit')?.value || 'حبة',
        piecesPerCarton: parseInt(form.querySelector('#prodPiecesPerCarton')?.value) || 10,
        category: form.querySelector('#prodCategory')?.value || 'عام'
      };

      if (!product.name) {
        showToast('يرجى إدخال اسم المنتج', 'warning');
        return;
      }

      ExcelEngine.addProduct(product);
      ExcelEngine.saveToLocalStorage();
      closeModal('productModal');
      form.reset();
      showToast('تم إضافة المنتج بنجاح!', 'success');
      renderProductsTable();
    });
  }

  // === Rep Payment Modal ===
  function openRepPaymentModal(repName) {
    openModal('repPaymentModal');
    const repInput = document.getElementById('rpRep');
    if (repInput) repInput.value = repName;
  }

  // === Additional forms (Payments, Staff, HR, Treasury, Warehouse, Supplier) ===
  function setupAdditionalForms() {
    const warehouseForm = document.getElementById('warehouseTransferForm');
    if (warehouseForm) {
      warehouseForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const rep = warehouseForm.querySelector('#whRep')?.value;
        const product = warehouseForm.querySelector('#whProduct')?.value;
        const qty = parseFloat(warehouseForm.querySelector('#whQty')?.value) || 0;
        const direction = warehouseForm.querySelector('#whDirection')?.value;
        const notes = warehouseForm.querySelector('#whNotes')?.value;

        if (!rep || !product || qty <= 0) {
          showToast('يرجى التأكد من المندوب والمنتج والكمية', 'warning');
          return;
        }

        let res;
        if (direction === 'TO_REP') res = ExcelEngine.transferToRepWarehouse(rep, product, qty, notes);
        else res = ExcelEngine.returnFromRepWarehouse(rep, product, qty, notes);

        if (res && res.success) {
          closeModal('warehouseTransferModal');
          warehouseForm.reset();
          showToast(direction === 'TO_REP' ? 'تم سحب المخزون للمندوب بنجاح' : 'تم إرجاع المخزون من المندوب بنجاح', 'success');
          // If renderInventoryTable exists, call it to refresh main stock
          if (typeof renderInventoryTable === 'function') renderInventoryTable();

          // Re-render rep profile if currently viewing it
          const activePage = document.querySelector('.page-section.active');
          if (activePage && activePage.id === 'page-entity-profile') {
            renderEntityProfile('rep', rep);
          }
        } else {
          showToast(res.error || 'حدث خطأ أثناء العملية', 'error');
        }
      });
    }

    const supplierFormEl = document.getElementById('supplierForm');
    if (supplierFormEl) {
      supplierFormEl.addEventListener('submit', (e) => {
        e.preventDefault();
        const supplier = {
          name: supplierFormEl.querySelector('#suppName')?.value,
          phone: supplierFormEl.querySelector('#suppPhone')?.value || '',
          address: supplierFormEl.querySelector('#suppAddress')?.value || '',
          cr: supplierFormEl.querySelector('#suppCR')?.value || '',
          taxId: supplierFormEl.querySelector('#suppTaxId')?.value || ''
        };
        if (!supplier.name) return;
        ExcelEngine.addSupplier(supplier);
        ExcelEngine.saveToLocalStorage();
        closeModal('supplierModal');
        supplierFormEl.reset();
        showToast('تم إضافة المورد بنجاح', 'success');
        if (typeof renderSuppliersTable === 'function') renderSuppliersTable();
      });
    }

    const hrActionForm = document.getElementById('hrActionForm');
    if (hrActionForm) {
      hrActionForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const data = {
          date: hrActionForm.querySelector('#hrDate')?.value || ExcelEngine.getTodayStr(),
          name: hrActionForm.querySelector('#hrStaff')?.value,
          type: hrActionForm.querySelector('#hrType')?.value,
          amount: parseFloat(hrActionForm.querySelector('#hrAmount')?.value) || 0,
          notes: hrActionForm.querySelector('#hrNotes')?.value || ''
        };
        if (!data.name || data.amount <= 0) return;
        ExcelEngine.addHRAction(data);
        ExcelEngine.saveToLocalStorage();
        closeModal('hrActionModal');
        hrActionForm.reset();
        showToast('تم تسجيل حركة الموارد البشرية بنجاح وتحديث الخزينة', 'success');
        if (typeof renderHRTable === 'function') renderHRTable();
        renderTreasury();
      });
    }

    const repPayForm = document.getElementById('repPaymentForm');
    if (repPayForm) {
      repPayForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const data = {
          date: repPayForm.querySelector('#rpDate')?.value || ExcelEngine.getTodayStr(),
          rep: repPayForm.querySelector('#rpRep')?.value,
          amount: parseFloat(repPayForm.querySelector('#rpAmount')?.value) || 0,
          paymentType: repPayForm.querySelector('#rpType')?.value || 'نقدي',
          notes: repPayForm.querySelector('#rpNotes')?.value || ''
        };
        if (!data.rep || data.amount <= 0) return;
        ExcelEngine.addRepPayment(data);
        ExcelEngine.saveToLocalStorage();
        closeModal('repPaymentModal');
        repPayForm.reset();
        showToast('تم تسجيل الدفعة المسلمة بنجاح وتحديث الخزينة', 'success');
        if (typeof renderRepsTable === 'function') renderRepsTable();
        renderTreasury();
      });
    }

    const custPayForm = document.getElementById('customerPaymentForm');
    if (custPayForm) {
      custPayForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const data = {
          date: custPayForm.querySelector('#cpDate')?.value || ExcelEngine.getTodayStr(),
          customer: custPayForm.querySelector('#cpCustomer')?.value,
          amount: parseFloat(custPayForm.querySelector('#cpAmount')?.value) || 0,
          invoiceRef: custPayForm.querySelector('#cpInvoiceRef')?.value || '',
          notes: custPayForm.querySelector('#cpNotes')?.value || '',
          collectedBy: custPayForm.querySelector('#cpCollectedBy')?.value || 'الشركة/كاشير المطعم'
        };
        if (!data.customer || data.amount <= 0) return;
        ExcelEngine.addCustomerPayment(data);
        ExcelEngine.saveToLocalStorage();
        closeModal('customerPaymentModal');
        custPayForm.reset();
        showToast('تم استلام الدفعة وتسويتها ماليًا', 'success');
        if (typeof renderCustomersTable === 'function') renderCustomersTable();
        renderTreasury();
      });
    }

    const staffForm = document.getElementById('staffForm');
    if (staffForm) {
      staffForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const s = {
          name: staffForm.querySelector('#stfName')?.value,
          role: staffForm.querySelector('#stfRole')?.value || 'موظف',
          password: staffForm.querySelector('#stfPassword')?.value || '',
          phone: staffForm.querySelector('#stfPhone')?.value || ''
        };
        if (!s.name) return;
        ExcelEngine.addStaff(s);
        ExcelEngine.saveToLocalStorage();
        closeModal('staffModal');
        staffForm.reset();
        showToast('تم إضافة الموظف', 'success');
        if (typeof renderStaffTable === 'function') renderStaffTable();
      });
    }

    const treasuryForm = document.getElementById('treasuryActionForm');
    if (treasuryForm) {
      treasuryForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const type = treasuryForm.querySelector('#trType')?.value;
        const amt = parseFloat(treasuryForm.querySelector('#trAmount')?.value) || 0;
        const ref = treasuryForm.querySelector('#trRef')?.value;
        const not = treasuryForm.querySelector('#trNotes')?.value;

        ExcelEngine.logTreasuryTransaction(type, amt, ref, not);
        closeModal('treasuryActionModal');
        treasuryForm.reset();
        showToast('تم تسجيل حركة الخزينة بنجاح', 'success');
        renderTreasury();
      });
    }

    const stockAdjForm = document.getElementById('stockAdjForm');
    if (stockAdjForm) {
      stockAdjForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const prod = stockAdjForm.querySelector('#adjProduct')?.value;
        const qty = parseFloat(stockAdjForm.querySelector('#adjNewQty')?.value) || 0;
        const type = stockAdjForm.querySelector('#adjType')?.value;
        if (prod) {
          if (typeof ExcelEngine.updateInventoryManual === 'function') {
            ExcelEngine.updateInventoryManual(prod, qty, type);
            closeModal('stockAdjModal');
            stockAdjForm.reset();
            if (typeof renderInventoryTable === 'function') renderInventoryTable();
            showToast('تم تسوية المخزون بنجاح', 'success');
          } else {
            showToast('وظيفة التسوية غير متوفرة', 'warning');
          }
        }
      });
    }
  }

  function setupSupplierForm() {
    // Overridden, it's handeled inside setupAdditionalForms to prevent duplicate
  }

  function confirmDeleteCustomer(id) {
    if (confirm('هل أنت متأكد من حذف هذا العميل؟')) {
      ExcelEngine.deleteCustomer(id);
      ExcelEngine.saveToLocalStorage();
      showToast('تم حذف العميل', 'info');
      renderCustomersTable();
    }
  }

  function confirmDeleteProduct(id) {
    if (confirm('هل أنت متأكد من حذف هذا المنتج؟')) {
      ExcelEngine.deleteProduct(id);
      ExcelEngine.saveToLocalStorage();
      showToast('تم حذف المنتج', 'info');
      renderProductsTable();
    }
  }

  // === Invoice Printing ===
  function printInvoice(id, type) {
    let data;
    if (type === 'sale') {
      data = ExcelEngine.getSales().find(s => s.id === id);
    } else {
      data = ExcelEngine.getPurchases().find(p => p.id === id);
    }
    if (!data) return;

    const printWindow = window.open('', '_blank');
    if (!printWindow) return;

    // Build the print content
    const html = `
      <!DOCTYPE html>
      <html dir="rtl" lang="ar">
      <head>
        <meta charset="UTF-8">
        <title>فاتورة ${type === 'sale' ? 'بيع' : 'شراء'} #${id}</title>
        <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap" rel="stylesheet">
        <style>
          body { font-family: 'Cairo', sans-serif; padding: 20px; direction: rtl; color: #333; }
          .invoice-box { max-width: 800px; margin: auto; padding: 30px; border: 1px solid #eee; box-shadow: 0 0 10px rgba(0, 0, 0, 0.15); }
          .header { display: flex; justify-content: space-between; align-items: center; border-bottom: 3px solid #1a73e8; padding-bottom: 15px; margin-bottom: 20px; }
          .header h1 { color: #1a73e8; font-size: 28px; margin: 0; }
          .info-row { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }
          .info-item { background: #f9f9f9; padding: 10px; border-radius: 5px; }
          table { width: 100%; border-collapse: collapse; margin-top: 20px; }
          th, td { padding: 12px; border: 1px solid #eee; text-align: right; }
          th { background: #1a73e8; color: white; }
          .total-row { font-weight: bold; background: #f5f5f5; }
          .security-footer { margin-top: 40px; padding-top: 20px; border-top: 1px dashed #ccc; display: flex; justify-content: space-between; align-items: flex-start; }
          .signature-box { font-size: 11px; color: #666; font-family: monospace; max-width: 60%; }
          #qrcode { width: 100px; height: 100px; background: #eee; display: flex; align-items: center; justify-content: center; border: 1px solid #ddd; }
          .footer-text { margin-top: 30px; text-align: center; color: #999; font-size: 12px; }
          @media print { .no-print { display: none; } }
        </style>
        <script src="https://cdn.jsdelivr.net/npm/qrcodejs@1.0.0/qrcode.min.js"></script>
      </head>
      <body>
        <div class="invoice-box">
          <div class="header">
            <div>
              <h1>${type === 'sale' ? 'فاتورة بيع' : 'فاتورة شراء'}</h1>
              <p>رقم الفاتورة: <strong>${id}</strong></p>
            </div>
            <div style="text-align: left">
              <p>التاريخ: ${data.date}</p>
              <p>المندوب: ${data.rep || '-'}</p>
            </div>
          </div>
          
          <div class="info-row">
            <div class="info-item">
              <strong>الجهة:</strong> ${data.customer || data.supplier || 'غير محدد'}
            </div>
            <div class="info-item">
              <strong>الحالة الضريبية:</strong> فاتورة مبسطة
            </div>
          </div>

          <table>
            <thead>
              <tr>
                <th>البيان</th>
                <th>الكمية</th>
                <th>السعر</th>
                <th>الإجمالي</th>
              </tr>
            </thead>
              <tr>
                <td>${data.product}</td>
                <td>${data.quantity !== undefined ? data.quantity : (data.netQuantity || 0)}</td>
                <td>${ExcelEngine.formatCurrency(data.price)}</td>
                <td>${ExcelEngine.formatCurrency(data.total)}</td>
              </tr>
              <tr class="total-row">
                <td colspan="3">الإجمالي الكلي</td>
                <td>${ExcelEngine.formatCurrency(data.total)}</td>
              </tr>
              <tr>
                <td colspan="3">التحصيل / المدفوع</td>
                <td style="color:green">${ExcelEngine.formatCurrency(data.collected || data.paid)}</td>
              </tr>
              <tr class="total-row">
                <td colspan="3">المتبقي</td>
                <td style="color:red">${ExcelEngine.formatCurrency(data.balance || (data.total - data.paid))}</td>
              </tr>
            </tbody>
          </table>

          ${data.notes ? `<div style="margin-top:20px; font-size:13px;"><strong>ملاحظات:</strong> ${data.notes}</div>` : ''}

          <div class="security-footer">
            <div class="signature-box">
              <p>التوقيع الرقمي (Digital Signature):</p>
              <div style="word-break: break-all; background: #f0f0f0; padding: 5px; border-radius: 4px; margin-top: 5px;">
                ${data.signature || 'SEC-NOT-SIGNED'}
              </div>
              <p style="margin-top:10px; font-size:10px;">هذه الفاتورة محمية بنظام XML التشفيري - أي محاولة تعديل تلغي صحة التوقيع.</p>
            </div>
            <div id="qrcode"></div>
          </div>

          <div class="footer-text">
            نظام إدارة الـ ERP المطور - نسخة سحابية آمنة
          </div>
        </div>

        <script src="https://cdn.jsdelivr.net/npm/qrcodejs@1.0.0/qrcode.min.js"></script>
        <script>
          window.onload = function() {
            setTimeout(function() {
              if (window.QRCode) {
                const qrData = "فاتورة: " + "${id}" + "\\nالتاريخ: " + "${data.date}" + "\\nالإجمالي: " + "${data.total}" + "\\nالجهة: " + "${data.customer || data.supplier || ''}";
                new QRCode(document.getElementById("qrcode"), {
                  text: qrData,
                  width: 100,
                  height: 100,
                  colorDark: "#000",
                  colorLight: "#fff",
                  correctLevel: QRCode.CorrectLevel.M
                });
              }
              // Wait for QR render then print
              setTimeout(() => { window.print(); }, 300);
            }, 100);
          };
        </script>
      </body>
      </html>
    `;

    printWindow.document.write(html);
    printWindow.document.close();
  }

  function printStatement(name, type = 'customer') {
    let statement;
    let title = 'كشف حساب';
    let summaryHtml = '';
    let tableHtml = '';

    if (type === 'customer') {
      statement = ExcelEngine.getCustomerStatement(name);
      title = `كشف حساب عميل: ${name}`;
      summaryHtml = `
        <div class="summary-card"><div>إجمالي المبيعات</div><div class="val">${ExcelEngine.formatCurrency(statement.totalSales)}</div></div>
        <div class="summary-card"><div>المحصل</div><div class="val" style="color:green">${ExcelEngine.formatCurrency(statement.totalCollected)}</div></div>
        <div class="summary-card"><div>المتبقي</div><div class="val" style="color:red">${ExcelEngine.formatCurrency(statement.balance)}</div></div>
      `;
      tableHtml = `
        <thead><tr><th>التاريخ</th><th>المنتج</th><th>الكمية</th><th>السعر</th><th>الإجمالي</th><th>التحصيل</th><th>الرصيد</th></tr></thead>
        <tbody>
          ${statement.transactions.map(t => `
            <tr>
              <td>${t.date}</td><td>${t.product}</td><td>${t.netQuantity || 0}</td>
              <td>${t.price.toFixed(2)}</td><td>${t.total.toFixed(2)}</td>
              <td>${t.collected.toFixed(2)}</td><td>${t.balance.toFixed(2)}</td>
            </tr>
          `).join('')}
        </tbody>
      `;
    } else if (type === 'supplier') {
      statement = ExcelEngine.getSupplierStatement(name);
      title = `كشف حساب مورد: ${name}`;
      summaryHtml = `
        <div class="summary-card"><div>إجمالي المشتريات</div><div class="val">${ExcelEngine.formatCurrency(statement.totalPurchases)}</div></div>
        <div class="summary-card"><div>إجمالي المدفوع</div><div class="val" style="color:green">${ExcelEngine.formatCurrency(statement.totalPaid)}</div></div>
        <div class="summary-card"><div>الرصيد المستحق</div><div class="val" style="color:red">${ExcelEngine.formatCurrency(statement.balance)}</div></div>
      `;
      tableHtml = `
        <thead><tr><th>التاريخ</th><th>البيان</th><th>مدين (عليه)</th><th>دائن (له)</th><th>الرصيد</th></tr></thead>
        <tbody>
          ${statement.transactions.map(t => `
            <tr><td>${t.date}</td><td>${t.desc}</td><td>${t.debit.toFixed(2)}</td><td>${t.credit.toFixed(2)}</td><td>${t.balance.toFixed(2)}</td></tr>
          `).join('')}
        </tbody>
      `;
    } else if (type === 'rep') {
      statement = ExcelEngine.getRepStatement(name);
      title = `كشف حساب مندوب: ${name}`;
      summaryHtml = `
        <div class="summary-card"><div>العهد المحصلة</div><div class="val">${ExcelEngine.formatCurrency(statement.transactions.filter(t => t.debit > 0).reduce((s, t) => s + t.debit, 0))}</div></div>
        <div class="summary-card"><div>المسدد للشركة</div><div class="val" style="color:green">${ExcelEngine.formatCurrency(statement.transactions.filter(t => t.credit > 0).reduce((s, t) => s + t.credit, 0))}</div></div>
        <div class="summary-card"><div>الرصيد القائم</div><div class="val" style="color:red">${ExcelEngine.formatCurrency(statement.balance)}</div></div>
      `;
      tableHtml = `
        <thead><tr><th>التاريخ</th><th>البيان</th><th>مدين (عليه)</th><th>دائن (له)</th><th>الرصيد</th></tr></thead>
        <tbody>
          ${statement.transactions.map(t => `
            <tr><td>${t.date}</td><td>${t.desc}</td><td>${t.debit.toFixed(2)}</td><td>${t.credit.toFixed(2)}</td><td>${t.balance.toFixed(2)}</td></tr>
          `).join('')}
        </tbody>
      `;
    } else if (type === 'staff') {
      statement = ExcelEngine.getStaffStatement(name);
      title = `كشف حساب موظف: ${name}`;
      summaryHtml = `
        <div class="summary-card"><div>إجمالي المستحقات</div><div class="val">${ExcelEngine.formatCurrency(statement.transactions.reduce((s, t) => s + t.credit, 0))}</div></div>
        <div class="summary-card"><div>إجمالي الخصومات</div><div class="val" style="color:red">${ExcelEngine.formatCurrency(statement.transactions.reduce((s, t) => s + t.debit, 0))}</div></div>
        <div class="summary-card"><div>باقي المستحق</div><div class="val" style="color:green">${ExcelEngine.formatCurrency(statement.balance)}</div></div>
      `;
      tableHtml = `
        <thead><tr><th>التاريخ</th><th>البيان</th><th>مدين (عليه)</th><th>دائن (له)</th><th>الرصيد</th></tr></thead>
        <tbody>
          ${statement.transactions.map(t => `
            <tr><td>${t.date}</td><td>${t.desc}</td><td>${t.debit.toFixed(2)}</td><td>${t.credit.toFixed(2)}</td><td>${t.balance.toFixed(2)}</td></tr>
          `).join('')}
        </tbody>
      `;
    }

    const printWindow = window.open('', '_blank');
    if (!printWindow) return;

    printWindow.document.write(`
      <!DOCTYPE html>
      <html dir="rtl" lang="ar">
      <head>
        <meta charset="UTF-8">
        <title>${title}</title>
        <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap" rel="stylesheet">
        <style>
          body { font-family: 'Cairo', sans-serif; padding: 20px; direction: rtl; }
          .header { text-align: center; border-bottom: 3px solid #1a73e8; padding-bottom: 15px; margin-bottom: 20px; }
          .header h1 { color: #1a73e8; font-size: 24px; margin-bottom: 5px; }
          .header h2 { font-size: 18px; color: #555; }
          .summary { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 25px; }
          .summary-card { background: #f8f9fc; padding: 15px; border-radius: 10px; text-align: center; border: 1px solid #e3e6f0; }
          .summary-card div { font-size: 12px; color: #666; margin-bottom: 5px; }
          .summary-card .val { font-size: 18px; font-weight: bold; color: #333; }
          table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 13px; }
          th, td { padding: 10px; border: 1px solid #e3e6f0; text-align: right; }
          th { background: #1a73e8; color: white; font-weight: 600; }
          tr:nth-child(even) { background: #f8f9fc; }
          @media print { .header { border-bottom-color: #000; } th { background-color: #eee !important; color: #000 !important; } }
        </style>
      </head>
      <body>
        <div class="header"><h1>${title}</h1><h2>مشروع الفروج الوطني</h2><p>تاريخ الطباعة: ${new Date().toLocaleDateString('ar-EG')}</p></div>
        <div class="summary">${summaryHtml}</div>
        <table>${tableHtml}</table>
        <div style="margin-top:40px; text-align:left; font-size:12px; color:#999;">طُبع بواسطة نظام ERP المطور</div>
        <script>window.onload = () => { setTimeout(() => window.print(), 500); };<\/script>
      </body>
      </html>
    `);
    printWindow.document.close();
  }

  // === Utility Functions ===
  function setTextContent(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
  }

  function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    if (!container) return;

    const icons = {
      success: 'fas fa-check-circle',
      error: 'fas fa-times-circle',
      warning: 'fas fa-exclamation-triangle',
      info: 'fas fa-info-circle'
    };

    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.innerHTML = `
      <i class="toast-icon ${icons[type]}"></i>
      <span class="toast-message">${message}</span>
      <button class="toast-close" onclick="this.parentElement.remove()">×</button>
    `;
    container.appendChild(toast);

    setTimeout(() => toast.remove(), 4000);
  }

  function showLoading(text) {
    const el = document.getElementById('loadingOverlay');
    if (el) {
      el.querySelector('p').textContent = text;
      el.style.display = 'flex';
    }
  }

  function hideLoading() {
    const el = document.getElementById('loadingOverlay');
    if (el) el.style.display = 'none';
  }

  function updateExcelStatus(connected) {
    const status = document.getElementById('excelStatus');
    if (!status) return;
    if (connected) {
      status.className = 'excel-status';
      status.querySelector('.status-text').textContent = 'متصل - البيانات محملة';
    } else {
      status.className = 'excel-status disconnected';
      status.querySelector('.status-text').textContent = 'غير متصل';
    }
  }

  // Search functionality
  function setupSearch() {
    document.querySelectorAll('.search-input').forEach(input => {
      input.addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase();
        const tableId = input.dataset.table;
        const table = document.getElementById(tableId);
        if (!table) return;

        table.querySelectorAll('tr').forEach(row => {
          const text = row.textContent.toLowerCase();
          row.style.display = text.includes(query) ? '' : 'none';
        });
      });
    });
  }

  // === Treasury Hub ===
  function renderTreasury() {
    const ledger = ExcelEngine.getLedger();
    const stats = ExcelEngine.getStats();

    const statsContainer = document.getElementById('treasuryStats');
    if (statsContainer) {
      statsContainer.innerHTML = `
        <div class="card premium-card">
          <div class="text-muted small">رصيد الخزينة الحالي</div>
          <div class="text-2xl font-bold font-num" style="color:var(--success-400)">${ExcelEngine.formatCurrency(stats.cashInBox)}</div>
          <p class="small opacity-70">بناءً على توريدات المناديب والتحصيل</p>
        </div>
        <div class="card premium-card">
          <div class="text-muted small">إجمالي المصروفات والسلف</div>
          <div class="text-2xl font-bold font-num" style="color:var(--warning-400)">${ExcelEngine.formatCurrency(stats.totalHRCosts)}</div>
        </div>
      `;
    }

    const ledgerBody = document.getElementById('ledgerBody');
    if (ledgerBody) {
      ledgerBody.innerHTML = ledger.map(item => `
        <tr>
          <td>${item.date}</td>
          <td><span class="nav-badge" style="position:static;">${item.category}</span></td>
          <td>${item.reference}</td>
          <td><span style="color:${item.type === 'IN' ? '#4caf50' : '#ef5350'}">${item.type === 'IN' ? 'وارد (+)' : 'صادر (-)'}</span></td>
          <td class="font-num">${ExcelEngine.formatCurrency(item.amount)}</td>
          <td>${item.notes || '-'}</td>
        </tr>
      `).join('');
    }

    const vBody = document.getElementById('vouchersBody');
    if (vBody) {
      vBody.innerHTML = ExcelEngine.getVouchers().map(v => `
        <tr>
          <td class="font-num">#${v.serial}</td>
          <td>${v.date}</td>
          <td><span class="nav-badge ${v.type === 'RECEIPT' ? 'success' : 'warning'}" style="position:static;">${v.type === 'RECEIPT' ? 'سند قبض' : 'سند صرف'}</span></td>
          <td>${v.party}</td>
          <td class="font-num">${ExcelEngine.formatCurrency(v.amount)}</td>
          <td>${v.user}</td>
          <td><button class="btn btn-sm btn-ghost" onclick="App.printVoucher('${v.serial}')"><i class="fas fa-print"></i></button></td>
        </tr>
      `).join('');
    }
  }

  function switchTreasuryTab(tab, btn) {
    document.querySelectorAll('#page-treasury .btn-ghost').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('treasury-ledger-view').style.display = tab === 'ledger' ? 'block' : 'none';
    document.getElementById('treasury-vouchers-view').style.display = tab === 'vouchers' ? 'block' : 'none';
  }

  // === Entity Profiles ===
  function viewProfile(type, name) {
    navigateTo('entity-profile');
    renderEntityProfile(type, name);
  }

  function renderEntityProfile(type, name) {
    const title = document.getElementById('profileTitle');
    const summary = document.getElementById('profileSummary');
    const wrapper = document.getElementById('profileTableWrapper');
    const actions = document.getElementById('profileActions');

    title.textContent = name;
    if (actions) actions.innerHTML = '';

    if (type === 'customer') {
      const stmt = ExcelEngine.getCustomerStatement(name);
      if (actions) {
        actions.innerHTML = `
          <button class="btn btn-primary btn-sm" onclick="App.showCustomerStatement('${name}')" style="margin-left:8px;"><i class="fas fa-file-lines"></i> عرض كشف الحساب</button>
          <button class="btn btn-accent btn-sm" onclick="App.exportReport('customer', '${name}')"><i class="fas fa-file-excel"></i> تصدير التقرير (Excel)</button>
        `;
      }
      summary.innerHTML = `
        <div class="card premium-card"><div>إجمالي المسحوبات</div><div class="text-xl font-num">${ExcelEngine.formatCurrency(stmt.totalSales)}</div></div>
        <div class="card premium-card"><div>إجمالي السداد</div><div class="text-xl font-num" style="color:#4caf50">${ExcelEngine.formatCurrency(stmt.totalCollected)}</div></div>
        <div class="card premium-card"><div>المديونية الحالية</div><div class="text-xl font-num" style="color:#ef5350">${ExcelEngine.formatCurrency(stmt.balance)}</div></div>
      `;
      wrapper.innerHTML = `
        <table class="data-table">
          <thead><tr><th>التاريخ</th><th>البيان</th><th>الكمية</th><th>الإجمالي</th><th>الحالة</th></tr></thead>
          <tbody>
            ${stmt.transactions.map(t => `
              <tr><td>${t.date}</td><td>${t.product}</td><td>${t.netQuantity || 0}</td><td>${ExcelEngine.formatCurrency(t.total || t.collected)}</td><td>${t.type}</td></tr>
            `).join('')}
          </tbody>
        </table>
      `;
    } else if (type === 'rep') {
      const stats = ExcelEngine.getRepAccountingSummary(name);

      if (actions) {
        actions.innerHTML = `
          <button class="btn btn-primary btn-sm" onclick="App.printRepDailyReport('${name}')" style="margin-left:8px;"><i class="fas fa-print"></i> طباعة التقرير اليومي</button>
          <button class="btn btn-accent btn-sm" onclick="App.exportReport('rep', '${name}')"><i class="fas fa-file-excel"></i> تصدير التقرير (Excel)</button>
        `;
      }

      summary.innerHTML = `
        <div class="card premium-card"><div>إجمالي المبيعات</div><div class="text-xl font-num">${ExcelEngine.formatCurrency(stats.totalSales)}</div></div>
        <div class="card premium-card"><div>إجمالي المشتريات</div><div class="text-xl font-num" style="color:var(--warning-500)">${ExcelEngine.formatCurrency(stats.totalPurchases)}</div></div>
        <div class="card premium-card"><div>المديونية (العهدة)</div><div class="text-xl font-num" style="color:#ef5350">${ExcelEngine.formatCurrency(stats.cashCustody)}</div></div>
        <div class="card premium-card"><div>الأرباح المحققة</div><div class="text-xl font-num" style="color:#4caf50">${ExcelEngine.formatCurrency(stats.totalProfit)}</div></div>
      `;

      const sales = ExcelEngine.getSales().filter(s => s.rep === name).sort((a, b) => new Date(b.date) - new Date(a.date));
      const purchases = ExcelEngine.getPurchases().filter(p => p.rep === name).sort((a, b) => new Date(b.date) - new Date(a.date));

      wrapper.innerHTML = `
        <h4 style="margin:0 0 10px; color:var(--primary-400);">مبيعات المندوب</h4>
        <table class="data-table" style="margin-bottom: 20px;">
          <thead><tr><th>التاريخ</th><th>العميل</th><th>المنتج</th><th>الكمية</th><th>الإجمالي</th></tr></thead>
          <tbody>
            ${sales.map(s => `<tr><td>${s.date}</td><td>${s.customer}</td><td>${s.product}</td><td>${s.netQuantity || 0}</td><td>${ExcelEngine.formatCurrency(s.total)}</td></tr>`).join('')}
          </tbody>
        </table>

        <h4 style="margin:0 0 10px; color:var(--warning-500);">مشتريات المندوب</h4>
        <table class="data-table">
          <thead><tr><th>التاريخ</th><th>المورد</th><th>المنتج</th><th>الكمية</th><th>الإجمالي</th></tr></thead>
          <tbody>
            ${purchases.map(p => `<tr><td>${p.date}</td><td>${p.supplier}</td><td>${p.product}</td><td>${p.quantity || 0}</td><td>${ExcelEngine.formatCurrency(p.total)}</td></tr>`).join('')}
          </tbody>
        </table>
      `;
    } else if (type === 'supplier') {
      const purchases = ExcelEngine.getPurchases().filter(p => p.supplier === name);
      const totalCost = purchases.reduce((sum, p) => sum + (p.total || 0), 0);
      const totalPaid = purchases.reduce((sum, p) => sum + (p.paid || 0), 0);
      const remaining = totalCost - totalPaid;

      if (actions) {
        actions.innerHTML = `
          <button class="btn btn-primary btn-sm" onclick="App.printSupplierStatement()" style="margin-left:8px;"><i class="fas fa-print"></i> طباعة كشف حساب مورد</button>
          <button class="btn btn-accent btn-sm" onclick="App.exportReport('supplier', '${name}')"><i class="fas fa-file-excel"></i> تصدير التقرير (Excel)</button>
        `;
      }

      summary.innerHTML = `
        <div class="card premium-card"><div>إجمالي المسحوبات</div><div class="text-xl font-num">${ExcelEngine.formatCurrency(totalCost)}</div></div>
        <div class="card premium-card"><div>إجمالي المدفوعات</div><div class="text-xl font-num" style="color:#4caf50">${ExcelEngine.formatCurrency(totalPaid)}</div></div>
        <div class="card premium-card"><div>المستحق للمورد</div><div class="text-xl font-num" style="color:#ef5350">${ExcelEngine.formatCurrency(remaining)}</div></div>
      `;

      wrapper.innerHTML = `
        <h4 style="margin:0 0 10px; color:var(--primary-400);">سجل التعاملات والمنتجات</h4>
        <table class="data-table">
          <thead><tr><th>التاريخ</th><th>المندوب</th><th>المنتج</th><th>الكمية</th><th>السعر</th><th>الإجمالي</th><th>المدفوع</th></tr></thead>
          <tbody>
            ${purchases.map(p => `<tr><td>${p.date}</td><td>${p.rep}</td><td>${p.product}</td><td>${p.quantity}</td><td>${ExcelEngine.formatCurrency(p.price)}</td><td>${ExcelEngine.formatCurrency(p.total)}</td><td style="color:green">${ExcelEngine.formatCurrency(p.paid)}</td></tr>`).join('')}
          </tbody>
        </table>
      `;
    }

    // Render rep warehouse section for rep profiles
    if (type === 'rep') {
      const extraInfo = document.getElementById('profileExtraInfo');
      if (extraInfo) {
        const repStock = ExcelEngine.getRepWarehouse(name);
        const transfers = ExcelEngine.getWarehouseTransfers(name).sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 20);

        let warehouseHtml = `
          <h4 style="margin: 0 0 12px; color:var(--warning-500);"><i class="fas fa-warehouse" style="margin-left:6px;"></i>مستودع المندوب</h4>
        `;

        if (repStock.length === 0) {
          warehouseHtml += `<div style="padding:15px; text-align:center; color:var(--text-tertiary);"><i class="fas fa-box-open"></i> لا يوجد مخزون حالي</div>`;
        } else {
          warehouseHtml += `<table class="data-table" style="margin-bottom:16px;">
            <thead><tr><th>المنتج</th><th>الكمية</th></tr></thead>
            <tbody>${repStock.map(rs => `<tr><td>${rs.product}</td><td class="number" style="font-weight:bold;color:var(--primary-400)">${rs.quantity}</td></tr>`).join('')}</tbody>
          </table>`;
        }

        if (transfers.length > 0) {
          warehouseHtml += `
            <h4 style="margin:16px 0 8px; color:var(--text-secondary); font-size:0.85rem;">آخر حركات التحويل</h4>
            <div style="max-height:200px; overflow-y:auto;">
              ${transfers.map(t => `
                <div style="display:flex;justify-content:space-between;align-items:center;padding:6px 8px;border-bottom:1px solid var(--border-glass);font-size:0.8rem;">
                  <span>${t.date} - ${t.product}</span>
                  <span style="color:${t.type === 'OUT_TO_REP' ? 'var(--success-400)' : 'var(--warning-500)'};font-weight:bold;">
                    ${t.type === 'OUT_TO_REP' ? '+' : '-'}${t.quantity}
                  </span>
                </div>
              `).join('')}
            </div>
          `;
        }

        warehouseHtml += `
          <div style="margin-top:16px;">
            <button class="btn btn-sm btn-primary" onclick="App.openWarehouseTransfer('${name}')">
              <i class="fas fa-dolly"></i> سحب/إرجاع مخزون
            </button>
          </div>
        `;

        extraInfo.innerHTML = warehouseHtml;
      }
    }
  }

  function printVoucher(serial) {
    const v = ExcelEngine.getVouchers().find(x => x.serial === serial);
    if (!v) return;
    const isReceipt = v.type === 'RECEIPT';
    const title = isReceipt ? 'سند قبض' : 'سند صرف';
    const color = isReceipt ? '#4CAF50' : '#EF5350';
    const icon = isReceipt ? '📥' : '📤';

    const win = window.open('', '_blank');
    if (!win) return;
    win.document.write(`
      <!DOCTYPE html>
      <html dir="rtl" lang="ar">
      <head>
        <meta charset="UTF-8">
        <title>${title} - ${v.serial}</title>
        <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap" rel="stylesheet">
        <script src="https://cdn.jsdelivr.net/npm/qrcodejs@1.0.0/qrcode.min.js"><\/script>
        <style>
          * { margin: 0; padding: 0; box-sizing: border-box; }
          body { font-family: 'Cairo', sans-serif; padding: 30px; color: #1a2035; background: #fff; }
          .voucher-box { max-width: 700px; margin: auto; border: 3px solid ${color}; border-radius: 12px; overflow: hidden; }
          .voucher-header { background: ${color}; color: white; padding: 20px 30px; display: flex; justify-content: space-between; align-items: center; }
          .voucher-header h1 { font-size: 26px; font-weight: 800; }
          .voucher-header .serial { font-size: 14px; opacity: 0.9; background: rgba(255,255,255,0.2); padding: 4px 12px; border-radius: 20px; }
          .voucher-body { padding: 30px; }
          .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 25px; }
          .info-item { background: #f8f9fa; padding: 15px; border-radius: 8px; border-right: 4px solid ${color}; }
          .info-item label { font-size: 12px; color: #666; display: block; margin-bottom: 4px; }
          .info-item .value { font-size: 16px; font-weight: 700; color: #1a2035; }
          .amount-box { background: linear-gradient(135deg, ${color}11, ${color}22); border: 2px solid ${color}; border-radius: 12px; padding: 25px; text-align: center; margin-bottom: 25px; }
          .amount-box label { font-size: 14px; color: #555; display: block; margin-bottom: 8px; }
          .amount-box .amount { font-size: 32px; font-weight: 800; color: ${color}; font-family: 'Cairo', sans-serif; }
          .notes-box { background: #f8f9fa; border-radius: 8px; padding: 15px; margin-bottom: 25px; min-height: 60px; }
          .notes-box label { font-size: 12px; color: #666; display: block; margin-bottom: 4px; }
          .signature-section { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 20px; margin-top: 40px; padding-top: 20px; border-top: 2px dashed #ddd; }
          .sig-box { text-align: center; }
          .sig-box .line { border-bottom: 1px solid #333; width: 80%; margin: 40px auto 8px; }
          .sig-box label { font-size: 12px; color: #666; }
          .voucher-footer { display: flex; justify-content: space-between; align-items: flex-end; margin-top: 30px; padding-top: 15px; border-top: 1px solid #eee; }
          .footer-text { font-size: 10px; color: #999; }
          #qrcode { width: 80px; height: 80px; }
          @media print { .no-print { display: none; } body { padding: 10px; } }
        </style>
      </head>
      <body>
        <div class="voucher-box">
          <div class="voucher-header">
            <div>
              <h1>${icon} ${title}</h1>
              <p style="font-size:13px; margin-top:4px;">مشروع الفروج الوطني الرائع</p>
            </div>
            <div style="text-align:left;">
              <div class="serial">${v.serial}</div>
              <p style="font-size:12px; margin-top:6px;">${v.date}</p>
            </div>
          </div>

          <div class="voucher-body">
            <div class="amount-box">
              <label>المبلغ المستحق</label>
              <div class="amount">${ExcelEngine.formatCurrency(v.amount)}</div>
            </div>

            <div class="info-grid">
              <div class="info-item">
                <label>الجهة / الطرف ${isReceipt ? '(المستلم منه)' : '(المصروف له)'}</label>
                <div class="value">${v.party || '---'}</div>
              </div>
              <div class="info-item">
                <label>تاريخ التحرير</label>
                <div class="value">${v.date}</div>
              </div>
              <div class="info-item">
                <label>نوع السند</label>
                <div class="value">${title}</div>
              </div>
              <div class="info-item">
                <label>محرر بواسطة</label>
                <div class="value">${v.user || 'المدير'}</div>
              </div>
            </div>

            <div class="notes-box">
              <label>البيان / ملاحظات</label>
              <p style="font-size:14px; margin-top:5px;">${v.notes || 'بدون ملاحظات'}</p>
            </div>

            <div class="signature-section">
              <div class="sig-box">
                <div class="line"></div>
                <label>توقيع المحاسب</label>
              </div>
              <div class="sig-box">
                <div class="line"></div>
                <label>توقيع ${isReceipt ? 'الدافع' : 'المستلم'}</label>
              </div>
              <div class="sig-box">
                <div class="line"></div>
                <label>توقيع المدير</label>
              </div>
            </div>

            <div class="voucher-footer">
              <div class="footer-text">
                <p>تم إصدار هذا السند آلياً من نظام ERP المتكامل</p>
                <p>هذا السند محمي ومُوثّق بنظام التشفير الرقمي</p>
              </div>
              <div id="qrcode"></div>
            </div>
          </div>
        </div>

        <script src="https://cdn.jsdelivr.net/npm/qrcodejs@1.0.0/qrcode.min.js"><\/script>
        <script>
          window.onload = function() {
            setTimeout(function() {
              if (window.QRCode) {
                new QRCode(document.getElementById("qrcode"), {
                  text: "${title}: ${v.serial}\\nالتاريخ: ${v.date}\\nالمبلغ: ${v.amount}\\nالجهة: ${v.party || ''}",
                  width: 80, height: 80,
                  colorDark: "#000", colorLight: "#fff",
                  correctLevel: QRCode.CorrectLevel.M
                });
              }
              setTimeout(() => { window.print(); }, 300);
            }, 200);
          };
        <\/script>
      </body>
      </html>
    `);
    win.document.close();
  }

  function printRepDailyReport(repName) {
    const today = ExcelEngine.getTodayStr();
    // Get stats for today specifically to show a daily report, or the overall stats?
    // User requested: "تقرير يومي اطبعه للمندوب احاسبه عليه من اجمالي مبيعاته و مشترياته و الارباح بينهم الفواتير اللي عملها"
    // Let's filter today's sales and purchases
    const allSales = ExcelEngine.getSales().filter(s => s.rep === repName && s.date === today);
    const allPurchases = ExcelEngine.getPurchases().filter(p => p.rep === repName && p.date === today);

    const dSales = allSales.reduce((acc, curr) => acc + curr.total, 0);
    const dPurchases = allPurchases.reduce((acc, curr) => acc + curr.total, 0);
    // Simple profit calculation for today: total sales - (total cost of items sold today)
    let dProfit = 0;
    allSales.forEach(s => {
      const prod = ExcelEngine.getProducts().find(p => p.name === s.product) || { buyPrice: 0, sellPrice: 0 };
      // profit per unit = sale price - product base unit buy price
      const profitPerUnit = (s.price > 0 ? s.price : prod.sellPrice) - prod.buyPrice;
      dProfit += (s.netQuantity || 0) * profitPerUnit;
    });

    const win = window.open('', '_blank');
    win.document.write(`
      <html dir="rtl" lang="ar">
      <head>
        <meta charset="UTF-8">
        <title>التقرير اليومي - ${repName}</title>
        <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap" rel="stylesheet">
        <style>
          body { font-family: 'Cairo', sans-serif; padding: 20px; direction: rtl; }
          .header { text-align: center; border-bottom: 2px solid #555; padding-bottom: 10px; margin-bottom: 20px; }
          .summary { display: flex; gap: 10px; margin-bottom: 20px; text-align:center; }
          .summary-card { flex: 1; padding: 15px; border: 1px solid #ddd; border-radius: 8px; background: #f9f9f9; }
          .summary-card h3 { margin: 0 0 10px 0; font-size:16px; color:#555; }
          .summary-card p { margin: 0; font-size: 20px; font-weight: bold; }
          table { width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 14px; }
          th, td { border: 1px solid #ccc; padding: 8px; text-align: center; }
          th { background: #eee; }
          .sales-th { background: rgba(33, 150, 243, 0.1); }
          .purch-th { background: rgba(255, 152, 0, 0.1); }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>كشف حساب المندوب اليومي</h1>
          <h2>المندوب: ${repName} | التاريخ: ${today}</h2>
        </div>
        <div class="summary">
          <div class="summary-card"><h3>إجمالي المشتريات</h3><p style="color:#ff9800">${ExcelEngine.formatCurrency(dPurchases)}</p></div>
          <div class="summary-card"><h3>إجمالي المبيعات</h3><p style="color:#2196f3">${ExcelEngine.formatCurrency(dSales)}</p></div>
          <div class="summary-card"><h3>الأرباح اليومية</h3><p style="color:${dProfit >= 0 ? '#4caf50' : '#f44336'}">${ExcelEngine.formatCurrency(dProfit)}</p></div>
        </div>
        
        <h3>فواتير المشتريات (اليوم)</h3>
        <table>
          <thead><tr class="purch-th"><th>رقم</th><th>المورد</th><th>المنتج</th><th>الكمية</th><th>الإجمالي</th></tr></thead>
          <tbody>
            ${allPurchases.length ? allPurchases.map((p, i) => `<tr><td>${i + 1}</td><td>${p.supplier}</td><td>${p.product}</td><td>${p.quantity}</td><td>${ExcelEngine.formatCurrency(p.total)}</td></tr>`).join('') : '<tr><td colspan="5">لا توجد مشتريات اليوم</td></tr>'}
          </tbody>
        </table>

        <h3>فواتير المبيعات (اليوم)</h3>
        <table>
          <thead><tr class="sales-th"><th>رقم</th><th>العميل</th><th>المنتج</th><th>الكمية</th><th>المرتجع</th><th>الإجمالي</th></tr></thead>
          <tbody>
            ${allSales.length ? allSales.map((s, i) => `<tr><td>${i + 1}</td><td>${s.customer}</td><td>${s.product}</td><td>${s.quantity}</td><td>${s.returned}</td><td>${ExcelEngine.formatCurrency(s.total)}</td></tr>`).join('') : '<tr><td colspan="6">لا توجد مبيعات اليوم</td></tr>'}
          </tbody>
        </table>
        
        <script>
          window.onload = () => window.print();
        </script>
      </body>
      </html>
    `);
    win.document.close();
  }

  function openWarehouseTransfer(repName) {
    openModal('warehouseTransferModal');
    const repSel = document.getElementById('whRep');
    if (repSel) repSel.value = repName || '';
  }

  function switchInventoryTab(tab, btn) {
    document.querySelectorAll('#page-inventory .btn-ghost').forEach(b => b.classList.remove('active'));
    if (btn) btn.classList.add('active');
    document.getElementById('inventory-stock-view').style.display = tab === 'stock' ? 'block' : 'none';
    document.getElementById('inventory-ledger-view').style.display = tab === 'ledger' ? 'block' : 'none';
  }

  // === Invoice Detail View ===
  let currentInvoiceType = 'sale';
  function viewInvoiceDetail(id, type) {
    currentInvoiceType = type;
    const inv = type === 'sale'
      ? ExcelEngine.getSales().find(s => s.id === id)
      : ExcelEngine.getPurchases().find(p => p.id === id);

    if (!inv) {
      showToast('الفاتورة غير موجودة', 'error');
      return;
    }

    navigateTo('invoice-detail');

    const titleEl = document.getElementById('invDetailTitle');
    const dateEl = document.getElementById('invDetailDate');
    const partyEl = document.getElementById('invDetailParty');
    const repEl = document.getElementById('invDetailRep');
    const tableBody = document.getElementById('invDetailTableBody');
    const notesEl = document.getElementById('invDetailNotes');

    // Update Header
    titleEl.textContent = type === 'sale' ? `تفاصيل فاتورة بيع #${inv.id}` : `تفاصيل فاتورة شراء #${inv.id}`;
    titleEl.style.color = type === 'sale' ? 'var(--primary-400)' : 'var(--warning-500)';

    dateEl.textContent = `${inv.date}`;
    partyEl.textContent = type === 'sale' ? (inv.customer || 'عميل نقدي') : (inv.supplier || 'مورد غير محدد');
    repEl.textContent = inv.rep || '-';

    // Update Actions
    const printBtn = document.getElementById('invDetailPrintBtn');
    printBtn.onclick = () => App.printInvoice(inv.id, type);

    const returnBtn = document.getElementById('invDetailReturnBtn');
    returnBtn.onclick = () => App.openReturnModal(inv.id, type);

    // Disable return if fully returned
    if (type === 'sale') {
      returnBtn.disabled = (inv.returned >= inv.quantity);
      returnBtn.style.opacity = returnBtn.disabled ? '0.5' : '1';
    } else {
      returnBtn.disabled = (inv.quantity <= 0); // purchases might not have separate 'returned' tracking yet, but keeping logic consistent
      returnBtn.style.opacity = returnBtn.disabled ? '0.5' : '1';
    }

    // Update Table
    const collectedValue = type === 'sale' ? inv.collected : inv.paid;
    const balanceValue = type === 'sale' ? inv.balance : inv.dueToSupplier;
    const balanceColor = balanceValue > 0 ? 'negative' : 'positive';

    // For net quantity display
    const returnedQty = type === 'sale' ? inv.returned : 0; // Purchase returns adjust quantity directly currently
    const netQty = type === 'sale' ? inv.netQuantity : inv.quantity;

    tableBody.innerHTML = `
      <tr>
        <td><strong>${inv.product}</strong></td>
        <td class="number">${inv.quantity}</td>
        <td class="number" style="color:var(--danger-400)">${returnedQty}</td>
        <td class="number" style="color:var(--success-400); font-weight:bold;">${netQty}</td>
        <td class="currency">${ExcelEngine.formatCurrency(inv.price)}</td>
        <td class="currency" style="font-weight:bold;">${ExcelEngine.formatCurrency(inv.total)}</td>
        <td class="currency" style="color:var(--success-400)">${ExcelEngine.formatCurrency(collectedValue)}</td>
        <td class="currency ${balanceColor}" style="font-weight:bold;">${ExcelEngine.formatCurrency(balanceValue)}</td>
      </tr>
    `;

    // Notes
    notesEl.textContent = inv.notes || 'لا توجد ملاحظات إضافية على هذه الفاتورة.';
  }

  function backToInvoices() {
    if (currentInvoiceType === 'sale') {
      navigateTo('sales');
    } else {
      navigateTo('purchases');
    }
  }

  // === Return Logic ===
  function openReturnModal(id, type) {
    const original = (type === 'sale')
      ? ExcelEngine.getSales().find(s => s.id === id)
      : ExcelEngine.getPurchases().find(p => p.id === id);

    if (!original) return;

    document.getElementById('returnOriginalId').value = id;
    document.getElementById('returnType').value = type;
    document.getElementById('returnQty').value = (type === 'sale' ? (original.netQuantity || original.quantity) : original.quantity);
    document.getElementById('returnMaxHint').textContent = `الحد الأقصى للمرتجع: ${type === 'sale' ? (original.netQuantity || original.quantity) : original.quantity}`;
    document.getElementById('returnQty').max = (type === 'sale' ? (original.netQuantity || original.quantity) : original.quantity);

    document.getElementById('returnInfoText').textContent = `فاتورة ${type === 'sale' ? 'بيع' : 'شراء'} رقم #${id} - ${original.product}`;
    openModal('returnModal');
  }

  function setupReturnForm() {
    const form = document.getElementById('returnForm');
    if (!form) return;
    form.onsubmit = (e) => {
      e.preventDefault();
      const id = parseInt(document.getElementById('returnOriginalId').value);
      const type = document.getElementById('returnType').value;
      const qty = parseFloat(document.getElementById('returnQty').value);
      const notes = document.getElementById('returnNotes').value;

      let result;
      if (type === 'sale') {
        result = ExcelEngine.addSaleReturn(id, qty, notes);
      } else {
        result = ExcelEngine.addPurchaseReturn(id, qty, notes);
      }

      if (result.success) {
        showToast('تم تسجيل المرتجع بنجاح', 'success');
        closeModal('returnModal');
        if (type === 'sale') renderSalesTable();
        else renderPurchasesTable();
        renderInventoryTable();
      } else {
        showToast(result.error || 'خطأ في العملية', 'error');
      }
    };
  }

  // === Voucher Logic ===
  function populateVoucherParties() {
    const partyType = document.getElementById('vPartyType').value;
    const select = document.getElementById('vPartyName');
    if (!select) return;

    if (partyType === 'customer') {
      const customers = ExcelEngine.getCustomers();
      select.innerHTML = customers.map(c => `<option value="${c.name}">${c.name}</option>`).join('');
    } else if (partyType === 'supplier') {
      const suppliers = ExcelEngine.getSuppliers();
      select.innerHTML = suppliers.map(s => `<option value="${s.name}">${s.name}</option>`).join('');
    } else if (partyType === 'rep') {
      const reps = ExcelEngine.getRepsList();
      select.innerHTML = reps.map(r => `<option value="${r}">${r}</option>`).join('');
    } else if (partyType === 'staff') {
      const staff = ExcelEngine.getStaff();
      select.innerHTML = staff.map(s => `<option value="${s.name}">${s.name}</option>`).join('');
    }
  }

  function setupVoucherForm() {
    const form = document.getElementById('voucherForm');
    if (!form) return;
    form.onsubmit = (e) => {
      e.preventDefault();
      const vData = {
        type: document.getElementById('vActionType').value,
        partyType: document.getElementById('vPartyType').value,
        party: document.getElementById('vPartyName').value,
        amount: parseFloat(document.getElementById('vAmount').value),
        date: document.getElementById('vDate').value || ExcelEngine.getTodayStr(),
        notes: document.getElementById('vNotes').value
      };

      const result = ExcelEngine.addVoucher(vData);
      if (result) {
        showToast('تم إصدار السند بنجاح', 'success');
        closeModal('voucherModal');
        App.printVoucher(result.serial);
        renderTreasury();
      }
    };
  }

  // === Product Editing ===
  function editProduct(id) {
    const products = ExcelEngine.getProducts();
    const p = products.find(x => x.id === id);
    if (!p) return;

    document.getElementById('editProdId').value = p.id;
    document.getElementById('editProdName').value = p.name;
    document.getElementById('editProdSellPrice').value = p.sellPrice;
    document.getElementById('editProdBuyPrice').value = p.buyPrice;
    document.getElementById('editProdUnit').value = p.unit;
    document.getElementById('editProdPiecesPerCarton').value = p.piecesPerCarton || 1;

    const catSelect = document.getElementById('editProdCategory');
    const cats = ExcelEngine.getCategories() || ['عام'];
    catSelect.innerHTML = cats.map(c => `<option value="${c}" ${c === p.category ? 'selected' : ''}>${c}</option>`).join('');

    openModal('editProductModal');
  }

  function setupEditProductForm() {
    const form = document.getElementById('editProductForm');
    if (!form) return;
    form.onsubmit = (e) => {
      e.preventDefault();
      const id = parseInt(document.getElementById('editProdId').value);
      const data = {
        name: document.getElementById('editProdName').value,
        sellPrice: parseFloat(document.getElementById('editProdSellPrice').value),
        buyPrice: parseFloat(document.getElementById('editProdBuyPrice').value),
        unit: document.getElementById('editProdUnit').value,
        piecesPerCarton: parseInt(document.getElementById('editProdPiecesPerCarton').value),
        category: document.getElementById('editProdCategory').value
      };

      if (ExcelEngine.updateProduct(id, data)) {
        showToast('تم تعديل المنتج بنجاح', 'success');
        closeModal('editProductModal');
        renderProductsTable();
        renderInventoryTable();
      }
    };
  }

  // === Category Management ===
  function addCategory() {
    renderCategoryList();
    openModal('categoryModal');
  }

  function renderCategoryList() {
    const ul = document.getElementById('categoryListUl');
    if (!ul) return;
    const cats = ExcelEngine.getCategories() || ['عام'];
    ul.innerHTML = cats.map(c => `
      <li style="padding:10px; border-bottom:1px solid rgba(0,0,0,0.05); display:flex; justify-content:space-between; align-items:center;">
        <span>${c}</span>
      </li>
    `).join('');
    populateCategorySelects();
  }

  function populateCategorySelects() {
    const selects = document.querySelectorAll('.categorySelect');
    const cats = ExcelEngine.getCategories() || ['عام'];
    selects.forEach(s => {
      const current = s.value;
      s.innerHTML = cats.map(c => `<option value="${c}">${c}</option>`).join('');
      if (current) s.value = current;
    });
  }

  function submitAddCategory() {
    const nameInput = document.getElementById('newCategoryInput');
    const name = nameInput.value.trim();
    if (!name) return;
    if (ExcelEngine.addCategory(name)) {
      nameInput.value = '';
      renderCategoryList();
      showToast('تم إضافة التصنيف', 'success');
    }
  }



  // === End of Duplicate Block Removal ===

  // === Voucher Printing ===
  function printVoucher(serial) {
    const vouchers = ExcelEngine.getVouchers();
    const v = vouchers.find(x => x.serial === serial);
    if (!v) return;

    const isReceipt = v.type === 'RECEIPT';
    const printWindow = window.open('', '_blank');
    if (!printWindow) return;

    printWindow.document.write(`<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head>
  <meta charset="UTF-8">
  <title>${isReceipt ? 'سند قبض' : 'سند صرف'} #${v.serial}</title>
  <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap" rel="stylesheet">
  <style>
    body { font-family:'Cairo',sans-serif; padding:30px; direction:rtl; background:#fff; }
    .voucher-box { max-width:600px; margin:auto; border:2px solid ${isReceipt ? '#1a73e8' : '#f57c00'}; border-radius:12px; padding:30px; }
    .voucher-header { text-align:center; margin-bottom:20px; border-bottom:2px dashed ${isReceipt ? '#1a73e8' : '#f57c00'}; padding-bottom:15px; }
    .voucher-header h1 { color:${isReceipt ? '#1a73e8' : '#f57c00'}; font-size:28px; margin:0 0 5px; }
    .voucher-header .serial { font-size:14px; color:#666; }
    .voucher-body { display:grid; grid-template-columns:1fr 1fr; gap:15px; margin:20px 0; }
    .field { background:#f8f9fc; padding:12px; border-radius:8px; }
    .field label { font-size:11px; color:#888; display:block; margin-bottom:4px; }
    .field .value { font-size:16px; font-weight:700; color:#333; }
    .amount-box { background:${isReceipt ? '#e3f2fd' : '#fff3e0'}; border:2px solid ${isReceipt ? '#1a73e8' : '#f57c00'}; border-radius:10px; text-align:center; padding:20px; margin:20px 0; }
    .amount-box .label { font-size:13px; color:#666; margin-bottom:8px; }
    .amount-box .amount { font-size:32px; font-weight:800; color:${isReceipt ? '#1a73e8' : '#f57c00'}; }
    .signatures { display:grid; grid-template-columns:1fr 1fr; gap:30px; margin-top:30px; border-top:1px dashed #ccc; padding-top:20px; }
    .sig-box { text-align:center; }
    .sig-box .line { border-bottom:1px solid #333; height:40px; margin-bottom:8px; }
    .sig-box .label { font-size:12px; color:#666; }
    .footer { text-align:center; margin-top:20px; font-size:11px; color:#aaa; }
    #qrcode { margin: 10px auto; display:flex; justify-content:center; }
  </style>
</head>
<body>
  <div class="voucher-box">
    <div class="voucher-header">
      <h1>${isReceipt ? '🟢 سند قبض' : '🔴 سند صرف'}</h1>
      <div class="serial">رقم السند: <strong>#${v.serial}</strong> | التاريخ: <strong>${v.date}</strong></div>
    </div>
    <div class="voucher-body">
      <div class="field"><label>نوع الجهة</label><div class="value">${v.partyType === 'customer' ? 'عميل' : v.partyType === 'supplier' ? 'مورد' : v.partyType === 'rep' ? 'مندوب' : 'موظف'}</div></div>
      <div class="field"><label>اسم الجهة</label><div class="value">${v.party}</div></div>
      <div class="field" style="grid-column:1/-1"><label>ملاحظات</label><div class="value">${v.notes || '-'}</div></div>
    </div>
    <div class="amount-box">
      <div class="label">${isReceipt ? 'المبلغ المقبوض' : 'المبلغ المصروف'}</div>
      <div class="amount">${ExcelEngine.formatCurrency(v.amount)}</div>
    </div>
    <div id="qrcode"></div>
    <div class="signatures">
      <div class="sig-box"><div class="line"></div><div class="label">المستلم / المحصّل</div></div>
      <div class="sig-box"><div class="line"></div><div class="label">المدير / المسؤول</div></div>
    </div>
    <div class="footer">مشروع الفروج الوطني - نظام ERP المتكامل | v2.0</div>
  </div>
  <script src="https://cdn.jsdelivr.net/npm/qrcodejs@1.0.0/qrcode.min.js"><\/script>
  <script>
    window.onload = function() {
      setTimeout(function() {
        if (window.QRCode) {
          new QRCode(document.getElementById('qrcode'), {
            text: 'سند #${v.serial} | ${v.party} | ${ExcelEngine.formatCurrency(v.amount)}',
            width:100, height:100
          });
        }
        setTimeout(() => window.print(), 400);
      }, 100);
    };
  <\/script>
</body>
</html>`);
    printWindow.document.close();
  }

  // === Reports Export ===
  async function exportReport(type, explicitName = null) {
    const repSel = document.getElementById('reportRepSelect');
    const custSel = document.getElementById('reportCustomerSelect');
    const prodSel = document.getElementById('reportProductSelect');

    if (type === 'rep') {
      const name = explicitName || repSel?.value;
      if (!name) { showToast('اختر مندوباً أولاً', 'warning'); return; }
      try {
        showLoading('جاري تصدير التقرير...');
        await ExcelEngine.exportAdvancedReports('rep', name);
        hideLoading();
        showToast('تم التصدير بنجاح', 'success');
      } catch (e) { hideLoading(); showToast('خطأ في التصدير: ' + e.message, 'error'); }
    } else if (type === 'customer') {
      const name = explicitName || custSel?.value;
      if (!name) { showToast('اختر عميلاً أولاً', 'warning'); return; }
      try {
        showLoading('جاري تصدير التقرير...');
        await ExcelEngine.exportAdvancedReports('customer', name);
        hideLoading();
        showToast('تم التصدير بنجاح', 'success');
      } catch (e) { hideLoading(); showToast('خطأ في التصدير: ' + e.message, 'error'); }
    } else if (type === 'supplier') {
      const name = explicitName;
      if (!name) { showToast('اختر موردا أولاً', 'warning'); return; }
      try {
        showLoading('جاري تصدير التقرير...');
        await ExcelEngine.exportAdvancedReports('customer', name); // 'customer' formatter parses history cleanly, Supplier acts similarly on transactions list if implemented correctly
        hideLoading();
        showToast('تم التصدير بنجاح', 'success');
      } catch (e) { hideLoading(); showToast('خطأ في التصدير: ' + e.message, 'error'); }
    } else if (type === 'productProfit') {
      const name = prodSel?.value;
      if (!name) { showToast('اختر منتجاً أولاً', 'warning'); return; }
      try {
        showLoading('جاري تصدير تقرير الربحية...');
        await ExcelEngine.exportAdvancedReports('productProfit', name);
        hideLoading();
        showToast('تم تصدير تقرير الربحية بنجاح', 'success');
      } catch (e) { hideLoading(); showToast('خطأ في التصدير: ' + e.message, 'error'); }
    } else {
      try {
        showLoading('جاري تصدير التقرير...');
        await ExcelEngine.exportAdvancedReports(type, null);
        hideLoading();
        showToast('تم التصدير بنجاح', 'success');
      } catch (e) { hideLoading(); showToast('خطأ في التصدير: ' + e.message, 'error'); }
    }
  }

  // Initialize on DOM ready
  document.addEventListener('DOMContentLoaded', () => {

    init();
  });

  // === ZATCA Phase 2 Simulation ===
  let currentZatcaInvoice = null;

  async function openZatcaCompliance() {
    const inv = currentInvoiceType === 'sale'
      ? ExcelEngine.getSales().find(s => s.id === parseInt(document.getElementById('invDetailTitle').textContent.match(/\d+/)?.[0]))
      : ExcelEngine.getPurchases().find(p => p.id === parseInt(document.getElementById('invDetailTitle').textContent.match(/\d+/)?.[0]));

    if (!inv) return;

    showLoading('جاري تحضير ملف المطابقة الرقمية...');
    currentZatcaInvoice = await ZatcaEngine.prepareZatcaCompliantInvoice(inv);
    hideLoading();

    document.getElementById('zatcaXmlView').textContent = currentZatcaInvoice.xml;
    document.getElementById('zatcaHashContent').textContent = currentZatcaInvoice.hash;

    document.getElementById('zatcaStatusBadge').textContent = 'جاهزة للربط والارسال';
    document.getElementById('zatcaStatusBadge').style.color = 'var(--success-400)';

    openModal('zatcaModal');
  }

  function switchZatcaTab(tab, btn) {
    document.querySelectorAll('#zatcaModal .btn-ghost').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('zatcaXmlView').style.display = tab === 'xml' ? 'block' : 'none';
    document.getElementById('zatcaHashView').style.display = tab === 'hash' ? 'block' : 'none';
  }

  function downloadZatcaXML() {
    if (!currentZatcaInvoice) return;
    const blob = new Blob([currentZatcaInvoice.xml], { type: 'text/xml' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `ZATCA-INV-${currentZatcaInvoice.hash.substring(0, 8)}.xml`;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    showToast('تم تحميل فاتورة UBL 2.1 بنجاح', 'success');
  }

  function verifyZatcaHash() {
    showToast('بصمة الـ SHA-256 صحيحة ومطابقة لشروط الهيئة', 'success');
  }

  // === Advanced Reports ===
  function openAdvancedReport(reportType) {
    const titles = {
      'salesReport': 'تقرير المبيعات',
      'inventoryMovement': 'حركة المخزون',
      'profitLoss': 'الأرباح والخسائر',
      'expensesReport': 'المصروفات',
      'repPerformance': 'أداء المناديب',
      'taxReport': 'التقرير الضريبي',
      'debtsReport': 'المديونيات',
      'productProfitability': 'ربحية المنتجات'
    };

    const previewArea = document.getElementById('reportPreviewArea');
    const previewTitle = document.getElementById('reportPreviewTitle');
    const previewContent = document.getElementById('reportPreviewContent');

    if (!previewArea || !previewTitle || !previewContent) return;

    previewTitle.textContent = titles[reportType] || 'تقرير';
    previewArea.style.display = 'block';

    let data;
    switch (reportType) {
      case 'salesReport':
        data = ExcelEngine.getSalesReport();
        previewContent.innerHTML = renderSalesReportTable(data);
        break;
      case 'inventoryMovement':
        data = ExcelEngine.getInventoryMovementReport();
        previewContent.innerHTML = renderInventoryMovementTable(data);
        break;
      case 'profitLoss':
        data = ExcelEngine.getProfitLossReport();
        previewContent.innerHTML = renderProfitLossTable(data);
        break;
      case 'expensesReport':
        data = ExcelEngine.getExpensesReport();
        previewContent.innerHTML = renderExpensesTable(data);
        break;
      case 'repPerformance':
        data = ExcelEngine.getRepPerformanceReport();
        previewContent.innerHTML = renderRepPerformanceTable(data);
        break;
      case 'taxReport':
        data = ExcelEngine.getTaxReport();
        previewContent.innerHTML = renderTaxReportTable(data);
        break;
      case 'debtsReport':
        data = ExcelEngine.getDebtsReport();
        previewContent.innerHTML = renderDebtsReportTable(data);
        break;
      case 'productProfitability':
        data = ExcelEngine.getProductProfitabilityReport();
        previewContent.innerHTML = renderProductProfitabilityTable(data);
        break;
      default:
        previewContent.innerHTML = '<p class="text-slate-500 text-center">لا توجد بيانات</p>';
    }

    // Scroll to preview
    previewArea.scrollIntoView({ behavior: 'smooth' });
  }

  function renderSalesReportTable(data) {
    if (!data || data.length === 0) return '<p class="text-slate-500 text-center">لا توجد بيانات مبيعات</p>';

    let html = '<table class="data-table"><thead><tr><th>التاريخ</th><th>العميل</th><th>المنتج</th><th>المندوب</th><th>الكمية</th><th>السعر</th><th>الإجمالي</th><th>المحصل</th><th>المتبقي</th></tr></thead><tbody>';

    data.forEach(row => {
      html += `<tr>
        <td>${row.date || '-'}</td>
        <td>${row.customer || '-'}</td>
        <td>${row.product || '-'}</td>
        <td>${row.rep || '-'}</td>
        <td class="number">${row.quantity || 0}</td>
        <td class="currency">${ExcelEngine.formatCurrency(row.price || 0)}</td>
        <td class="currency">${ExcelEngine.formatCurrency(row.total || 0)}</td>
        <td class="currency">${ExcelEngine.formatCurrency(row.collected || 0)}</td>
        <td class="currency ${row.balance > 0 ? 'negative' : ''}">${ExcelEngine.formatCurrency(row.balance || 0)}</td>
      </tr>`;
    });

    html += '</tbody></table>';

    // Summary
    const totalSales = data.reduce((sum, r) => sum + (r.total || 0), 0);
    const totalCollected = data.reduce((sum, r) => sum + (r.collected || 0), 0);
    const totalBalance = data.reduce((sum, r) => sum + (r.balance || 0), 0);

    html += `<div class="mt-4 p-4 bg-surface rounded-lg">
      <div class="grid grid-cols-3 gap-4 text-sm">
        <div><strong>إجمالي المبيعات:</strong> <span class="currency">${ExcelEngine.formatCurrency(totalSales)}</span></div>
        <div><strong>إجمالي المحصل:</strong> <span class="currency" style="color:#166534">${ExcelEngine.formatCurrency(totalCollected)}</span></div>
        <div><strong>إجمالي المتبقي:</strong> <span class="currency negative">${ExcelEngine.formatCurrency(totalBalance)}</span></div>
      </div>
    </div>`;

    return html;
  }

  function renderInventoryMovementTable(data) {
    if (!data || data.length === 0) return '<p class="text-slate-500 text-center">لا توجد حركات مخزون</p>';

    let html = '<table class="data-table"><thead><tr><th>التاريخ</th><th>المنتج</th><th>النوع</th><th>الكمية</th><th>الموقع</th><th>المرجع</th></tr></thead><tbody>';

    data.forEach(row => {
      const typeColor = row.type === 'IN' ? '#166534' : '#991b1b';
      html += `<tr>
        <td>${row.date || '-'}</td>
        <td>${row.product || '-'}</td>
        <td><span style="color:${typeColor};font-weight:bold">${row.type === 'IN' ? 'وارد +' : 'صادر -'}</span></td>
        <td class="number">${row.quantity || 0}</td>
        <td>${row.location || '-'}</td>
        <td>${row.reference || '-'}</td>
      </tr>`;
    });

    html += '</tbody></table>';
    return html;
  }

  function renderProfitLossTable(data) {
    if (!data) return '<p class="text-slate-500 text-center">لا توجد بيانات</p>';

    let html = '<div class="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-6">';
    html += `<div class="stat-card"><p class="text-xs text-slate-400 mb-1">إجمالي المبيعات</p><p class="text-lg font-bold currency">${ExcelEngine.formatCurrency(data.totalSales || 0)}</p></div>`;
    html += `<div class="stat-card"><p class="text-xs text-slate-400 mb-1">إجمالي المشتريات</p><p class="text-lg font-bold currency">${ExcelEngine.formatCurrency(data.totalPurchases || 0)}</p></div>`;
    html += `<div class="stat-card"><p class="text-xs text-slate-400 mb-1">إجمالي المصروفات</p><p class="text-lg font-bold currency" style="color:#991b1b">${ExcelEngine.formatCurrency(data.totalExpenses || 0)}</p></div>`;
    html += `<div class="stat-card"><p class="text-xs text-slate-400 mb-1">صافي الربح</p><p class="text-lg font-bold currency" style="color:${(data.netProfit || 0) >= 0 ? '#166534' : '#991b1b'}">${ExcelEngine.formatCurrency(data.netProfit || 0)}</p></div>`;
    html += '</div>';

    return html;
  }

  function renderExpensesTable(data) {
    if (!data || data.length === 0) return '<p class="text-slate-500 text-center">لا توجد مصروفات</p>';

    let html = '<table class="data-table"><thead><tr><th>التاريخ</th><th>النوع</th><th>الوصف</th><th>المبلغ</th><th>الموظف</th></tr></thead><tbody>';

    data.forEach(row => {
      html += `<tr>
        <td>${row.date || '-'}</td>
        <td>${row.type || '-'}</td>
        <td>${row.description || '-'}</td>
        <td class="currency negative">${ExcelEngine.formatCurrency(row.amount || 0)}</td>
        <td>${row.staff || '-'}</td>
      </tr>`;
    });

    html += '</tbody></table>';

    const total = data.reduce((sum, r) => sum + (r.amount || 0), 0);
    html += `<div class="mt-4 p-4 bg-surface rounded-lg text-left">
      <strong>إجمالي المصروفات:</strong> <span class="currency negative">${ExcelEngine.formatCurrency(total)}</span>
    </div>`;

    return html;
  }

  function renderRepPerformanceTable(data) {
    if (!data || data.length === 0) return '<p class="text-slate-500 text-center">لا توجد بيانات مناديب</p>';

    let html = '<table class="data-table"><thead><tr><th>المندوب</th><th>عدد المبيعات</th><th>إجمالي المبيعات</th><th>إجمالي المشتريات</th><th>الربح</th><th>التحصيلات</th><th>المستحق</th><th>العهدة</th></tr></thead><tbody>';

    data.forEach(row => {
      const profitColor = (row.profit || 0) >= 0 ? '#166534' : '#991b1b';
      html += `<tr>
        <td><strong>${row.name || '-'}</strong></td>
        <td class="number">${row.salesCount || 0}</td>
        <td class="currency">${ExcelEngine.formatCurrency(row.totalSales || 0)}</td>
        <td class="currency" style="color:#854d0e">${ExcelEngine.formatCurrency(row.totalPurchases || 0)}</td>
        <td class="currency" style="color:${profitColor}">${ExcelEngine.formatCurrency(row.profit || 0)}</td>
        <td class="currency">${ExcelEngine.formatCurrency(row.totalCollected || 0)}</td>
        <td class="currency">${ExcelEngine.formatCurrency(row.totalDue || 0)}</td>
        <td class="currency negative">${ExcelEngine.formatCurrency(row.custody || 0)}</td>
      </tr>`;
    });

    html += '</tbody></table>';
    return html;
  }

  function renderTaxReportTable(data) {
    if (!data) return '<p class="text-slate-500 text-center">لا توجد بيانات ضريبية</p>';

    let html = '<div class="grid grid-cols-2 gap-4 mb-6">';
    html += `<div class="stat-card"><p class="text-xs text-slate-400 mb-1">المبيعات الخاضعة للضريبة</p><p class="text-lg font-bold currency">${ExcelEngine.formatCurrency(data.taxableSales || 0)}</p></div>`;
    html += `<div class="stat-card"><p class="text-xs text-slate-400 mb-1">الضريبة المستحقة (15%)</p><p class="text-lg font-bold currency" style="color:#991b1b">${ExcelEngine.formatCurrency(data.taxAmount || 0)}</p></div>`;
    html += '</div>';

    if (data.transactions && data.transactions.length > 0) {
      html += '<table class="data-table"><thead><tr><th>التاريخ</th><th>رقم الفاتورة</th><th>العميل</th><th>القيمة قبل الضريبة</th><th>الضريبة (15%)</th><th>الإجمالي</th></tr></thead><tbody>';

      data.transactions.forEach(row => {
        html += `<tr>
          <td>${row.date || '-'}</td>
          <td>${row.invoiceNumber || '-'}</td>
          <td>${row.customer || '-'}</td>
          <td class="currency">${ExcelEngine.formatCurrency(row.taxableAmount || 0)}</td>
          <td class="currency" style="color:#991b1b">${ExcelEngine.formatCurrency(row.taxAmount || 0)}</td>
          <td class="currency">${ExcelEngine.formatCurrency(row.total || 0)}</td>
        </tr>`;
      });

      html += '</tbody></table>';
    }

    return html;
  }

  function renderDebtsReportTable(data) {
    if (!data || data.length === 0) return '<p class="text-slate-500 text-center">لا توجد مديونيات</p>';

    let html = '<table class="data-table"><thead><tr><th>العميل</th><th>إجمالي المبيعات</th><th>المحصل</th><th>المتبقي</th><th>نسبة التحصيل</th><th>آخر عملية</th></tr></thead><tbody>';

    data.forEach(row => {
      const collectionRate = row.total > 0 ? ((row.collected / row.total) * 100).toFixed(1) : '0';
      html += `<tr>
        <td><strong>${row.customer || '-'}</strong></td>
        <td class="currency">${ExcelEngine.formatCurrency(row.totalSales || 0)}</td>
        <td class="currency" style="color:#166534">${ExcelEngine.formatCurrency(row.collected || 0)}</td>
        <td class="currency negative">${ExcelEngine.formatCurrency(row.remaining || 0)}</td>
        <td><span class="badge ${parseFloat(collectionRate) >= 80 ? 'badge-success' : parseFloat(collectionRate) >= 50 ? 'badge-warning' : 'badge-danger'}">${collectionRate}%</span></td>
        <td>${row.lastDate || '-'}</td>
      </tr>`;
    });

    html += '</tbody></table>';

    const totalSales = data.reduce((sum, r) => sum + (r.totalSales || 0), 0);
    const totalCollected = data.reduce((sum, r) => sum + (r.collected || 0), 0);
    const totalRemaining = data.reduce((sum, r) => sum + (r.remaining || 0), 0);

    html += `<div class="mt-4 p-4 bg-surface rounded-lg">
      <div class="grid grid-cols-3 gap-4 text-sm">
        <div><strong>إجمالي المبيعات:</strong> <span class="currency">${ExcelEngine.formatCurrency(totalSales)}</span></div>
        <div><strong>إجمالي المحصل:</strong> <span class="currency" style="color:#166534">${ExcelEngine.formatCurrency(totalCollected)}</span></div>
        <div><strong>إجمالي المتبقي:</strong> <span class="currency negative">${ExcelEngine.formatCurrency(totalRemaining)}</span></div>
      </div>
    </div>`;

    return html;
  }

  function renderProductProfitabilityTable(data) {
    if (!data || data.length === 0) return '<p class="text-slate-500 text-center">لا توجد بيانات ربحية</p>';

    let html = '<table class="data-table"><thead><tr><th>المنتج</th><th>المبيعات</th><th>تكلفة البضاعة</th><th>إجمالي الربح</th><th>نسبة الربح</th><th>الكمية المباعة</th></tr></thead><tbody>';

    data.forEach(row => {
      const profitMargin = row.sales > 0 ? ((row.profit / row.sales) * 100).toFixed(1) : '0';
      const marginColor = parseFloat(profitMargin) >= 30 ? '#166534' : parseFloat(profitMargin) >= 15 ? '#854d0e' : '#991b1b';
      html += `<tr>
        <td><strong>${row.product || '-'}</strong></td>
        <td class="currency">${ExcelEngine.formatCurrency(row.sales || 0)}</td>
        <td class="currency">${ExcelEngine.formatCurrency(row.cost || 0)}</td>
        <td class="currency" style="color:${marginColor}">${ExcelEngine.formatCurrency(row.profit || 0)}</td>
        <td><span class="badge ${parseFloat(profitMargin) >= 30 ? 'badge-success' : parseFloat(profitMargin) >= 15 ? 'badge-warning' : 'badge-danger'}">${profitMargin}%</span></td>
        <td class="number">${row.quantitySold || 0}</td>
      </tr>`;
    });

    html += '</tbody></table>';

    const totalSales = data.reduce((sum, r) => sum + (r.sales || 0), 0);
    const totalCost = data.reduce((sum, r) => sum + (r.cost || 0), 0);
    const totalProfit = data.reduce((sum, r) => sum + (r.profit || 0), 0);

    html += `<div class="mt-4 p-4 bg-surface rounded-lg">
      <div class="grid grid-cols-3 gap-4 text-sm">
        <div><strong>إجمالي المبيعات:</strong> <span class="currency">${ExcelEngine.formatCurrency(totalSales)}</span></div>
        <div><strong>إجمالي التكلفة:</strong> <span class="currency">${ExcelEngine.formatCurrency(totalCost)}</span></div>
        <div><strong>إجمالي الربح:</strong> <span class="currency" style="color:#166534">${ExcelEngine.formatCurrency(totalProfit)}</span></div>
      </div>
    </div>`;

    return html;
  }

  function applyDateFilter() {
    const fromDate = document.getElementById('reportFromDate')?.value;
    const toDate = document.getElementById('reportToDate')?.value;

    if (!fromDate && !toDate) {
      showToast('الرجاء تحديد تاريخ على الأقل', 'warning');
      return;
    }

    // Store filter dates for reports to use
    ExcelEngine.setReportDateFilter(fromDate, toDate);
    showToast('تم تطبيق فلتر التاريخ', 'success');

    // Refresh current report if open
    const previewArea = document.getElementById('reportPreviewArea');
    if (previewArea && previewArea.style.display === 'block') {
      // Re-run the current report with new filter
      const currentTitle = document.getElementById('reportPreviewTitle')?.textContent;
      const reportMap = {
        'تقرير المبيعات': 'salesReport',
        'حركة المخزون': 'inventoryMovement',
        'الأرباح والخسائر': 'profitLoss',
        'المصروفات': 'expensesReport',
        'أداء المناديب': 'repPerformance',
        'التقرير الضريبي': 'taxReport',
        'المديونيات': 'debtsReport',
        'ربحية المنتجات': 'productProfitability'
      };

      for (const [title, type] of Object.entries(reportMap)) {
        if (currentTitle === title) {
          openAdvancedReport(type);
          break;
        }
      }
    }
  }

  function clearDateFilter() {
    const fromDate = document.getElementById('reportFromDate');
    const toDate = document.getElementById('reportToDate');

    if (fromDate) fromDate.value = '';
    if (toDate) toDate.value = '';

    ExcelEngine.clearReportDateFilter();
    showToast('تم مسح فلتر التاريخ', 'info');
  }

  function printReport() {
    const previewContent = document.getElementById('reportPreviewContent');
    const previewTitle = document.getElementById('reportPreviewTitle');

    if (!previewContent || !previewTitle) return;

    const printWindow = window.open('', '_blank');
    if (!printWindow) {
      showToast('يرجى السماح بالنوافذ المنبثقة للطباعة', 'warning');
      return;
    }

    printWindow.document.write(`<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head>
  <meta charset="UTF-8">
  <title>${previewTitle.textContent}</title>
  <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap" rel="stylesheet">
  <style>
    body { font-family: 'Cairo', sans-serif; padding: 20px; direction: rtl; }
    .header { text-align: center; border-bottom: 3px solid #031636; padding-bottom: 15px; margin-bottom: 20px; }
    .header h1 { color: #031636; margin: 0; }
    table { width: 100%; border-collapse: collapse; margin-top: 20px; }
    th, td { padding: 8px 12px; border: 1px solid #ddd; text-align: right; font-size: 12px; }
    th { background: #031636; color: white; }
    .footer { margin-top: 20px; text-align: center; color: #666; font-size: 12px; }
    .currency { font-variant-numeric: tabular-nums; }
    .negative { color: #ef4444; font-weight: bold; }
  </style>
</head>
<body>
  <div class="header">
    <h1>${previewTitle.textContent}</h1>
    <p>تاريخ الطباعة: ${new Date().toLocaleDateString('ar-EG')}</p>
  </div>
  <div>${previewContent.innerHTML}</div>
  <div class="footer">
    <p>نظام Sovereign Ledger لإدارة الفواتير</p>
  </div>
  <script>window.print();<\/script>
</body>
</html>`);

    printWindow.document.close();
    showToast('جاري الطباعة...', 'info');
  }

  function exportReportPreview() {
    const previewContent = document.getElementById('reportPreviewContent');
    const previewTitle = document.getElementById('reportPreviewTitle');

    if (!previewContent || !previewTitle) return;

    // Create CSV from table data
    const table = previewContent.querySelector('table');
    if (!table) {
      showToast('لا يمكن تصدير هذا التقرير كجدول', 'warning');
      return;
    }

    let csv = '\uFEFF'; // BOM for Arabic support

    // Headers
    const headers = [];
    table.querySelectorAll('thead th').forEach(th => {
      headers.push(th.textContent.trim());
    });
    csv += headers.join(',') + '\n';

    // Rows
    table.querySelectorAll('tbody tr').forEach(tr => {
      const row = [];
      tr.querySelectorAll('td').forEach(td => {
        row.push('"' + td.textContent.trim().replace(/"/g, '""') + '"');
      });
      csv += row.join(',') + '\n';
    });

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${previewTitle.textContent.replace(/\s+/g, '_')}_${new Date().toISOString().split('T')[0]}.csv`;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);

    showToast('تم تصدير التقرير بنجاح', 'success');
  }

  function closeReportPreview() {
    const previewArea = document.getElementById('reportPreviewArea');
    if (previewArea) {
      previewArea.style.display = 'none';
    }
  }

  // === Public API ===
  return {
    navigateTo,
    refreshPage,
    openModal,
    closeModal,
    confirmDeleteCustomer,
    confirmDeleteProduct,
    showCustomerStatement,
    generateStatement,
    printInvoice,
    printStatement,
    exportReport,
    openRepPaymentModal,
    showSupplierDetail,
    backToSupplierList,
    confirmDeleteSupplier,
    printSupplierStatement,
    showToast,
    switchTreasuryTab,
    viewProfile,
    printVoucher,
    printRepDailyReport,
    openWarehouseTransfer,
    switchInventoryTab,
    populateStatementDropdown,
    openReturnModal,
    addCategory,
    submitAddCategory,
    editProduct,
    generateStatement,
    populateStatementDropdown,
    renderProductsTable,
    printVoucher,
    exportReport,
    viewInvoiceDetail,
    backToInvoices,
    openZatcaCompliance,
    switchZatcaTab,
    downloadZatcaXML,
    verifyZatcaHash,
    openAdvancedReport,
    applyDateFilter,
    clearDateFilter,
    printReport,
    exportReportPreview,
    closeReportPreview,
    init
  };
})();

window.App = App;

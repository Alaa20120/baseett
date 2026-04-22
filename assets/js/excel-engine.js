// ============================================
// محرك Excel - Excel Engine
// قراءة وكتابة ملفات Excel باستخدام SheetJS
// ============================================

const ExcelEngine = (() => {
  // Sheet names
  const SHEETS = {
    SALES: 'المبيعات',
    PURCHASES: 'المشتريات',
    INVENTORY: 'المخزون',
    CUSTOMERS: 'العملاء',
    REPS: 'المندوب',
    REP_FINANCIALS: 'ماليات المندوب',
    REP_PAYMENTS: 'مدفوعات المناديب',
    PRODUCTS: 'المنتجات',
    SUPPLIERS: 'الموردين',
    STATEMENTS: 'كشوفات الحساب',
    SETTINGS: 'الإعدادات',
    STAFF: 'الموظفين',
    HR: 'الموارد البشرية',
    CUST_PAY: 'دفعات العملاء'
  };

  // Column definitions per sheet
  const COLUMNS = {
    SALES: ['التاريخ', 'اسم العميل', 'المنتج', 'اسم المندوب', 'الكمية', 'المرتجع', 'صافي الكمية', 'السعر', 'الإجمالي', 'التحصيل', 'رصيد من حساب العميل', 'ملاحظات'],
    PURCHASES: ['التاريخ', 'المنتج', 'الكمية', 'اسم المندوب', 'السعر', 'الإجمالي', 'المورد', 'المستحق للمورد', 'المدفوع'],
    INVENTORY: ['المنتج', 'الرصيد السابق', 'المشترى', 'المسحوب', 'الصافي'],
    CUSTOMERS: ['اسم العميل', 'الرقم المرجعي', 'المحصل', 'الإجمالي', 'المتبقي', 'آخر تعامل'],
    REPS: ['التاريخ', 'الكمية', 'الصنف', 'اسم المندوب'],
    REP_PAYMENTS: ['التاريخ', 'اسم المندوب', 'المبلغ', 'نوع الدفع', 'ملاحظات'],
    PRODUCTS: ['اسم المنتج', 'سعر البيع', 'سعر الشراء', 'الوحدة', 'الصنف'],
    SUPPLIERS: ['اسم المورد', 'رقم الهاتف', 'العنوان', 'الإجمالي', 'المدفوع', 'المتبقي']
  };

  // Sample products from the screenshots
  const DEFAULT_PRODUCTS = [
    { name: 'بيض المروج بلاستيك', sellPrice: 155, buyPrice: 142.40, unit: 'كرتون', category: 'بيض' },
    { name: 'بيض المروج m', sellPrice: 155, buyPrice: 142.40, unit: 'كرتون', category: 'بيض' },
    { name: 'بيض المروج l2', sellPrice: 155, buyPrice: 142.40, unit: 'كرتون', category: 'بيض' },
    { name: 'بيض المروج l', sellPrice: 165, buyPrice: 157.40, unit: 'كرتون', category: 'بيض' },
    { name: 'بيض المروج xl', sellPrice: 175, buyPrice: 165, unit: 'كرتون', category: 'بيض' },
    { name: 'بيض المروج xxl', sellPrice: 185, buyPrice: 175, unit: 'كرتون', category: 'بيض' },
    { name: 'دجاج 800', sellPrice: 12, buyPrice: 11.50, unit: 'حبة', category: 'دجاج' },
    { name: 'دجاج 900 اصيلة ابيض', sellPrice: 13, buyPrice: 12, unit: 'حبة', category: 'دجاج' },
    { name: 'دجاج 1000 اصيلة ابيض', sellPrice: 13.30, buyPrice: 12.50, unit: 'حبة', category: 'دجاج' },
    { name: 'دجاج 1000 الفروج الوطني', sellPrice: 13.30, buyPrice: 11.70, unit: 'حبة', category: 'دجاج' },
    { name: 'دجاج 900 الفروج الوطني', sellPrice: 13, buyPrice: 11.70, unit: 'حبة', category: 'دجاج' },
    { name: 'دجاج 1100', sellPrice: 14, buyPrice: 13, unit: 'حبة', category: 'دجاج' },
    { name: 'دجاج 1200', sellPrice: 15, buyPrice: 14, unit: 'حبة', category: 'دجاج' },
    { name: 'شاورما', sellPrice: 195, buyPrice: 195, unit: 'كيلو', category: 'شاورما' },
    { name: 'اجل', sellPrice: 11.50, buyPrice: 11.50, unit: 'حبة', category: 'أخرى' },
    { name: 'بيض المروج معلف 10 اطباق', sellPrice: 50, buyPrice: 41.5, unit: 'كرتون', category: 'بيض' },
    { name: 'GB', sellPrice: 0, buyPrice: 0, unit: 'حبة', category: 'أخرى' }
  ];

  // Sample reps from screenshots
  const DEFAULT_REPS = ['محمد عادل', 'خالد', 'حمود', 'وليد'];

  // Sample customers from screenshots
  const DEFAULT_CUSTOMERS = [
    'الحصاد الوفير', 'روشات الجودة', 'الحروف', 'بيت الدولة الحمراء', 'بدروز المدينة',
    'السقا', 'حافظ الديك', 'حلويكي', 'زهور المداري', 'شولة حايل', 'مطبق بلدي',
    'شقردة الهجرة للتجارة', 'ركن التوفير', 'احسان ابراهيم القت', 'بروست الجودة',
    'قدموس احمد', 'نسمة المشويات', 'الرنجاد البخاري', 'الخيال الذهبي البخاري',
    'حنين البخاري', 'شاولة الروق البخاري', 'فروج النوال', 'فروح الشام',
    'اسواق كلات المركزية', 'اسواق وادي الرولة', 'بن مقبول',
    'منتى الاسوري', 'فيحاء الشام', 'بيت اجلا', 'جرائد هايبر',
    'زوادة الحادلة', 'سما الأميل المجارة', 'شهو البخاري',
    'زهور المدارن', 'ابو فرج', 'الزوادة بيدر', 'هرم التموين',
    'اسواق العريش المركزية', 'حاويات دسرتت', 'الازعه سبرون',
    'عميل 1', 'عميل 2', 'تمونيات التوفير المركزية', 'اصالة الماعدة',
    'كوبر الفيروز', 'اسواق وادي الوادة', 'عميل 3',
    'وصالة الازوق البجاري', 'عميل 4', 'واحة الخير', 'شاولة الزوق البجاري',
    'الهجرة للتجارة', 'المروج الزراعية', 'المعال المعالي',
    'مؤسسة خلود', 'معمل هاجر العامودي', 'الناصرية', 'البقالة', 'ابو حوريه'
  ];

  const DEFAULT_SUPPLIERS = [
    'المروج الزراعية', 'الناصرية', 'المعال المعالي', 'مؤسسة خلود', 'معمل هاجر العامودي'
  ];

  // Current workbook data (in-memory state)
  let workbookData = {
    sales: [],
    purchases: [],
    inventory: [],
    customers: [],
    reps: [],
    repPayments: [],
    products: [],
    suppliers: [],
    staff: [],
    hr: [],
    customerPayments: [],
    treasury: [],
    vouchers: [],
    expenses: [],
    repWarehouse: [],
    warehouseTransfers: [],
    categories: ['عام', 'بيض', 'دجاج', 'شاورما', 'أخرى']
  };

  let currentUser = null;
  let isLoaded = false;

  // === Initialize with default data ===
  function initializeDefaults() {
    // Products
    workbookData.products = DEFAULT_PRODUCTS.map((p, i) => ({
      id: i + 1,
      name: p.name,
      sellPrice: p.sellPrice,
      buyPrice: p.buyPrice,
      unit: p.unit,
      category: p.category
    }));

    // Customers
    workbookData.customers = DEFAULT_CUSTOMERS.map((name, i) => ({
      id: i + 1,
      name: name,
      refNumber: `C${String(i + 1).padStart(4, '0')}`,
      collected: 0,
      total: 0,
      remaining: 0,
      lastDate: ''
    }));

    // Inventory
    workbookData.inventory = DEFAULT_PRODUCTS.map((p, i) => ({
      id: i + 1,
      product: p.name,
      previousStock: 0,
      purchased: 0,
      withdrawn: 0,
      net: 0
    }));

    // Suppliers 
    workbookData.suppliers = DEFAULT_SUPPLIERS.map((name, i) => ({
      id: i + 1,
      name: name,
      phone: '',
      address: '',
      total: 0,
      paid: 0,
      remaining: 0
    }));

    workbookData.sales = [];
    workbookData.purchases = [];
    workbookData.reps = [];
    workbookData.repPayments = [];

    // Staff
    workbookData.staff = [
      { id: 999, name: 'admin', role: 'مشرف', password: 'admin' },
      { id: 1, name: 'محمد عادل', role: 'مندوب', password: '123' },
      { id: 2, name: 'أحمد علي', role: 'مندوب', password: '123' },
      { id: 3, name: 'سيد محمود', role: 'مندوب', password: '123' }
    ];

    workbookData.treasury = [];
    workbookData.vouchers = [];
    workbookData.customerPayments = [];
    workbookData.hr = [];
    workbookData.repWarehouse = [];
    workbookData.warehouseTransfers = [];
    workbookData.categories = ['عام', 'بيض', 'دجاج', 'شاورما', 'أخرى'];

    isLoaded = true;
    saveToLocalStorage();
  }

  // === Read Excel File ===
  async function loadFromFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: 'array', cellDates: true });

          // Parse Sales
          if (wb.SheetNames.includes(SHEETS.SALES) || wb.SheetNames.includes('المبيعات')) {
            const sheetName = wb.SheetNames.find(n => n.includes('مبيعات') || n === SHEETS.SALES) || wb.SheetNames[0];
            const ws = wb.Sheets[sheetName];
            const raw = XLSX.utils.sheet_to_json(ws, { defval: '' });
            workbookData.sales = raw.map((row, i) => ({
              id: i + 1,
              date: formatDateFromExcel(row['التاريخ'] || row['تاريخ'] || ''),
              customer: row['اسم العميل'] || row['العميل'] || '',
              product: row['المنتج'] || row['الصنف'] || '',
              rep: row['اسم المندوب'] || row['المندوب'] || '',
              quantity: parseNum(row['الكمية'] || row['كمية']),
              returned: parseNum(row['المرتجع']),
              netQuantity: parseNum(row['صافي الكمية'] || row['صافي']),
              price: parseNum(row['السعر'] || row['سعر']),
              total: parseNum(row['الإجمالي'] || row['اجمالي'] || row['الاجمالي']),
              collected: parseNum(row['التحصيل'] || row['تحصيل']),
              balance: parseNum(row['رصيد من حساب العميل'] || row['الرصيد']),
              notes: row['ملاحظات'] || ''
            }));
          }

          // Parse Purchases
          const purchSheet = wb.SheetNames.find(n => n.includes('مشتر') || n.includes('شراء'));
          if (purchSheet) {
            const ws = wb.Sheets[purchSheet];
            const raw = XLSX.utils.sheet_to_json(ws, { defval: '' });
            workbookData.purchases = raw.map((row, i) => ({
              id: i + 1,
              date: formatDateFromExcel(row['التاريخ'] || row['تاريخ'] || ''),
              product: row['المنتج'] || row['الصنف'] || '',
              quantity: parseNum(row['الكمية']),
              rep: row['اسم المندوب'] || row['المندوب'] || '',
              price: parseNum(row['السعر'] || row['سعر']),
              total: parseNum(row['الإجمالي'] || row['اجمالي'] || row['الاجمالي']),
              supplier: row['المورد'] || '',
              dueToSupplier: parseNum(row['المستحق للمورد'] || row['المستحق']),
              paid: parseNum(row['المدفوع'])
            }));
          }

          // Parse Customers
          const custSheet = wb.SheetNames.find(n => n.includes('عملا') || n.includes('عميل') || n.includes('بيانات'));
          if (custSheet) {
            const ws = wb.Sheets[custSheet];
            const raw = XLSX.utils.sheet_to_json(ws, { defval: '' });
            workbookData.customers = raw.map((row, i) => ({
              id: i + 1,
              name: row['اسم العميل'] || row['العميل'] || '',
              refNumber: row['الرقم المرجعي'] || row['الرقم الهريدي'] || `C${String(i + 1).padStart(4, '0')}`,
              collected: parseNum(row['المحصل'] || row['المحصّل']),
              total: parseNum(row['الإجمالي'] || row['الاجمالي']),
              remaining: parseNum(row['المتبقي']),
              lastDate: formatDateFromExcel(row['التاريخ'] || row['آخر تعامل'] || '')
            }));
          }

          // Parse Inventory  
          const invSheet = wb.SheetNames.find(n => n.includes('مخزون'));
          if (invSheet) {
            const ws = wb.Sheets[invSheet];
            const raw = XLSX.utils.sheet_to_json(ws, { defval: '' });
            workbookData.inventory = raw.map((row, i) => ({
              id: i + 1,
              product: row['المنتج'] || row['الصنف'] || '',
              previousStock: parseNum(row['السابق'] || row['الرصيد السابق']),
              purchased: parseNum(row['المشترى'] || row['المشتري']),
              withdrawn: parseNum(row['المسحوب']),
              net: parseNum(row['الصافي'] || row['المتبقي'] || row['الصافى'])
            }));
          }

          // Parse Rep data
          const repSheet = wb.SheetNames.find(n => n === 'المندوب' || n.includes('مندوب'));
          if (repSheet) {
            const ws = wb.Sheets[repSheet];
            const raw = XLSX.utils.sheet_to_json(ws, { defval: '' });
            workbookData.reps = raw.map((row, i) => ({
              id: i + 1,
              date: formatDateFromExcel(row['التاريخ'] || ''),
              quantity: parseNum(row['الكمية']),
              product: row['الصنف'] || row['المنتج'] || '',
              rep: row['اسم المندوب'] || row['المندوب'] || ''
            }));
          }

          if (wb.SheetNames.includes(SHEETS.STAFF)) {
            workbookData.staff = XLSX.utils.sheet_to_json(wb.Sheets[SHEETS.STAFF]);
          } else {
            // Fallback: extract from default reps if null
            workbookData.staff = DEFAULT_REPS.map((n, i) => ({ id: i + 1, name: n, role: 'مندوب', phone: '' }));
          }
          if (wb.SheetNames.includes(SHEETS.HR)) {
            workbookData.hr = XLSX.utils.sheet_to_json(wb.Sheets[SHEETS.HR]);
          } else { workbookData.hr = []; }
          if (wb.SheetNames.includes(SHEETS.CUST_PAY)) {
            workbookData.customerPayments = XLSX.utils.sheet_to_json(wb.Sheets[SHEETS.CUST_PAY]);
          } else { workbookData.customerPayments = []; }

          // Build products list from inventory if not loaded
          if (workbookData.products.length === 0 && workbookData.inventory.length > 0) {
            workbookData.products = workbookData.inventory.map((item, i) => ({
              id: i + 1,
              name: item.product,
              sellPrice: 0,
              buyPrice: 0,
              unit: 'حبة',
              category: 'عام'
            }));
          }

          isLoaded = true;
          resolve(workbookData);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  // === Export to Excel ===
  function exportToFile(filename = 'فواتير_النظام.xlsx') {
    const wb = XLSX.utils.book_new();

    // Sales sheet
    if (workbookData.sales.length > 0) {
      const salesData = workbookData.sales.map(s => ({
        'التاريخ': s.date,
        'اسم العميل': s.customer,
        'المنتج': s.product,
        'اسم المندوب': s.rep,
        'الكمية': s.quantity,
        'المرتجع': s.returned,
        'صافي الكمية': s.netQuantity,
        'السعر': s.price,
        'الإجمالي': s.total,
        'التحصيل': s.collected,
        'رصيد من حساب العميل': s.balance,
        'ملاحظات': s.notes
      }));
      const ws = XLSX.utils.json_to_sheet(salesData);
      applySheetStyles(ws, salesData.length, 12);
      XLSX.utils.book_append_sheet(wb, ws, SHEETS.SALES);
    } else {
      const ws = XLSX.utils.aoa_to_sheet([COLUMNS.SALES]);
      XLSX.utils.book_append_sheet(wb, ws, SHEETS.SALES);
    }

    // Purchases sheet
    if (workbookData.purchases.length > 0) {
      const purchData = workbookData.purchases.map(p => ({
        'التاريخ': p.date,
        'المنتج': p.product,
        'الكمية': p.quantity,
        'اسم المندوب': p.rep,
        'السعر': p.price,
        'الإجمالي': p.total,
        'المورد': p.supplier,
        'المستحق للمورد': p.dueToSupplier,
        'المدفوع': p.paid
      }));
      const ws = XLSX.utils.json_to_sheet(purchData);
      applySheetStyles(ws, purchData.length, 9);
      XLSX.utils.book_append_sheet(wb, ws, SHEETS.PURCHASES);
    } else {
      const ws = XLSX.utils.aoa_to_sheet([COLUMNS.PURCHASES]);
      XLSX.utils.book_append_sheet(wb, ws, SHEETS.PURCHASES);
    }

    // Inventory sheet
    if (workbookData.inventory.length > 0) {
      const invData = workbookData.inventory.map(inv => ({
        'المنتج': inv.product,
        'الرصيد السابق': inv.previousStock,
        'المشترى': inv.purchased,
        'المسحوب': inv.withdrawn,
        'الصافي': inv.net
      }));
      const ws = XLSX.utils.json_to_sheet(invData);
      applySheetStyles(ws, invData.length, 5);
      XLSX.utils.book_append_sheet(wb, ws, SHEETS.INVENTORY);
    }

    // Customers sheet
    if (workbookData.customers.length > 0) {
      const custData = workbookData.customers.map(c => ({
        'اسم العميل': c.name,
        'الرقم المرجعي': c.refNumber,
        'المحصل': c.collected,
        'الإجمالي': c.total,
        'المتبقي': c.remaining,
        'آخر تعامل': c.lastDate
      }));
      const ws = XLSX.utils.json_to_sheet(custData);
      applySheetStyles(ws, custData.length, 6);
      XLSX.utils.book_append_sheet(wb, ws, SHEETS.CUSTOMERS);
    }

    // Reps sheet
    if (workbookData.reps.length > 0) {
      const repData = workbookData.reps.map(r => ({
        'التاريخ': r.date,
        'الكمية': r.quantity,
        'الصنف': r.product,
        'اسم المندوب': r.rep
      }));
      const ws = XLSX.utils.json_to_sheet(repData);
      XLSX.utils.book_append_sheet(wb, ws, SHEETS.REPS);
    }

    // Rep Payments sheet
    if (workbookData.repPayments.length > 0) {
      const rpData = workbookData.repPayments.map(rp => ({
        'التاريخ': rp.date,
        'اسم المندوب': rp.rep,
        'المبلغ': rp.amount,
        'نوع الدفع': rp.paymentType,
        'ملاحظات': rp.notes
      }));
      const ws = XLSX.utils.json_to_sheet(rpData);
      XLSX.utils.book_append_sheet(wb, ws, SHEETS.REP_PAYMENTS);
    }

    // Products sheet
    if (workbookData.products.length > 0) {
      const prodData = workbookData.products.map(p => ({
        'اسم المنتج': p.name,
        'سعر البيع': p.sellPrice,
        'سعر الشراء': p.buyPrice,
        'الوحدة': p.unit,
        'الصنف': p.category
      }));
      const ws = XLSX.utils.json_to_sheet(prodData);
      applySheetStyles(ws, prodData.length, 5);
      XLSX.utils.book_append_sheet(wb, ws, SHEETS.PRODUCTS);
    }

    // Suppliers sheet
    if (workbookData.suppliers.length > 0) {
      const suppData = workbookData.suppliers.map(s => ({
        'اسم المورد': s.name,
        'رقم الهاتف': s.phone,
        'العنوان': s.address,
        'الإجمالي': s.total,
        'المدفوع': s.paid,
        'المتبقي': s.remaining
      }));
      const ws = XLSX.utils.json_to_sheet(suppData);
      applySheetStyles(ws, suppData.length, 6);
      XLSX.utils.book_append_sheet(wb, ws, SHEETS.SUPPLIERS);
    }

    // Staff / HR / Payments
    if (workbookData.staff && workbookData.staff.length > 0) {
      const ws = XLSX.utils.json_to_sheet(workbookData.staff);
      XLSX.utils.book_append_sheet(wb, ws, SHEETS.STAFF);
    }
    if (workbookData.hr && workbookData.hr.length > 0) {
      const ws = XLSX.utils.json_to_sheet(workbookData.hr);
      XLSX.utils.book_append_sheet(wb, ws, SHEETS.HR);
    }
    if (workbookData.customerPayments && workbookData.customerPayments.length > 0) {
      const ws = XLSX.utils.json_to_sheet(workbookData.customerPayments);
      XLSX.utils.book_append_sheet(wb, ws, SHEETS.CUST_PAY);
    }

    XLSX.writeFile(wb, filename);
  }

  // Apply column widths
  function applySheetStyles(ws, rowCount, colCount) {
    ws['!cols'] = Array(colCount).fill({ wch: 18 });
    if (!ws['!ref']) return;
  }

  // === Rep Warehouse System (مستودع المندوب) ===

  function transferToRepWarehouse(repName, product, quantity, notes) {
    if (!repName || !product || quantity <= 0) return { success: false, error: 'بيانات غير صحيحة' };

    // Check main warehouse stock
    const inv = workbookData.inventory.find(i => i.product === product);
    if (!inv || (inv.net || 0) < quantity) {
      return { success: false, error: `المخزون الرئيسي لا يكفي! المتاح: ${inv ? inv.net : 0}` };
    }

    // Deduct from main warehouse
    updateInventoryWithdrawal(product, quantity);

    // Add to rep warehouse
    let repItem = workbookData.repWarehouse.find(rw => rw.rep === repName && rw.product === product);
    if (!repItem) {
      repItem = { rep: repName, product: product, quantity: 0 };
      workbookData.repWarehouse.push(repItem);
    }
    repItem.quantity += quantity;

    // Log the transfer
    const transfer = {
      id: getNextId(workbookData.warehouseTransfers),
      date: getTodayStr(),
      rep: repName,
      product: product,
      quantity: quantity,
      type: 'OUT_TO_REP',
      notes: notes || '',
      user: currentUser ? currentUser.name : 'System'
    };
    workbookData.warehouseTransfers.push(transfer);

    saveToLocalStorage();
    return { success: true, transfer };
  }

  function returnFromRepWarehouse(repName, product, quantity, notes) {
    if (!repName || !product || quantity <= 0) return { success: false, error: 'بيانات غير صحيحة' };

    // Check rep warehouse stock
    const repItem = workbookData.repWarehouse.find(rw => rw.rep === repName && rw.product === product);
    if (!repItem || repItem.quantity < quantity) {
      return { success: false, error: `مستودع المندوب لا يكفي! المتاح: ${repItem ? repItem.quantity : 0}` };
    }

    // Deduct from rep warehouse
    repItem.quantity -= quantity;

    // Add back to main warehouse
    updateInventoryPurchase(product, quantity);

    // Log the transfer
    const transfer = {
      id: getNextId(workbookData.warehouseTransfers),
      date: getTodayStr(),
      rep: repName,
      product: product,
      quantity: quantity,
      type: 'RETURN_FROM_REP',
      notes: notes || 'إرجاع من مستودع المندوب',
      user: currentUser ? currentUser.name : 'System'
    };
    workbookData.warehouseTransfers.push(transfer);

    saveToLocalStorage();
    return { success: true, transfer };
  }

  function getRepWarehouse(repName) {
    return (workbookData.repWarehouse || []).filter(rw => rw.rep === repName && rw.quantity > 0);
  }

  function getRepWarehouseAll() {
    return workbookData.repWarehouse || [];
  }

  function getWarehouseTransfers(repName) {
    if (repName) {
      return (workbookData.warehouseTransfers || []).filter(t => t.rep === repName);
    }
    return workbookData.warehouseTransfers || [];
  }

  function getRepProductStock(repName, productName) {
    const item = (workbookData.repWarehouse || []).find(rw => rw.rep === repName && rw.product === productName);
    return item ? item.quantity : 0;
  }

  // === CRUD Operations ===

  // -- Sales --
  function addSale(sale) {
    const isCarton = sale.uom === 'carton';
    // Find product to get piecesPerCarton
    const prod = workbookData.products.find(p => p.name === sale.product);
    const piecesPerCarton = (prod && prod.piecesPerCarton) ? parseInt(prod.piecesPerCarton) : 1;
    const qtyMultiplier = isCarton ? piecesPerCarton : 1;

    const netQty = (sale.quantity || 0) - (sale.returned || 0);
    const stockQty = netQty * qtyMultiplier; // Real quantity for stock

    let total = netQty * (sale.price || 0);
    if (sale.applyVat) total = total - (total * 0.15);
    const newSale = {
      id: getNextId(workbookData.sales),
      date: sale.date || getTodayStr(),
      customer: sale.customer || '',
      product: sale.product || '',
      rep: sale.rep || '',
      quantity: sale.quantity || 0,
      returned: sale.returned || 0,
      uom: sale.uom || 'piece',
      piecesPerCarton: piecesPerCarton,
      netQuantity: netQty,
      stockQuantity: stockQty,
      price: sale.price || 0,
      total: total,
      collected: sale.collected || 0,
      balance: total - (sale.collected || 0),
      notes: sale.notes || '',
      impactStock: sale.impactStock !== false
    };

    // Rep Warehouse: if rep is specified, deduct from rep's warehouse instead of main
    if (newSale.impactStock && newSale.rep) {
      const repStock = getRepProductStock(newSale.rep, newSale.product);
      if (repStock < newSale.stockQuantity) {
        // Not enough in rep warehouse
        return { success: false, error: `مستودع المندوب (${newSale.rep}) لا يحتوي على كمية كافية من ${newSale.product}. المتاح: ${repStock}` };
      }
      // Deduct from rep warehouse
      const repItem = workbookData.repWarehouse.find(rw => rw.rep === newSale.rep && rw.product === newSale.product);
      if (repItem) repItem.quantity -= newSale.stockQuantity;
    } else if (newSale.impactStock && !newSale.rep) {
      // No rep specified, deduct from main warehouse as before
      const inv = workbookData.inventory.find(i => i.product === newSale.product);
      if (inv && (inv.net || 0) < newSale.stockQuantity) {
        // Warning but still process
      }
      updateInventoryWithdrawal(newSale.product, newSale.stockQuantity);
    }

    // Security: XML & Hash
    const security = generateInvoiceSecurity(newSale);
    newSale.xml = security.xml;
    newSale.signature = security.hash;

    workbookData.sales.push(newSale);

    // Update customer balance
    updateCustomerBalance(newSale.customer, newSale.total, newSale.collected);

    // Auto-log to Treasury if immediate collection
    if (newSale.collected > 0) {
      logTreasuryTransaction('IN', newSale.collected, `مبيعات: ${newSale.customer} (فاتورة #${newSale.id})`, newSale.notes);
    }

    saveToLocalStorage();
    newSale.success = true;
    return newSale;
  }

  // Immutable: deleteSale removed for security

  // -- Purchases --
  function addPurchase(purchase) {
    const isCarton = purchase.uom === 'carton';
    const prod = workbookData.products.find(p => p.name === purchase.product);
    const piecesPerCarton = (prod && prod.piecesPerCarton) ? parseInt(prod.piecesPerCarton) : 1;
    const qtyMultiplier = isCarton ? piecesPerCarton : 1;

    let total = (purchase.quantity || 0) * (purchase.price || 0);
    if (purchase.applyVat) total = total - (total * 0.15);
    const newPurchase = {
      id: getNextId(workbookData.purchases),
      date: purchase.date || getTodayStr(),
      product: purchase.product || '',
      quantity: purchase.quantity || 0,
      uom: purchase.uom || 'piece',
      piecesPerCarton: piecesPerCarton,
      stockQuantity: (purchase.quantity || 0) * qtyMultiplier,
      rep: purchase.rep || '',
      price: purchase.price || 0,
      total: total,
      supplier: purchase.supplier || '',
      dueToSupplier: total,
      paid: purchase.paid || 0,
      impactStock: purchase.impactStock !== false
    };

    // Security
    const security = generateInvoiceSecurity(newPurchase);
    newPurchase.xml = security.xml;
    newPurchase.signature = security.hash;

    workbookData.purchases.push(newPurchase);

    // Update inventory: if rep specified, add to rep warehouse; otherwise main warehouse
    if (newPurchase.impactStock && newPurchase.rep) {
      // Add to rep's warehouse
      let repItem = workbookData.repWarehouse.find(rw => rw.rep === newPurchase.rep && rw.product === newPurchase.product);
      if (!repItem) {
        repItem = { rep: newPurchase.rep, product: newPurchase.product, quantity: 0 };
        workbookData.repWarehouse.push(repItem);
      }
      repItem.quantity += newPurchase.stockQuantity;
    } else if (newPurchase.impactStock) {
      updateInventoryPurchase(newPurchase.product, newPurchase.stockQuantity);
    }

    // Update supplier balance
    updateSupplierBalance(newPurchase.supplier, newPurchase.total, newPurchase.paid);

    // Auto-log to Treasury if payment happened
    if (newPurchase.paid > 0) {
      logTreasuryTransaction('OUT', newPurchase.paid, `مشتريات: ${newPurchase.supplier} (فاتورة #${newPurchase.id})`, `فاتورة شراء منتج ${newPurchase.product}`);
    }

    saveToLocalStorage();
    return newPurchase;
  }

  // Immutable: deletePurchase removed for security

  // -- Customers --
  function addCustomer(customer) {
    const newCust = {
      id: getNextId(workbookData.customers),
      name: customer.name || '',
      refNumber: customer.refNumber || `C${String(workbookData.customers.length + 1).padStart(4, '0')}`,
      collected: 0,
      total: 0,
      remaining: 0,
      lastDate: ''
    };
    workbookData.customers.push(newCust);
    saveToLocalStorage();
    return newCust;
  }

  function deleteCustomer(id) {
    const idx = workbookData.customers.findIndex(c => c.id === id);
    if (idx > -1) {
      workbookData.customers.splice(idx, 1);
      saveToLocalStorage();
    }
  }

  // -- Products --
  function addProduct(product) {
    const newProd = {
      id: getNextId(workbookData.products),
      name: product.name || '',
      sellPrice: product.sellPrice || 0,
      buyPrice: product.buyPrice || 0,
      unit: product.unit || 'حبة',
      piecesPerCarton: parseInt(product.piecesPerCarton) || 10,
      category: product.category || 'عام'
    };
    workbookData.products.push(newProd);

    // Also add to inventory
    workbookData.inventory.push({
      id: getNextId(workbookData.inventory),
      product: newProd.name,
      previousStock: 0,
      purchased: 0,
      withdrawn: 0,
      net: 0
    });

    saveToLocalStorage();
    return newProd;
  }

  function deleteProduct(id) {
    const idx = workbookData.products.findIndex(p => p.id === id);
    if (idx > -1) {
      workbookData.products.splice(idx, 1);
      saveToLocalStorage();
    }
  }

  function updateProduct(id, updatedData) {
    const prod = workbookData.products.find(p => p.id === id);
    if (prod) {
      if (updatedData.name !== undefined) prod.name = updatedData.name;
      if (updatedData.sellPrice !== undefined) prod.sellPrice = parseFloat(updatedData.sellPrice) || 0;
      if (updatedData.buyPrice !== undefined) prod.buyPrice = parseFloat(updatedData.buyPrice) || 0;
      if (updatedData.category !== undefined) prod.category = updatedData.category;
      if (updatedData.unit !== undefined) prod.unit = updatedData.unit;
      if (updatedData.piecesPerCarton !== undefined) prod.piecesPerCarton = parseFloat(updatedData.piecesPerCarton) || 1;
      saveToLocalStorage();
      return true;
    }
    return false;
  }

  // -- Suppliers --
  function addSupplier(supplier) {
    const newSupp = {
      id: getNextId(workbookData.suppliers),
      name: supplier.name || '',
      phone: supplier.phone || '',
      address: supplier.address || '',
      total: 0,
      paid: 0,
      remaining: 0
    };
    workbookData.suppliers.push(newSupp);
    saveToLocalStorage();
    return newSupp;
  }

  function deleteSupplier(id) {
    const idx = workbookData.suppliers.findIndex(s => s.id === id);
    if (idx > -1) {
      workbookData.suppliers.splice(idx, 1);
      saveToLocalStorage();
    }
  }

  // -- Rep Payments --
  function addRepPayment(payment) {
    const newPayment = {
      id: getNextId(workbookData.repPayments),
      date: payment.date || getTodayStr(),
      rep: payment.rep || '',
      amount: payment.amount || 0,
      paymentType: payment.paymentType || 'نقدي',
      notes: payment.notes || ''
    };
    workbookData.repPayments.push(newPayment);

    // Auto-log to Treasury (Delivered to restaurant)
    logTreasuryTransaction('IN', newPayment.amount, `توريد عهدة: ${newPayment.rep}`, newPayment.notes || 'توريد مبالغ محصلة');

    saveToLocalStorage();
    return newPayment;
  }

  function addStaff(staff) {
    const newStaff = {
      id: getNextId(workbookData.staff),
      name: staff.name,
      username: staff.username || staff.name,
      role: staff.role || 'موظف',
      phone: staff.phone || '',
      password: staff.password || '1234',
      status: 'active' // active | suspended
    };
    workbookData.staff.push(newStaff);
    saveToLocalStorage();
    return newStaff;
  }

  function toggleStaffStatus(staffId) {
    const s = workbookData.staff.find(x => x.id === staffId);
    if (s) {
      s.status = (s.status === 'active') ? 'suspended' : 'active';
      saveToLocalStorage();
    }
    return s;
  }

  function addCategory(name) {
    if (!name) return false;
    if (!workbookData.categories.includes(name)) {
      workbookData.categories.push(name);
      saveToLocalStorage();
      return true;
    }
    return false;
  }

  function getCategories() {
    return workbookData.categories || ['عام', 'بيض', 'دجاج', 'شاورما', 'أخرى'];
  }

  function addVoucher(v) {
    const serial = getNextId(workbookData.vouchers);
    const newVoucher = {
      serial: serial,
      id: serial, // for consistency
      date: v.date || getTodayStr(),
      type: v.type, // 'RECEIPT' or 'PAYMENT'
      partyType: v.partyType, // 'customer', 'supplier', 'rep', 'staff'
      party: v.party,
      amount: parseFloat(v.amount) || 0,
      notes: v.notes || '',
      user: currentUser ? currentUser.name : 'Unknown'
    };

    workbookData.vouchers.push(newVoucher);

    // Impact Treasury
    const treasuryType = v.type === 'RECEIPT' ? 'IN' : 'OUT';
    logTreasuryTransaction(treasuryType, newVoucher.amount, `سند ${v.type === 'RECEIPT' ? 'قبض' : 'صرف'} #${serial}: ${v.party}`, v.notes);

    // Impact Entity Balance
    if (v.partyType === 'customer') {
      // Add as a customer payment
      addCustomerPayment({
        date: newVoucher.date,
        customer: v.party,
        amount: newVoucher.amount,
        notes: `سند قبض #${serial}: ${v.notes}`,
        collectedBy: v.collectedBy || 'الشركة/كاشير المطعم'
      });
    } else if (v.partyType === 'supplier') {
      // Update supplier paid balance
      const supp = workbookData.suppliers.find(s => s.name === v.party);
      if (supp) {
        supp.paid = (supp.paid || 0) + newVoucher.amount;
        supp.remaining = (supp.total || 0) - supp.paid;
      }
    } else if (v.partyType === 'rep') {
      // Add as a rep payment (if Receipt, it's money collected from rep)
      if (v.type === 'RECEIPT') {
        addRepPayment({
          date: newVoucher.date,
          rep: v.party,
          amount: newVoucher.amount,
          notes: `سند قبض #${serial}: ${v.notes}`
        });
      }
    } else if (v.partyType === 'staff') {
      // Add as HR action
      addHRAction({
        date: newVoucher.date,
        name: v.party,
        amount: newVoucher.amount,
        type: v.type === 'RECEIPT' ? 'سلفة مستردة' : 'سلفة',
        notes: `سند ${v.type === 'RECEIPT' ? 'قبض' : 'صرف'} #${serial}: ${v.notes}`
      });
    }

    saveToLocalStorage();
    return newVoucher;
  }

  function getSupplierStatement(name) {
    const purchases = workbookData.purchases.filter(p => p.supplier === name);
    const vouchers = workbookData.vouchers.filter(v => v.party === name && v.partyType === 'supplier');

    // Create a chronological ledger
    let ledger = [];
    purchases.forEach(p => ledger.push({ date: p.date, type: 'PURCHASE', desc: `شراء: ${p.product}`, amount: p.total, debit: p.total, credit: 0 }));
    vouchers.forEach(v => ledger.push({ date: v.date, type: 'PAYMENT', desc: `سند صرف #${v.serial}`, amount: v.amount, debit: 0, credit: v.amount }));

    ledger.sort((a, b) => new Date(a.date) - new Date(b.date));

    let balance = 0;
    ledger = ledger.map(item => {
      balance += (item.debit - item.credit);
      return { ...item, balance };
    });

    const totalPurchases = purchases.reduce((s, p) => s + (p.total || 0), 0);
    const totalPaid = (purchases.reduce((s, p) => s + (p.paid || 0), 0)) + vouchers.reduce((s, v) => s + (v.amount || 0), 0);

    return {
      name,
      totalPurchases,
      totalPaid,
      balance,
      transactions: ledger
    };
  }

  function getRepStatement(name) {
    const sales = workbookData.sales.filter(s => s.rep === name);
    const payments = workbookData.repPayments.filter(p => p.rep === name);
    const custPayments = (workbookData.customerPayments || []).filter(p => p.collectedBy === name);

    let ledger = [];
    sales.forEach(s => ledger.push({ date: s.date, type: 'SALE', desc: `بيع: ${s.customer} - ${s.product}`, amount: s.collected, debit: s.collected, credit: 0 }));
    custPayments.forEach(p => ledger.push({ date: p.date, type: 'COLLECTION', desc: `تحصيل من عميل: ${p.customer}`, amount: p.amount, debit: p.amount, credit: 0 }));
    payments.forEach(p => ledger.push({ date: p.date, type: 'DELIVERY', desc: `توريد للمطعم`, amount: p.amount, debit: 0, credit: p.amount }));

    ledger.sort((a, b) => new Date(a.date) - new Date(b.date));

    let balance = 0;
    ledger = ledger.map(item => {
      balance += (item.debit - item.credit);
      return { ...item, balance };
    });

    return {
      name,
      balance,
      transactions: ledger
    };
  }

  function getStaffStatement(name) {
    const actions = (workbookData.hr || []).filter(h => h.name === name);
    let ledger = [];
    actions.forEach(a => {
      const isDebit = a.type === 'سلفة' || a.type === 'خصم';
      ledger.push({
        date: a.date,
        type: a.type,
        desc: a.notes || a.type,
        debit: isDebit ? a.amount : 0,
        credit: isDebit ? 0 : a.amount
      });
    });

    ledger.sort((a, b) => new Date(a.date) - new Date(b.date));

    let balance = 0;
    ledger = ledger.map(item => {
      balance += (item.credit - item.debit); // Balance for staff is usually what we owe them
      return { ...item, balance };
    });

    return {
      name,
      balance,
      transactions: ledger
    };
  }

  function addSaleReturn(originalId, returnQty, returnNotes) {
    const original = workbookData.sales.find(s => s.id === originalId);
    if (!original) return { success: false, error: 'الفاتورة الأصلية غير موجودة' };

    if (returnQty > original.netQuantity) {
      return { success: false, error: 'الكمية المرتجعة أكبر من الكمية المباعة' };
    }

    // Create a reverse entry
    const returnSale = {
      ...original,
      id: getNextId(workbookData.sales),
      date: getTodayStr(),
      quantity: 0,
      returned: returnQty,
      netQuantity: -returnQty,
      stockQuantity: -(returnQty * (original.piecesPerCarton || 1)),
      total: -(returnQty * original.price * (original.applyVat ? 0.85 : 1)),
      collected: 0, // Usually return doesn't return cash immediately in the log
      balance: -(returnQty * original.price * (original.applyVat ? 0.85 : 1)),
      notes: `مرتجع لفاتورة #${originalId}: ${returnNotes || ''}`,
      parentId: originalId
    };

    // Update stock if original impacted stock
    if (original.impactStock) {
      if (original.rep) {
        const repItem = workbookData.repWarehouse.find(rw => rw.rep === original.rep && rw.product === original.product);
        if (repItem) repItem.quantity += Math.abs(returnSale.stockQuantity);
      } else {
        updateInventoryWithdrawal(original.product, returnSale.stockQuantity); // Negative withdrawn = add to stock
      }
    }

    workbookData.sales.push(returnSale);
    updateCustomerBalance(original.customer, returnSale.total, 0);
    saveToLocalStorage();
    return { success: true, sale: returnSale };
  }

  function addPurchaseReturn(originalId, returnQty, returnNotes) {
    const original = workbookData.purchases.find(p => p.id === originalId);
    if (!original) return { success: false, error: 'الفاتورة الأصلية غير موجودة' };

    if (returnQty > original.quantity) {
      return { success: false, error: 'الكمية المرتجعة أكبر من الكمية المشتراة' };
    }

    const returnPurchase = {
      ...original,
      id: getNextId(workbookData.purchases),
      date: getTodayStr(),
      quantity: -returnQty,
      stockQuantity: -(returnQty * (original.piecesPerCarton || 1)),
      total: -(returnQty * original.price * (original.applyVat ? 0.85 : 1)),
      paid: 0,
      dueToSupplier: -(returnQty * original.price * (original.applyVat ? 0.85 : 1)),
      notes: `مرتجع مشتريات لفاتورة #${originalId}: ${returnNotes || ''}`,
      parentId: originalId
    };

    if (original.impactStock) {
      updateInventoryWithdrawal(original.product, -returnPurchase.stockQuantity); // Adding to withrdawn = minus from stock. Wait.
      // updateInventoryWithdrawal(prod, qty) adds to withdrawn. 
      // For purchase return, we want to MINUS from stock.
      // Purchase originally added to inventory.purchased.
      // So return should minus from inventory.purchased.
      const inv = workbookData.inventory.find(i => i.product === original.product);
      if (inv) {
        inv.purchased -= Math.abs(returnPurchase.stockQuantity);
        inv.net = (inv.previousStock + inv.purchased) - inv.withdrawn;
      }
    }

    workbookData.purchases.push(returnPurchase);
    // Update supplier balance (assuming remaining balance accounts for this)
    const supp = workbookData.suppliers.find(s => s.name === original.supplier);
    if (supp) {
      supp.total += returnPurchase.total;
      supp.remaining = supp.total - supp.paid;
    }

    saveToLocalStorage();
    return { success: true, purchase: returnPurchase };
  }

  function updateInventoryManual(productName, type, amount, notes) {
    const inv = workbookData.inventory.find(i => i.product === productName);
    if (!inv) return false;

    if (type === 'set') {
      inv.previousStock = amount;
      inv.purchased = 0;
      inv.withdrawn = 0;
    } else if (type === 'add') {
      inv.previousStock += amount;
    } else if (type === 'sub') {
      inv.previousStock -= amount;
    }

    inv.net = (inv.previousStock + inv.purchased) - inv.withdrawn;
    saveToLocalStorage();
    return true;
  }

  function removeStaff(name) {
    workbookData.staff = workbookData.staff.filter(s => s.name !== name);
    saveToLocalStorage();
  }

  function addHRAction(action) {
    if (!action.date) action.date = getTodayStr();
    workbookData.hr.push(action);

    // Auto-log to Treasury (Expense)
    if (action.type === 'راتب' || action.type === 'سلفة' || action.type === 'مكافأة' || action.type === 'مصاريف أخرى') {
      logTreasuryTransaction('OUT', action.amount, `موارد بشرية: ${action.name} (${action.type})`, action.notes);
    }

    saveToLocalStorage();
  }

  function addCustomerPayment(payment) {
    if (!payment.date) payment.date = getTodayStr();
    payment.id = getNextId(workbookData.customerPayments);
    workbookData.customerPayments.push(payment);

    // Update customer balance (reduce their remaining debt)
    let cust = workbookData.customers.find(c => c.name === payment.customer);
    if (cust) {
      cust.collected += (payment.amount || 0);
      cust.remaining = cust.total - cust.collected;
      cust.lastDate = payment.date;
    }

    // Auto-log to Treasury if it reached restaurant/company OR force it to record as paid
    const refText = payment.invoiceRef ? ` - فاتورة: ${payment.invoiceRef}` : '';
    if (payment.collectedBy === 'الشركة/كاشير المطعم') {
      logTreasuryTransaction('IN', payment.amount, `سداد عميل: ${payment.customer}${refText}`, payment.notes);
      generateVoucher('RECEIPT', payment.customer, payment.amount, `تحصيل دفعة${refText} - ${payment.notes || ''}`);
    } else {
      // If collected by rep, it increases rep's custody. To ensure it hits treasury eventually,
      // it waits for Rep Payment. But we still need the voucher.
      generateVoucher('RECEIPT', payment.customer, payment.amount, `محصل ع/ط المندوب: ${payment.collectedBy}${refText} - ${payment.notes || ''}`);
    }

    saveToLocalStorage();
  }

  // --- Vouchers (السندات) ---
  function generateVoucher(type, party, amount, notes) {
    const v = {
      id: getNextId(workbookData.vouchers),
      serial: `SOND-${getTodayStr().replace(/-/g, '')}-${String(workbookData.vouchers.length + 1).padStart(4, '0')}`,
      date: getTodayStr(),
      type: type, // 'RECEIPT' (قبض) or 'PAYMENT' (صرف)
      party: party,
      amount: amount,
      notes: notes || '',
      user: currentUser ? currentUser.name : 'System'
    };
    workbookData.vouchers.push(v);
    saveToLocalStorage();
    return v;
  }

  function getVouchers() { return workbookData.vouchers; }

  // --- Inbound/Outbound Ledger (كشف الوارد والصادر) ---
  function getLedger() {
    // Combine Treasury movements, Customer Payments, Rep Payments into a unified stream
    const events = [];

    workbookData.treasury.forEach(t => events.push({ ...t, category: 'خزينة' }));
    workbookData.customerPayments.forEach(p => events.push({
      id: p.id, date: p.date, type: 'IN', amount: p.amount, reference: `عميل: ${p.customer}`, notes: p.notes, category: 'سداد عملاء'
    }));
    workbookData.repPayments.forEach(p => events.push({
      id: p.id, date: p.date, type: 'IN', amount: p.amount, reference: `عهد: ${p.rep}`, notes: p.notes, category: 'توريد مناديب'
    }));

    return events.sort((a, b) => new Date(b.date) - new Date(a.date));
  }

  // --- Inventory Ledger (سجل حركة المخزون) ---
  function getInventoryLedger(productName) {
    const events = [];

    // Sales (Withdrawals/Returns)
    workbookData.sales.forEach(s => {
      if (s.impactStock !== false && (!productName || s.product === productName)) {
        events.push({
          date: s.date,
          id: s.id,
          product: s.product,
          type: s.returned > 0 ? 'IN' : 'OUT',
          category: s.returned > 0 ? 'مرتجع مبيعات' : 'مبيعات',
          amount: s.stockQuantity || s.netQuantity || 0,
          party: s.customer,
          reference: `فاتورة بيع #${s.id}`,
          location: s.rep ? `مستودع مندوب: ${s.rep}` : 'المخزن الرئيسي'
        });
      }
    });

    // Purchases (Additions/Returns)
    workbookData.purchases.forEach(p => {
      if (p.impactStock !== false && (!productName || p.product === productName)) {
        events.push({
          date: p.date,
          id: p.id,
          product: p.product,
          type: p.quantity < 0 ? 'OUT' : 'IN',
          category: p.quantity < 0 ? 'مرتجع مشتريات' : 'مشتريات',
          amount: Math.abs(p.stockQuantity || p.quantity || 0),
          party: p.supplier,
          reference: `فاتورة شراء #${p.id}`,
          location: p.rep ? `مستودع مندوب: ${p.rep}` : 'المخزن الرئيسي'
        });
      }
    });

    // Warehouse Transfers
    if (workbookData.warehouseTransfers) {
      workbookData.warehouseTransfers.forEach(t => {
        if (!productName || t.product === productName) {
          events.push({
            date: t.date,
            id: t.id,
            product: t.product,
            // Transfer OUT of main, IN to rep
            type: t.type === 'OUT_TO_REP' ? 'OUT' : 'IN',
            category: t.type === 'OUT_TO_REP' ? 'سحب لمندوب' : 'إرجاع من مندوب',
            amount: t.quantity,
            party: t.rep,
            reference: `حركة مخزن #${t.id || 'N/A'}`,
            location: 'المخزن الرئيسي'
          });
        }
      });
    }

    // Sort by date desc
    events.sort((a, b) => {
      const d1 = new Date(a.date).getTime();
      const d2 = new Date(b.date).getTime();
      if (d1 !== d2) return d2 - d1;
      return (b.id || 0) - (a.id || 0); // fallback
    });
    return events;
  }

  // === Inventory Updates ===
  function updateInventoryWithdrawal(productName, qty) {
    const inv = workbookData.inventory.find(i => i.product === productName);
    if (inv) {
      inv.withdrawn += qty;
      inv.net = inv.previousStock + inv.purchased - inv.withdrawn;
    }
  }

  function updateInventoryPurchase(productName, qty) {
    const inv = workbookData.inventory.find(i => i.product === productName);
    if (inv) {
      inv.purchased += qty;
      inv.net = inv.previousStock + inv.purchased - inv.withdrawn;
    }
  }

  // === Customer Balance Updates ===
  function updateCustomerBalance(customerName, saleTotal, collected) {
    let cust = workbookData.customers.find(c => c.name === customerName);
    if (cust) {
      cust.total += saleTotal;
      cust.collected += collected;
      cust.remaining = cust.total - cust.collected;
      cust.lastDate = getTodayStr();
    }
  }

  // === Supplier Balance Updates ===
  function updateSupplierBalance(supplierName, purchaseTotal, paid) {
    let supp = workbookData.suppliers.find(s => s.name === supplierName);
    if (supp) {
      supp.total += purchaseTotal;
      supp.paid += paid;
      supp.remaining = supp.total - supp.paid;
    }
  }

  // === Statistics/Analytics ===
  function getStats() {
    const totalSales = workbookData.sales.reduce((sum, s) => sum + (s.total || 0), 0);
    const totalPurchases = workbookData.purchases.reduce((sum, p) => sum + (p.total || 0), 0);
    const totalCollected = workbookData.sales.reduce((sum, s) => sum + (s.collected || 0), 0);
    const totalDue = workbookData.customers.reduce((sum, c) => sum + (c.remaining || 0), 0);
    const profit = totalSales - totalPurchases;

    // Top products by sales
    const productSales = {};
    workbookData.sales.forEach(s => {
      if (!productSales[s.product]) productSales[s.product] = { qty: 0, total: 0 };
      productSales[s.product].qty += s.netQuantity || 0;
      productSales[s.product].total += s.total || 0;
    });

    const topProducts = Object.entries(productSales)
      .map(([name, data]) => ({ name, ...data }))
      .sort((a, b) => b.total - a.total)
      .slice(0, 10);

    // Low stock items
    const lowStock = workbookData.inventory
      .filter(i => i.net < 10 && i.net >= 0)
      .sort((a, b) => a.net - b.net);

    // Top customers by debt
    const topDebtors = workbookData.customers
      .filter(c => c.remaining > 0)
      .sort((a, b) => b.remaining - a.remaining)
      .slice(0, 10);

    // Sales by date
    const salesByDate = {};
    workbookData.sales.forEach(s => {
      if (!salesByDate[s.date]) salesByDate[s.date] = 0;
      salesByDate[s.date] += s.total || 0;
    });

    // Purchases by date
    const purchasesByDate = {};
    workbookData.purchases.forEach(p => {
      if (!purchasesByDate[p.date]) purchasesByDate[p.date] = 0;
      purchasesByDate[p.date] += p.total || 0;
    });

    // Rep performance
    const repPerformance = {};
    workbookData.sales.forEach(s => {
      if (!repPerformance[s.rep]) repPerformance[s.rep] = { sales: 0, count: 0 };
      repPerformance[s.rep].sales += s.total || 0;
      repPerformance[s.rep].count++;
    });

    // === Advanced financial calculations ===
    // Real profit using product cost prices
    let realProfit = 0;
    workbookData.sales.forEach(s => {
      const prod = workbookData.products.find(p => p.name === s.product);
      const costPerUnit = prod ? prod.buyPrice : 0;
      const cost = (s.netQuantity || 0) * costPerUnit;
      realProfit += ((s.total || 0) - cost);
    });

    // Customer payments total
    const totalCustomerPayments = (workbookData.customerPayments || []).reduce((s, p) => s + (p.amount || 0), 0);

    // Rep payments (delivered to restaurant)
    const totalRepDelivered = workbookData.repPayments.reduce((s, p) => s + (p.amount || 0), 0);

    // Direct customer payments to restaurant (not through a rep)
    const directCustPayments = (workbookData.customerPayments || [])
      .filter(p => p.collectedBy === 'الشركة/كاشير المطعم')
      .reduce((s, p) => s + (p.amount || 0), 0);

    // Cash in Box = rep deliveries + direct customer payments to restaurant
    const cashInBox = totalRepDelivered + directCustPayments;

    // Customer payments collected by reps (not directly to restaurant)
    const repCustPayments = (workbookData.customerPayments || [])
      .filter(p => p.collectedBy !== 'الشركة/كاشير المطعم')
      .reduce((s, p) => s + (p.amount || 0), 0);

    // Total Rep Custody = what they collected from sales + from customer payments - what they delivered
    const totalRepCustody = totalCollected + repCustPayments - totalRepDelivered;

    // HR costs
    const totalHRCosts = (workbookData.hr || []).reduce((s, h) => s + (h.amount || 0), 0);

    return {
      totalSales,
      totalPurchases,
      totalCollected,
      totalDue,
      profit,
      realProfit,
      cashInBox,
      totalRepCustody,
      totalRepDelivered,
      totalCustomerPayments,
      totalHRCosts,
      salesCount: workbookData.sales.length,
      purchasesCount: workbookData.purchases.length,
      customersCount: workbookData.customers.length,
      productsCount: workbookData.products.length,
      topProducts,
      lowStock,
      topDebtors,
      salesByDate,
      purchasesByDate,
      repPerformance
    };
  }

  // === Account Statement ===
  function getCustomerStatement(customerName) {
    const customer = workbookData.customers.find(c => c.name === customerName);
    const salesRecords = workbookData.sales.filter(s => s.customer === customerName);
    const custPayments = (workbookData.customerPayments || []).filter(p => p.customer === customerName);

    let transactions = [];
    let runningBalance = 0;

    salesRecords.forEach(s => {
      transactions.push({
        date: s.date,
        type: 'مبيعات',
        product: s.product,
        netQuantity: s.netQuantity,
        price: s.price,
        total: s.total,
        collected: s.collected,
        balance: 0
      });
    });

    custPayments.forEach(p => {
      transactions.push({
        date: p.date,
        type: 'سداد دفعة',
        product: 'سداد نقدي' + (p.collectedBy ? ' (محصل: ' + p.collectedBy + ')' : ''),
        netQuantity: 0,
        price: 0,
        total: 0,
        collected: p.amount,
        balance: 0
      });
    });

    // Chronological order
    transactions.sort((a, b) => new Date(a.date) - new Date(b.date));

    transactions.forEach(t => {
      if (t.type === 'مبيعات') {
        runningBalance += (t.total - t.collected);
      } else {
        runningBalance -= t.collected;
      }
      t.balance = runningBalance;
    });

    const totalSales = salesRecords.reduce((s, r) => s + (r.total || 0), 0);
    const totalCollectedFromSales = salesRecords.reduce((s, r) => s + (r.collected || 0), 0);
    const totalPayments = custPayments.reduce((s, p) => s + (p.amount || 0), 0);

    return {
      customer: customer || { name: customerName },
      transactions: transactions,
      totalSales: totalSales,
      totalCollected: totalCollectedFromSales + totalPayments,
      balance: runningBalance
    };
  }

  function getSupplierStatement(supplierName) {
    const supplier = workbookData.suppliers.find(s => s.name === supplierName);
    const purchaseRecords = workbookData.purchases.filter(p => p.supplier === supplierName);
    const vouchers = workbookData.vouchers.filter(v => v.party === supplierName && v.partyType === 'supplier');

    let transactions = [];
    let runningBalance = 0;

    purchaseRecords.forEach(p => {
      transactions.push({
        date: p.date,
        desc: `فاتورة شراء #${p.id} - ${p.product}`,
        debit: p.total, // Debt to supplier
        credit: p.paid, // Portion paid during purchase
        balance: 0
      });
    });

    vouchers.forEach(v => {
      transactions.push({
        date: v.date,
        desc: `${v.type === 'PAYMENT' ? 'سند صرف' : 'سند قبض'} #${v.serial} - ${v.notes}`,
        debit: v.type === 'RECEIPT' ? v.amount : 0, // Rare: Receipt from supplier?
        credit: v.type === 'PAYMENT' ? v.amount : 0, // We paid supplier
        balance: 0
      });
    });

    transactions.sort((a, b) => new Date(a.date) - new Date(b.date));

    transactions.forEach(t => {
      runningBalance += (t.debit - t.credit);
      t.balance = runningBalance;
    });

    return {
      supplier: supplier || { name: supplierName },
      transactions: transactions,
      totalPurchases: transactions.reduce((s, t) => s + t.debit, 0),
      totalPaid: transactions.reduce((s, t) => s + t.credit, 0),
      balance: runningBalance
    };
  }

  function getRepStatement(repName) {
    const sales = workbookData.sales.filter(s => s.rep === repName);
    const vouchers = workbookData.vouchers.filter(v => v.party === repName && v.partyType === 'rep');
    const custPayments = (workbookData.customerPayments || []).filter(p => p.collectedBy === repName);

    let transactions = [];
    let runningBalance = 0;

    // Rep is "Debtor" for what they collect, "Creditor" for what they deliver
    sales.forEach(s => {
      transactions.push({
        date: s.date,
        desc: `تحصيل فاتورة بيع #${s.id} - ${s.customer}`,
        debit: s.collected, // They hold this cash
        credit: 0,
        balance: 0
      });
    });

    custPayments.forEach(p => {
      transactions.push({
        date: p.date,
        desc: `تحصيل دفعة عميل: ${p.customer}`,
        debit: p.amount, // They hold this cash
        credit: 0,
        balance: 0
      });
    });

    vouchers.forEach(v => {
      transactions.push({
        date: v.date,
        desc: `${v.type === 'PAYMENT' ? 'صرف للمندوب' : 'توريد من المندوب'} #${v.serial} - ${v.notes}`,
        debit: v.type === 'PAYMENT' ? v.amount : 0, // We gave them cash
        credit: v.type === 'RECEIPT' ? v.amount : 0, // They gave us cash
        balance: 0
      });
    });

    transactions.sort((a, b) => new Date(a.date) - new Date(b.date));
    transactions.forEach(t => {
      runningBalance += (t.debit - t.credit);
      t.balance = runningBalance;
    });

    return { transactions, balance: runningBalance };
  }

  function getStaffStatement(staffName) {
    const hr = (workbookData.hr || []).filter(h => h.name === staffName);
    const vouchers = workbookData.vouchers.filter(v => v.party === staffName && v.partyType === 'staff');

    let transactions = [];
    let runningBalance = 0;

    // Staff is "Creditor" for what they earn (Salary), "Debtor" for what they receive (Voucher/Advance)
    hr.forEach(h => {
      transactions.push({
        date: h.date,
        desc: `${h.type} - ${h.notes}`,
        debit: h.type === 'سلفة' ? h.amount : 0,
        credit: (h.type === 'راتب' || h.type === 'مكافأة') ? h.amount : 0,
        balance: 0
      });
    });

    vouchers.forEach(v => {
      transactions.push({
        date: v.date,
        desc: `${v.type === 'PAYMENT' ? 'صرف مالي' : 'قبض مالي'} #${v.serial} - ${v.notes}`,
        debit: v.type === 'PAYMENT' ? v.amount : 0,
        credit: v.type === 'RECEIPT' ? v.amount : 0,
        balance: 0
      });
    });

    transactions.sort((a, b) => new Date(a.date) - new Date(b.date));
    transactions.forEach(t => {
      runningBalance += (t.credit - t.debit); // Positive balance means we owe them
      t.balance = runningBalance;
    });

    return { transactions, balance: runningBalance };
  }

  // === Vouchers & Returns (NEW) ===
  function addVoucher(v) {
    const serial = `V-${Date.now().toString().slice(-6)}`;
    const voucher = {
      serial,
      date: v.date || getTodayStr(),
      type: v.type, // RECEIPT or PAYMENT
      partyType: v.partyType, // customer, supplier, rep, staff
      party: v.party,
      amount: parseNum(v.amount),
      notes: v.notes || '',
      user: 'المدير'
    };

    workbookData.vouchers.push(voucher);

    // Impact entity balances
    if (v.partyType === 'customer') {
      const cust = workbookData.customers.find(c => c.name === v.party);
      if (cust) {
        if (v.type === 'RECEIPT') cust.collected += voucher.amount;
        else cust.collected -= voucher.amount; // Rare
        cust.remaining = cust.total - cust.collected;
      }
    } else if (v.partyType === 'supplier') {
      const supp = workbookData.suppliers.find(s => s.name === v.party);
      if (supp) {
        if (v.type === 'PAYMENT') supp.paid += voucher.amount;
        else supp.paid -= voucher.amount;
        supp.remaining = supp.total - supp.paid;
      }
    }

    // Log to Treasury
    logTreasuryTransaction(
      v.type === 'RECEIPT' ? 'IN' : 'OUT',
      voucher.amount,
      `سند ${v.type === 'RECEIPT' ? 'قبض' : 'صرف'} #${serial}`,
      `${v.party} - ${v.notes}`
    );

    saveToLocalStorage();
    return voucher;
  }

  function addSaleReturn(saleId, qty, notes) {
    const sale = workbookData.sales.find(s => s.id === saleId);
    if (!sale) return { success: false, error: 'الفاتورة غير موجودة' };

    const returnQty = parseNum(qty);
    if (returnQty > (sale.netQuantity || sale.quantity)) {
      return { success: false, error: 'الكمية المرتجعة أكبر من المتاح في الفاتورة' };
    }

    // Update original record or add a reverse record? 
    // Best practice: Add a "Return" entry to maintain audit trail.
    const price = sale.price;
    const totalReturnVal = returnQty * price;

    // Update original
    sale.returned = (sale.returned || 0) + returnQty;
    sale.netQuantity = sale.quantity - sale.returned;
    sale.total = sale.netQuantity * price;
    // Note: collected remains as is, but debt (remaining) will decrease

    // Update inventory/stock
    updateInventoryWithdrawal(sale.product, -returnQty); // Negative withdrawal = return to stock

    // Update customer balance
    const cust = workbookData.customers.find(c => c.name === sale.customer);
    if (cust) {
      cust.total -= totalReturnVal;
      cust.remaining = cust.total - cust.collected;
    }

    // Log treasury impact if we refunded cash
    // For now, assume it's a balance correction. If cash, user uses Voucher.

    saveToLocalStorage();
    return { success: true };
  }

  function addPurchaseReturn(purchaseId, qty, notes) {
    const purch = workbookData.purchases.find(p => p.id === purchaseId);
    if (!purch) return { success: false, error: 'الفاتورة غير موجودة' };

    const returnQty = parseNum(qty);
    if (returnQty > (purch.quantity)) {
      return { success: false, error: 'الكمية المرتجعة أكبر من المشتراة' };
    }

    const price = purch.price;
    const totalReturnVal = returnQty * price;

    // Update original
    purch.quantity -= returnQty;
    purch.total = purch.quantity * price;

    // Update inventory/stock
    updateInventoryPurchase(purch.product, -returnQty); // Negative purchase = out from stock

    // Update supplier balance
    const supp = workbookData.suppliers.find(s => s.name === purch.supplier);
    if (supp) {
      supp.total -= totalReturnVal;
      supp.remaining = supp.total - supp.paid;
    }

    saveToLocalStorage();
    return { success: true };
  }

  // === Helpers ===
  function parseNum(val) {
    if (val === null || val === undefined || val === '') return 0;
    const str = String(val).replace(/[,$]/g, '').trim();
    const num = parseFloat(str);
    return isNaN(num) ? 0 : num;
  }

  function formatDateFromExcel(val) {
    if (!val) return '';
    if (val instanceof Date) {
      return val.toISOString().split('T')[0];
    }
    // Handle various date strings
    const str = String(val).trim();
    if (str.includes('-Apr') || str.match(/\d+-[A-Za-z]+/)) {
      // Format like "14-Apr"
      const year = new Date().getFullYear();
      const d = new Date(`${str}-${year}`);
      if (!isNaN(d.getTime())) {
        return d.toISOString().split('T')[0];
      }
    }
    return str;
  }

  function getTodayStr() {
    return new Date().toISOString().split('T')[0];
  }

  function getNextId(arr) {
    if (arr.length === 0) return 1;
    return Math.max(...arr.map(item => item.id || 0)) + 1;
  }

  function formatCurrency(val) {
    return new Intl.NumberFormat('ar-SA', {
      style: 'currency',
      currency: 'SAR',
      minimumFractionDigits: 2
    }).format(val || 0);
  }

  function formatNumber(val) {
    return new Intl.NumberFormat('ar-SA').format(val || 0);
  }

  // === Data Getters ===
  function getData() {
    return workbookData;
  }

  function getSales() { return workbookData.sales; }
  function getPurchases() { return workbookData.purchases; }
  function getInventory() { return workbookData.inventory; }
  function getCustomers(filterByRep) {
    if (filterByRep && filterByRep !== 'الشركة/كاشير المطعم') {
      // Filter customers who have sales with this rep
      const repCustomerNames = new Set(
        workbookData.sales.filter(s => s.rep === filterByRep).map(s => s.customer)
      );
      return workbookData.customers.filter(c => repCustomerNames.has(c.name));
    }
    return workbookData.customers;
  }
  function getReps() {
    const staffReps = (workbookData.staff || []).filter(s => s.role === 'مندوب').map(s => s.name);
    return staffReps.length > 0 ? staffReps : DEFAULT_REPS;
  }
  function getRepsList() {
    const staffReps = (workbookData.staff || []).filter(s => s.role === 'مندوب').map(s => s.name);
    return staffReps.length > 0 ? staffReps : DEFAULT_REPS;
  }
  function getProducts() { return workbookData.products; }
  function getSuppliers() { return workbookData.suppliers; }
  function getRepPayments() { return workbookData.repPayments; }
  function getIsLoaded() { return isLoaded; }
  function getStaff() { return workbookData.staff || []; }
  function getHR() { return workbookData.hr || []; }
  function getCustomerPayments() { return workbookData.customerPayments || []; }

  // === Rep Accounting Summary ===
  function getRepAccountingSummary(repName) {
    const sales = workbookData.sales.filter(s => s.rep === repName);
    const purchases = workbookData.purchases.filter(p => p.rep === repName);
    const payments = workbookData.repPayments.filter(p => p.rep === repName);
    const custPayments = (workbookData.customerPayments || []).filter(p => p.collectedBy === repName);
    const hrActions = (workbookData.hr || []).filter(h => h.name === repName);
    const products = workbookData.products;

    // Sales totals
    const totalSales = sales.reduce((s, sale) => s + (sale.total || 0), 0);
    const totalCollectedFromSales = sales.reduce((s, sale) => s + (sale.collected || 0), 0);

    // Purchases totals
    const totalPurchases = purchases.reduce((s, p) => s + (p.total || 0), 0);

    // Real profit using cost price
    let totalProfit = 0;
    sales.forEach(sale => {
      const prod = products.find(p => p.name === sale.product);
      const costPerUnit = prod ? prod.buyPrice : 0;
      const cost = (sale.netQuantity || 0) * costPerUnit;
      totalProfit += ((sale.total || 0) - cost);
    });

    // Payments delivered to restaurant
    const totalDelivered = payments.reduce((s, p) => s + (p.amount || 0), 0);

    // Customer payments collected by this rep
    const totalCustPaymentsCollected = custPayments.reduce((s, p) => s + (p.amount || 0), 0);

    // Cash custody = (collected from sales + collected from customer payments) - delivered to restaurant
    const cashCustody = totalCollectedFromSales + totalCustPaymentsCollected - totalDelivered;

    // HR deductions (salaries, advances)
    let totalSalaries = 0;
    let totalAdvances = 0;
    hrActions.forEach(h => {
      if (h.type === 'راتب' || h.type === 'مكافأة') totalSalaries += (h.amount || 0);
      if (h.type === 'سلفة') totalAdvances += (h.amount || 0);
    });

    return {
      name: repName,
      salesCount: sales.length,
      totalSales,
      totalCollectedFromSales,
      totalPurchases,
      totalProfit,
      totalDelivered,
      totalCustPaymentsCollected,
      cashCustody,
      totalSalaries,
      totalAdvances,
      netBalance: cashCustody - totalAdvances
    };
  }

  // Save to localStorage
  function saveToLocalStorage() {
    try {
      localStorage.setItem('invoiceSystemData', JSON.stringify(workbookData));
    } catch (e) {
      console.error('Error saving to localStorage:', e);
    }
  }

  // Load from localStorage
  function loadFromLocalStorage() {
    try {
      const data = localStorage.getItem('invoiceSystemData');
      if (data) {
        workbookData = JSON.parse(data);
        if (!workbookData.treasury) workbookData.treasury = [];
        if (!workbookData.vouchers) workbookData.vouchers = [];
        if (!workbookData.staff) workbookData.staff = [];
        if (!workbookData.repWarehouse) workbookData.repWarehouse = [];
        if (!workbookData.warehouseTransfers) workbookData.warehouseTransfers = [];
        if (!workbookData.categories) workbookData.categories = ['عام', 'بيض', 'دجاج', 'شاورما', 'أخرى'];

        // Ensure admin always exists in staff
        const hasAdmin = workbookData.staff.find(s => s.name === 'admin');
        if (!hasAdmin) {
          workbookData.staff.push({ id: 999, name: 'admin', role: 'مشرف', password: 'admin' });
        }
        isLoaded = true;

        // Restore session if exists
        const session = sessionStorage.getItem('currentUser');
        if (session) {
          currentUser = JSON.parse(session);
        }
        return true;
      }
    } catch (e) {
      console.error('Error loading from localStorage:', e);
    }
    return false;
  }

  // === Authentication ===
  function login(name, password) {
    // Support login by username or name
    const user = workbookData.staff.find(s =>
      (s.username === name || s.name === name) && s.password === password
    );
    if (!user) return false;
    if (user.status === 'suspended') return 'suspended';
    currentUser = user;
    sessionStorage.setItem('currentUser', JSON.stringify(user));
    return true;
  }

  function logout() {
    currentUser = null;
    sessionStorage.removeItem('currentUser');
  }

  function getCurrentUser() {
    return currentUser;
  }

  // === Treasury (الخزينة) ===
  function logTreasuryTransaction(type, amount, reference, notes) {
    const trx = {
      id: getNextId(workbookData.treasury),
      date: getTodayStr(),
      type: type, // 'IN' or 'OUT'
      amount: amount,
      reference: reference || '',
      notes: notes || '',
      user: currentUser ? currentUser.name : 'Unknown'
    };
    workbookData.treasury.push(trx);
    saveToLocalStorage();
    return trx;
  }

  function getTreasury() {
    return workbookData.treasury;
  }

  // === Advanced Reports with ExcelJS ===
  async function exportAdvancedReports(type, query) {
    if (!window.ExcelJS) throw new Error('مكتبة ExcelJS غير محملة');

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'نظام الفواتير المتقدم';
    const companyTitle = 'مشروع الفروج الوطني الرائع';
    const companySubtitle = 'نظام إدارة المبيعات والمخزون المتكامل';
    const logoHex = 'FF1A73E8'; // Blue color

    async function saveWorkbook(name) {
      try {
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = name;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
      } catch (err) {
        console.error('Download failed', err);
        alert('فشل تصدير الملف: ' + err.message);
      }
    }

    function setupSheet(ws, title, columns) {
      ws.views = [{ rightToLeft: true }];

      // Header branding
      ws.mergeCells(`A1:${String.fromCharCode(64 + columns.length)}1`);
      const headCell = ws.getCell('A1');
      headCell.value = companyTitle;
      headCell.font = { name: 'Cairo', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
      headCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: logoHex } };
      headCell.alignment = { vertical: 'middle', horizontal: 'center' };
      ws.getRow(1).height = 35;

      ws.mergeCells(`A2:${String.fromCharCode(64 + columns.length)}2`);
      const subCell = ws.getCell('A2');
      subCell.value = title + ' - تاريخ التقرير: ' + getTodayStr();
      subCell.font = { name: 'Cairo', size: 11, bold: true, color: { argb: 'FF1A2035' } };
      subCell.alignment = { vertical: 'middle', horizontal: 'center' };
      ws.getRow(2).height = 25;

      ws.getRow(3).height = 10; // Spacer

      // Columns
      ws.columns = columns;

      const headerRow = ws.getRow(4);
      headerRow.height = 24;
      headerRow.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Cairo', size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF42A5F5' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      });

      return ws;
    }

    // Helper to format cells gracefully
    function styleRow(row) {
      row.height = 20;
      row.eachCell(c => {
        c.font = { name: 'Cairo', size: 10 };
        c.alignment = { vertical: 'middle', horizontal: 'center' };
        c.border = { top: { style: 'hair' }, left: { style: 'hair' }, bottom: { style: 'hair' }, right: { style: 'hair' } };
      });
    }

    if (type === 'customer') {
      const stmt = getCustomerStatement(query);
      const ws = workbook.addWorksheet('كشف حساب');

      const columns = [
        { header: 'التاريخ', key: 'date', width: 15 },
        { header: 'المنتج', key: 'product', width: 25 },
        { header: 'الكمية', key: 'qty', width: 10 },
        { header: 'السعر', key: 'price', width: 15 },
        { header: 'الإجمالي', key: 'total', width: 15 },
        { header: 'التحصيل', key: 'collected', width: 15 },
        { header: 'الرصيد المتبقي', key: 'balance', width: 15 }
      ];

      setupSheet(ws, `كشف حساب العميل: ${query}`, columns);

      stmt.transactions.forEach(t => {
        const row = ws.addRow({
          date: t.date, product: t.product, qty: t.netQuantity,
          price: t.price, total: t.total, collected: t.collected, balance: t.balance
        });
        styleRow(row);
        row.getCell('price').numFmt = '#,##0.00';
        row.getCell('total').numFmt = '#,##0.00';
        row.getCell('collected').numFmt = '#,##0.00';
        row.getCell('balance').numFmt = '#,##0.00';
      });

      ws.addRow([]);
      const summaryRow = ws.addRow({ date: 'الإجمالي الكلي', total: stmt.totalSales, collected: stmt.totalCollected, balance: stmt.balance });
      styleRow(summaryRow);
      summaryRow.font = { bold: true, name: 'Cairo', size: 10 };
      summaryRow.getCell('total').numFmt = '#,##0.00';
      summaryRow.getCell('collected').numFmt = '#,##0.00';
      summaryRow.getCell('balance').numFmt = '#,##0.00';

      await saveWorkbook(`كشف_حساب_${query}.xlsx`);
    }

    else if (type === 'supplier') {
      const stmt = getSupplierStatement(query);
      const ws = workbook.addWorksheet('كشف المورد');

      const columns = [
        { header: 'التاريخ', key: 'date', width: 15 },
        { header: 'المنتج', key: 'product', width: 25 },
        { header: 'الكمية', key: 'qty', width: 10 },
        { header: 'السعر', key: 'price', width: 15 },
        { header: 'الإجمالي', key: 'total', width: 15 },
        { header: 'المدفوع', key: 'paid', width: 15 },
        { header: 'المتبقي', key: 'balance', width: 15 }
      ];

      setupSheet(ws, `كشف حساب المورد: ${query}`, columns);

      let runBalance = 0;
      stmt.transactions.forEach(t => {
        runBalance += (t.total - t.paid);
        const row = ws.addRow({
          date: t.date, product: t.product, qty: t.quantity,
          price: t.price, total: t.total, paid: t.paid, balance: runBalance
        });
        styleRow(row);
        row.getCell('price').numFmt = '#,##0.00';
        row.getCell('total').numFmt = '#,##0.00';
        row.getCell('paid').numFmt = '#,##0.00';
        row.getCell('balance').numFmt = '#,##0.00';
      });

      ws.addRow([]);
      const summaryRow = ws.addRow({ date: 'الإجمالي', total: stmt.totalPurchases, paid: stmt.totalPaid, balance: stmt.balance });
      styleRow(summaryRow);
      summaryRow.font = { bold: true, name: 'Cairo', size: 10 };
      summaryRow.getCell('total').numFmt = '#,##0.00';
      summaryRow.getCell('paid').numFmt = '#,##0.00';
      summaryRow.getCell('balance').numFmt = '#,##0.00';

      await saveWorkbook(`كشف_مورد_${query}.xlsx`);
    }

    else if (type === 'rep') {
      const allSales = workbookData.sales.filter(s => s.rep === query);
      const allPurchases = workbookData.purchases.filter(p => p.rep === query);
      const allPayments = workbookData.repPayments.filter(p => p.rep === query);

      // Sheet 1: Sales and Profits
      const ws = workbook.addWorksheet('مبيعات وأرباح المندوب');
      const columns = [
        { header: 'التاريخ', key: 'date', width: 15 },
        { header: 'العميل', key: 'customer', width: 20 },
        { header: 'المنتج', key: 'product', width: 20 },
        { header: 'الكمية', key: 'qty', width: 10 },
        { header: 'سعر البيع', key: 'salePrice', width: 15 },
        { header: 'إجمالي المبيعات', key: 'sales', width: 15 },
        { header: 'سعر الشراء (التكلفة)', key: 'buyPrice', width: 18 },
        { header: 'الربح / الخسارة', key: 'profit', width: 15 }
      ];

      setupSheet(ws, `مبيعات وأرباح المندوب: ${query}`, columns);

      let totalSalesAmt = 0;
      let totalProfitAmt = 0;
      const products = getProducts();

      allSales.forEach(s => {
        totalSalesAmt += s.total;

        const prd = products.find(px => px.name === s.product);
        const costPrice = prd ? prd.buyPrice : 0;
        const totalCost = s.netQuantity * costPrice;
        const profit = s.total - totalCost;
        totalProfitAmt += profit;

        const row = ws.addRow({
          date: s.date, customer: s.customer, product: s.product, qty: s.netQuantity,
          salePrice: s.price, sales: s.total, buyPrice: costPrice, profit: profit
        });
        styleRow(row);
        row.getCell('salePrice').numFmt = '#,##0.00';
        row.getCell('sales').numFmt = '#,##0.00';
        row.getCell('buyPrice').numFmt = '#,##0.00';

        const profitCell = row.getCell('profit');
        profitCell.numFmt = '#,##0.00';
        if (profit < 0) {
          profitCell.font = { name: 'Cairo', size: 10, color: { argb: 'FFEF5350' }, bold: true }; // Red for loss
        } else if (profit > 0) {
          profitCell.font = { name: 'Cairo', size: 10, color: { argb: 'FF4CAF50' }, bold: true }; // Green for profit
        }
      });

      ws.addRow([]);
      const summarySalesRow = ws.addRow({ customer: 'الإجمالي', sales: totalSalesAmt, profit: totalProfitAmt });
      styleRow(summarySalesRow);
      summarySalesRow.font = { bold: true, name: 'Cairo', size: 11 };
      summarySalesRow.getCell('sales').numFmt = '#,##0.00';
      summarySalesRow.getCell('profit').numFmt = '#,##0.00';
      if (totalProfitAmt < 0) summarySalesRow.getCell('profit').font = { color: { argb: 'FFEF5350' }, bold: true };
      else summarySalesRow.getCell('profit').font = { color: { argb: 'FF4CAF50' }, bold: true };

      // Sheet 2: Purchases
      const wsPurchases = workbook.addWorksheet('مشتريات المندوب');
      const colsPurchases = [
        { header: 'التاريخ', key: 'date', width: 15 },
        { header: 'المورد', key: 'supplier', width: 20 },
        { header: 'المنتج', key: 'product', width: 20 },
        { header: 'الكمية', key: 'qty', width: 10 },
        { header: 'السعر', key: 'price', width: 15 },
        { header: 'إجمالي المشتروات', key: 'total', width: 18 }
      ];
      setupSheet(wsPurchases, `مشتريات المندوب: ${query}`, colsPurchases);

      let totalPurchAmnt = 0;
      allPurchases.forEach(p => {
        totalPurchAmnt += p.total;
        const row = wsPurchases.addRow({ date: p.date, supplier: p.supplier, product: p.product, qty: p.quantity, price: p.price, total: p.total });
        styleRow(row);
        row.getCell('price').numFmt = '#,##0.00';
        row.getCell('total').numFmt = '#,##0.00';
      });
      wsPurchases.addRow([]);
      const sumPurchRow = wsPurchases.addRow({ supplier: 'الإجمالي', total: totalPurchAmnt });
      styleRow(sumPurchRow);
      sumPurchRow.font = { bold: true, name: 'Cairo', size: 11 };
      sumPurchRow.getCell('total').numFmt = '#,##0.00';

      // Sheet 3: Financial Summary
      const wsSummary = workbook.addWorksheet('الملخص المالي له');
      const colsSummary = [
        { header: 'البند المالي', key: 'item', width: 30 },
        { header: 'المبلغ', key: 'amount', width: 20 }
      ];
      setupSheet(wsSummary, `الملخص المالي والدفعات للمندوب: ${query}`, colsSummary);

      let totalPay = 0;
      allPayments.forEach(p => totalPay += p.amount);

      const items = [
        { item: 'إجمالي مبيعات المندوب', amount: totalSalesAmt },
        { item: 'إجمالي مشتريات المندوب', amount: totalPurchAmnt },
        { item: 'صافي الأرباح المحققة', amount: totalProfitAmt },
        { item: '', amount: '' },
        { item: 'إجمالي الدفعات المسلمة له', amount: totalPay },
        { item: 'المتبقي (المبيعات - الدفعات)', amount: totalSalesAmt - totalPay }
      ];

      items.forEach(it => {
        const row = wsSummary.addRow(it);
        styleRow(row);
        if (typeof it.amount === 'number') {
          row.getCell('amount').numFmt = '#,##0.00';
        }
      });
      wsSummary.getRow(5).height = 10; // Empty row height
      wsSummary.getRow(8).font = { bold: true, name: 'Cairo', size: 12 };

      await saveWorkbook(`تقرير_شامل_للمندوب_${query}.xlsx`);
    }

    else if (type === 'full') {
      const sheetsConf = [
        { name: 'الملخص المالي', title: 'لوحة التحكم والنتائج المالية', data: [], cols: [{ h: 'البند', k: 'item', w: 30 }, { h: 'المبلغ المستحق', k: 'amount', w: 20 }] },
        { name: SHEETS.SALES, title: 'سجل المبيعات الشامل والتحصيلات', data: workbookData.sales, cols: [{ h: 'التاريخ', k: 'date', w: 12 }, { h: 'العميل', k: 'customer', w: 20 }, { h: 'المنتج', k: 'product', w: 20 }, { h: 'المندوب', k: 'rep', w: 15 }, { h: 'الكمية', k: 'netQuantity', w: 10 }, { h: 'السعر', k: 'price', w: 12 }, { h: 'الإجمالي', k: 'total', w: 15 }, { h: 'التحصيل', k: 'collected', w: 15 }] },
        { name: SHEETS.PURCHASES, title: 'سجل المشتريات والمصروفات المباشرة', data: workbookData.purchases, cols: [{ h: 'التاريخ', k: 'date', w: 12 }, { h: 'المنتج', k: 'product', w: 20 }, { h: 'المندوب', k: 'rep', w: 15 }, { h: 'الكمية', k: 'quantity', w: 10 }, { h: 'السعر', k: 'price', w: 12 }, { h: 'الإجمالي', k: 'total', w: 15 }, { h: 'المورد', k: 'supplier', w: 20 }, { h: 'المدفوع', k: 'paid', w: 15 }] },
        { name: SHEETS.INVENTORY, title: 'تقرير الجرد وحركة المخزون', data: workbookData.inventory, cols: [{ h: 'المنتج', k: 'product', w: 25 }, { h: 'الرصيد السابق', k: 'previousStock', w: 12 }, { h: 'المشترى', k: 'purchased', w: 12 }, { h: 'المسحوب', k: 'withdrawn', w: 12 }, { h: 'الصافي الحالي', k: 'net', w: 15 }] },
        { name: 'الخزينة والسندات', title: 'حركة الخزينة وسجل السندات', data: workbookData.treasury, cols: [{ h: 'التاريخ', k: 'date', w: 15 }, { h: 'النوع', k: 'type', w: 10 }, { h: 'المبلغ', k: 'amount', w: 15 }, { h: 'المرجع', k: 'reference', w: 25 }, { h: 'ملاحظات', k: 'notes', w: 30 }, { h: 'المستخدم', k: 'user', w: 15 }] },
        { name: 'الموارد البشرية', title: 'سجل الموظفين والرواتب والعهد', data: workbookData.hr, cols: [{ h: 'التاريخ', k: 'date', w: 15 }, { h: 'الموظف', k: 'name', w: 20 }, { h: 'النوع', k: 'type', w: 15 }, { h: 'المبلغ', k: 'amount', w: 15 }, { h: 'ملاحظات', k: 'notes', w: 30 }] }
      ];

      for (let conf of sheetsConf) {
        const ws = workbook.addWorksheet(conf.name);
        setupSheet(ws, conf.title, conf.cols.map(c => ({ header: c.h, key: c.k, width: c.w })));

        if (conf.name === 'الملخص المالي') {
          const stats = getStats();
          const rows = [
            { item: 'إجمالي المبيعات', amount: stats.totalSales },
            { item: 'إجمالي المشتريات', amount: stats.totalPurchases },
            { item: 'صافي الربح المبدئي', amount: stats.realProfit },
            { item: 'إجمالي تكاليف العمالة (HR)', amount: stats.totalHRCosts },
            { item: 'الربح النهائي للمشروع', amount: stats.realProfit - stats.totalHRCosts },
            { item: 'كاش الصندوق (توريد المناديب)', amount: stats.cashInBox },
            { item: 'إجمالي الديون العالقة (عملاء)', amount: stats.totalDue }
          ];
          rows.forEach(r => {
            const row = ws.addRow(r);
            styleRow(row);
            row.getCell('amount').numFmt = '#,##0.00';
            if (r.item.includes('الربح') && r.amount < 0) row.getCell('amount').font = { color: { argb: 'FFEF5350' }, bold: true };
            if (r.item.includes('الربح') && r.amount > 0) row.getCell('amount').font = { color: { argb: 'FF4CAF50' }, bold: true };
          });
        } else {
          conf.data.forEach(item => {
            const rowData = {};
            conf.cols.forEach(c => rowData[c.k] = item[c.k]);
            const row = ws.addRow(rowData);
            styleRow(row);
            row.eachCell(c => {
              if (typeof c.value === 'number') c.numFmt = '#,##0.00';
            });
          });
        }
      }
      await saveWorkbook(`التقرير_الشامل_للمشروع_${getTodayStr()}.xlsx`);
    }

    else if (type === 'inventory') {
      const ws = workbook.addWorksheet('المخزون');
      const columns = [
        { header: 'المنتج', key: 'product', width: 25 },
        { header: 'الرصيد السابق', key: 'previousStock', width: 15 },
        { header: 'المشتريات', key: 'purchased', width: 15 },
        { header: 'المسحوبات', key: 'withdrawn', width: 15 },
        { header: 'الصافي الحالي', key: 'net', width: 15 },
        { header: 'قيمة المخزون (بيع)', key: 'valSale', width: 20 }
      ];
      setupSheet(ws, 'تقرير الجرد والمخزون التفصيلي', columns);

      workbookData.inventory.forEach(inv => {
        const prod = workbookData.products.find(p => p.name === inv.product);
        const valSale = (inv.net || 0) * (prod ? (prod.sellPrice || 0) : 0);
        const row = ws.addRow({ ...inv, valSale });
        styleRow(row);
        row.getCell('valSale').numFmt = '#,##0.00';
      });
      await saveWorkbook(`تقرير_المخزون_${getTodayStr()}.xlsx`);
    }

    else if (type === 'profits') {
      const ws = workbook.addWorksheet('الأرباح');
      const columns = [{ h: 'البيان', k: 'item', w: 35 }, { h: 'المبلغ', k: 'amount', w: 20 }];
      setupSheet(ws, 'تقرير الأرباح والخسائر المالي', columns.map(c => ({ header: c.h, key: c.k, width: c.w })));

      const stats = getStats();
      const rows = [
        { item: 'إجمالي المبيعات', amount: stats.totalSales || 0 },
        { item: 'إجمالي تكلفة المشتريات', amount: stats.totalPurchases || 0 },
        { item: 'مجمل الربح', amount: stats.profit || 0 },
        { item: 'صافي الربح (بعد التكلفة)', amount: stats.realProfit || 0 },
        { item: 'إجمالي مصاريف الرواتب والعهد (HR)', amount: stats.totalHRCosts || 0 },
        { item: 'الربح الصافي النهائي', amount: (stats.realProfit || 0) - (stats.totalHRCosts || 0) }
      ];

      rows.forEach(r => {
        const row = ws.addRow(r);
        styleRow(row);
        row.getCell('amount').numFmt = '#,##0.00';
        if (typeof r.amount === 'number' && r.amount < 0) row.getCell('amount').font = { color: { argb: 'FFEF5350' }, bold: true };
      });
      await saveWorkbook(`تقرير_الأرباح_${getTodayStr()}.xlsx`);
    }

    // === NEW ADVANCED REPORTS ===

    else if (type === 'salesReport') {
      const ws = workbook.addWorksheet('تقرير المبيعات');
      const columns = [
        { header: 'التاريخ', key: 'date', width: 15 },
        { header: 'رقم الفاتورة', key: 'id', width: 12 },
        { header: 'العميل', key: 'customer', width: 25 },
        { header: 'المنتج', key: 'product', width: 25 },
        { header: 'المندوب', key: 'rep', width: 20 },
        { header: 'الكمية', key: 'quantity', width: 10 },
        { header: 'المرتجع', key: 'returned', width: 10 },
        { header: 'الصافي', key: 'netQty', width: 10 },
        { header: 'السعر', key: 'price', width: 12 },
        { header: 'الإجمالي', key: 'total', width: 15 },
        { header: 'المحصل', key: 'collected', width: 15 },
        { header: 'المتبقي', key: 'balance', width: 15 }
      ];
      setupSheet(ws, 'تقرير المبيعات التفصيلي', columns);

      let totalQty = 0, totalAmount = 0, totalCollected = 0, totalBalance = 0;
      workbookData.sales.forEach(s => {
        totalQty += s.netQuantity || 0;
        totalAmount += s.total || 0;
        totalCollected += s.collected || 0;
        totalBalance += s.balance || 0;
        const row = ws.addRow({
          date: s.date, id: s.id, customer: s.customer, product: s.product,
          rep: s.rep, quantity: s.quantity, returned: s.returned, netQty: s.netQuantity,
          price: s.price, total: s.total, collected: s.collected, balance: s.balance
        });
        styleRow(row);
        ['price', 'total', 'collected', 'balance'].forEach(k => {
          row.getCell(k).numFmt = '#,##0.00';
        });
      });

      ws.addRow([]);
      const summaryRow = ws.addRow({
        date: 'الإجمالي', quantity: totalQty, netQty: totalQty,
        price: '-', total: totalAmount, collected: totalCollected, balance: totalBalance
      });
      styleRow(summaryRow);
      summaryRow.font = { bold: true, name: 'Cairo', size: 11 };
      ['total', 'collected', 'balance'].forEach(k => {
        summaryRow.getCell(k).numFmt = '#,##0.00';
      });

      await saveWorkbook(`تقرير_المبيعات_${getTodayStr()}.xlsx`);
    }

    else if (type === 'inventoryMovement') {
      const ws = workbook.addWorksheet('حركة المخزون');
      const columns = [
        { header: 'التاريخ', key: 'date', width: 15 },
        { header: 'رقم الحركة', key: 'id', width: 12 },
        { header: 'المنتج', key: 'product', width: 30 },
        { header: 'النوع', key: 'type', width: 15 },
        { header: 'التصنيف', key: 'category', width: 20 },
        { header: 'الكمية', key: 'amount', width: 12 },
        { header: 'الطرف', key: 'party', width: 25 },
        { header: 'المرجع', key: 'reference', width: 25 },
        { header: 'الموقع', key: 'location', width: 25 }
      ];
      setupSheet(ws, 'تقرير حركة المخزون الشامل', columns);

      const ledger = getInventoryLedger();
      ledger.forEach(item => {
        const row = ws.addRow(item);
        styleRow(row);
        row.getCell('amount').numFmt = '#,##0.00';
        if (item.type === 'OUT') {
          row.getCell('type').font = { color: { argb: 'FFEF5350' }, bold: true, name: 'Cairo', size: 10 };
        } else {
          row.getCell('type').font = { color: { argb: 'FF4CAF50' }, bold: true, name: 'Cairo', size: 10 };
        }
      });

      await saveWorkbook(`تقرير_حركة_المخزون_${getTodayStr()}.xlsx`);
    }

    else if (type === 'profitLoss') {
      const ws = workbook.addWorksheet('الأرباح والخسائر');
      const columns = [
        { header: 'البند', key: 'item', width: 40 },
        { header: 'التفاصيل', key: 'details', width: 30 },
        { header: 'المبلغ', key: 'amount', width: 20 }
      ];
      setupSheet(ws, 'تقرير الأرباح والخسائر الشامل', columns);

      const stats = getStats();
      const products = getProducts();

      // حساب تفصيلي للأرباح حسب المنتج
      let detailedProfitRows = [];
      products.forEach(p => {
        const productSales = workbookData.sales.filter(s => s.product === p.name);
        const totalSalesQty = productSales.reduce((sum, s) => sum + (s.netQuantity || 0), 0);
        const totalSalesAmt = productSales.reduce((sum, s) => sum + (s.total || 0), 0);
        const totalCost = totalSalesQty * (p.buyPrice || 0);
        const profit = totalSalesAmt - totalCost;

        if (totalSalesQty > 0) {
          detailedProfitRows.push({
            item: `مبيعات ${p.name}`,
            details: `الكمية: ${totalSalesQty} | سعر البيع: ${p.sellPrice} | التكلفة: ${p.buyPrice}`,
            amount: profit
          });
        }
      });

      const rows = [
        { item: '=== إيرادات المبيعات ===', details: '', amount: '' },
        { item: 'إجمالي المبيعات', details: `${workbookData.sales.length} فاتورة`, amount: stats.totalSales },
        { item: 'إجمالي تكلفة البضاعة المباعة', details: '', amount: -stats.totalPurchases },
        { item: 'مجمل الربح', details: '', amount: stats.profit },
        { item: '', details: '', amount: '' },
        { item: '=== صافي الربح الحقيقي ===', details: 'حسب تكلفة كل منتج', amount: stats.realProfit },
        { item: '', details: '', amount: '' },
        { item: '=== المصروفات التشغيلية ===', details: '', amount: '' },
        { item: 'إجمالي مصاريف الموارد البشرية', details: `رواتب، سلف، مكافآت`, amount: -stats.totalHRCosts },
        { item: '', details: '', amount: '' },
        { item: '=== الربح الصافي النهائي ===', details: 'بعد جميع المصروفات', amount: stats.realProfit - stats.totalHRCosts },
        { item: '', details: '', amount: '' },
        { item: '=== مؤشرات مالية إضافية ===', details: '', amount: '' },
        { item: 'رصيد الخزينة', details: 'كاش الصندوق', amount: stats.cashInBox },
        { item: 'عهدة المناديب', details: 'أموال تحت التحصيل', amount: stats.totalRepCustody },
        { item: 'الديون العالقة', details: `${stats.topDebtors.length} عميل عليه مديونية`, amount: -stats.totalDue }
      ];

      rows.forEach(r => {
        const row = ws.addRow(r);
        styleRow(row);
        if (typeof r.amount === 'number') {
          row.getCell('amount').numFmt = '#,##0.00';
          if (r.amount < 0) {
            row.getCell('amount').font = { color: { argb: 'FFEF5350' }, bold: true, name: 'Cairo', size: 10 };
          } else if (r.amount > 0 && (r.item.includes('الربح') || r.item.includes('إجمالي'))) {
            row.getCell('amount').font = { color: { argb: 'FF4CAF50' }, bold: true, name: 'Cairo', size: 10 };
          }
        }
        if (r.item.startsWith('===') && r.item.endsWith('===')) {
          row.font = { bold: true, name: 'Cairo', size: 11 };
          row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE3F2FD' } };
        }
      });

      // إضافة جدول تفصيلي لأرباح المنتجات
      ws.addRow([]);
      const detailHeader = ws.addRow({ item: '=== تفصيل أرباح المنتجات ===', details: '', amount: '' });
      detailHeader.font = { bold: true, name: 'Cairo', size: 11 };
      detailHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE3F2FD' } };

      detailedProfitRows.sort((a, b) => b.amount - a.amount).forEach(r => {
        const row = ws.addRow(r);
        styleRow(row);
        row.getCell('amount').numFmt = '#,##0.00';
        if (r.amount < 0) {
          row.getCell('amount').font = { color: { argb: 'FFEF5350' }, bold: true, name: 'Cairo', size: 10 };
        } else {
          row.getCell('amount').font = { color: { argb: 'FF4CAF50' }, bold: true, name: 'Cairo', size: 10 };
        }
      });

      await saveWorkbook(`تقرير_الأرباح_والخسائر_${getTodayStr()}.xlsx`);
    }

    else if (type === 'expensesReport') {
      const ws = workbook.addWorksheet('المصروفات');
      const columns = [
        { header: 'التاريخ', key: 'date', width: 15 },
        { header: 'النوع', key: 'type', width: 20 },
        { header: 'الموظف/الطرف', key: 'name', width: 25 },
        { header: 'المبلغ', key: 'amount', width: 18 },
        { header: 'ملاحظات', key: 'notes', width: 40 }
      ];
      setupSheet(ws, 'تقرير المصروفات التفصيلي', columns);

      // جمع كل المصروفات من HR والخزينة
      let expenses = [];

      // مصروفات الموارد البشرية
      workbookData.hr.forEach(h => {
        expenses.push({
          date: h.date,
          type: h.type || 'مصروف',
          name: h.name || '',
          amount: h.amount || 0,
          notes: h.notes || ''
        });
      });

      // مصروفات الخزينة (الصادر)
      workbookData.treasury.forEach(t => {
        if (t.type === 'OUT') {
          expenses.push({
            date: t.date,
            type: 'مصروف خزينة',
            name: t.reference || '',
            amount: t.amount || 0,
            notes: t.notes || ''
          });
        }
      });

      // ترتيب حسب التاريخ
      expenses.sort((a, b) => new Date(b.date) - new Date(a.date));

      let totalExpenses = 0;
      expenses.forEach(e => {
        totalExpenses += e.amount || 0;
        const row = ws.addRow(e);
        styleRow(row);
        row.getCell('amount').numFmt = '#,##0.00';
      });

      ws.addRow([]);
      const summaryRow = ws.addRow({
        date: 'الإجمالي',
        type: `${expenses.length} مصروف`,
        name: '',
        amount: totalExpenses,
        notes: ''
      });
      styleRow(summaryRow);
      summaryRow.font = { bold: true, name: 'Cairo', size: 11 };
      summaryRow.getCell('amount').numFmt = '#,##0.00';
      summaryRow.getCell('amount').font = { color: { argb: 'FFEF5350' }, bold: true };

      await saveWorkbook(`تقرير_المصروفات_${getTodayStr()}.xlsx`);
    }

    else if (type === 'repPerformance') {
      const ws = workbook.addWorksheet('أداء المناديب');
      const columns = [
        { header: 'المندوب', key: 'name', width: 25 },
        { header: 'عدد المبيعات', key: 'salesCount', width: 15 },
        { header: 'إجمالي المبيعات', key: 'totalSales', width: 20 },
        { header: 'إجمالي المشتريات', key: 'totalPurchases', width: 20 },
        { header: 'صافي الربح', key: 'totalProfit', width: 20 },
        { header: 'الموَرَّد', key: 'totalDelivered', width: 18 },
        { header: 'عهدة نقدية', key: 'cashCustody', width: 18 },
        { header: 'رواتب/سلف', key: 'totalSalaries', width: 18 },
        { header: 'صافي المستحق', key: 'netBalance', width: 20 }
      ];
      setupSheet(ws, 'تقرير مقارنة أداء المناديب', columns);

      const reps = getRepsList();
      let repData = [];

      reps.forEach(repName => {
        const summary = getRepAccountingSummary(repName);
        repData.push(summary);
      });

      // ترتيب حسب إجمالي المبيعات
      repData.sort((a, b) => b.totalSales - a.totalSales);

      repData.forEach(r => {
        const row = ws.addRow(r);
        styleRow(row);
        ['totalSales', 'totalPurchases', 'totalProfit', 'totalDelivered', 'cashCustody', 'totalSalaries', 'netBalance'].forEach(k => {
          row.getCell(k).numFmt = '#,##0.00';
        });
        // تلوين الربح
        if (r.totalProfit < 0) {
          row.getCell('totalProfit').font = { color: { argb: 'FFEF5350' }, bold: true, name: 'Cairo', size: 10 };
        } else if (r.totalProfit > 0) {
          row.getCell('totalProfit').font = { color: { argb: 'FF4CAF50' }, bold: true, name: 'Cairo', size: 10 };
        }
        // تلوين العهدة
        if (r.cashCustody > 0) {
          row.getCell('cashCustody').font = { color: { argb: 'FFEF5350' }, bold: true, name: 'Cairo', size: 10 };
        }
      });

      // إضافة ملخص
      if (repData.length > 0) {
        ws.addRow([]);
        const totals = repData.reduce((acc, r) => {
          acc.salesCount += r.salesCount;
          acc.totalSales += r.totalSales;
          acc.totalPurchases += r.totalPurchases;
          acc.totalProfit += r.totalProfit;
          acc.totalDelivered += r.totalDelivered;
          acc.cashCustody += r.cashCustody;
          acc.totalSalaries += r.totalSalaries;
          acc.netBalance += r.netBalance;
          return acc;
        }, { salesCount: 0, totalSales: 0, totalPurchases: 0, totalProfit: 0, totalDelivered: 0, cashCustody: 0, totalSalaries: 0, netBalance: 0 });

        const summaryRow = ws.addRow({
          name: 'الإجمالي',
          ...totals
        });
        styleRow(summaryRow);
        summaryRow.font = { bold: true, name: 'Cairo', size: 11 };
        ['totalSales', 'totalPurchases', 'totalProfit', 'totalDelivered', 'cashCustody', 'totalSalaries', 'netBalance'].forEach(k => {
          summaryRow.getCell(k).numFmt = '#,##0.00';
        });
      }

      await saveWorkbook(`تقرير_أداء_المناديب_${getTodayStr()}.xlsx`);
    }

    else if (type === 'taxReport') {
      const ws = workbook.addWorksheet('التقرير الضريبي');
      const columns = [
        { header: 'التاريخ', key: 'date', width: 15 },
        { header: 'رقم الفاتورة', key: 'id', width: 12 },
        { header: 'النوع', key: 'invoiceType', width: 15 },
        { header: 'الطرف', key: 'party', width: 25 },
        { header: 'المنتج', key: 'product', width: 25 },
        { header: 'القيمة قبل الضريبة', key: 'preTax', width: 20 },
        { header: 'الضريبة (15%)', key: 'tax', width: 18 },
        { header: 'الإجمالي', key: 'total', width: 18 }
      ];
      setupSheet(ws, 'التقرير الضريبي (VAT 15%)', columns);

      let transactions = [];

      // المبيعات (ضريبة مخرجات)
      workbookData.sales.forEach(s => {
        const preTax = s.total || 0;
        const tax = preTax * 0.15;
        transactions.push({
          date: s.date,
          id: s.id,
          invoiceType: 'مبيعات',
          party: s.customer,
          product: s.product,
          preTax: preTax,
          tax: tax,
          total: preTax + tax
        });
      });

      // المشتريات (ضريبة مدخلات)
      workbookData.purchases.forEach(p => {
        const preTax = p.total || 0;
        const tax = preTax * 0.15;
        transactions.push({
          date: p.date,
          id: p.id,
          invoiceType: 'مشتريات',
          party: p.supplier,
          product: p.product,
          preTax: preTax,
          tax: tax,
          total: preTax + tax
        });
      });

      // ترتيب حسب التاريخ
      transactions.sort((a, b) => new Date(a.date) - new Date(b.date));

      let totalSalesTax = 0, totalPurchaseTax = 0;
      transactions.forEach(t => {
        if (t.invoiceType === 'مبيعات') totalSalesTax += t.tax;
        else totalPurchaseTax += t.tax;

        const row = ws.addRow(t);
        styleRow(row);
        ['preTax', 'tax', 'total'].forEach(k => {
          row.getCell(k).numFmt = '#,##0.00';
        });
      });

      // ملخص ضريبي
      ws.addRow([]);
      const netTax = totalSalesTax - totalPurchaseTax;

      const summaryRows = [
        { date: 'الملخص الضريبي', invoiceType: '', party: '', product: '', preTax: '', tax: '', total: '' },
        { date: '', invoiceType: 'ضريبة المبيعات (مخرجات)', party: '', product: '', preTax: '', tax: totalSalesTax, total: '' },
        { date: '', invoiceType: 'ضريبة المشتريات (مدخلات)', party: '', product: '', preTax: '', tax: -totalPurchaseTax, total: '' },
        { date: '', invoiceType: 'صافي الضريبة المستحقة', party: '', product: '', preTax: '', tax: netTax, total: '' }
      ];

      summaryRows.forEach(r => {
        const row = ws.addRow(r);
        styleRow(row);
        row.getCell('tax').numFmt = '#,##0.00';
        if (r.invoiceType === 'صافي الضريبة المستحقة') {
          row.font = { bold: true, name: 'Cairo', size: 11 };
          row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBBDEFB' } };
          if (netTax > 0) {
            row.getCell('tax').font = { color: { argb: 'FFEF5350' }, bold: true };
          } else {
            row.getCell('tax').font = { color: { argb: 'FF4CAF50' }, bold: true };
          }
        }
      });

      await saveWorkbook(`التقرير_الضريبي_${getTodayStr()}.xlsx`);
    }

    else if (type === 'debtsReport') {
      const ws = workbook.addWorksheet('المديونيات');
      const columns = [
        { header: 'العميل', key: 'name', width: 30 },
        { header: 'رقم المرجع', key: 'refNumber', width: 15 },
        { header: 'إجمالي المبيعات', key: 'total', width: 20 },
        { header: 'المحصل', key: 'collected', width: 20 },
        { header: 'المتبقي (المديونية)', key: 'remaining', width: 25 },
        { header: 'نسبة التحصيل', key: 'collectionRate', width: 15 },
        { header: 'آخر تعامل', key: 'lastDate', width: 15 }
      ];
      setupSheet(ws, 'تقرير تحليل المديونيات والديون', columns);

      // العملاء اللي عليهم ديون
      let debtors = workbookData.customers
        .filter(c => (c.remaining || 0) > 0)
        .map(c => ({
          name: c.name,
          refNumber: c.refNumber,
          total: c.total || 0,
          collected: c.collected || 0,
          remaining: c.remaining || 0,
          collectionRate: c.total > 0 ? ((c.collected / c.total) * 100).toFixed(1) + '%' : '0%',
          lastDate: c.lastDate || '-'
        }))
        .sort((a, b) => b.remaining - a.remaining);

      let totalDebt = 0;
      debtors.forEach(d => {
        totalDebt += d.remaining;
        const row = ws.addRow(d);
        styleRow(row);
        ['total', 'collected', 'remaining'].forEach(k => {
          row.getCell(k).numFmt = '#,##0.00';
        });
        // تلوين حسب نسبة التحصيل
        const rate = parseFloat(d.collectionRate) || 0;
        if (rate < 50) {
          row.getCell('remaining').font = { color: { argb: 'FFEF5350' }, bold: true, name: 'Cairo', size: 10 };
        } else if (rate >= 80) {
          row.getCell('remaining').font = { color: { argb: 'FF4CAF50' }, bold: true, name: 'Cairo', size: 10 };
        }
      });

      if (debtors.length === 0) {
        ws.addRow({ name: 'لا توجد مديونيات على العملاء', refNumber: '', total: '', collected: '', remaining: '', collectionRate: '', lastDate: '' });
      }

      ws.addRow([]);
      const summaryRow = ws.addRow({
        name: `الإجمالي (${debtors.length} عميل عليه ديون)`,
        refNumber: '',
        total: debtors.reduce((s, d) => s + d.total, 0),
        collected: debtors.reduce((s, d) => s + d.collected, 0),
        remaining: totalDebt,
        collectionRate: '',
        lastDate: ''
      });
      styleRow(summaryRow);
      summaryRow.font = { bold: true, name: 'Cairo', size: 11 };
      ['total', 'collected', 'remaining'].forEach(k => {
        summaryRow.getCell(k).numFmt = '#,##0.00';
      });
      summaryRow.getCell('remaining').font = { color: { argb: 'FFEF5350' }, bold: true, name: 'Cairo', size: 11 };

      // مديونيات الموردين
      ws.addRow([]);
      const supplierHeader = ws.addRow({ name: '=== مديونيات الموردين ===', refNumber: '', total: '', collected: '', remaining: '', collectionRate: '', lastDate: '' });
      supplierHeader.font = { bold: true, name: 'Cairo', size: 11 };
      supplierHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE3F2FD' } };

      let suppliersDebt = workbookData.suppliers
        .filter(s => (s.remaining || 0) > 0)
        .map(s => ({
          name: s.name,
          refNumber: '',
          total: s.total || 0,
          collected: s.paid || 0,
          remaining: s.remaining || 0,
          collectionRate: '',
          lastDate: ''
        }))
        .sort((a, b) => b.remaining - a.remaining);

      let totalSupplierDebt = 0;
      suppliersDebt.forEach(s => {
        totalSupplierDebt += s.remaining;
        const row = ws.addRow(s);
        styleRow(row);
        ['total', 'collected', 'remaining'].forEach(k => {
          row.getCell(k).numFmt = '#,##0.00';
        });
      });

      if (suppliersDebt.length > 0) {
        const suppSummaryRow = ws.addRow({
          name: `إجمالي مديونيات الموردين (${suppliersDebt.length} مورد)`,
          refNumber: '',
          total: suppliersDebt.reduce((s, d) => s + d.total, 0),
          collected: suppliersDebt.reduce((s, d) => s + d.collected, 0),
          remaining: totalSupplierDebt,
          collectionRate: '',
          lastDate: ''
        });
        styleRow(suppSummaryRow);
        suppSummaryRow.font = { bold: true, name: 'Cairo', size: 11 };
        ['total', 'collected', 'remaining'].forEach(k => {
          suppSummaryRow.getCell(k).numFmt = '#,##0.00';
        });
      }

      await saveWorkbook(`تقرير_المديونيات_${getTodayStr()}.xlsx`);
    }

    else if (type === 'productProfitability') {
      const ws = workbook.addWorksheet('ربحية المنتجات');
      const columns = [
        { header: 'المنتج', key: 'name', width: 30 },
        { header: 'القسم', key: 'category', width: 15 },
        { header: 'سعر الشراء', key: 'buyPrice', width: 15 },
        { header: 'سعر البيع', key: 'sellPrice', width: 15 },
        { header: 'هامش الربح', key: 'margin', width: 15 },
        { header: 'نسبة الربح', key: 'marginPercent', width: 15 },
        { header: 'الكمية المباعة', key: 'soldQty', width: 15 },
        { header: 'إجمالي المبيعات', key: 'totalSales', width: 20 },
        { header: 'إجمالي التكلفة', key: 'totalCost', width: 20 },
        { header: 'صافي الربح', key: 'netProfit', width: 20 },
        { header: 'الرصيد الحالي', key: 'currentStock', width: 15 }
      ];
      setupSheet(ws, 'تقرير ربحية المنتجات الشامل', columns);

      const products = getProducts();
      let productData = [];

      products.forEach(p => {
        const sales = workbookData.sales.filter(s => s.product === p.name);
        const totalSoldQty = sales.reduce((sum, s) => sum + (s.netQuantity || 0), 0);
        const totalSalesAmt = sales.reduce((sum, s) => sum + (s.total || 0), 0);
        const totalCost = totalSoldQty * (p.buyPrice || 0);
        const netProfit = totalSalesAmt - totalCost;
        const margin = (p.sellPrice || 0) - (p.buyPrice || 0);
        const marginPercent = (p.buyPrice || 0) > 0 ? ((margin / p.buyPrice) * 100).toFixed(1) + '%' : '0%';

        const inv = workbookData.inventory.find(i => i.product === p.name);
        const currentStock = inv ? (inv.net || 0) : 0;

        productData.push({
          name: p.name,
          category: p.category || 'عام',
          buyPrice: p.buyPrice || 0,
          sellPrice: p.sellPrice || 0,
          margin: margin,
          marginPercent: marginPercent,
          soldQty: totalSoldQty,
          totalSales: totalSalesAmt,
          totalCost: totalCost,
          netProfit: netProfit,
          currentStock: currentStock
        });
      });

      // ترتيب حسب صافي الربح
      productData.sort((a, b) => b.netProfit - a.netProfit);

      let totalProfit = 0;
      productData.forEach(p => {
        totalProfit += p.netProfit;
        const row = ws.addRow(p);
        styleRow(row);
        ['buyPrice', 'sellPrice', 'margin', 'totalSales', 'totalCost', 'netProfit'].forEach(k => {
          row.getCell(k).numFmt = '#,##0.00';
        });

        // تلوين الربح
        if (p.netProfit < 0) {
          row.getCell('netProfit').font = { color: { argb: 'FFEF5350' }, bold: true, name: 'Cairo', size: 10 };
        } else if (p.netProfit > 0) {
          row.getCell('netProfit').font = { color: { argb: 'FF4CAF50' }, bold: true, name: 'Cairo', size: 10 };
        }

        // تحذير للمنتجات اللي فيها خسارة
        if (p.netProfit < 0) {
          row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '33FFEBEE' } };
        }
      });

      // ملخص
      ws.addRow([]);
      const summaryRow = ws.addRow({
        name: 'الإجمالي',
        category: `${productData.length} منتج`,
        buyPrice: '',
        sellPrice: '',
        margin: '',
        marginPercent: '',
        soldQty: productData.reduce((s, p) => s + p.soldQty, 0),
        totalSales: productData.reduce((s, p) => s + p.totalSales, 0),
        totalCost: productData.reduce((s, p) => s + p.totalCost, 0),
        netProfit: totalProfit,
        currentStock: productData.reduce((s, p) => s + p.currentStock, 0)
      });
      styleRow(summaryRow);
      summaryRow.font = { bold: true, name: 'Cairo', size: 11 };
      ['totalSales', 'totalCost', 'netProfit'].forEach(k => {
        summaryRow.getCell(k).numFmt = '#,##0.00';
      });
      if (totalProfit > 0) {
        summaryRow.getCell('netProfit').font = { color: { argb: 'FF4CAF50' }, bold: true };
      } else {
        summaryRow.getCell('netProfit').font = { color: { argb: 'FFEF5350' }, bold: true };
      }

      await saveWorkbook(`تقرير_ربحية_المنتجات_${getTodayStr()}.xlsx`);
    }

    else if (type === 'productProfit') {
      const ws = workbook.addWorksheet('أرباح المنتج');
      const columns = [
        { header: 'البيان', key: 'item', width: 35 },
        { header: 'الكمية', key: 'qty', width: 15 },
        { header: 'المبلغ الإجمالي', key: 'amount', width: 25 }
      ];
      setupSheet(ws, `تقرير ربحية وحركة المنتج: ${query}`, columns);

      const sales = workbookData.sales.filter(s => s.product === query);
      const purchases = workbookData.purchases.filter(p => p.product === query);
      const inv = workbookData.inventory.find(i => i.product === query);

      const totalSalesQty = sales.reduce((sum, s) => sum + (s.netQuantity || 0), 0);
      const totalSalesAmt = sales.reduce((sum, s) => sum + (s.total || 0), 0);

      const totalPurchQty = purchases.reduce((sum, p) => sum + (p.quantity || 0), 0);
      const totalPurchAmt = purchases.reduce((sum, p) => sum + (p.total || 0), 0);

      const netStock = inv ? (inv.net || 0) : 0;
      const profit = totalSalesAmt - totalPurchAmt;

      const rows = [
        { item: 'إجمالي المشتريات من المنتج', qty: totalPurchQty, amount: totalPurchAmt },
        { item: 'إجمالي مبيعات المنتج', qty: totalSalesQty, amount: totalSalesAmt },
        { item: 'الرصيد المتبقي في المخزن', qty: netStock, amount: '---' },
        { item: 'الفرق (الربح / الخسارة)', qty: '---', amount: profit },
      ];

      rows.forEach(r => {
        const row = ws.addRow(r);
        styleRow(row);
        if (typeof r.amount === 'number') row.getCell('amount').numFmt = '#,##0.00';
        if (r.item.includes('الفرق') && typeof r.amount === 'number') {
          row.font = { bold: true, name: 'Cairo', size: 12 };
          if (r.amount < 0) row.getCell('amount').font = { color: { argb: 'FFEF5350' }, bold: true };
          if (r.amount > 0) row.getCell('amount').font = { color: { argb: 'FF4CAF50' }, bold: true };
        }
      });
      await saveWorkbook(`ربحية_منتج_${query}.xlsx`);
    }
  }

  // --- Security Helpers ---
  function generateInvoiceSecurity(invoice) {
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<invoice>
  <id>${invoice.id}</id>
  <date>${invoice.date}</date>
  <customer>${invoice.customer || invoice.supplier || 'N/A'}</customer>
  <product>${invoice.product}</product>
  <quantity>${invoice.quantity || invoice.netQuantity}</quantity>
  <price>${invoice.price}</price>
  <total>${invoice.total}</total>
  <timestamp>${new Date().toISOString()}</timestamp>
  <status>SECURE_STAMP_G2</status>
</invoice>`;

    // Simple digital signature based on content
    let hash = 0;
    const content = xml + "ERP_SIGN_KEY_PRO_2026";
    for (let i = 0; i < content.length; i++) {
      hash = ((hash << 5) - hash) + content.charCodeAt(i);
      hash |= 0;
    }
    const signature = "SEC-" + Math.abs(hash).toString(16).toUpperCase();
    return { xml, hash: signature };
  }

  function secureExistingInvoices() {
    let count = 0;
    [...workbookData.sales, ...workbookData.purchases].forEach(inv => {
      if (!inv.signature) {
        const security = generateInvoiceSecurity(inv);
        inv.xml = security.xml;
        inv.signature = security.hash;
        count++;
      }
    });
    if (count > 0) saveToLocalStorage();
    return count;
  }

  // --- Public API Integration ---
  const PublicAPI = {
    SHEETS,
    COLUMNS,
    DEFAULT_REPS,
    initializeDefaults,
    loadFromFile,
    exportToFile,
    addSale,
    addPurchase,
    getCustomerStatement,
    addCustomer,
    deleteCustomer,
    addProduct,
    updateProduct,
    deleteProduct,
    addSupplier,
    deleteSupplier,
    addRepPayment,
    getStats,
    getCustomerStatement,
    getSupplierStatement,
    getData,
    getSales,
    getPurchases,
    getInventory,
    getCustomers,
    getReps,
    getRepsList,
    getProducts,
    getSuppliers,
    getStaff,
    getHR,
    getCustomerPayments,
    addStaff,
    removeStaff,
    toggleStaffStatus,
    addHRAction,
    addCustomerPayment,
    getRepPayments,
    getIsLoaded,
    exportAdvancedReports,
    formatCurrency,
    formatNumber,
    saveToLocalStorage,
    loadFromLocalStorage,
    parseNum,
    getTodayStr,
    getRepAccountingSummary,
    login,
    logout,
    getCurrentUser,
    getTreasury,
    getVouchers,
    getLedger,
    updateInventoryManual,
    generateInvoiceSecurity,
    secureExistingInvoices,
    logTreasuryTransaction,
    transferToRepWarehouse,
    returnFromRepWarehouse,
    getRepWarehouse,
    getRepWarehouseAll,
    getWarehouseTransfers,
    getRepProductStock,
    getInventoryLedger,
    addCategory,
    getCategories,
    addVoucher,
    getRepStatement,
    getStaffStatement,
    addSaleReturn,
    addPurchaseReturn
  };

  return PublicAPI;
})();

window.ExcelEngine = ExcelEngine;


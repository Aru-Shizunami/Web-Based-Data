const APP_CONFIG = {
  spreadsheetProperty: 'DEBT_TRACKER_SPREADSHEET_ID',
  spreadsheetName: 'Orbit | Catatan Hutang',
  sheets: {
    summary: { name: 'Ringkasan', headers: [] },
    debts: {
      name: 'Hutang',
      headers: ['ID Hutang', 'Tanggal', 'Nama', 'Catatan', 'Total', 'Dibayar', 'Sisa', 'Status', 'Diperbarui'],
    },
    details: {
      name: 'Detail Hutang',
      headers: ['ID Detail', 'ID Hutang', 'Barang', 'Jumlah', 'Harga Satuan', 'Subtotal', 'Harga Beli', 'Total Modal'],
    },
    customers: {
      name: 'Pelanggan',
      headers: ['ID Pelanggan', 'Nama', 'Kontak', 'Catatan'],
    },
    products: {
      name: 'Barang',
      headers: ['ID Barang', 'Nama Barang', 'Harga Beli', 'Harga Jual', 'Diperbarui'],
    },
    payments: {
      name: 'Pembayaran',
      headers: ['ID Pembayaran', 'ID Hutang', 'Tanggal Bayar', 'Jumlah', 'Jenis', 'Catatan'],
    },
    purchases: {
      name: 'Belanja',
      headers: ['ID Belanja', 'Tanggal', 'Barang', 'Jumlah', 'Harga Beli', 'Total Belanja', 'Pemasok', 'Catatan'],
    },
    fullData: {
      name: 'Data Lengkap',
      headers: ['Tanggal Transaksi', 'ID Hutang', 'Nama', 'Kontak', 'Catatan', 'Status', 'Total Hutang', 'Total Dibayar', 'Sisa Hutang', 'ID Detail', 'Barang', 'Jumlah', 'Harga Jual / Unit', 'Subtotal Jual', 'Harga Beli / Unit', 'Total Modal', 'Laba Kotor'],
    },
  },
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Orbit | Catatan Hutang')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupSpreadsheet() {
  const properties = PropertiesService.getScriptProperties();
  const existingId = properties.getProperty(APP_CONFIG.spreadsheetProperty);
  if (existingId) {
    const existing = SpreadsheetApp.openById(existingId);
    upgradeSpreadsheet_(existing);
    rebuildMonthlyReports_(existing);
    return `Spreadsheet sudah tersedia: ${existing.getUrl()}`;
  }

  const spreadsheet = SpreadsheetApp.create(APP_CONFIG.spreadsheetName);
  properties.setProperty(APP_CONFIG.spreadsheetProperty, spreadsheet.getId());

  const summary = spreadsheet.getSheets()[0];
  summary.setName(APP_CONFIG.sheets.summary.name);
  configureSheet_(summary, []);

  Object.keys(APP_CONFIG.sheets).forEach((key) => {
    if (key !== 'summary') {
      const config = APP_CONFIG.sheets[key];
      configureSheet_(spreadsheet.insertSheet(config.name), config.headers);
    }
  });

  summary.getRange('A1:H1').merge().setValue('ORBIT | RINGKASAN HUTANG');
  summary.getRange('A3').setValue('Total hutang');
  summary.getRange('B3').setFormula("=SUM('Hutang'!E2:E)");
  summary.getRange('D3').setValue('Sudah dibayar');
  summary.getRange('E3').setFormula("=SUM('Hutang'!F2:F)");
  summary.getRange('G3').setValue('Sisa hutang');
  summary.getRange('H3').setFormula("=SUM('Hutang'!G2:G)");
  summary.getRange('A5').setValue('Jumlah catatan');
  summary.getRange('B5').setFormula("=COUNTA('Hutang'!A2:A)");
  summary.getRange('A1:H1')
    .setBackground('#111b2c')
    .setFontColor('#f4f7fb')
    .setFontSize(16)
    .setFontWeight('bold');
  summary.getRange('A3:H5').setFontColor('#223047');
  summary.getRangeList(['A3', 'D3', 'G3', 'A5']).setFontWeight('bold');
  summary.getRangeList(['B3', 'E3', 'H3']).setNumberFormat('"Rp" #,##0');
  summary.setColumnWidths(1, 8, 140);

  upgradeSpreadsheet_(spreadsheet);
  configureDataFormats_(spreadsheet);
  rebuildMonthlyReports_(spreadsheet);
  return `Spreadsheet berhasil dibuat: ${spreadsheet.getUrl()}`;
}

function getAppData() {
  const spreadsheet = getSpreadsheet_();
  const debtRows = readRows_(spreadsheet.getSheetByName(APP_CONFIG.sheets.debts.name));
  const detailRows = readRows_(spreadsheet.getSheetByName(APP_CONFIG.sheets.details.name));
  const purchaseRows = readRows_(spreadsheet.getSheetByName(APP_CONFIG.sheets.purchases.name));
  const customers = readRows_(spreadsheet.getSheetByName(APP_CONFIG.sheets.customers.name))
    .map((row) => ({ id: row[0], name: row[1], contact: row[2], notes: row[3] }));
  const products = readRows_(spreadsheet.getSheetByName(APP_CONFIG.sheets.products.name))
    .map((row) => ({ id: row[0], name: row[1], costPrice: Number(row[2]) || 0, price: Number(row[3]) || 0 }));
  const detailsByDebt = {};

  detailRows.forEach((row) => {
    const debtId = String(row[1]);
    if (!detailsByDebt[debtId]) detailsByDebt[debtId] = [];
    detailsByDebt[debtId].push({
      name: row[2],
      quantity: Number(row[3]) || 0,
      price: Number(row[4]) || 0,
      subtotal: Number(row[5]) || 0,
      costPrice: Number(row[6]) || 0,
      costSubtotal: Number(row[7]) || 0,
    });
  });

  const debts = debtRows.map((row) => {
    const total = Number(row[4]) || 0;
    const paid = Number(row[5]) || 0;
    const remaining = Number(row[6]) || 0;
    return {
      id: String(row[0]),
      date: dateForInput_(row[1]),
      name: String(row[2] || ''),
      notes: String(row[3] || ''),
      total,
      paid,
      remaining,
      status: remaining <= 0 ? 'LUNAS' : (paid > 0 ? 'DICICIL' : 'BELUM LUNAS'),
      updated: dateForDisplay_(row[8]),
      items: detailsByDebt[String(row[0])] || [],
    };
  }).sort((first, second) => second.date.localeCompare(first.date));

  const purchases = purchaseRows.map((row) => ({
    id: String(row[0]),
    date: dateForInput_(row[1]),
    name: String(row[2] || ''),
    quantity: Number(row[3]) || 0,
    costPrice: Number(row[4]) || 0,
    total: Number(row[5]) || 0,
    supplier: String(row[6] || ''),
    notes: String(row[7] || ''),
  })).sort((first, second) => second.date.localeCompare(first.date));

  return {
    debts,
    purchases,
    customers,
    products,
    spreadsheetUrl: spreadsheet.getUrl(),
    summary: debts.reduce((result, debt) => {
      result.total += debt.total;
      result.paid += debt.paid;
      result.remaining += debt.remaining;
      return result;
    }, { total: 0, paid: 0, remaining: 0, profit: detailRows.reduce((sum, row) => sum + (Number(row[5]) || 0) - (Number(row[7]) || 0), 0) }),
  };
}

function saveTransaction(payload) {
  if (!payload || !String(payload.name || '').trim()) throw new Error('Isi nama orang yang berhutang.');
  if (!payload.date || !Array.isArray(payload.items) || payload.items.length === 0) {
    throw new Error('Tanggal dan minimal satu barang wajib diisi.');
  }

  const date = new Date(`${payload.date}T12:00:00`);
  if (Number.isNaN(date.getTime())) throw new Error('Tanggal tidak valid.');
  const items = payload.items.map((item) => {
    const name = String(item.name || '').trim();
    const quantity = Number(item.quantity);
    const price = Number(item.price);
    const costPrice = Number(item.costPrice);
    if (!name || !Number.isFinite(quantity) || quantity <= 0 || !Number.isFinite(price) || price < 0 || !Number.isFinite(costPrice) || costPrice < 0) {
      throw new Error('Periksa nama barang, jumlah, harga jual, dan harga beli.');
    }
    return { name, quantity, price, subtotal: quantity * price, costPrice, costSubtotal: quantity * costPrice };
  });
  const total = items.reduce((sum, item) => sum + item.subtotal, 0);
  const paid = Number(payload.paid) || 0;
  if (!Number.isFinite(paid) || paid < 0 || paid > total) {
    throw new Error('Jumlah dibayar harus antara Rp0 dan total hutang.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const spreadsheet = getSpreadsheet_();
    const debtSheet = spreadsheet.getSheetByName(APP_CONFIG.sheets.debts.name);
    const detailSheet = spreadsheet.getSheetByName(APP_CONFIG.sheets.details.name);
    const paymentSheet = spreadsheet.getSheetByName(APP_CONFIG.sheets.payments.name);
    const customers = readRows_(spreadsheet.getSheetByName(APP_CONFIG.sheets.customers.name));
    const products = readRows_(spreadsheet.getSheetByName(APP_CONFIG.sheets.products.name));
    const name = String(payload.name).trim();
    ensureMasterEntry_(spreadsheet.getSheetByName(APP_CONFIG.sheets.customers.name), customers, name, 'customer');
    items.forEach((item) => ensureMasterEntry_(spreadsheet.getSheetByName(APP_CONFIG.sheets.products.name), products, item.name, 'product', item.costPrice, item.price));

    const debtId = String(payload.id || Utilities.getUuid());
    const currentRows = readRows_(debtSheet);
    const rowIndex = currentRows.findIndex((current) => String(current[0]) === debtId);
    const oldPaid = rowIndex < 0 ? 0 : Number(currentRows[rowIndex][5]) || 0;
    const row = [debtId, date, name, String(payload.notes || '').trim(), total, paid, total - paid, '', new Date()];
    if (rowIndex < 0) {
      debtSheet.appendRow(row);
    } else {
      debtSheet.getRange(rowIndex + 2, 1, 1, row.length).setValues([row]);
    }

    const details = readRows_(detailSheet).filter((current) => String(current[1]) !== debtId);
    items.forEach((item) => details.push([
      Utilities.getUuid(), debtId, item.name, item.quantity, item.price, item.subtotal, item.costPrice, item.costSubtotal,
    ]));
    writeBodyRows_(detailSheet, details);

    const paymentDelta = paid - oldPaid;
    if (paymentDelta !== 0) {
      const paymentDate = rowIndex < 0 ? date : new Date();
      paymentSheet.appendRow([
        Utilities.getUuid(), debtId, paymentDate, paymentDelta,
        paymentDelta < 0 ? 'Koreksi' : (rowIndex < 0 ? 'Pembayaran awal' : 'Pembayaran'),
        paymentDelta < 0 ? 'Koreksi dari perubahan catatan hutang' : '',
      ]);
    }
    rebuildMonthlyReports_(spreadsheet);
    return { id: debtId, message: 'Catatan hutang tersimpan.' };
  } finally {
    lock.releaseLock();
  }
}

function deleteTransaction(debtId) {
  const id = String(debtId || '');
  if (!id) throw new Error('ID hutang tidak ditemukan.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const spreadsheet = getSpreadsheet_();
    const debtSheet = spreadsheet.getSheetByName(APP_CONFIG.sheets.debts.name);
    const debtRows = readRows_(debtSheet);
    const match = debtRows.findIndex((row) => String(row[0]) === id);
    if (match < 0) throw new Error('Catatan hutang sudah tidak ada.');
    debtRows.splice(match, 1);
    writeBodyRows_(debtSheet, debtRows);

    const detailSheet = spreadsheet.getSheetByName(APP_CONFIG.sheets.details.name);
    writeBodyRows_(detailSheet, readRows_(detailSheet).filter((row) => String(row[1]) !== id));
    const paymentSheet = spreadsheet.getSheetByName(APP_CONFIG.sheets.payments.name);
    writeBodyRows_(paymentSheet, readRows_(paymentSheet).filter((row) => String(row[1]) !== id));
    rebuildMonthlyReports_(spreadsheet);
    return { message: 'Catatan hutang dihapus.' };
  } finally {
    lock.releaseLock();
  }
}

function settleCustomerDebts(customerName) {
  const normalizedName = normalizePersonName_(customerName);
  if (!normalizedName) throw new Error('Nama pelanggan tidak ditemukan.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const spreadsheet = getSpreadsheet_();
    const debtSheet = spreadsheet.getSheetByName(APP_CONFIG.sheets.debts.name);
    const debtRows = readRows_(debtSheet);
    const paymentSheet = spreadsheet.getSheetByName(APP_CONFIG.sheets.payments.name);
    const paymentRows = readRows_(paymentSheet);
    const paidAt = new Date();
    let settledCount = 0;
    let settledTotal = 0;

    debtRows.forEach((row) => {
      if (normalizePersonName_(row[2]) !== normalizedName) return;
      const total = Number(row[4]) || 0;
      const paid = Number(row[5]) || 0;
      const remaining = Math.max(0, total - paid);
      if (!remaining) return;
      row[5] = total;
      row[6] = 0;
      row[7] = 'LUNAS';
      row[8] = paidAt;
      paymentRows.push([
        Utilities.getUuid(), row[0], paidAt, remaining, 'Pelunasan gabungan',
        'Pelunasan semua hutang pelanggan melalui website',
      ]);
      settledCount += 1;
      settledTotal += remaining;
    });

    if (!settledCount) throw new Error('Tidak ada sisa hutang untuk nama ini.');
    writeBodyRows_(debtSheet, debtRows);
    writeBodyRows_(paymentSheet, paymentRows);
    rebuildMonthlyReports_(spreadsheet);
    return { settledCount, settledTotal, message: `${settledCount} catatan dilunasi (${formatRupiah_(settledTotal)}).` };
  } finally {
    lock.releaseLock();
  }
}

function savePurchase(payload) {
  if (!payload || !String(payload.name || '').trim()) throw new Error('Pilih atau isi nama barang.');
  if (!payload.date) throw new Error('Tanggal belanja wajib diisi.');
  const date = new Date(`${payload.date}T12:00:00`);
  const quantity = Number(payload.quantity);
  const costPrice = Number(payload.costPrice);
  if (Number.isNaN(date.getTime()) || !Number.isFinite(quantity) || quantity <= 0 || !Number.isFinite(costPrice) || costPrice < 0) {
    throw new Error('Periksa tanggal, jumlah, dan harga beli.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const spreadsheet = getSpreadsheet_();
    const sheet = spreadsheet.getSheetByName(APP_CONFIG.sheets.purchases.name);
    const rows = readRows_(sheet);
    const id = String(payload.id || Utilities.getUuid());
    const row = [id, date, String(payload.name).trim(), quantity, costPrice, quantity * costPrice, String(payload.supplier || '').trim(), String(payload.notes || '').trim()];
    const existingIndex = rows.findIndex((current) => String(current[0]) === id);
    if (existingIndex < 0) rows.push(row);
    else rows[existingIndex] = row;
    writeBodyRows_(sheet, rows);
    rebuildMonthlyReports_(spreadsheet);
    return { message: 'Catatan belanja tersimpan.' };
  } finally {
    lock.releaseLock();
  }
}

function deletePurchase(purchaseId) {
  const id = String(purchaseId || '');
  if (!id) throw new Error('ID belanja tidak ditemukan.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const spreadsheet = getSpreadsheet_();
    const sheet = spreadsheet.getSheetByName(APP_CONFIG.sheets.purchases.name);
    const rows = readRows_(sheet);
    const match = rows.findIndex((row) => String(row[0]) === id);
    if (match < 0) throw new Error('Catatan belanja sudah tidak ada.');
    rows.splice(match, 1);
    writeBodyRows_(sheet, rows);
    rebuildMonthlyReports_(spreadsheet);
    return { message: 'Catatan belanja dihapus.' };
  } finally {
    lock.releaseLock();
  }
}

function saveMaster(type, payload) {
  const isProduct = type === 'product';
  if (!isProduct && type !== 'customer') throw new Error('Jenis data tidak dikenal.');
  const name = String(payload && payload.name || '').trim();
  if (!name) throw new Error('Nama wajib diisi.');
  if (isProduct && (!String(payload.costPrice == null ? '' : payload.costPrice).trim() || !String(payload.price == null ? '' : payload.price).trim())) {
    throw new Error('Harga beli dan harga jual barang wajib diisi.');
  }
  const price = Number(payload && payload.price);
  const costPrice = Number(payload && payload.costPrice);
  if (isProduct && (!Number.isFinite(price) || price < 0 || !Number.isFinite(costPrice) || costPrice < 0)) {
    throw new Error('Harga beli dan harga jual harus berupa angka nol atau lebih.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const spreadsheet = getSpreadsheet_();
    const config = isProduct ? APP_CONFIG.sheets.products : APP_CONFIG.sheets.customers;
    const sheet = spreadsheet.getSheetByName(config.name);
    const rows = readRows_(sheet);
    const id = String(payload.id || Utilities.getUuid());
    const duplicate = rows.find((row) => String(row[0]) !== id && String(row[1]).toLowerCase() === name.toLowerCase());
    if (duplicate) throw new Error('Nama tersebut sudah ada di daftar.');

    const row = isProduct
      ? [id, name, costPrice, price, new Date()]
      : [id, name, String(payload.contact || '').trim(), String(payload.notes || '').trim()];
    const existingIndex = rows.findIndex((current) => String(current[0]) === id);
    if (existingIndex < 0) rows.push(row);
    else rows[existingIndex] = row;
    writeBodyRows_(sheet, rows);
    if (isProduct) rebuildMonthlyReports_(spreadsheet);
    return { message: isProduct ? 'Data barang tersimpan.' : 'Data pelanggan tersimpan.' };
  } finally {
    lock.releaseLock();
  }
}

function upgradeSpreadsheet_(spreadsheet) {
  Object.keys(APP_CONFIG.sheets).forEach((key) => {
    const config = APP_CONFIG.sheets[key];
    if (key === 'summary') return;
    let sheet = spreadsheet.getSheetByName(config.name);
    if (!sheet) sheet = spreadsheet.insertSheet(config.name);
    if (sheet.getLastRow() === 0) configureSheet_(sheet, config.headers);
  });

  const productSheet = spreadsheet.getSheetByName(APP_CONFIG.sheets.products.name);
  const productHeaders = productSheet.getRange(1, 1, 1, productSheet.getLastColumn()).getDisplayValues()[0];
  if (productHeaders[2] === 'Harga' && productHeaders[3] === 'Diperbarui') {
    const oldRows = readRows_(productSheet);
    productSheet.clear();
    configureSheet_(productSheet, APP_CONFIG.sheets.products.headers);
    if (oldRows.length) {
      const upgradedRows = oldRows.map((row) => [row[0], row[1], 0, Number(row[2]) || 0, row[3] || '']);
      productSheet.getRange(2, 1, upgradedRows.length, upgradedRows[0].length).setValues(upgradedRows);
    }
  }

  const detailSheet = spreadsheet.getSheetByName(APP_CONFIG.sheets.details.name);
  ensureTrailingHeaders_(detailSheet, APP_CONFIG.sheets.details.headers);
  const paymentSheet = spreadsheet.getSheetByName(APP_CONFIG.sheets.payments.name);
  ensureTrailingHeaders_(paymentSheet, APP_CONFIG.sheets.payments.headers);

  const properties = PropertiesService.getScriptProperties();
  const migrationKey = 'DEBT_TRACKER_PAYMENT_MIGRATION_V1';
  if (properties.getProperty(migrationKey) !== 'done') {
    const existingPaymentDebtIds = new Set(readRows_(paymentSheet).map((row) => String(row[1])));
    const legacyPayments = readRows_(spreadsheet.getSheetByName(APP_CONFIG.sheets.debts.name))
      .filter((row) => (Number(row[5]) || 0) > 0 && !existingPaymentDebtIds.has(String(row[0])))
      .map((row) => [Utilities.getUuid(), row[0], row[1], Number(row[5]), 'Migrasi', 'Pembayaran lama; tanggal memakai tanggal transaksi']);
    if (legacyPayments.length) {
      paymentSheet.getRange(paymentSheet.getLastRow() + 1, 1, legacyPayments.length, legacyPayments[0].length).setValues(legacyPayments);
    }
    properties.setProperty(migrationKey, 'done');
  }

  const summary = spreadsheet.getSheetByName(APP_CONFIG.sheets.summary.name);
  const titleRange = summary.getRange('A1:H1');
  if (!titleRange.isPartOfMerge()) titleRange.merge();
  summary.getRange('A1').setValue('ORBIT | RINGKASAN HUTANG');
  summary.getRange('A3').setValue('Total hutang');
  summary.getRange('B3').setFormula("=SUM('Hutang'!E2:E)");
  summary.getRange('D3').setValue('Sudah dibayar');
  summary.getRange('E3').setFormula("=SUM('Hutang'!F2:F)");
  summary.getRange('G3').setValue('Sisa hutang');
  summary.getRange('H3').setFormula("=SUM('Hutang'!G2:G)");
  summary.getRange('A5').setValue('Jumlah catatan');
  summary.getRange('B5').setFormula("=COUNTA('Hutang'!A2:A)");
  summary.getRange('A1:H1').setBackground('#111b2c').setFontColor('#f4f7fb').setFontSize(16).setFontWeight('bold');
  summary.getRangeList(['A3', 'D3', 'G3', 'A5']).setFontWeight('bold');
  summary.setColumnWidths(1, 8, 140);
  configureDataFormats_(spreadsheet);
}

function ensureTrailingHeaders_(sheet, requiredHeaders) {
  const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  requiredHeaders.forEach((header) => {
    if (currentHeaders.indexOf(header) >= 0) return;
    const nextColumn = sheet.getLastColumn() + 1;
    sheet.getRange(1, nextColumn)
      .setValue(header)
      .setBackground('#111b2c')
      .setFontColor('#f4f7fb')
      .setFontWeight('bold');
  });
  sheet.setFrozenRows(1);
}

function rebuildMonthlyReports_(spreadsheet) {
  const debtRows = readRows_(spreadsheet.getSheetByName(APP_CONFIG.sheets.debts.name));
  const detailRows = readRows_(spreadsheet.getSheetByName(APP_CONFIG.sheets.details.name));
  const paymentRows = readRows_(spreadsheet.getSheetByName(APP_CONFIG.sheets.payments.name));
  const purchaseRows = readRows_(spreadsheet.getSheetByName(APP_CONFIG.sheets.purchases.name));
  rebuildFullData_(spreadsheet, debtRows, detailRows);
  const currentMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const monthKeys = debtRows.map((row) => monthKey_(row[1]))
    .concat(paymentRows.map((row) => monthKey_(row[2])))
    .concat(purchaseRows.map((row) => monthKey_(row[1])))
    .concat(spreadsheet.getSheets().map((sheet) => sheet.getName()).filter((name) => /^\d{4}-\d{2}$/.test(name)))
    .filter(Boolean);
  monthKeys.push(currentMonth);
  monthKeys.sort();

  const firstMonth = monthKeys[0] || currentMonth;
  const lastMonth = monthKeys[monthKeys.length - 1] || currentMonth;
  const monthKeysToBuild = [];
  const monthCursor = firstOfMonth_(firstMonth);
  const monthEnd = firstOfMonth_(lastMonth);
  while (monthCursor <= monthEnd) {
    monthKeysToBuild.push(Utilities.formatDate(monthCursor, Session.getScriptTimeZone(), 'yyyy-MM'));
    monthCursor.setMonth(monthCursor.getMonth() + 1);
  }

  const detailsByDebt = {};
  detailRows.forEach((row) => {
    const debtId = String(row[1]);
    if (!detailsByDebt[debtId]) detailsByDebt[debtId] = [];
    detailsByDebt[debtId].push(row);
  });

  monthKeysToBuild.forEach((key) => {
    const report = spreadsheet.getSheetByName(key) || spreadsheet.insertSheet(key);
    renderMonthlyReport_(report, key, debtRows, detailsByDebt, paymentRows, purchaseRows);
  });
}

function rebuildFullData_(spreadsheet, debtRows, detailRows) {
  const sheet = spreadsheet.getSheetByName(APP_CONFIG.sheets.fullData.name);
  const customerRows = readRows_(spreadsheet.getSheetByName(APP_CONFIG.sheets.customers.name));
  const customerContacts = {};
  customerRows.forEach((row) => {
    customerContacts[normalizePersonName_(row[1])] = row[2] || '';
  });
  const detailsByDebt = {};
  detailRows.forEach((row) => {
    const id = String(row[1]);
    if (!detailsByDebt[id]) detailsByDebt[id] = [];
    detailsByDebt[id].push(row);
  });

  const rows = [];
  debtRows.forEach((debt) => {
    const details = detailsByDebt[String(debt[0])] || [null];
    const totalDebt = Number(debt[4]) || 0;
    const paid = Number(debt[5]) || 0;
    const remaining = Number(debt[6]) || 0;
    const status = remaining <= 0 ? 'LUNAS' : (paid > 0 ? 'DICICIL' : 'BELUM LUNAS');
    details.forEach((detail, index) => {
      const subtotal = detail ? Number(detail[5]) || 0 : 0;
      const costSubtotal = detail ? Number(detail[7]) || 0 : 0;
      rows.push([
        debt[1], debt[0], debt[2], customerContacts[normalizePersonName_(debt[2])] || '', debt[3], status,
        index === 0 ? totalDebt : '',
        index === 0 ? paid : '',
        index === 0 ? remaining : '',
        detail ? detail[0] : '', detail ? detail[2] : '', detail ? Number(detail[3]) || 0 : '',
        detail ? Number(detail[4]) || 0 : '', subtotal,
        detail ? Number(detail[6]) || 0 : '', costSubtotal, subtotal - costSubtotal,
      ]);
    });
  });

  writeBodyRows_(sheet, rows);
  sheet.getRange('A2:A').setNumberFormat('dd mmm yyyy');
  sheet.getRange('G2:I').setNumberFormat('"Rp" #,##0');
  sheet.getRange('M2:Q').setNumberFormat('"Rp" #,##0');
  sheet.setFrozenRows(1);
}

function renderMonthlyReport_(sheet, key, debtRows, detailsByDebt, paymentRows, purchaseRows) {
  const start = firstOfMonth_(key);
  const end = new Date(start.getFullYear(), start.getMonth() + 1, 0, 23, 59, 59, 999);
  const monthDebts = debtRows.filter((row) => {
    const date = asDate_(row[1]);
    return date && date >= start && date <= end;
  });
  const monthPayments = paymentRows.filter((row) => {
    const date = asDate_(row[2]);
    return date && date >= start && date <= end;
  });
  const monthPurchases = purchaseRows.filter((row) => {
    const date = asDate_(row[1]);
    return date && date >= start && date <= end;
  });
  const allDetails = [];
  monthDebts.forEach((debt) => (detailsByDebt[String(debt[0])] || []).forEach((detail) => allDetails.push([
    dateForInput_(debt[1]), debt[0], debt[2], detail[2], Number(detail[3]) || 0,
    Number(detail[4]) || 0, Number(detail[5]) || 0, Number(detail[6]) || 0, Number(detail[7]) || 0,
    (Number(detail[5]) || 0) - (Number(detail[7]) || 0),
  ])));

  const paymentsByDebt = {};
  paymentRows.forEach((row) => {
    const paymentDate = asDate_(row[2]);
    if (!paymentDate || paymentDate > end) return;
    const debtId = String(row[1]);
    paymentsByDebt[debtId] = (paymentsByDebt[debtId] || 0) + (Number(row[3]) || 0);
  });
  const oldOutstanding = debtRows.filter((row) => {
    const date = asDate_(row[1]);
    return date && date < start && date <= end && (Number(row[4]) || 0) - (paymentsByDebt[String(row[0])] || 0) > 0;
  }).map((row) => {
    const total = Number(row[4]) || 0;
    const paid = paymentsByDebt[String(row[0])] || 0;
    return [monthKey_(row[1]), dateForInput_(row[1]), row[0], row[2], total, paid, Math.max(0, total - paid)];
  });
  const monthEndOutstanding = debtRows.reduce((sum, row) => {
    const date = asDate_(row[1]);
    if (!date || date > end) return sum;
    return sum + Math.max(0, (Number(row[4]) || 0) - (paymentsByDebt[String(row[0])] || 0));
  }, 0);
  const sales = monthDebts.reduce((sum, row) => sum + (Number(row[4]) || 0), 0);
  const cost = allDetails.reduce((sum, row) => sum + row[8], 0);
  const cashReceived = monthPayments.reduce((sum, row) => sum + (Number(row[3]) || 0), 0);
  const purchaseTotal = monthPurchases.reduce((sum, row) => sum + (Number(row[5]) || 0), 0);
  const monthLabel = Utilities.formatDate(start, Session.getScriptTimeZone(), 'MMMM yyyy');
  const carryStart = Math.max(9, 9 + allDetails.length);
  const purchaseStart = carryStart + Math.max(oldOutstanding.length, 1) + 3;

  sheet.clear();
  sheet.getRange('A1:L1').merge().setValue(`ORBIT | LAPORAN BULANAN — ${monthLabel.toUpperCase()}`);
  sheet.getRange('A3:L3').setValues([[
    'Penjualan', '', 'Modal terjual', '', 'Laba kotor', '', 'Kas masuk', '', 'Belanja barang', '', 'Sisa hutang', '',
  ]]);
  sheet.getRange('A4:L4').setValues([[
    null, sales, null, cost, null, sales - cost, null, cashReceived, null, purchaseTotal, null, monthEndOutstanding,
  ]]);
  sheet.getRange('A6:J6').merge().setValue('RINCIAN PENJUALAN BULAN INI');
  sheet.getRange('A7:J7').setValues([[
    'Tanggal', 'ID Hutang', 'Nama', 'Barang', 'Jumlah', 'Harga Jual', 'Subtotal Jual', 'Harga Beli', 'Total Modal', 'Laba Kotor',
  ]]);
  if (allDetails.length) sheet.getRange(8, 1, allDetails.length, 10).setValues(allDetails);

  sheet.getRange(carryStart, 1, 1, 7).merge().setValue('HUTANG DARI BULAN SEBELUMNYA YANG MASIH TERSISA');
  sheet.getRange(carryStart + 1, 1, 1, 7).setValues([[
    'Bulan Asal', 'Tanggal', 'ID Hutang', 'Nama', 'Total Awal', 'Dibayar s/d Akhir Bulan', 'Sisa Akhir Bulan',
  ]]);
  if (oldOutstanding.length) sheet.getRange(carryStart + 2, 1, oldOutstanding.length, 7).setValues(oldOutstanding);
  else sheet.getRange(carryStart + 2, 1).setValue('Tidak ada hutang bawaan yang masih tersisa.');

  sheet.getRange(purchaseStart, 1, 1, 8).merge().setValue('RINCIAN BELANJA BARANG BULAN INI');
  sheet.getRange(purchaseStart + 1, 1, 1, 8).setValues([[
    'Tanggal', 'ID Belanja', 'Barang', 'Jumlah', 'Harga Beli', 'Total Belanja', 'Pemasok', 'Catatan',
  ]]);
  const monthlyPurchaseDetails = monthPurchases.map((row) => [
    dateForInput_(row[1]), row[0], row[2], Number(row[3]) || 0, Number(row[4]) || 0,
    Number(row[5]) || 0, row[6], row[7],
  ]);
  if (monthlyPurchaseDetails.length) {
    sheet.getRange(purchaseStart + 2, 1, monthlyPurchaseDetails.length, 8).setValues(monthlyPurchaseDetails);
  } else {
    sheet.getRange(purchaseStart + 2, 1).setValue('Tidak ada belanja barang yang dicatat bulan ini.');
  }

  sheet.getRange('A1:L1').setBackground('#111b2c').setFontColor('#f4f7fb').setFontSize(15).setFontWeight('bold');
  sheet.getRange('A3:L3').setBackground('#172437').setFontColor('#9aaabd').setFontWeight('bold');
  sheet.getRange('A4:L4').setBackground('#121c2b').setFontColor('#9ce4c5').setFontWeight('bold');
  sheet.getRange('A6:J6').setBackground('#172437').setFontColor('#9ce4c5').setFontWeight('bold');
  sheet.getRange('A7:J7').setBackground('#26384b').setFontColor('#f4f7fb').setFontWeight('bold');
  sheet.getRange(carryStart, 1, 1, 7).setBackground('#172437').setFontColor('#f2c77c').setFontWeight('bold');
  sheet.getRange(carryStart + 1, 1, 1, 7).setBackground('#26384b').setFontColor('#f4f7fb').setFontWeight('bold');
  sheet.getRange(purchaseStart, 1, 1, 8).setBackground('#172437').setFontColor('#9ce4c5').setFontWeight('bold');
  sheet.getRange(purchaseStart + 1, 1, 1, 8).setBackground('#26384b').setFontColor('#f4f7fb').setFontWeight('bold');
  sheet.getRange('B4:L4').setNumberFormat('"Rp" #,##0');
  sheet.getRange('F8:J').setNumberFormat('"Rp" #,##0');
  sheet.getRange(carryStart + 2, 5, Math.max(1, oldOutstanding.length), 3).setNumberFormat('"Rp" #,##0');
  sheet.getRange(purchaseStart + 2, 5, Math.max(1, monthlyPurchaseDetails.length), 2).setNumberFormat('"Rp" #,##0');
  sheet.setFrozenRows(7);
  sheet.setColumnWidths(1, 12, 135);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 170);
}

function monthKey_(value) {
  const date = asDate_(value);
  return date ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM') : '';
}

function asDate_(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (!value) return null;
  const text = String(value);
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12);
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function firstOfMonth_(key) {
  const parts = key.split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, 1);
}

function normalizePersonName_(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toLocaleLowerCase();
}

function formatRupiah_(value) {
  return `Rp${new Intl.NumberFormat('id-ID', { maximumFractionDigits: 0 }).format(value)}`;
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(APP_CONFIG.spreadsheetProperty);
  if (!id) throw new Error('Spreadsheet belum dibuat. Jalankan setupSpreadsheet() satu kali dari editor Apps Script.');
  return SpreadsheetApp.openById(id);
}

function configureSheet_(sheet, headers) {
  sheet.clear();
  if (headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#111b2c')
      .setFontColor('#f4f7fb')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
}

function configureDataFormats_(spreadsheet) {
  spreadsheet.getSheetByName(APP_CONFIG.sheets.debts.name).getRange('E2:G').setNumberFormat('"Rp" #,##0');
  spreadsheet.getSheetByName(APP_CONFIG.sheets.debts.name).getRange('B2:B').setNumberFormat('dd mmm yyyy');
  spreadsheet.getSheetByName(APP_CONFIG.sheets.details.name).getRange('E2:F').setNumberFormat('"Rp" #,##0');
  spreadsheet.getSheetByName(APP_CONFIG.sheets.details.name).getRange('G2:H').setNumberFormat('"Rp" #,##0');
  spreadsheet.getSheetByName(APP_CONFIG.sheets.products.name).getRange('C2:D').setNumberFormat('"Rp" #,##0');
  spreadsheet.getSheetByName(APP_CONFIG.sheets.payments.name).getRange('C2:C').setNumberFormat('dd mmm yyyy');
  spreadsheet.getSheetByName(APP_CONFIG.sheets.payments.name).getRange('D2:D').setNumberFormat('"Rp" #,##0');
  spreadsheet.getSheetByName(APP_CONFIG.sheets.purchases.name).getRange('B2:B').setNumberFormat('dd mmm yyyy');
  spreadsheet.getSheetByName(APP_CONFIG.sheets.purchases.name).getRange('E2:F').setNumberFormat('"Rp" #,##0');
  spreadsheet.getSheetByName(APP_CONFIG.sheets.fullData.name).getRange('G2:I').setNumberFormat('"Rp" #,##0');
  spreadsheet.getSheetByName(APP_CONFIG.sheets.fullData.name).getRange('M2:Q').setNumberFormat('"Rp" #,##0');
  spreadsheet.getSheetByName(APP_CONFIG.sheets.summary.name).getRange('B3').setNumberFormat('"Rp" #,##0');
  spreadsheet.getSheetByName(APP_CONFIG.sheets.summary.name).getRange('E3').setNumberFormat('"Rp" #,##0');
  spreadsheet.getSheetByName(APP_CONFIG.sheets.summary.name).getRange('H3').setNumberFormat('"Rp" #,##0');
}

function readRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
}

function writeBodyRows_(sheet, rows) {
  const lastRow = sheet.getLastRow();
  const columnCount = sheet.getLastColumn();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, columnCount).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, columnCount).setValues(rows);
}

function ensureMasterEntry_(sheet, rows, name, type, costPrice, salePrice) {
  const exists = rows.some((row) => String(row[1]).toLowerCase() === name.toLowerCase());
  if (exists) return;
  if (type === 'product') sheet.appendRow([Utilities.getUuid(), name, costPrice, salePrice, new Date()]);
  else sheet.appendRow([Utilities.getUuid(), name, '', '']);
  rows.push([null, name]);
}

function dateForInput_(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value || '').slice(0, 10);
}

function dateForDisplay_(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd MMM yyyy');
  }
  return String(value || '');
}
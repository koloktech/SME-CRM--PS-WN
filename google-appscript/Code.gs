/**
 * Wan Boutique Order Manager API
 *
 * 1. Paste this file into Apps Script attached to the target Google Sheet.
 * 2. Run setupDatabase() once and authorize it.
 * 3. Deploy as a Web App, then paste the /exec URL into the web app Settings page.
 */

const SHEETS = {
  ORDERS: 'Orders', ITEMS: 'OrderItems', CUSTOMERS: 'Customers', PAYMENTS: 'Payments',
  SETTINGS: 'Settings', AUDIT: 'AuditLog', LEGACY: 'LegacyImport'
};
const HEADERS = {
  Orders: ['Order ID','Invoice Number','Created Date','Updated Date','Customer ID','Customer Name','Phone','Address','Country','Currency','Items Total','Order Service Fee','Shipping Fee','Discount','Grand Total','Paid Amount','Balance','Payment Status','Payment Method','Order Status','Notes','Version','Deleted'],
  OrderItems: ['Order Item ID','Order ID','Brand','Product Name','Colour','Size','Quantity','Pricing Mode','Unit Price','Merchandise Total','Service Fee','Line Total','Notes','Created Date'],
  Customers: ['Customer ID','Name','Normalized Phone','Display Phone','Address','Country','Created Date','Updated Date'],
  Payments: ['Payment ID','Order ID','Payment Date','Amount','Method','Reference','Status','Created Date'],
  Settings: ['Key','Value','Description'],
  AuditLog: ['Audit ID','Timestamp','Action','Entity Type','Entity ID','Details JSON'],
  LegacyImport: ['Legacy Row','Legacy ID','Date','Name','Contact No','Brand','Item','Size','Remark','Imported At']
};

function doGet(e) {
  return respond_(handle_(e && e.parameter ? e.parameter.action : 'GET_BOOTSTRAP', e && e.parameter ? e.parameter : {}));
}

function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : (e.parameter || {});
    return respond_(handle_(body.action, body.payload || body, body.token));
  } catch (error) {
    return respond_({ success: false, error: { message: error.message || String(error) } });
  }
}

function respond_(body) {
  return ContentService.createTextOutput(JSON.stringify(body)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(action, payload, token) {
  try {
    ensureAuthorized_(token);
    switch (String(action || 'GET_BOOTSTRAP').toUpperCase()) {
      case 'GET_BOOTSTRAP': return ok_(getBootstrap_());
      case 'LIST_ORDERS': return ok_(getBootstrap_());
      case 'GET_ORDER': return ok_(getOrder_(payload.orderId));
      case 'CREATE_ORDER': return ok_(createOrder_(payload));
      case 'UPDATE_ORDER': return ok_(updateOrder_(payload));
      case 'DELETE_ORDER': return ok_(deleteOrder_(payload.orderId));
      case 'CREATE_PAYMENT': return ok_(createPayment_(payload));
      case 'SEARCH_CUSTOMERS': return ok_(searchCustomers_(payload.query));
      case 'GET_DASHBOARD': return ok_(getDashboard_());
      case 'SAVE_SETTINGS': return ok_(saveSettings_(payload));
      default: throw new Error('Unsupported action: ' + action);
    }
  } catch (error) {
    return { success: false, error: { message: error.message || String(error) } };
  }
}

function ok_(data) { return { success: true, data: data }; }

/** Create all CRM tabs and headers. Safe to run again; existing data is retained. */
function setupDatabase() {
  const spreadsheet = getSpreadsheet_();
  Object.keys(HEADERS).forEach(function(name) {
    const sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
    if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS[name]);
    else if (sheet.getRange(1, 1, 1, HEADERS[name].length).getValues()[0].join('|') !== HEADERS[name].join('|')) {
      throw new Error('The ' + name + ' sheet has unexpected headers. Do not overwrite data; create a new sheet or restore the expected headers.');
    }
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold').setBackground('#8f3054').setFontColor('#ffffff');
    sheet.autoResizeColumns(1, HEADERS[name].length);
  });
  const settings = getRows_(SHEETS.SETTINGS);
  if (!settings.length) {
    appendRows_(SHEETS.SETTINGS, [
      ['BUSINESS_NAME', 'Wan Boutique', 'Displayed on invoices'],
      ['INVOICE_PREFIX', 'WB', 'Prefix for generated invoice numbers'],
      ['DEFAULT_CURRENCY', 'RM', 'Display currency'],
      ['SCHEMA_VERSION', '1', 'Do not edit'],
      ['APP_TOKEN', '', 'Optional simple request token; do not put a secret in public frontend code']
    ]);
  }
  return 'Database setup complete in ' + spreadsheet.getName();
}

/** Run in a standalone Apps Script project if it is not bound to the target Sheet. */
function setSpreadsheetId(spreadsheetId) {
  if (!spreadsheetId) throw new Error('A Google Sheet ID is required.');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheetId);
  return 'Spreadsheet ID saved.';
}

function getBootstrap_() {
  const orders = getRows_(SHEETS.ORDERS).filter(function(order) { return !toBoolean_(order['Deleted']); }).map(orderToClient_);
  const items = getRows_(SHEETS.ITEMS).map(itemToClient_);
  const payments = getRows_(SHEETS.PAYMENTS).map(paymentToClient_);
  const customers = getRows_(SHEETS.CUSTOMERS).map(customerToClient_);
  const settings = getSettings_();
  const paymentTotals = {};
  payments.forEach(function(payment) {
    const sign = payment.status === 'Refunded' ? -1 : 1;
    paymentTotals[payment.orderId] = (paymentTotals[payment.orderId] || 0) + sign * asNumber_(payment.amount);
  });
  orders.forEach(function(order) {
    order.paidAmount = paymentTotals[order.orderId] || 0;
    order.balance = Math.max(0, asNumber_(order.grandTotal) - order.paidAmount);
    order.paymentStatus = paymentStatus_(order.grandTotal, order.paidAmount);
  });
  return { orders: orders, items: items, payments: payments, customers: customers, settings: settings };
}

function getOrder_(orderId) {
  if (!orderId) throw new Error('Order ID is required.');
  const data = getBootstrap_();
  const order = data.orders.filter(function(row) { return row.orderId === orderId; })[0];
  if (!order) throw new Error('Order not found.');
  return { order: order, items: data.items.filter(function(row) { return row.orderId === orderId; }), payments: data.payments.filter(function(row) { return row.orderId === orderId; }) };
}

function createOrder_(payload) {
  validateOrder_(payload);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const calc = calculateOrder_(payload);
    const now = new Date();
    const customer = upsertCustomer_(payload, now);
    const orderId = Utilities.getUuid();
    const invoiceNumber = nextInvoiceNumber_(now);
    const orderRow = [orderId, invoiceNumber, now, now, customer.customerId, clean_(payload.customerName), normalizePhone_(payload.phone), clean_(payload.address), clean_(payload.country || 'Malaysia'), 'RM', calc.itemsTotal, calc.orderServiceFee, calc.shippingFee, calc.discount, calc.grandTotal, 0, calc.grandTotal, 'Unpaid', clean_(payload.paymentMethod), clean_(payload.orderStatus || 'Draft'), clean_(payload.notes), 1, false];
    appendRows_(SHEETS.ORDERS, [orderRow]);
    const itemRows = calc.items.map(function(item) { return [Utilities.getUuid(), orderId, clean_(item.brand), clean_(item.productName), clean_(item.colour), clean_(item.size), item.quantity, item.pricingMode, item.unitPrice, item.merchandiseTotal, item.serviceFee, item.lineTotal, clean_(item.notes), now]; });
    appendRows_(SHEETS.ITEMS, itemRows);
    if (asNumber_(payload.paidAmount) > 0) appendRows_(SHEETS.PAYMENTS, [[Utilities.getUuid(), orderId, now, asNumber_(payload.paidAmount), clean_(payload.paymentMethod || 'Bank Transfer'), 'Initial payment', 'Paid', now]]);
    refreshOrderPayment_(orderId);
    audit_('CREATE_ORDER', 'Order', orderId, { invoiceNumber: invoiceNumber, itemCount: itemRows.length });
    return getOrder_(orderId);
  } finally { lock.releaseLock(); }
}

function updateOrder_(payload) {
  if (!payload.orderId) throw new Error('Order ID is required.');
  validateOrder_(payload);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const orderSheet = getSheet_(SHEETS.ORDERS); const rowIndex = findRow_(orderSheet, 'Order ID', payload.orderId);
    if (!rowIndex) throw new Error('Order not found.');
    const existing = rowToObject_(orderSheet.getRange(rowIndex, 1, 1, HEADERS.Orders.length).getValues()[0], HEADERS.Orders);
    if (String(existing['Version']) !== String(payload.version)) throw new Error('This order was changed elsewhere. Refresh and try again.');
    const calc = calculateOrder_(payload); const now = new Date(); const customer = upsertCustomer_(payload, now);
    const row = [payload.orderId, existing['Invoice Number'], existing['Created Date'], now, customer.customerId, clean_(payload.customerName), normalizePhone_(payload.phone), clean_(payload.address), clean_(payload.country || 'Malaysia'), existing['Currency'] || 'RM', calc.itemsTotal, calc.orderServiceFee, calc.shippingFee, calc.discount, calc.grandTotal, existing['Paid Amount'] || 0, Math.max(0, calc.grandTotal - asNumber_(existing['Paid Amount'])), paymentStatus_(calc.grandTotal, existing['Paid Amount']), clean_(payload.paymentMethod), clean_(payload.orderStatus || 'Draft'), clean_(payload.notes), asNumber_(existing['Version']) + 1, false];
    orderSheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    const itemSheet = getSheet_(SHEETS.ITEMS); const allItems = getRows_(SHEETS.ITEMS); const keep = allItems.filter(function(item) { return item['Order ID'] !== payload.orderId; }).map(function(item) { return HEADERS.OrderItems.map(function(header) { return item[header]; }); });
    itemSheet.clearContents(); itemSheet.getRange(1, 1, 1, HEADERS.OrderItems.length).setValues([HEADERS.OrderItems]); if (keep.length) itemSheet.getRange(2, 1, keep.length, HEADERS.OrderItems.length).setValues(keep);
    appendRows_(SHEETS.ITEMS, calc.items.map(function(item) { return [Utilities.getUuid(), payload.orderId, clean_(item.brand), clean_(item.productName), clean_(item.colour), clean_(item.size), item.quantity, item.pricingMode, item.unitPrice, item.merchandiseTotal, item.serviceFee, item.lineTotal, clean_(item.notes), now]; }));
    refreshOrderPayment_(payload.orderId); audit_('UPDATE_ORDER', 'Order', payload.orderId, { version: asNumber_(existing['Version']) + 1 }); return getOrder_(payload.orderId);
  } finally { lock.releaseLock(); }
}

function deleteOrder_(orderId) {
  const sheet = getSheet_(SHEETS.ORDERS); const row = findRow_(sheet, 'Order ID', orderId); if (!row) throw new Error('Order not found.');
  sheet.getRange(row, HEADERS.Orders.indexOf('Deleted') + 1).setValue(true); sheet.getRange(row, HEADERS.Orders.indexOf('Updated Date') + 1).setValue(new Date()); audit_('DELETE_ORDER', 'Order', orderId, {}); return { orderId: orderId };
}

function createPayment_(payload) {
  if (!payload.orderId) throw new Error('Order is required.'); if (asNumber_(payload.amount) <= 0) throw new Error('Payment amount must be greater than zero.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const order = getOrder_(payload.orderId).order; if (!order) throw new Error('Order not found.'); const now = new Date(); const id = Utilities.getUuid();
    appendRows_(SHEETS.PAYMENTS, [[id, payload.orderId, now, asNumber_(payload.amount), clean_(payload.method || 'Bank Transfer'), clean_(payload.reference), clean_(payload.status || 'Paid'), now]]);
    refreshOrderPayment_(payload.orderId); audit_('CREATE_PAYMENT', 'Payment', id, { orderId: payload.orderId, amount: asNumber_(payload.amount) }); return { payment: paymentToClient_(rowToObject_([id, payload.orderId, now, asNumber_(payload.amount), clean_(payload.method || 'Bank Transfer'), clean_(payload.reference), clean_(payload.status || 'Paid'), now], HEADERS.Payments)), order: getOrder_(payload.orderId).order };
  } finally { lock.releaseLock(); }
}

function refreshOrderPayment_(orderId) {
  const sheet = getSheet_(SHEETS.ORDERS); const row = findRow_(sheet, 'Order ID', orderId); if (!row) return;
  const order = rowToObject_(sheet.getRange(row, 1, 1, HEADERS.Orders.length).getValues()[0], HEADERS.Orders);
  const paid = getRows_(SHEETS.PAYMENTS).filter(function(payment) { return payment['Order ID'] === orderId; }).reduce(function(sum, payment) { return sum + (payment['Status'] === 'Refunded' ? -1 : 1) * asNumber_(payment['Amount']); }, 0);
  const total = asNumber_(order['Grand Total']); sheet.getRange(row, HEADERS.Orders.indexOf('Paid Amount') + 1, 1, 3).setValues([[paid, Math.max(0, total - paid), paymentStatus_(total, paid)]]);
}

function calculateOrder_(payload) {
  const items = (payload.items || []).map(function(item) { const quantity = Math.max(1, Math.floor(asNumber_(item.quantity) || 1)); const pricingMode = item.pricingMode === 'UNIT' ? 'UNIT' : 'TOTAL'; const unitPrice = asNumber_(item.unitPrice); const merchandiseTotal = pricingMode === 'UNIT' ? quantity * unitPrice : asNumber_(item.merchandiseTotal); const serviceFee = asNumber_(item.serviceFee); return { brand: item.brand, productName: item.productName, colour: item.colour, size: item.size, quantity: quantity, pricingMode: pricingMode, unitPrice: unitPrice, merchandiseTotal: merchandiseTotal, serviceFee: serviceFee, lineTotal: merchandiseTotal + serviceFee, notes: item.notes }; });
  const itemsTotal = items.reduce(function(sum, item) { return sum + item.merchandiseTotal; }, 0); const itemServiceFees = items.reduce(function(sum, item) { return sum + item.serviceFee; }, 0); const orderServiceFee = asNumber_(payload.orderServiceFee); const shippingFee = asNumber_(payload.shippingFee); const discount = asNumber_(payload.discount); return { items: items, itemsTotal: itemsTotal, orderServiceFee: orderServiceFee, shippingFee: shippingFee, discount: discount, grandTotal: Math.max(0, itemsTotal + itemServiceFees + orderServiceFee + shippingFee - discount) };
}

function validateOrder_(payload) { if (!clean_(payload.customerName)) throw new Error('Customer name is required.'); if (!payload.items || !payload.items.length) throw new Error('At least one item is required.'); payload.items.forEach(function(item, index) { if (!clean_(item.productName)) throw new Error('Item ' + (index + 1) + ' needs a product name.'); }); }
function upsertCustomer_(payload, now) { const phone = normalizePhone_(payload.phone); const sheet = getSheet_(SHEETS.CUSTOMERS); const rows = getRows_(SHEETS.CUSTOMERS); let row = phone ? rows.filter(function(customer) { return customer['Normalized Phone'] === phone; })[0] : null; const data = [row ? row['Customer ID'] : Utilities.getUuid(), clean_(payload.customerName), phone, clean_(payload.phone), clean_(payload.address), clean_(payload.country || 'Malaysia'), row ? row['Created Date'] : now, now]; if (row) { const index = findRow_(sheet, 'Customer ID', row['Customer ID']); sheet.getRange(index, 1, 1, data.length).setValues([data]); } else appendRows_(SHEETS.CUSTOMERS, [data]); return { customerId: data[0] }; }
function nextInvoiceNumber_(date) { const props = PropertiesService.getScriptProperties(); const timezone = getSpreadsheet_().getSpreadsheetTimeZone() || Session.getScriptTimeZone(); const day = Utilities.formatDate(date, timezone, 'yyyyMMdd'); const key = 'INVOICE_SEQUENCE_' + day; const next = asNumber_(props.getProperty(key)) + 1; props.setProperty(key, String(next)); const prefix = getSettings_().INVOICE_PREFIX || 'WB'; return prefix + '-' + day + '-' + ('000' + next).slice(-3); }
function searchCustomers_(query) { const term = clean_(query).toLowerCase(); return getRows_(SHEETS.CUSTOMERS).map(customerToClient_).filter(function(customer) { return !term || [customer.name, customer.phone, customer.address].join(' ').toLowerCase().indexOf(term) > -1; }).slice(0, 25); }
function getDashboard_() { const data = getBootstrap_(); const active = data.orders.filter(function(order) { return !order.deleted; }); return { totalOrders: active.length, customers: data.customers.length, paidRevenue: active.reduce(function(sum, order) { return sum + asNumber_(order.paidAmount); }, 0), outstanding: active.reduce(function(sum, order) { return sum + asNumber_(order.balance); }, 0) }; }
function saveSettings_(payload) { const sheet = getSheet_(SHEETS.SETTINGS); const rows = getRows_(SHEETS.SETTINGS); Object.keys(payload || {}).forEach(function(key) { if (payload[key] === undefined || payload[key] === null) return; const existing = rows.filter(function(row) { return row.Key === key; })[0]; if (existing) sheet.getRange(findRow_(sheet, 'Key', key), 2).setValue(String(payload[key])); else appendRows_(SHEETS.SETTINGS, [[key, String(payload[key]), 'Updated from application']]); }); return { settings: getSettings_() }; }
function getSettings_() { const result = {}; getRows_(SHEETS.SETTINGS).forEach(function(row) { result[row.Key] = row.Value; }); return result; }
function ensureAuthorized_(token) { const expected = getSettings_().APP_TOKEN || PropertiesService.getScriptProperties().getProperty('APP_TOKEN'); if (expected && token !== expected) throw new Error('Unauthorised request.'); }

function getSpreadsheet_() { const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'); if (id) return SpreadsheetApp.openById(id); const active = SpreadsheetApp.getActiveSpreadsheet(); if (!active) throw new Error('No active spreadsheet. Bind this script to a Sheet or run setSpreadsheetId(sheetId).'); return active; }
function getSheet_(name) { const sheet = getSpreadsheet_().getSheetByName(name); if (!sheet) throw new Error('Missing sheet ' + name + '. Run setupDatabase() first.'); return sheet; }
function getRows_(name) { const sheet = getSheet_(name); if (sheet.getLastRow() < 2) return []; const values = sheet.getRange(1, 1, sheet.getLastRow(), HEADERS[name].length).getValues(); return values.slice(1).filter(function(row) { return row.some(function(value) { return value !== ''; }); }).map(function(row) { return rowToObject_(row, HEADERS[name]); }); }
function appendRows_(name, rows) { if (!rows || !rows.length) return; const sheet = getSheet_(name); sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows); }
function findRow_(sheet, header, value) { const column = HEADERS[sheet.getName()].indexOf(header) + 1; if (column < 1 || sheet.getLastRow() < 2) return 0; const values = sheet.getRange(2, column, sheet.getLastRow() - 1, 1).getValues(); for (let i = 0; i < values.length; i += 1) if (String(values[i][0]) === String(value)) return i + 2; return 0; }
function rowToObject_(row, headers) { const object = {}; headers.forEach(function(header, index) { object[header] = row[index]; }); return object; }
function orderToClient_(row) { return { orderId: row['Order ID'], invoiceNumber: row['Invoice Number'], createdAt: iso_(row['Created Date']), updatedAt: iso_(row['Updated Date']), customerId: row['Customer ID'], customerName: row['Customer Name'], phone: row['Phone'], address: row['Address'], country: row['Country'], currency: row['Currency'], itemsTotal: asNumber_(row['Items Total']), orderServiceFee: asNumber_(row['Order Service Fee']), shippingFee: asNumber_(row['Shipping Fee']), discount: asNumber_(row['Discount']), grandTotal: asNumber_(row['Grand Total']), paidAmount: asNumber_(row['Paid Amount']), balance: asNumber_(row['Balance']), paymentStatus: row['Payment Status'], paymentMethod: row['Payment Method'], orderStatus: row['Order Status'], notes: row['Notes'], version: asNumber_(row['Version']), deleted: toBoolean_(row['Deleted']) }; }
function itemToClient_(row) { return { orderItemId: row['Order Item ID'], orderId: row['Order ID'], brand: row.Brand, productName: row['Product Name'], colour: row.Colour, size: row.Size, quantity: asNumber_(row.Quantity), pricingMode: row['Pricing Mode'], unitPrice: asNumber_(row['Unit Price']), merchandiseTotal: asNumber_(row['Merchandise Total']), serviceFee: asNumber_(row['Service Fee']), lineTotal: asNumber_(row['Line Total']), notes: row.Notes, createdAt: iso_(row['Created Date']) }; }
function paymentToClient_(row) { return { paymentId: row['Payment ID'], orderId: row['Order ID'], paymentDate: iso_(row['Payment Date']), amount: asNumber_(row.Amount), method: row.Method, reference: row.Reference, status: row.Status, createdAt: iso_(row['Created Date']) }; }
function customerToClient_(row) { return { customerId: row['Customer ID'], name: row.Name, phone: row['Normalized Phone'] || row['Display Phone'], address: row.Address, country: row.Country, createdAt: iso_(row['Created Date']) }; }
function paymentStatus_(total, paid) { total = asNumber_(total); paid = asNumber_(paid); if (paid <= 0) return 'Unpaid'; if (paid >= total && total > 0) return 'Paid'; return 'Partially Paid'; }
function normalizePhone_(value) { let digits = String(value || '').replace(/\D/g, ''); if (digits && digits.charAt(0) === '0') digits = '6' + digits; return digits; }
function asNumber_(value) { const number = Number(value); return isFinite(number) ? Math.max(0, number) : 0; }
function clean_(value) { return String(value || '').trim(); }
function toBoolean_(value) { return value === true || String(value).toLowerCase() === 'true'; }
function iso_(value) { return value instanceof Date ? value.toISOString() : value; }
function audit_(action, entityType, entityId, details) { appendRows_(SHEETS.AUDIT, [[Utilities.getUuid(), new Date(), action, entityType, entityId, JSON.stringify(details || {})]]); }

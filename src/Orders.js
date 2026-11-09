/**
 * Importing orders and reading/writing the Orders and Line items sheets.
 */

function menuImportOrders() {
  var result = importOrders();
  notify_("Imported: " + result.added + " new, " + result.updated + " updated.");
}

/** Fetches orders and writes them to the sheet. New orders are appended; known ones are refreshed in place. */
function importOrders() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var settings = readSettings_();
    var orders = fetchOrders_(settings).map(normalizeOrder);
    var sheet = sheet_(SHEETS.ORDERS);
    var rows = readOrderRows_();
    var plan = planImport(rows.map(function (r) { return r.id; }), orders);

    plan.update.forEach(function (order) {
      var row = rows.filter(function (r) { return r.id === order.id; })[0];
      var values = orderRow(order, row);
      values[ORDER_COLUMNS.indexOf("Status")] = invoiceStatus(merge_(row, { cancelled: order.cancelled }), today_());
      sheet.getRange(row.rowNumber, 1, 1, ORDER_COLUMNS.length).setValues([values]);
    });
    if (plan.add.length) {
      var newRows = plan.add.map(function (order) {
        var values = orderRow(order, {});
        values[ORDER_COLUMNS.indexOf("Status")] = invoiceStatus({ cancelled: order.cancelled }, today_());
        return values;
      });
      var start = lastOrderRow_() + 1;
      sheet.getRange(start, 1, newRows.length, ORDER_COLUMNS.length).setValues(newRows);
    }
    writeLineItems_(orders);
    refreshSummary_();
    log_("Import", plan.add.length + " new, " + plan.update.length + " updated from " +
      (settings.shopifyStore || settings.ordersUrl));
    return { added: plan.add.length, updated: plan.update.length };
  } finally {
    lock.releaseLock();
  }
}

/** Raw orders from Shopify (when a store is set) or from the Orders URL. */
function fetchOrders_(settings) {
  if (settings.shopifyStore) return fetchShopifyOrders_(settings.shopifyStore);
  if (!settings.ordersUrl) throw new Error("Set an Orders URL or a Shopify store in Settings.");
  if (settings.ordersUrl.toLowerCase() === DEMO_SOURCE) return DEMO_ORDERS;
  var response = UrlFetchApp.fetch(settings.ordersUrl, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw new Error("Orders URL answered " + response.getResponseCode() + ".");
  }
  var data = JSON.parse(response.getContentText());
  return Array.isArray(data) ? data : data.orders || [];
}

/** Shopify Admin REST API, all statuses, following the Link header page by page. */
function fetchShopifyOrders_(store) {
  var token = PropertiesService.getScriptProperties().getProperty("SHOPIFY_TOKEN");
  if (!token) throw new Error("Add the Shopify Admin API token as the script property SHOPIFY_TOKEN.");
  var domain = store.replace(/^https?:\/\//, "").replace(/\/.*$/, "");
  var url = "https://" + domain + "/admin/api/2024-10/orders.json?status=any&limit=250";
  var all = [];
  for (var page = 0; url && page < 20; page++) {
    var response = UrlFetchApp.fetch(url, { headers: { "X-Shopify-Access-Token": token }, muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) throw new Error("Shopify answered " + response.getResponseCode() + ".");
    all = all.concat(JSON.parse(response.getContentText()).orders || []);
    var headers = response.getHeaders();
    url = nextPageUrl(headers.Link || headers.link);
  }
  return all;
}

/** Order rows with what the script needs to know, plus their row number. */
function readOrderRows_() {
  var sheet = sheet_(SHEETS.ORDERS);
  var last = lastOrderRow_();
  if (last < 2) return [];
  var values = sheet.getRange(2, 1, last - 1, ORDER_COLUMNS.length).getValues();
  var col = function (name) { return ORDER_COLUMNS.indexOf(name); };
  return values.map(function (v, i) {
    return {
      rowNumber: i + 2,
      id: String(v[col("Order ID")]),
      name: String(v[col("Order")]),
      created: String(v[col("Created")]),
      customerName: String(v[col("Customer")]),
      email: String(v[col("Email")]),
      address: String(v[col("Address")]),
      currency: String(v[col("Currency")]),
      total: toCents(v[col("Total")]),
      cancelled: v[col("Shop status")] === "cancelled",
      invoiceNumber: String(v[col("Invoice no.")]),
      invoicedOn: String(v[col("Invoiced on")]),
      dueOn: String(v[col("Due on")]),
      pdfUrl: String(v[col("Invoice PDF")]),
      sentOn: String(v[col("Sent on")]),
      paid: v[col("Paid")] === true,
      paidOn: String(v[col("Paid on")]),
      status: String(v[col("Status")]),
    };
  }).filter(function (r) { return r.id; });
}

/** Writes one or more columns of one order row by column name. */
function setOrderCells_(rowNumber, values) {
  var sheet = sheet_(SHEETS.ORDERS);
  Object.keys(values).forEach(function (name) {
    sheet.getRange(rowNumber, ORDER_COLUMNS.indexOf(name) + 1).setValue(values[name]);
  });
}

/** Replaces the line items of the imported orders, keeping lines of orders not in this import. */
function writeLineItems_(orders) {
  var sheet = sheet_(SHEETS.LINES);
  var ids = {};
  orders.forEach(function (o) { ids[o.id] = true; });
  var last = sheet.getLastRow();
  var kept = last < 2 ? [] : sheet.getRange(2, 1, last - 1, LINE_COLUMNS.length).getValues()
    .filter(function (r) { return r[0] !== "" && !ids[String(r[0])]; });
  var fresh = [];
  orders.forEach(function (o) { fresh = fresh.concat(lineRows(o)); });
  var all = kept.concat(fresh);
  if (last >= 2) sheet.getRange(2, 1, last - 1, LINE_COLUMNS.length).clearContent();
  if (all.length) sheet.getRange(2, 1, all.length, LINE_COLUMNS.length).setValues(all);
}

/** Line items of one order, read back from the sheet (used when the invoice is made). */
function lineItemsFor_(orderId) {
  var sheet = sheet_(SHEETS.LINES);
  var last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, LINE_COLUMNS.length).getValues()
    .filter(function (r) { return String(r[0]) === orderId; })
    .map(function (r) {
      return { sku: String(r[2]), title: String(r[3]), quantity: Number(r[4]), unitPrice: toCents(r[5]), lineTotal: toCents(r[6]) };
    });
}

function sheet_(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" is missing. Run Order to Invoice > Set up.');
  return sheet;
}

function merge_(a, b) {
  var out = {};
  Object.keys(a || {}).forEach(function (k) { out[k] = a[k]; });
  Object.keys(b || {}).forEach(function (k) { out[k] = b[k]; });
  return out;
}

/**
 * Last row that has an Order ID. sheet.getLastRow() is not used here: the Paid column holds a checkbox
 * on every row, and Sheets counts an unticked box (FALSE) as content.
 */
function lastOrderRow_() {
  var sheet = sheet_(SHEETS.ORDERS);
  var values = sheet.getRange(1, 1, sheet.getMaxRows(), 1).getValues();
  for (var i = values.length - 1; i >= 1; i--) if (values[i][0] !== "") return i + 1;
  return 1;
}

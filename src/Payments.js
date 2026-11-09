/**
 * Payments, statuses, the Summary sheet and the daily automation.
 */

/**
 * Simple trigger: ticking "Paid" fills "Paid on" with today and sets the status; unticking clears it.
 * Only touches the Orders sheet, so it needs no extra permissions.
 */
function onEdit(e) {
  var range = e && e.range;
  if (!range || range.getSheet().getName() !== SHEETS.ORDERS || range.getRow() < 2) return;
  var paidCol = ORDER_COLUMNS.indexOf("Paid") + 1;
  if (range.getColumn() > paidCol || range.getLastColumn() < paidCol) return;
  var sheet = range.getSheet();
  for (var r = range.getRow(); r <= range.getLastRow(); r++) {
    var paid = sheet.getRange(r, paidCol).getValue() === true;
    sheet.getRange(r, ORDER_COLUMNS.indexOf("Paid on") + 1).setValue(paid ? today_() : "");
    var row = readOrderRow_(sheet, r);
    sheet.getRange(r, ORDER_COLUMNS.indexOf("Status") + 1).setValue(invoiceStatus(row, today_()));
  }
  refreshSummary_(); // keep "Paid" and "Outstanding" in the Summary sheet current
}

function readOrderRow_(sheet, rowNumber) {
  var v = sheet.getRange(rowNumber, 1, 1, ORDER_COLUMNS.length).getValues()[0];
  var col = function (name) { return v[ORDER_COLUMNS.indexOf(name)]; };
  return {
    cancelled: col("Shop status") === "cancelled",
    paid: col("Paid") === true,
    invoiceNumber: String(col("Invoice no.")),
    dueOn: String(col("Due on")),
  };
}

function menuRefresh() {
  var changed = refreshStatuses_();
  refreshSummary_();
  notify_("Status updated on " + changed + " row(s). Summary refreshed.");
}

/** Recomputes every status (Unpaid turns Overdue after the due date). Returns how many changed. */
function refreshStatuses_() {
  var sheet = sheet_(SHEETS.ORDERS);
  var today = today_();
  var rows = readOrderRows_();
  if (!rows.length) return 0;
  var statusCol = ORDER_COLUMNS.indexOf("Status") + 1;
  var values = sheet.getRange(2, statusCol, lastOrderRow_() - 1, 1).getValues();
  var changed = 0;
  rows.forEach(function (row) {
    var status = invoiceStatus(row, today);
    if (values[row.rowNumber - 2][0] !== status) {
      values[row.rowNumber - 2][0] = status;
      changed += 1;
    }
  });
  if (changed) sheet.getRange(2, statusCol, values.length, 1).setValues(values);
  return changed;
}

/** Rewrites the Summary sheet: one row per month, amounts in the shop currency of the first order. */
function refreshSummary_() {
  var rows = readOrderRows_().map(function (r) { return merge_(r, { status: invoiceStatus(r, today_()) }); });
  var sheet = sheet_(SHEETS.SUMMARY);
  var currency = rows.length ? rows[0].currency : "USD";
  var out = monthlySummary(rows).map(function (m) {
    return [m.month, m.orders, formatMoney(m.sales, currency), formatMoney(m.invoiced, currency),
      formatMoney(m.paid, currency), formatMoney(m.outstanding, currency), m.overdue];
  });
  var last = sheet.getLastRow();
  if (last >= 2) sheet.getRange(2, 1, last - 1, 7).clearContent();
  if (out.length) {
    sheet.getRange(2, 1, out.length, 1).setNumberFormat("@");
    sheet.getRange(2, 1, out.length, 7).setValues(out);
  }
}

// ---- daily automation ----

var DAILY_HANDLER = "dailyRun";

/** What the daily trigger does: import, invoice, send, then refresh. Reminders stay a manual choice. */
function dailyRun() {
  try {
    importOrders();
    createInvoices();
    sendInvoices();
    refreshStatuses_();
    refreshSummary_();
  } catch (error) {
    log_("Daily run failed", String(error && error.message || error));
    throw error; // Apps Script then emails the owner about the failed trigger.
  }
}

function menuInstallTrigger() {
  removeDailyTriggers_();
  ScriptApp.newTrigger(DAILY_HANDLER).timeBased().everyDays(1).atHour(7).create();
  notify_("Daily automation is on (around 7:00 in the spreadsheet's time zone).");
}

function menuRemoveTrigger() {
  var removed = removeDailyTriggers_();
  notify_(removed ? "Daily automation is off." : "Daily automation was not on.");
}

function removeDailyTriggers_() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === DAILY_HANDLER) {
      ScriptApp.deleteTrigger(t);
      removed += 1;
    }
  });
  return removed;
}

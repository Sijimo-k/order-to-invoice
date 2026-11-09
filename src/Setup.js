/**
 * Menu, one-click setup and settings. Everything a reviewer needs is created by "Set up".
 */

var SHEETS = { ORDERS: "Orders", LINES: "Line items", SETTINGS: "Settings", SUMMARY: "Summary", LOG: "Log" };

/** "demo" means the fictional orders built into the script (DemoData.js), so a first try needs no setup. */
var DEMO_SOURCE = "demo";

/** Settings rows: [label, default value, note]. The label is the key the code reads. */
var SETTING_ROWS = [
  ["Shop name", "Harbor Leaf Tea (demo)", "Shown on invoices and emails."],
  ["Shop address", "1 Demo Way, Portland, OR 97201", ""],
  ["Shop email", "billing@harborleaf.example.com", "Printed on invoices and used as reply-to. Leave empty to use your own address."],
  ["Orders URL", DEMO_SOURCE, '"demo" = built-in fictional orders. Or a URL that returns JSON in the Shopify orders shape. Ignored when a Shopify store is set below.'],
  ["Shopify store", "", "Optional, e.g. your-store.myshopify.com. Put the Admin API token in Script Properties as SHOPIFY_TOKEN."],
  ["Payment terms (days)", 14, "Due date = invoice date + this many days."],
  ["Test mode", true, "When ticked, every email goes to you, not to customers."],
  ["Invoice template", "", "Filled in by Set up (Google Doc ID)."],
  ["Invoice folder", "", "Filled in by Set up (Drive folder ID)."],
  ["Last invoice no.", "", "Updated automatically."],
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Order to Invoice")
    .addItem("Import orders", "menuImportOrders")
    .addItem("Create invoices", "menuCreateInvoices")
    .addItem("Send invoices", "menuSendInvoices")
    .addItem("Send reminders (overdue)", "menuSendReminders")
    .addItem("Refresh status and summary", "menuRefresh")
    .addSeparator()
    .addItem("Turn on daily automation", "menuInstallTrigger")
    .addItem("Turn off daily automation", "menuRemoveTrigger")
    .addSeparator()
    .addItem("Set up (first time)", "menuSetUp")
    .addItem("Reset demo data", "menuResetDemo")
    .addToUi();
}

function menuResetDemo() {
  var ui = SpreadsheetApp.getUi();
  var answer = ui.alert("Reset demo data?",
    "This clears Orders, Line items, Summary and the invoice counter, and moves the invoice PDFs to the Drive trash. " +
    "Settings and the invoice template stay.", ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  var removed = resetDemo();
  notify_("Cleared. " + removed + " PDF(s) moved to the trash.");
}

/** Empties the data sheets and trashes the generated PDFs, so the demo can be run again from the start. */
function resetDemo() {
  [SHEETS.ORDERS, SHEETS.LINES, SHEETS.SUMMARY].forEach(function (name) {
    var sheet = sheet_(name);
    var rows = sheet.getMaxRows() - 1;
    sheet.getRange(2, 1, rows, sheet.getMaxColumns()).clearContent();
  });
  // Keep the Paid column as unticked checkboxes.
  sheet_(SHEETS.ORDERS).getRange(2, ORDER_COLUMNS.indexOf("Paid") + 1, sheet_(SHEETS.ORDERS).getMaxRows() - 1, 1).uncheck();
  setSetting_("Last invoice no.", "");
  var removed = 0;
  var files = DriveApp.getFolderById(String(getSetting_("Invoice folder"))).getFilesByType(MimeType.PDF);
  while (files.hasNext()) {
    files.next().setTrashed(true);
    removed += 1;
  }
  log_("Reset", removed + " PDF(s) trashed");
  return removed;
}

function menuSetUp() {
  var created = setUp();
  SpreadsheetApp.getUi().alert(
    "Ready.\n\n" + created.join("\n") +
    "\n\nNext: Order to Invoice > Import orders. Test mode is on, so emails only go to you."
  );
}

/** Creates the sheets, the invoice template and the PDF folder if they are missing. Safe to run again. */
function setUp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var created = [];
  var settings = ensureSheet_(ss, SHEETS.SETTINGS, ["Setting", "Value", "Note"], created);
  if (settings.getLastRow() < 2) {
    settings.getRange(2, 1, SETTING_ROWS.length, 3).setValues(SETTING_ROWS);
    settings.getRange(2 + SETTING_ROWS.findIndex(function (r) { return r[0] === "Test mode"; }), 2).insertCheckboxes().check();
    settings.setColumnWidth(1, 170).setColumnWidth(2, 420).setColumnWidth(3, 520);
  }

  var orders = ensureSheet_(ss, SHEETS.ORDERS, ORDER_COLUMNS, created);
  // Text format keeps IDs and YYYY-MM-DD dates exactly as written (no auto date conversion).
  ["Order ID", "Created", "Invoiced on", "Due on", "Sent on", "Paid on"].forEach(function (name) {
    orders.getRange(2, ORDER_COLUMNS.indexOf(name) + 1, orders.getMaxRows() - 1, 1).setNumberFormat("@");
  });
  ["Subtotal", "Discount", "Shipping", "Tax", "Total"].forEach(function (name) {
    orders.getRange(2, ORDER_COLUMNS.indexOf(name) + 1, orders.getMaxRows() - 1, 1).setNumberFormat("#,##0.00");
  });
  orders.getRange(2, ORDER_COLUMNS.indexOf("Paid") + 1, orders.getMaxRows() - 1, 1).insertCheckboxes();
  addStatusColors_(orders);

  var lines = ensureSheet_(ss, SHEETS.LINES, LINE_COLUMNS, created);
  lines.getRange("A2:A").setNumberFormat("@");
  lines.getRange("F2:G").setNumberFormat("#,##0.00");
  ensureSheet_(ss, SHEETS.SUMMARY, ["Month", "Orders", "Sales", "Invoiced", "Paid", "Outstanding", "Overdue invoices"], created);
  ensureSheet_(ss, SHEETS.LOG, ["When", "Action", "Details"], created);

  if (!getSetting_("Invoice folder")) {
    var folder = DriveApp.createFolder(ss.getName() + " - invoices");
    setSetting_("Invoice folder", folder.getId());
    created.push("Folder for invoice PDFs: " + folder.getName());
  }
  if (!getSetting_("Invoice template")) {
    var templateId = createInvoiceTemplate_(getSetting_("Invoice folder"));
    setSetting_("Invoice template", templateId);
    created.push("Invoice template (Google Doc) - edit its look freely, keep the {{placeholders}}.");
  }
  // Remove the empty default sheet ("Sheet1" in English, other names in other languages).
  var ours = Object.keys(SHEETS).map(function (k) { return SHEETS[k]; });
  ss.getSheets().forEach(function (s) {
    if (ours.indexOf(s.getName()) < 0 && s.getLastRow() === 0 && s.getLastColumn() === 0) ss.deleteSheet(s);
  });
  ss.setActiveSheet(orders);
  if (!created.length) created.push("Everything was already set up.");
  return created;
}

function ensureSheet_(ss, name, header, created) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    created.push("Sheet: " + name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight("bold").setBackground("#eef2f6");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function addStatusColors_(sheet) {
  var col = ORDER_COLUMNS.indexOf("Status") + 1;
  var range = sheet.getRange(2, col, sheet.getMaxRows() - 1, 1);
  var rule = function (text, background, color) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text)
      .setBackground(background).setFontColor(color).setRanges([range]).build();
  };
  sheet.setConditionalFormatRules([
    rule(STATUS.PAID, "#e6f4ea", "#1a7f37"),
    rule(STATUS.OVERDUE, "#fde8e8", "#b42318"),
    rule(STATUS.UNPAID, "#fff4e5", "#b35900"),
    rule(STATUS.CANCELLED, "#f1f3f4", "#5f6368"),
  ]);
}

/** A plain, editable invoice layout. The table row with {{Item}} is repeated once per line item. */
function createInvoiceTemplate_(folderId) {
  var doc = DocumentApp.create("Invoice template - Order to Invoice");
  var body = doc.getBody();
  body.clear();
  body.setMarginTop(48).setMarginBottom(48).setMarginLeft(56).setMarginRight(56);
  body.appendParagraph("{{ShopName}}").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph("{{ShopAddress}}\n{{ShopEmail}}");
  body.appendParagraph("INVOICE").setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendTable([
    ["Invoice no.", "{{InvoiceNumber}}", "Bill to", "{{CustomerName}}"],
    ["Invoice date", "{{InvoiceDate}}", "", "{{CustomerAddress}}"],
    ["Due date", "{{DueDate}}", "", "{{CustomerEmail}}"],
    ["Order", "{{OrderName}}", "", ""],
  ]).setBorderWidth(0);
  body.appendParagraph("");
  var items = body.appendTable([
    ["Item", "SKU", "Qty", "Unit price", "Amount"],
    ["{{Item}}", "{{Sku}}", "{{Qty}}", "{{UnitPrice}}", "{{LineTotal}}"],
  ]);
  items.getRow(0).editAsText().setBold(true);
  body.appendTable([
    ["Subtotal", "{{Subtotal}}"],
    ["Discount", "{{Discount}}"],
    ["Shipping", "{{Shipping}}"],
    ["Tax", "{{Tax}}"],
    ["Total", "{{Total}}"],
  ]).setBorderWidth(0).getRow(4).editAsText().setBold(true);
  body.appendParagraph("Payment is due by {{DueDate}}. Thank you for your order.");
  doc.saveAndClose();
  var file = DriveApp.getFileById(doc.getId());
  file.moveTo(DriveApp.getFolderById(folderId));
  return doc.getId();
}

// ---- settings ----

function settingsSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  if (!sheet) throw new Error('Run "Order to Invoice > Set up" first.');
  return sheet;
}

function settingRow_(label) {
  var labels = settingsSheet_().getRange("A:A").getValues();
  for (var i = 0; i < labels.length; i++) if (labels[i][0] === label) return i + 1;
  throw new Error('Setting "' + label + '" is missing. Run Set up again.');
}

function getSetting_(label) {
  return settingsSheet_().getRange(settingRow_(label), 2).getValue();
}

function setSetting_(label, value) {
  settingsSheet_().getRange(settingRow_(label), 2).setValue(value);
}

function readSettings_() {
  return {
    shop: {
      name: String(getSetting_("Shop name")),
      address: String(getSetting_("Shop address")),
      email: String(getSetting_("Shop email") || Session.getEffectiveUser().getEmail()),
    },
    ordersUrl: String(getSetting_("Orders URL")).trim(),
    shopifyStore: String(getSetting_("Shopify store")).trim(),
    termsDays: Number(getSetting_("Payment terms (days)")) || 14,
    testMode: getSetting_("Test mode") === true,
    templateId: String(getSetting_("Invoice template")),
    folderId: String(getSetting_("Invoice folder")),
  };
}

/** Today in the spreadsheet's time zone, as YYYY-MM-DD. */
function today_() {
  return Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd");
}

function log_(action, details) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOG);
  if (sheet) sheet.appendRow([new Date(), action, details]);
}

/** Shows a short message: a toast when run from the menu, nothing when run by a trigger. */
function notify_(message) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, "Order to Invoice", 8);
  } catch (e) {
    // No UI when a time-driven trigger runs the script.
  }
}

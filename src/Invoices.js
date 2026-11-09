/**
 * Invoice PDFs from the Docs template, and the emails that carry them.
 */

var MAX_PER_RUN = 25; // keeps one run well inside the Apps Script time limit

function menuCreateInvoices() {
  var result = createInvoices();
  notify_("Created " + result.created + " invoice(s)." + (result.more ? " Run again for the rest." : ""));
}

function menuSendInvoices() {
  var result = sendInvoices();
  notify_("Sent " + result.sent + " invoice email(s)" + (readSettings_().testMode ? " (test mode: to you only)." : "."));
}

function menuSendReminders() {
  var result = sendReminders();
  notify_("Sent " + result.sent + " reminder(s)" + (readSettings_().testMode ? " (test mode: to you only)." : "."));
}

/** Makes a PDF for every order that is not invoiced, not cancelled and not zero. */
function createInvoices() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var settings = readSettings_();
    var today = today_();
    var todo = readOrderRows_().filter(function (r) {
      return !r.cancelled && !r.invoiceNumber && r.total > 0;
    });
    var batch = todo.slice(0, MAX_PER_RUN);
    batch.forEach(function (row) {
      var number = nextInvoiceNumber(getSetting_("Last invoice no."), Number(today.slice(0, 4)));
      var invoice = { number: number, date: today, dueOn: addDays(today, settings.termsDays) };
      var pdf = renderInvoicePdf_(row, invoice, settings);
      // Save the number first so a failure later never reuses it.
      setSetting_("Last invoice no.", number);
      setOrderCells_(row.rowNumber, {
        "Invoice no.": number,
        "Invoiced on": invoice.date,
        "Due on": invoice.dueOn,
        "Invoice PDF": pdf.getUrl(),
        "Status": invoiceStatus(merge_(row, { invoiceNumber: number, dueOn: invoice.dueOn }), today),
      });
      log_("Invoice", number + " for order " + row.name);
    });
    refreshSummary_();
    return { created: batch.length, more: todo.length > batch.length };
  } finally {
    lock.releaseLock();
  }
}

/** Copies the template, fills it, saves a PDF in the invoice folder and removes the working copy. */
function renderInvoicePdf_(row, invoice, settings) {
  var folder = DriveApp.getFolderById(settings.folderId);
  var copy = DriveApp.getFileById(settings.templateId).makeCopy(invoice.number + " (working copy)", folder);
  try {
    var doc = DocumentApp.openById(copy.getId());
    var body = doc.getBody();
    var order = orderForInvoice_(row);
    fillItemsTable_(body, order);
    var values = invoiceValues(order, invoice, settings.shop);
    Object.keys(values).forEach(function (key) {
      replacePlaceholder_(body, key, values[key]);
    });
    var leftover = body.findText("\\{\\{[A-Za-z0-9_ ]+\\}\\}");
    if (leftover) log_("Template", "Unknown placeholder left in " + invoice.number + ": " + leftover.getElement().asText().getText());
    doc.saveAndClose();
    var pdf = copy.getAs(MimeType.PDF).setName(invoice.number + " - " + row.name.replace("#", "Order ") + ".pdf");
    return folder.createFile(pdf);
  } finally {
    copy.setTrashed(true);
  }
}

/** The order as Core expects it, rebuilt from the sheet so manual edits in the sheet are respected. */
function orderForInvoice_(row) {
  var sheet = sheet_(SHEETS.ORDERS);
  var v = sheet.getRange(row.rowNumber, 1, 1, ORDER_COLUMNS.length).getValues()[0];
  var col = function (name) { return v[ORDER_COLUMNS.indexOf(name)]; };
  var lines = lineItemsFor_(row.id);
  return {
    id: row.id,
    name: row.name,
    customerName: row.customerName,
    email: row.email,
    currency: row.currency || "USD",
    address: row.address,
    subtotal: toCents(col("Subtotal")),
    discount: toCents(col("Discount")),
    shipping: toCents(col("Shipping")),
    tax: toCents(col("Tax")),
    total: toCents(col("Total")),
    lines: lines,
  };
}

/** Repeats the table row that holds {{Item}} once per line item, then removes the template row. */
function fillItemsTable_(body, order) {
  var found = body.findText("\\{\\{Item\\}\\}");
  if (!found) return;
  var templateRow = found.getElement().getParent().getParent().getParent(); // Text > Paragraph > Cell > Row
  if (templateRow.getType() !== DocumentApp.ElementType.TABLE_ROW) return;
  var table = templateRow.getParentTable();
  var index = table.getChildIndex(templateRow);
  order.lines.forEach(function (line, i) {
    var row = table.insertTableRow(index + 1 + i, templateRow.copy());
    var values = {
      Item: line.title,
      Sku: line.sku,
      Qty: String(line.quantity),
      UnitPrice: formatMoney(line.unitPrice, order.currency),
      LineTotal: formatMoney(line.lineTotal, order.currency),
    };
    Object.keys(values).forEach(function (key) {
      replacePlaceholder_(row, key, values[key]);
    });
  });
  table.removeRow(index);
}

/**
 * Replaces every {{key}} inside an element. Done by delete + insert instead of replaceText, because
 * replaceText treats "$" in the replacement as a group reference and money values contain "$".
 */
function replacePlaceholder_(element, key, value) {
  var pattern = "\\{\\{\\s*" + key + "\\s*\\}\\}";
  var text = value === undefined || value === null ? "" : String(value);
  for (var found = element.findText(pattern), guard = 0; found && guard < 200; found = element.findText(pattern), guard++) {
    var target = found.getElement().asText();
    var start = found.getStartOffset();
    target.deleteText(start, found.getEndOffsetInclusive());
    if (text) target.insertText(start, text);
  }
}

/** Emails each invoice PDF that has not been sent yet. */
function sendInvoices() {
  var settings = readSettings_();
  var owner = Session.getEffectiveUser().getEmail();
  var sent = 0;
  readOrderRows_().filter(function (r) {
    return r.pdfUrl && !r.sentOn && !r.cancelled && !r.paid;
  }).slice(0, MAX_PER_RUN).forEach(function (row) {
    sendInvoiceEmail_(row, settings, owner, "invoice");
    setOrderCells_(row.rowNumber, { "Sent on": today_() });
    sent += 1;
  });
  return { sent: sent };
}

/** Emails a reminder for every overdue invoice. */
function sendReminders() {
  refreshStatuses_();
  var settings = readSettings_();
  var owner = Session.getEffectiveUser().getEmail();
  var sent = 0;
  readOrderRows_().filter(function (r) { return r.status === STATUS.OVERDUE; })
    .slice(0, MAX_PER_RUN).forEach(function (row) {
      sendInvoiceEmail_(row, settings, owner, "reminder");
      sent += 1;
    });
  return { sent: sent };
}

function sendInvoiceEmail_(row, settings, owner, kind) {
  var order = orderForInvoice_(row);
  var values = invoiceValues(order, { number: row.invoiceNumber, date: row.invoicedOn, dueOn: row.dueOn }, settings.shop);
  var mail = invoiceEmail(values, kind);
  var recipient = resolveRecipient(row.email, settings.testMode, owner);
  var pdf = DriveApp.getFileById(fileIdFromUrl_(row.pdfUrl)).getBlob();
  MailApp.sendEmail({
    to: recipient.to,
    subject: recipient.subjectPrefix + mail.subject,
    body: mail.body,
    name: settings.shop.name,
    replyTo: settings.shop.email,
    attachments: [pdf],
  });
  log_(kind === "reminder" ? "Reminder" : "Email", row.invoiceNumber + " to " + recipient.to +
    (settings.testMode ? " (test mode, meant for " + row.email + ")" : ""));
}

function fileIdFromUrl_(url) {
  var match = /\/d\/([A-Za-z0-9_-]+)/.exec(url) || /[?&]id=([A-Za-z0-9_-]+)/.exec(url);
  if (!match) throw new Error("Cannot read the PDF link: " + url);
  return match[1];
}

/**
 * Pure logic for Order to Invoice: no Google services here, so it runs (and is tested) in Node too.
 * Apps Script loads every file into one global scope; Node gets the same functions via module.exports.
 *
 * Money is kept in integer cents to avoid floating point drift. Dates are "YYYY-MM-DD" strings
 * (the shop's calendar day), so time zones never shift a due date.
 */

/** Columns of the Orders sheet, in order. The sheet header is written from this list. */
var ORDER_COLUMNS = [
  "Order ID", "Order", "Created", "Customer", "Email", "Address", "Currency",
  "Subtotal", "Discount", "Shipping", "Tax", "Total",
  "Shop status", "Invoice no.", "Invoiced on", "Due on", "Invoice PDF", "Sent on",
  "Paid", "Paid on", "Status", "Note",
];

var LINE_COLUMNS = ["Order ID", "Order", "SKU", "Item", "Qty", "Unit price", "Line total"];

var STATUS = {
  NOT_INVOICED: "Not invoiced",
  UNPAID: "Unpaid",
  OVERDUE: "Overdue",
  PAID: "Paid",
  CANCELLED: "Cancelled",
};

/** "12.50" / 12.5 / "" -> 1250. Throws on text that is not a number, so bad data is not silently zero. */
function toCents(value) {
  if (value === null || value === undefined || value === "") return 0;
  var text = String(value).trim().replace(/,/g, "");
  if (!/^-?\d+(\.\d+)?$/.test(text)) throw new Error('Not an amount: "' + value + '"');
  var negative = text.charAt(0) === "-";
  var parts = text.replace("-", "").split(".");
  var fraction = (parts[1] || "").concat("00");
  var cents = Number(parts[0]) * 100 + Number(fraction.slice(0, 2));
  // Round half up on the third decimal (Shopify sends two, but be safe).
  if (fraction.length > 4 && Number(fraction.charAt(2)) >= 5) cents += 1;
  return negative ? -cents : cents;
}

/** 1250 -> "12.50" (plain number text for sheets). */
function centsToAmount(cents) {
  var sign = cents < 0 ? "-" : "";
  var abs = Math.abs(cents);
  var fraction = String(abs % 100);
  return sign + Math.floor(abs / 100) + "." + (fraction.length < 2 ? "0" + fraction : fraction);
}

var CURRENCY_SYMBOLS = { USD: "$", CAD: "CA$", AUD: "A$", EUR: "€", GBP: "£", JPY: "¥" };

/** 123450, "USD" -> "$1,234.50". JPY has no minor unit, so it is shown without decimals. */
function formatMoney(cents, currency) {
  var symbol = CURRENCY_SYMBOLS[currency] || (currency ? currency + " " : "");
  var amount = currency === "JPY" ? String(Math.round(cents / 100)) : centsToAmount(Math.abs(cents));
  var withCommas = amount.replace(/^(\d+)/, function (whole) {
    return whole.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  });
  return (cents < 0 ? "-" : "") + symbol + withCommas;
}

/** Shopify order (Admin REST shape) -> flat order used by the sheet. */
function normalizeOrder(raw) {
  if (!raw || raw.id === undefined || raw.id === null) throw new Error("Order without an id.");
  var customer = raw.customer || {};
  var billing = raw.billing_address || {};
  var name = (billing.name || [customer.first_name, customer.last_name].filter(Boolean).join(" ")).trim();
  var shippingSet = raw.total_shipping_price_set && raw.total_shipping_price_set.shop_money;
  var lines = (raw.line_items || []).map(function (item) {
    var unit = toCents(item.price);
    var qty = Number(item.quantity) || 0;
    return {
      sku: item.sku || "",
      title: item.variant_title ? item.title + " - " + item.variant_title : item.title,
      quantity: qty,
      unitPrice: unit,
      lineTotal: unit * qty,
    };
  });
  return {
    id: String(raw.id),
    name: raw.name || "#" + raw.order_number,
    created: String(raw.created_at || "").slice(0, 10),
    customerName: name,
    email: raw.email || customer.email || "",
    currency: raw.currency || "USD",
    subtotal: lines.reduce(function (sum, l) { return sum + l.lineTotal; }, 0),
    discount: toCents(raw.total_discounts),
    shipping: toCents(shippingSet ? shippingSet.amount : 0),
    tax: toCents(raw.total_tax),
    total: toCents(raw.total_price),
    shopStatus: raw.cancelled_at ? "cancelled" : raw.financial_status || "",
    cancelled: Boolean(raw.cancelled_at),
    address: [billing.address1, billing.address2, billing.city, billing.province_code || billing.province, billing.zip, billing.country]
      .filter(Boolean).join(", "),
    lines: lines,
  };
}

/**
 * Check that the parts add up to the shop's total (subtotal - discount + shipping + tax).
 * Returns "" when they do, otherwise a note for the sheet. Taxes included in prices are reported, not "fixed".
 */
function totalsNote(order) {
  var computed = order.subtotal - order.discount + order.shipping + order.tax;
  if (computed === order.total) return "";
  if (order.subtotal - order.discount + order.shipping === order.total) return "Tax is included in prices.";
  return "Total from the shop (" + formatMoney(order.total, order.currency) + ") differs from the parts (" +
    formatMoney(computed, order.currency) + "). Check the order.";
}

/** Which incoming orders are new and which update rows already in the sheet. */
function planImport(existingIds, orders) {
  var known = {};
  existingIds.forEach(function (id) { known[String(id)] = true; });
  var seen = {};
  var plan = { add: [], update: [] };
  orders.forEach(function (order) {
    if (seen[order.id]) return; // the same order twice in one feed
    seen[order.id] = true;
    (known[order.id] ? plan.update : plan.add).push(order);
  });
  return plan;
}

/** Last number used ("INV-2026-0007" or "") and the year -> next number. Restarts at 0001 each year. */
function nextInvoiceNumber(lastNumber, year) {
  var match = /^INV-(\d{4})-(\d+)$/.exec(String(lastNumber || "").trim());
  var next = match && Number(match[1]) === year ? Number(match[2]) + 1 : 1;
  var digits = String(next);
  while (digits.length < 4) digits = "0" + digits;
  return "INV-" + year + "-" + digits;
}

/** "2026-11-09" + 14 -> "2026-11-23" (calendar days, no time zone). */
function addDays(isoDate, days) {
  var parts = isoDate.split("-").map(Number);
  var date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2] + days));
  return date.toISOString().slice(0, 10);
}

/** Status of one order row from what the sheet knows. */
function invoiceStatus(row, today) {
  if (row.cancelled) return STATUS.CANCELLED;
  if (row.paid) return STATUS.PAID;
  if (!row.invoiceNumber) return STATUS.NOT_INVOICED;
  if (row.dueOn && row.dueOn < today) return STATUS.OVERDUE;
  return STATUS.UNPAID;
}

/** Values for one Orders row (same order as ORDER_COLUMNS). Invoice/payment columns keep what the sheet had. */
function orderRow(order, kept) {
  kept = kept || {};
  return [
    order.id, order.name, order.created, order.customerName, order.email, order.address, order.currency,
    centsToAmount(order.subtotal), centsToAmount(order.discount), centsToAmount(order.shipping),
    centsToAmount(order.tax), centsToAmount(order.total),
    order.shopStatus, kept.invoiceNumber || "", kept.invoicedOn || "", kept.dueOn || "", kept.pdfUrl || "",
    kept.sentOn || "", Boolean(kept.paid), kept.paidOn || "", kept.status || "", totalsNote(order),
  ];
}

function lineRows(order) {
  return order.lines.map(function (l) {
    return [order.id, order.name, l.sku, l.title, l.quantity, centsToAmount(l.unitPrice), centsToAmount(l.lineTotal)];
  });
}

/** Replace {{Placeholders}} in text. Missing values become "" and are listed so the caller can warn. */
function fillPlaceholders(text, values) {
  var missing = [];
  var out = String(text).replace(/\{\{\s*([A-Za-z][A-Za-z0-9_]*)\s*\}\}/g, function (_, key) {
    if (Object.prototype.hasOwnProperty.call(values, key)) return String(values[key]);
    if (missing.indexOf(key) < 0) missing.push(key);
    return "";
  });
  return { text: out, missing: missing };
}

/** Placeholder values for the invoice template and the email. */
function invoiceValues(order, invoice, shop) {
  var money = function (cents) { return formatMoney(cents, order.currency); };
  return {
    ShopName: shop.name,
    ShopAddress: shop.address,
    ShopEmail: shop.email,
    InvoiceNumber: invoice.number,
    InvoiceDate: invoice.date,
    DueDate: invoice.dueOn,
    OrderName: order.name,
    CustomerName: order.customerName,
    CustomerEmail: order.email,
    CustomerAddress: order.address,
    Subtotal: money(order.subtotal),
    Discount: order.discount ? "-" + money(order.discount) : money(0),
    Shipping: money(order.shipping),
    Tax: money(order.tax),
    Total: money(order.total),
  };
}

/** Email subject and plain-text body for a new invoice or a reminder. */
function invoiceEmail(values, kind) {
  var reminder = kind === "reminder";
  var subject = (reminder ? "Reminder: " : "") + "Invoice " + values.InvoiceNumber + " from " + values.ShopName;
  var body = [
    "Hi " + (values.CustomerName || "there") + ",",
    "",
    reminder
      ? "This is a friendly reminder that invoice " + values.InvoiceNumber + " for order " + values.OrderName +
        " (" + values.Total + ") was due on " + values.DueDate + "."
      : "Thank you for your order " + values.OrderName + ". Your invoice " + values.InvoiceNumber +
        " for " + values.Total + " is attached. Payment is due by " + values.DueDate + ".",
    reminder ? "If you have already paid, please ignore this message." : "",
    "",
    values.ShopName,
    values.ShopEmail,
  ].filter(function (line, i, all) { return line !== "" || all[i - 1] !== ""; }).join("\n");
  return { subject: subject, body: body };
}

/**
 * Where an email really goes. In test mode every email goes to the script owner, with the intended
 * recipient in the subject, so a demo never mails a customer.
 */
function resolveRecipient(customerEmail, testMode, ownerEmail) {
  if (testMode) return { to: ownerEmail, subjectPrefix: "[TEST to " + (customerEmail || "no email") + "] " };
  if (!customerEmail) throw new Error("The order has no email address.");
  return { to: customerEmail, subjectPrefix: "" };
}

/** Rows like { created, total, paid, invoiceNumber, status } -> one summary row per month, newest first. */
function monthlySummary(rows) {
  var byMonth = {};
  rows.forEach(function (r) {
    if (r.status === STATUS.CANCELLED) return;
    var month = String(r.created).slice(0, 7);
    var m = byMonth[month] || (byMonth[month] = { month: month, orders: 0, sales: 0, invoiced: 0, paid: 0, outstanding: 0, overdue: 0 });
    m.orders += 1;
    m.sales += r.total;
    if (r.invoiceNumber) m.invoiced += r.total;
    if (r.paid) m.paid += r.total;
    else if (r.invoiceNumber) m.outstanding += r.total;
    if (r.status === STATUS.OVERDUE) m.overdue += 1;
  });
  return Object.keys(byMonth).sort().reverse().map(function (k) { return byMonth[k]; });
}

/** Shopify paginates with a Link header; return the URL of the next page or "". */
function nextPageUrl(linkHeader) {
  var parts = String(linkHeader || "").split(",");
  for (var i = 0; i < parts.length; i++) {
    var match = /<([^>]+)>\s*;\s*rel="next"/.exec(parts[i]);
    if (match) return match[1];
  }
  return "";
}

if (typeof module !== "undefined") {
  module.exports = {
    ORDER_COLUMNS: ORDER_COLUMNS, LINE_COLUMNS: LINE_COLUMNS, STATUS: STATUS,
    toCents: toCents, centsToAmount: centsToAmount, formatMoney: formatMoney,
    normalizeOrder: normalizeOrder, totalsNote: totalsNote, planImport: planImport,
    nextInvoiceNumber: nextInvoiceNumber, addDays: addDays, invoiceStatus: invoiceStatus,
    orderRow: orderRow, lineRows: lineRows, fillPlaceholders: fillPlaceholders,
    invoiceValues: invoiceValues, invoiceEmail: invoiceEmail, resolveRecipient: resolveRecipient,
    monthlySummary: monthlySummary, nextPageUrl: nextPageUrl,
  };
}

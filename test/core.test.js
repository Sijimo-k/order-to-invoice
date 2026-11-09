import { describe, expect, it } from "vitest";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const core = require("../src/Core.js");
const orders = require("../demo/orders.json").orders;

const sample = {
  id: 5001,
  name: "#1001",
  created_at: "2026-11-02T10:15:00-05:00",
  email: "maya.lin@example.com",
  currency: "USD",
  financial_status: "pending",
  total_discounts: "5.00",
  total_tax: "3.60",
  total_price: "52.60",
  total_shipping_price_set: { shop_money: { amount: "6.00", currency_code: "USD" } },
  customer: { first_name: "Maya", last_name: "Lin" },
  billing_address: { name: "Maya Lin", address1: "12 Harbor St", city: "Portland", province_code: "OR", zip: "97201", country: "United States" },
  line_items: [
    { sku: "TEA-SEN-50", title: "Morning Sencha", variant_title: "50 g", quantity: 2, price: "12.50" },
    { sku: "TEA-OOL-50", title: "High Mountain Oolong", variant_title: null, quantity: 1, price: "23.00" },
  ],
};

describe("money", () => {
  it("parses amounts into cents without float drift", () => {
    expect(core.toCents("12.50")).toBe(1250);
    expect(core.toCents("0.1")).toBe(10);
    expect(core.toCents("1,234.5")).toBe(123450);
    expect(core.toCents("")).toBe(0);
    expect(core.toCents("-3.00")).toBe(-300);
    expect(() => core.toCents("12 USD")).toThrow("Not an amount");
  });

  it("formats with symbol and thousands separators", () => {
    expect(core.formatMoney(123450, "USD")).toBe("$1,234.50");
    expect(core.formatMoney(-500, "EUR")).toBe("-€5.00");
    expect(core.formatMoney(5, "USD")).toBe("$0.05");
    expect(core.centsToAmount(1250)).toBe("12.50");
  });
});

describe("importing orders", () => {
  it("maps a Shopify order to a flat order", () => {
    const o = core.normalizeOrder(sample);
    expect(o).toMatchObject({
      id: "5001", name: "#1001", created: "2026-11-02", customerName: "Maya Lin", currency: "USD",
      subtotal: 4800, discount: 500, shipping: 600, tax: 360, total: 5260, shopStatus: "pending",
      address: "12 Harbor St, Portland, OR, 97201, United States",
    });
    expect(o.lines[0]).toEqual({ sku: "TEA-SEN-50", title: "Morning Sencha - 50 g", quantity: 2, unitPrice: 1250, lineTotal: 2500 });
    expect(core.totalsNote(o)).toBe("");
  });

  it("notes totals that do not add up, and tax-inclusive prices", () => {
    const wrong = core.normalizeOrder({ ...sample, total_price: "60.00" });
    expect(core.totalsNote(wrong)).toContain("differs");
    const inclusive = core.normalizeOrder({ ...sample, total_price: "49.00" });
    expect(core.totalsNote(inclusive)).toBe("Tax is included in prices.");
  });

  it("marks cancelled orders", () => {
    const o = core.normalizeOrder({ ...sample, cancelled_at: "2026-11-03T09:00:00Z" });
    expect(o.cancelled).toBe(true);
    expect(o.shopStatus).toBe("cancelled");
  });

  it("adds new orders, updates known ones, and ignores repeats in the same feed", () => {
    const a = core.normalizeOrder(sample);
    const b = core.normalizeOrder({ ...sample, id: 5002, name: "#1002" });
    const plan = core.planImport(["5001"], [a, b, b]);
    expect(plan.update.map((o) => o.id)).toEqual(["5001"]);
    expect(plan.add.map((o) => o.id)).toEqual(["5002"]);
  });

  it("keeps invoice and payment columns when a row is refreshed", () => {
    const row = core.orderRow(core.normalizeOrder(sample), { invoiceNumber: "INV-2026-0003", paid: true, paidOn: "2026-11-05" });
    const at = (name) => row[core.ORDER_COLUMNS.indexOf(name)];
    expect(row).toHaveLength(core.ORDER_COLUMNS.length);
    expect(at("Invoice no.")).toBe("INV-2026-0003");
    expect(at("Paid")).toBe(true);
    expect(at("Total")).toBe("52.60");
    expect(core.lineRows(core.normalizeOrder(sample))[1]).toEqual(["5001", "#1001", "TEA-OOL-50", "High Mountain Oolong", 1, "23.00", "23.00"]);
  });
});

describe("invoices", () => {
  it("numbers invoices per year", () => {
    expect(core.nextInvoiceNumber("", 2026)).toBe("INV-2026-0001");
    expect(core.nextInvoiceNumber("INV-2026-0009", 2026)).toBe("INV-2026-0010");
    expect(core.nextInvoiceNumber("INV-2025-0412", 2026)).toBe("INV-2026-0001");
    expect(core.nextInvoiceNumber("INV-2026-9999", 2026)).toBe("INV-2026-10000");
  });

  it("adds calendar days across month and year ends", () => {
    expect(core.addDays("2026-11-09", 14)).toBe("2026-11-23");
    expect(core.addDays("2026-12-25", 14)).toBe("2027-01-08");
    expect(core.addDays("2028-02-20", 10)).toBe("2028-03-01");
  });

  it("works out the status", () => {
    const today = "2026-11-20";
    expect(core.invoiceStatus({}, today)).toBe("Not invoiced");
    expect(core.invoiceStatus({ invoiceNumber: "INV-2026-0001", dueOn: "2026-11-23" }, today)).toBe("Unpaid");
    expect(core.invoiceStatus({ invoiceNumber: "INV-2026-0001", dueOn: "2026-11-19" }, today)).toBe("Overdue");
    expect(core.invoiceStatus({ invoiceNumber: "INV-2026-0001", dueOn: "2026-11-20" }, today)).toBe("Unpaid"); // due today is not late
    expect(core.invoiceStatus({ invoiceNumber: "INV-2026-0001", dueOn: "2026-11-19", paid: true }, today)).toBe("Paid");
    expect(core.invoiceStatus({ cancelled: true, paid: true }, today)).toBe("Cancelled");
  });

  it("fills template placeholders and reports missing ones", () => {
    const out = core.fillPlaceholders("Invoice {{InvoiceNumber}} for {{ CustomerName }} {{Unknown}}", { InvoiceNumber: "INV-2026-0001", CustomerName: "Maya Lin" });
    expect(out.text).toBe("Invoice INV-2026-0001 for Maya Lin ");
    expect(out.missing).toEqual(["Unknown"]);
  });

  it("builds the email text", () => {
    const values = core.invoiceValues(core.normalizeOrder(sample), { number: "INV-2026-0001", date: "2026-11-09", dueOn: "2026-11-23" },
      { name: "Harbor Leaf Tea", address: "1 Demo Way", email: "hello@harborleaf.example.com" });
    expect(values.Total).toBe("$52.60");
    expect(values.Discount).toBe("-$5.00");
    const mail = core.invoiceEmail(values, "invoice");
    expect(mail.subject).toBe("Invoice INV-2026-0001 from Harbor Leaf Tea");
    expect(mail.body).toContain("Payment is due by 2026-11-23.");
    expect(mail.body).not.toMatch(/\n\n\n/);
    expect(core.invoiceEmail(values, "reminder").subject).toMatch(/^Reminder: /);
  });

  it("sends only to the owner in test mode", () => {
    expect(core.resolveRecipient("maya.lin@example.com", true, "me@example.com"))
      .toEqual({ to: "me@example.com", subjectPrefix: "[TEST to maya.lin@example.com] " });
    expect(core.resolveRecipient("maya.lin@example.com", false, "me@example.com").to).toBe("maya.lin@example.com");
    expect(() => core.resolveRecipient("", false, "me@example.com")).toThrow("no email");
  });
});

describe("summary", () => {
  it("groups by month and leaves out cancelled orders", () => {
    const rows = [
      { created: "2026-11-02", total: 5000, invoiceNumber: "INV-1", paid: true, status: "Paid" },
      { created: "2026-11-05", total: 3000, invoiceNumber: "INV-2", paid: false, status: "Overdue" },
      { created: "2026-11-06", total: 2000, invoiceNumber: "", paid: false, status: "Not invoiced" },
      { created: "2026-11-07", total: 9900, invoiceNumber: "", paid: false, status: "Cancelled" },
      { created: "2026-10-30", total: 1000, invoiceNumber: "INV-0", paid: false, status: "Unpaid" },
    ];
    const [nov, oct] = core.monthlySummary(rows);
    expect(nov).toEqual({ month: "2026-11", orders: 3, sales: 10000, invoiced: 8000, paid: 5000, outstanding: 3000, overdue: 1 });
    expect(oct.month).toBe("2026-10");
  });
});

describe("demo data", () => {
  it("every demo order imports and adds up (or is marked)", () => {
    expect(orders.length).toBeGreaterThanOrEqual(8);
    const normalized = orders.map(core.normalizeOrder);
    expect(new Set(normalized.map((o) => o.id)).size).toBe(orders.length);
    for (const o of normalized) {
      expect(o.email).toMatch(/@example\.com$/);
      expect(core.totalsNote(o)).toBe("");
    }
    expect(normalized.some((o) => o.cancelled)).toBe(true);
  });
});

describe("Shopify pagination", () => {
  it("finds the next page in the Link header", () => {
    const link = '<https://shop.example.com/admin/api/2024-10/orders.json?page_info=abc&limit=250>; rel="previous", <https://shop.example.com/admin/api/2024-10/orders.json?page_info=def&limit=250>; rel="next"';
    expect(core.nextPageUrl(link)).toBe("https://shop.example.com/admin/api/2024-10/orders.json?page_info=def&limit=250");
    expect(core.nextPageUrl('<https://x.example.com/a>; rel="previous"')).toBe("");
    expect(core.nextPageUrl(undefined)).toBe("");
  });
});

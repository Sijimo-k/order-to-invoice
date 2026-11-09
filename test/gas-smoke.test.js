// Runs the Apps Script files in Node against small in-memory fakes of the Google services,
// to catch wiring mistakes (wrong column, missing function) before trying it in a real spreadsheet.
import { describe, expect, it } from "vitest";
import { readFileSync, readdirSync } from "node:fs";
import vm from "node:vm";

const demo = readFileSync(new URL("../demo/orders.json", import.meta.url), "utf8");

function makeSheet(name) {
  const cells = []; // cells[row][col], 1-based
  const sheet = {
    name,
    cells,
    getName: () => name,
    getLastRow: () => {
      for (let r = cells.length - 1; r >= 1; r--) if (cells[r] && cells[r].some((v) => v !== "" && v !== undefined)) return r; // like Sheets, FALSE counts as content
      return 0;
    },
    getLastColumn: () => Math.max(0, ...cells.map((r) => (r ? r.length - 1 : 0))),
    getMaxRows: () => 1000,
    getMaxColumns: () => 26,
    getRange: (a, b, c, d) => {
      if (typeof a === "string") {
        const col = a.charCodeAt(0) - 64;
        return makeRange(sheet, a.includes("2") ? 2 : 1, col, 999, 1);
      }
      return makeRange(sheet, a, b, c || 1, d || 1);
    },
    appendRow: (values) => makeRange(sheet, sheet.getLastRow() + 1, 1, 1, values.length).setValues([values]),
    setFrozenRows: () => sheet,
    setColumnWidth: () => sheet,
    setConditionalFormatRules: () => sheet,
  };
  return sheet;
}

function makeRange(sheet, row, col, rows, cols) {
  const get = (r, c) => (sheet.cells[r] && sheet.cells[r][c] !== undefined ? sheet.cells[r][c] : "");
  const range = {
    getRow: () => row,
    getLastRow: () => row + rows - 1,
    getColumn: () => col,
    getLastColumn: () => col + cols - 1,
    getSheet: () => sheet,
    getValues: () => Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => get(row + i, col + j))),
    getValue: () => get(row, col),
    setValues: (values) => {
      values.forEach((v, i) => v.forEach((x, j) => {
        sheet.cells[row + i] = sheet.cells[row + i] || [];
        // Sheets turns "12.50" into a number; mimic that for plain numeric text.
        sheet.cells[row + i][col + j] = typeof x === "string" && /^-?\d+(\.\d+)?$/.test(x) && !sheet.textCols?.has(col + j) ? Number(x) : x;
      }));
      return range;
    },
    setValue: (v) => range.setValues([[v]]),
    clearContent: () => range.setValues(Array.from({ length: rows }, () => Array(cols).fill(""))),
    setNumberFormat: (f) => {
      if (f === "@") for (let j = 0; j < cols; j++) (sheet.textCols = sheet.textCols || new Set()).add(col + j);
      return range;
    },
    insertCheckboxes: () => range.setValues(Array.from({ length: rows }, () => Array(cols).fill(false))),
    uncheck: () => range.setValues(Array.from({ length: rows }, () => Array(cols).fill(false))),
    check: () => range.setValue(true),
    setFontWeight: () => range,
    setBackground: () => range,
  };
  return range;
}

function loadScript() {
  const sheets = {};
  const files = {};
  const mails = [];
  const props = {};
  let id = 0;
  const ss = {
    getName: () => "Demo shop",
    getSheetByName: (n) => sheets[n] || null,
    insertSheet: (n) => (sheets[n] = makeSheet(n)),
    getSheets: () => Object.values(sheets),
    deleteSheet: (s) => delete sheets[s.getName()],
    setActiveSheet: () => {},
    getSpreadsheetTimeZone: () => "America/Los_Angeles",
    toast: () => {},
  };
  const textEl = (doc) => ({ findText: () => null, replaceText: () => {} , doc });
  const fakeDoc = (docId) => {
    const body = {
      clear: () => body, setMarginTop: () => body, setMarginBottom: () => body, setMarginLeft: () => body, setMarginRight: () => body,
      appendParagraph: () => ({ setHeading: () => {} }),
      appendTable: () => ({ setBorderWidth: () => ({ getRow: () => ({ editAsText: () => ({ setBold: () => {} }) }) }), getRow: () => ({ editAsText: () => ({ setBold: () => {} }) }) }),
      findText: () => null,
    };
    return { getId: () => docId, getBody: () => body, saveAndClose: () => {} };
  };
  const file = (name, extra = {}) => {
    const fid = "f" + ++id;
    const f = {
      getId: () => fid, getUrl: () => `https://drive.google.com/file/d/${fid}/view`, getName: () => name,
      makeCopy: (n) => file(n), moveTo: () => f, setTrashed: () => {}, getAs: () => ({ setName: (n) => ({ name: n }) }),
      getBlob: () => ({ name }), ...extra,
    };
    files[fid] = f;
    return f;
  };
  const pdfs = [];
  const folder = {
    getId: () => "folder1", getName: () => "Demo shop - invoices",
    createFile: (blob) => { const f = file(blob.name); pdfs.push(f); return f; },
    getFilesByType: () => { const list = pdfs.splice(0); return { hasNext: () => list.length > 0, next: () => list.shift() }; },
  };
  const context = {
    console,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      getUi: () => ({ alert: () => {}, createMenu: () => ({ addItem() { return this; }, addSeparator() { return this; }, addToUi() {} }) }),
      newConditionalFormatRule: () => { const b = { whenTextEqualTo: () => b, setBackground: () => b, setFontColor: () => b, setRanges: () => b, build: () => ({}) }; return b; },
    },
    DriveApp: {
      createFolder: () => folder, getFolderById: () => folder,
      getFileById: (fid) => files[fid] || file("template"),
    },
    DocumentApp: { create: () => fakeDoc("template1"), openById: (d) => fakeDoc(d), ParagraphHeading: {}, ElementType: { TABLE_ROW: "TABLE_ROW" } },
    UrlFetchApp: { fetch: () => { throw new Error("demo source must not use the network"); } },
    MailApp: { sendEmail: (m) => mails.push(m) },
    Session: { getEffectiveUser: () => ({ getEmail: () => "owner@example.com" }) },
    Utilities: { formatDate: () => "2026-11-30" },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => props[k] || null }) },
    ScriptApp: { getProjectTriggers: () => [], newTrigger: () => ({}) },
    MimeType: { PDF: "application/pdf" },
  };
  vm.createContext(context);
  const dir = new URL("../src/", import.meta.url);
  for (const name of readdirSync(dir).filter((n) => n.endsWith(".js")).sort()) {
    // Core.js ends with a Node-only export guard; "module" is undefined here, as in Apps Script.
    vm.runInContext(readFileSync(new URL(name, dir), "utf8"), context, { filename: name });
  }
  return { context, sheets, mails };
}

describe("Apps Script flow (with fakes)", () => {
  it("sets up, imports, invoices, emails in test mode, and tracks payment", () => {
    const { context: gas, sheets, mails } = loadScript();
    gas.setUp();
    expect(Object.keys(sheets).sort()).toEqual(["Line items", "Log", "Orders", "Settings", "Summary"]);

    expect(gas.importOrders()).toEqual({ added: 10, updated: 0 });
    expect(gas.importOrders()).toEqual({ added: 0, updated: 10 }); // running twice adds nothing
    const rows = gas.readOrderRows_();
    expect(rows).toHaveLength(10);
    expect(rows[0].rowNumber).toBe(2); // right under the header, not after the 1,000 checkbox rows
    expect(rows.find((r) => r.name === "#1004").status).toBe("Cancelled");
    expect(rows.find((r) => r.name === "#1001").address).toContain("Portland");
    expect(sheets["Line items"].getLastRow() - 1).toBe(18);

    expect(gas.createInvoices()).toEqual({ created: 9, more: false }); // the cancelled one is skipped
    const invoiced = gas.readOrderRows_().filter((r) => r.invoiceNumber);
    expect(invoiced.map((r) => r.invoiceNumber)).toContain("INV-2026-0009");
    expect(invoiced[0].dueOn).toBe("2026-12-14");
    expect(gas.createInvoices().created).toBe(0); // nothing twice

    expect(gas.sendInvoices().sent).toBe(9);
    expect(mails.every((m) => m.to === "owner@example.com" && m.subject.startsWith("[TEST to "))).toBe(true);
    expect(gas.sendInvoices().sent).toBe(0);

    // Tick "Paid" on the first invoiced row through the simple trigger.
    const row = invoiced[0].rowNumber;
    const paidCol = gas.ORDER_COLUMNS.indexOf("Paid") + 1;
    sheets.Orders.getRange(row, paidCol).setValue(true);
    gas.onEdit({ range: sheets.Orders.getRange(row, paidCol) });
    const paid = gas.readOrderRows_().find((r) => r.rowNumber === row);
    expect(paid.status).toBe("Paid");
    expect(paid.paidOn).toBe("2026-11-30");

    const summary = sheets.Summary.getRange(2, 1, 2, 7).getValues();
    expect(summary.map((r) => r[0])).toEqual(["2026-11", "2026-10"]);

    expect(gas.resetDemo()).toBe(9); // the 9 PDFs go to the trash
    expect(gas.readOrderRows_()).toHaveLength(0);
    expect(gas.importOrders()).toEqual({ added: 10, updated: 0 });
    expect(gas.readOrderRows_()[0].rowNumber).toBe(2);
    gas.createInvoices();
    expect(gas.readOrderRows_()[0].invoiceNumber).toBe("INV-2026-0001"); // counter restarted
  });
});

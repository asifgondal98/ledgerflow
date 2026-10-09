/**
 * Full "Statement of Account" Excel (.xlsx) export.
 *
 * Mirrors everything the Print / PDF statement shows: business header,
 * statement title/period/date, client details, outstanding summary boxes,
 * outstanding invoices table, account ledger (with opening balance) and the
 * closing summary — all in one sheet so it reads top to bottom like the PDF.
 */

export interface StatementExcelInput {
  filename: string;
  business: {
    name: string;
    address: string;
    email: string;
    phone: string;
    vatNumber: string;
    companyNumber: string;
  };
  client: {
    company: string;
    name: string;
    address: string;
    vatNumber: string;
    accountReference: string;
  };
  endClientLabel?: string | undefined;
  periodLabel?: string | undefined;
  /** ISO yyyy-mm-dd */
  statementDate: string;
  outstandingTotals: { total: number; overdue: number; due: number };
  outstandingInvoices: {
    invoiceDate: string;
    number: string;
    outstanding: number;
    dueDate: string;
    status: "Due" | "Overdue";
    ageing: number | null;
    description: string;
    poReference: string;
    note: string;
  }[];
  openingBalance?: number | undefined;
  ledger: {
    date: string;
    type: string;
    reference: string;
    description: string;
    debit: number;
    credit: number;
    balance: number;
  }[];
  summary: {
    invoiced: number;
    paid: number;
    creditNotesNet: number;
    balanceDue: number;
  };
}

const MONEY_FMT = '"£"#,##0.00;[Red]-"£"#,##0.00';
const DATE_FMT = "dd mmm yyyy";
const HEAD_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8EDF3" } } as const;
const THIN = { style: "thin", color: { argb: "FF9AA5B1" } } as const;

/** Stops spreadsheet apps treating text starting with = + - @ as a formula. */
function safeText(v: string): string {
  return /^[=+\-@\t\r]/.test(v) ? `'${v}` : v;
}

/** ISO date -> Date at UTC midnight so Excel shows the same calendar day in any timezone. */
function isoToDate(iso: string): Date | string {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || "");
  if (!m) return iso || "";
  return new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
}

export async function downloadStatementXlsx(input: StatementExcelInput): Promise<void> {
  if (typeof window === "undefined") return;
  const ExcelJS = (await import("exceljs")).default;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Statement");

  // A Date | B Type/Status | C Reference | D Description | E Debit/Outstanding
  // F Credit/Due date | G Balance/Ageing | H PO No. | I Notes
  ws.columns = [
    { width: 15 },
    { width: 15 },
    { width: 20 },
    { width: 58 },
    { width: 18 },
    { width: 18 },
    { width: 18 },
    { width: 16 },
    { width: 38 },
  ];
  const LAST_COL = 9;

  const textRow = (value: string, opts: { bold?: boolean; size?: number } = {}) => {
    const r = ws.addRow([safeText(value)]);
    r.getCell(1).font = { bold: opts.bold ?? false, size: opts.size ?? 11 };
    r.getCell(1).alignment = { vertical: "top", wrapText: false };
    return r;
  };
  const blank = () => ws.addRow([]);
  const sectionTitle = (title: string) => {
    const r = ws.addRow([title]);
    r.getCell(1).font = { bold: true, size: 12 };
    return r;
  };
  const styleHeader = (row: import("exceljs").Row, cols: number) => {
    for (let c = 1; c <= cols; c++) {
      const cell = row.getCell(c);
      cell.font = { bold: true };
      cell.fill = HEAD_FILL;
      cell.border = { bottom: THIN };
      cell.alignment = { vertical: "middle", wrapText: true };
    }
  };

  // ---- Business header ---------------------------------------------------
  const b = input.business;
  if (b.name.trim()) textRow(b.name, { bold: true, size: 15 });
  for (const line of b.address.split(/\r?\n/).filter((l) => l.trim())) textRow(line);
  if (b.email.trim()) textRow(b.email);
  if (b.phone.trim()) textRow(b.phone);
  if (b.vatNumber.trim()) textRow(`VAT No. ${b.vatNumber}`);
  if (b.companyNumber.trim()) textRow(`Company No. ${b.companyNumber}`);
  if (ws.rowCount > 0) blank();

  // ---- Statement title block --------------------------------------------
  textRow("Statement of Account", { bold: true, size: 14 });
  if (input.periodLabel) textRow(`Period: ${input.periodLabel}`);
  {
    const r = ws.addRow(["Statement date:", isoToDate(input.statementDate)]);
    r.getCell(1).font = { bold: true };
    r.getCell(2).numFmt = DATE_FMT;
    r.getCell(2).alignment = { horizontal: "left" };
  }
  if (input.client.accountReference) {
    const r = ws.addRow(["Account ref:", safeText(input.client.accountReference)]);
    r.getCell(1).font = { bold: true };
    r.getCell(2).alignment = { horizontal: "left" };
  }
  blank();

  // ---- Client details ----------------------------------------------------
  const c = input.client;
  {
    const r = ws.addRow(["Statement for"]);
    r.getCell(1).font = { bold: true, color: { argb: "FF6B7785" } };
  }
  textRow(c.company, { bold: true, size: 13 });
  if (c.name.trim()) textRow(c.name);
  if (input.endClientLabel) textRow(`Client: ${input.endClientLabel}`);
  for (const line of c.address.split(/\r?\n/).filter((l) => l.trim())) textRow(line);
  if (c.vatNumber) textRow(`VAT No. ${c.vatNumber}`);
  blank();

  // ---- Outstanding summary + outstanding invoices -----------------------
  if (input.outstandingTotals.total > 0.004) {
    sectionTitle("Outstanding Balance");
    const heads = ws.addRow(["Total Outstanding", "", "Overdue Amount", "", "Due Amount"]);
    styleHeader(heads, 5);
    const vals = ws.addRow([
      input.outstandingTotals.total,
      "",
      input.outstandingTotals.overdue,
      "",
      input.outstandingTotals.due,
    ]);
    for (const col of [1, 3, 5]) {
      vals.getCell(col).numFmt = MONEY_FMT;
      vals.getCell(col).font = { bold: true, size: 12 };
      vals.getCell(col).alignment = { horizontal: "left" };
    }
    blank();

    if (input.outstandingInvoices.length > 0) {
      sectionTitle("Outstanding Invoices");
      const h = ws.addRow([
        "Invoice Date",
        "Status",
        "Invoice No.",
        "Description",
        "Outstanding",
        "Due Date",
        "Ageing (days)",
        "PO No.",
        "Notes",
      ]);
      styleHeader(h, LAST_COL);
      h.getCell(5).alignment = { horizontal: "right", vertical: "middle" };
      h.getCell(7).alignment = { horizontal: "right", vertical: "middle" };

      for (const o of input.outstandingInvoices) {
        const r = ws.addRow([
          isoToDate(o.invoiceDate),
          o.status,
          safeText(o.number),
          safeText(o.description || "—"),
          o.outstanding,
          isoToDate(o.dueDate),
          o.ageing ?? "—",
          safeText(o.poReference || "—"),
          safeText(o.note),
        ]);
        r.getCell(1).numFmt = DATE_FMT;
        r.getCell(1).alignment = { horizontal: "left", vertical: "top" };
        r.getCell(4).alignment = { wrapText: true, vertical: "top" };
        r.getCell(5).numFmt = MONEY_FMT;
        r.getCell(5).alignment = { horizontal: "right", vertical: "top" };
        r.getCell(6).numFmt = DATE_FMT;
        r.getCell(6).alignment = { horizontal: "left", vertical: "top" };
        r.getCell(7).alignment = { horizontal: "right", vertical: "top" };
        r.getCell(9).alignment = { wrapText: true, vertical: "top" };
        r.getCell(2).font = {
          bold: true,
          color: { argb: o.status === "Overdue" ? "FFC0392B" : "FFB9770E" },
        };
      }
      blank();
    }
  }

  // ---- Account ledger ----------------------------------------------------
  if (input.ledger.length === 0 && input.openingBalance === undefined) {
    textRow(
      input.periodLabel
        ? `There is no account activity in ${input.periodLabel} for this client.`
        : "There is no account activity to show for this client.",
    );
  } else {
    sectionTitle("Account Ledger");
    const h = ws.addRow(["Date", "Type", "Reference", "Description", "Debit", "Credit", "Balance"]);
    styleHeader(h, 7);
    for (const col of [5, 6, 7]) h.getCell(col).alignment = { horizontal: "right", vertical: "middle" };

    if (input.openingBalance !== undefined) {
      const r = ws.addRow(["", "—", "", "Opening balance (brought forward)", "", "", input.openingBalance]);
      r.getCell(7).numFmt = MONEY_FMT;
      r.getCell(7).alignment = { horizontal: "right" };
      r.font = { italic: true };
    }
    for (const l of input.ledger) {
      const r = ws.addRow([
        isoToDate(l.date),
        l.type,
        safeText(l.reference),
        safeText(l.description),
        l.debit ? l.debit : "",
        l.credit ? l.credit : "",
        l.balance,
      ]);
      r.getCell(1).numFmt = DATE_FMT;
      r.getCell(1).alignment = { horizontal: "left", vertical: "top" };
      r.getCell(4).alignment = { wrapText: true, vertical: "top" };
      for (const col of [5, 6, 7]) {
        r.getCell(col).numFmt = MONEY_FMT;
        r.getCell(col).alignment = { horizontal: "right", vertical: "top" };
      }
    }
  }
  blank();

  // ---- Closing summary ---------------------------------------------------
  sectionTitle("Summary");
  const summaryLine = (label: string, value: number, bold = false) => {
    const r = ws.addRow([label, "", "", "", value]);
    r.getCell(5).numFmt = MONEY_FMT;
    r.getCell(5).alignment = { horizontal: "right" };
    r.getCell(1).font = { bold };
    r.getCell(5).font = { bold };
    return r;
  };
  summaryLine("Total invoiced", input.summary.invoiced);
  summaryLine("Total paid", input.summary.paid);
  if (Math.abs(input.summary.creditNotesNet) > 0.004) {
    summaryLine("Credit notes (net)", input.summary.creditNotesNet);
  }
  const total = summaryLine("Balance due", input.summary.balanceDue, true);
  for (let col = 1; col <= 5; col++) total.getCell(col).border = { top: THIN };

  if (b.email.trim() || b.phone.trim()) {
    blank();
    const contact = [b.email, b.phone].filter((v) => v.trim()).join(" or ");
    textRow(`If anything on this statement looks wrong, please contact ${contact}.`);
  }

  // ---- Print setup (so printing the xlsx also looks right) --------------
  ws.pageSetup = {
    paperSize: 9,
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
  };

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = input.filename.endsWith(".xlsx") ? input.filename : `${input.filename}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

import { activeSession, qboQuery } from "../server/qbo.js";

const DATASETS = {
  items: {
    label: "Products & Services",
    sheetName: "Products_Services",
    entity: "Item",
    map: (item) => [
      item.Id,
      item.Name || "",
      item.Type || "",
      item.Sku || "",
      item.Description || "",
      item.UnitPrice || 0,
      item.Active !== false,
    ],
    headers: ["QuickBooks ID", "Name", "Type", "SKU", "Description", "Unit Price", "Active"],
  },
  employees: {
    label: "Employees",
    sheetName: "Employees",
    entity: "Employee",
    map: (item) => [
      item.Id,
      item.DisplayName || `${item.GivenName || ""} ${item.FamilyName || ""}`.trim(),
      item.PrimaryEmailAddr?.Address || "",
      item.PrimaryPhone?.FreeFormNumber || "",
      item.Active !== false,
    ],
    headers: ["QuickBooks ID", "Name", "Email", "Phone", "Active"],
  },
  locations: {
    label: "Locations",
    sheetName: "Locations",
    entity: "Department",
    map: (item) => [item.Id, item.FullyQualifiedName || item.Name || "", item.SubDepartment || false, item.Active !== false],
    headers: ["QuickBooks ID", "Name", "Sub-location", "Active"],
  },
  terms: {
    label: "Payment Terms",
    sheetName: "Payment_Terms",
    entity: "Term",
    map: (item) => [
      item.Id,
      item.Name || "",
      item.Type || "",
      item.DueDays || "",
      item.DiscountDays || "",
      item.DiscountPercent || "",
      item.Active !== false,
    ],
    headers: ["QuickBooks ID", "Name", "Type", "Due Days", "Discount Days", "Discount %", "Active"],
  },
  paymentMethods: {
    label: "Payment Methods",
    sheetName: "Payment_Methods",
    entity: "PaymentMethod",
    map: (item) => [item.Id, item.Name || "", item.Type || "", item.Active !== false],
    headers: ["QuickBooks ID", "Name", "Type", "Active"],
  },
  taxCodes: {
    label: "Tax Codes",
    sheetName: "Tax_Codes",
    entity: "TaxCode",
    map: (item) => [
      item.Id,
      item.Name || "",
      item.Description || "",
      item.Taxable || false,
      item.Active !== false,
    ],
    headers: ["QuickBooks ID", "Name", "Description", "Taxable", "Active"],
  },
  invoices: {
    label: "Invoices",
    sheetName: "Invoices",
    entity: "Invoice",
    transaction: true,
    map: (item) => [
      item.Id,
      item.DocNumber || "",
      item.TxnDate || "",
      item.CustomerRef?.name || "",
      item.TotalAmt || 0,
      item.Balance || 0,
      item.DueDate || "",
    ],
    headers: ["QuickBooks ID", "Document #", "Date", "Customer", "Total", "Balance", "Due Date"],
  },
  bills: {
    label: "Bills",
    sheetName: "Bills",
    entity: "Bill",
    transaction: true,
    map: (item) => [
      item.Id,
      item.DocNumber || "",
      item.TxnDate || "",
      item.VendorRef?.name || "",
      item.TotalAmt || 0,
      item.Balance || 0,
      item.DueDate || "",
    ],
    headers: ["QuickBooks ID", "Document #", "Date", "Vendor", "Total", "Balance", "Due Date"],
  },
  payments: {
    label: "Customer Payments",
    sheetName: "Customer_Payments",
    entity: "Payment",
    transaction: true,
    map: (item) => [
      item.Id,
      item.TxnDate || "",
      item.CustomerRef?.name || "",
      item.TotalAmt || 0,
      item.PaymentMethodRef?.name || "",
      item.PaymentRefNum || "",
    ],
    headers: ["QuickBooks ID", "Date", "Customer", "Amount", "Payment Method", "Reference"],
  },
  expenses: {
    label: "Expenses",
    sheetName: "Expenses",
    entity: "Purchase",
    transaction: true,
    map: (item) => [
      item.Id,
      item.TxnDate || "",
      item.EntityRef?.name || "",
      item.PaymentType || "",
      item.TotalAmt || 0,
      item.AccountRef?.name || "",
      item.DocNumber || "",
    ],
    headers: ["QuickBooks ID", "Date", "Payee", "Payment Type", "Total", "Account", "Document #"],
  },
  deposits: {
    label: "Deposits",
    sheetName: "Deposits",
    entity: "Deposit",
    transaction: true,
    map: (item) => [
      item.Id,
      item.TxnDate || "",
      item.DepositToAccountRef?.name || "",
      item.TotalAmt || 0,
      item.PrivateNote || "",
    ],
    headers: ["QuickBooks ID", "Date", "Deposit Account", "Total", "Memo"],
  },
  purchaseOrders: {
    label: "Purchase Orders",
    sheetName: "Purchase_Orders",
    entity: "PurchaseOrder",
    transaction: true,
    map: (item) => [
      item.Id,
      item.DocNumber || "",
      item.TxnDate || "",
      item.VendorRef?.name || "",
      item.TotalAmt || 0,
      item.POStatus || "",
    ],
    headers: ["QuickBooks ID", "Document #", "Date", "Vendor", "Total", "Status"],
  },
  journalEntries: {
    label: "Journal Entries",
    sheetName: "Journal_Entries",
    entity: "JournalEntry",
    transaction: true,
    map: (item) => [
      item.Id,
      item.DocNumber || "",
      item.TxnDate || "",
      item.TotalAmt || 0,
      item.PrivateNote || "",
      item.Adjustment || false,
    ],
    headers: ["QuickBooks ID", "Document #", "Date", "Total", "Memo", "Adjustment"],
  },
};

function isoDate(date) {
  return date.toISOString().slice(0, 10);
}

function addDays(date, days) {
  const copy = new Date(date);
  copy.setUTCDate(copy.getUTCDate() + days);
  return copy;
}

function dateRangePreset(key) {
  const today = new Date();
  const end = isoDate(today);

  if (key === "last30") {
    return {
      start: isoDate(addDays(today, -30)),
      end,
      label: "Last 30 days",
    };
  }

  if (key === "ytd") {
    return {
      start: `${today.getUTCFullYear()}-01-01`,
      end,
      label: "Year to date",
    };
  }

  if (key === "all") {
    return {
      start: "",
      end: "",
      label: "All sandbox data",
    };
  }

  return {
    start: isoDate(addDays(today, -90)),
    end,
    label: "Last 90 days",
  };
}

function buildQuery(dataset, dateRangeKey) {
  if (!dataset.transaction) {
    return {
      query: `select * from ${dataset.entity} maxresults 1000`,
      dateRangeLabel: "",
    };
  }

  const range = dateRangePreset(dateRangeKey);
  if (!range.start || !range.end) {
    return {
      query: `select * from ${dataset.entity} maxresults 1000`,
      dateRangeLabel: range.label,
    };
  }

  return {
    query: `select * from ${dataset.entity} where TxnDate >= '${range.start}' and TxnDate <= '${range.end}' maxresults 1000`,
    dateRangeLabel: range.label,
  };
}

export default async function handler(req, res) {
  if (req.method !== "GET") {
    res.setHeader("Allow", "GET");
    return res.status(405).json({ error: "Method not allowed." });
  }

  const dataset = DATASETS[String(req.query.dataset || "")];
  if (!dataset) return res.status(400).json({ error: "Unsupported QuickBooks dataset." });

  try {
    const session = await activeSession(req, res);
    if (!session) return res.status(401).json({ error: "QuickBooks is not connected." });

    const { query, dateRangeLabel } = buildQuery(dataset, String(req.query.dateRange || "last90"));
    const result = await qboQuery(session, query);
    const rows = (result[dataset.entity] || []).map(dataset.map);
    return res.status(200).json({
      success: true,
      label: dataset.label,
      sheetName: dataset.sheetName,
      headers: dataset.headers,
      rows,
      count: rows.length,
      dateRangeLabel,
    });
  } catch (error) {
    return res.status(502).json({ error: error.message });
  }
}

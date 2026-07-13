import { activeSession, qboCreate, qboQuery } from "../server/qbo.js";

function parseJsonBody(req) {
  if (!req.body) return {};
  if (typeof req.body === "object") return req.body;
  try {
    return JSON.parse(req.body);
  } catch {
    return {};
  }
}

function normalizeText(value) {
  return String(value || "").trim();
}

function normalizeKey(value) {
  return normalizeText(value).toLowerCase();
}

function normalizeAmount(value) {
  if (value === "" || value === null || value === undefined) return 0;
  const amount = Number(typeof value === "string" ? value.replace(/[$,\s]/g, "") : value);
  return Number.isFinite(amount) ? amount : null;
}

function normalizeDate(value) {
  if (!value) return null;

  if (typeof value === "number") {
    const excelEpoch = Date.UTC(1899, 11, 30);
    return new Date(excelEpoch + value * 24 * 60 * 60 * 1000).toISOString().slice(0, 10);
  }

  const parsed = new Date(value);
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);

  return null;
}

function roundMoney(value) {
  return Math.round(value * 100) / 100;
}

function hasMoreThanTwoDecimals(value) {
  if (!Number.isFinite(value)) return false;
  return Math.abs(value * 100 - Math.round(value * 100)) > 0.000001;
}

function validateSupportedJournalLines(lines) {
  const allowedFields = new Set([
    "lineNo",
    "rowNumber",
    "account",
    "vendor",
    "className",
    "description",
    "debit",
    "credit",
    "date",
  ]);

  const errors = [];
  const normalizedLines = [];
  let totalDebit = 0;
  let totalCredit = 0;
  let txnDate = null;

  if (!Array.isArray(lines)) {
    return { valid: false, errors: ["Journal lines must be submitted as an array."] };
  }

  if (lines.length === 0) {
    return { valid: false, errors: ["No journal lines were submitted."] };
  }

  // 🔥 NEW: Remove blank rows before validation
  lines = lines.filter((line) => {
    const account = normalizeText(line.account);
    const debit = normalizeAmount(line.debit);
    const credit = normalizeAmount(line.credit);

    const isBlank =
      !account &&
      (!debit || debit === 0) &&
      (!credit || credit === 0);

    return !isBlank;
  });

  if (lines.length === 0) {
    return { valid: false, errors: ["No journal lines with data were submitted."] };
  }

  if (lines.length > 200) {
    errors.push("Journal entries are limited to 200 lines in this workflow.");
  }

  lines.forEach((line, index) => {
    const rawLine = line && typeof line === "object" && !Array.isArray(line) ? line : {};
    const submittedRowNumber = Number(rawLine.rowNumber);
    const rowNumber =
      Number.isInteger(submittedRowNumber) && submittedRowNumber > 0 ? submittedRowNumber : index + 2;
    const unknownFields = Object.keys(rawLine).filter((key) => !allowedFields.has(key));

    if (unknownFields.length) {
      errors.push(`Row ${rowNumber}: unsupported field(s): ${unknownFields.join(", ")}.`);
    }

    const account = normalizeText(rawLine.account);
    const vendor = normalizeText(rawLine.vendor);
    const className = normalizeText(rawLine.className);
    const description = normalizeText(rawLine.description);
    const debit = normalizeAmount(rawLine.debit);
    const credit = normalizeAmount(rawLine.credit);
    const rowDate = normalizeDate(rawLine.date);

    if (!account) {
      errors.push(`Row ${rowNumber}: account is required.`);
    }

    if (account.length > 100) {
      errors.push(`Row ${rowNumber}: account is too long.`);
    }

    if (vendor.length > 100) {
      errors.push(`Row ${rowNumber}: vendor is too long.`);
    }

    if (className.length > 100) {
      errors.push(`Row ${rowNumber}: class is too long.`);
    }

    if (description.length > 4000) {
      errors.push(`Row ${rowNumber}: description is too long.`);
    }

    if (!rowDate) {
      errors.push(`Row ${rowNumber}: date is required and must be valid.`);
    } else if (!txnDate) {
      txnDate = rowDate;
    } else if (rowDate !== txnDate) {
      errors.push(`Row ${rowNumber}: all lines must use the same journal date (${txnDate}).`);
    }

    if (debit === null || credit === null) {
      errors.push(`Row ${rowNumber}: debit and credit must be numeric.`);
      return;
    }

    if (debit < 0 || credit < 0) {
      errors.push(`Row ${rowNumber}: negative amounts are not allowed.`);
    }

    if (hasMoreThanTwoDecimals(debit) || hasMoreThanTwoDecimals(credit)) {
      errors.push(`Row ${rowNumber}: debit and credit can have at most two decimal places.`);
    }

    if ((debit > 0 && credit > 0) || (debit <= 0 && credit <= 0)) {
      errors.push(`Row ${rowNumber}: enter an amount on exactly one side.`);
    }

    totalDebit = roundMoney(totalDebit + debit);
    totalCredit = roundMoney(totalCredit + credit);

    normalizedLines.push({
      rowNumber,
      account,
      vendor,
      className,
      description,
      debit,
      credit,
      date: rowDate,
    });
  });

  if (totalDebit <= 0 && totalCredit <= 0) {
    errors.push("Journal total must be greater than zero.");
  }

  if (Math.abs(totalDebit - totalCredit) > 0.01) {
    errors.push(`Journal entry is unbalanced: debits ${totalDebit.toFixed(2)}, credits ${totalCredit.toFixed(2)}.`);
  }

  return {
    valid: errors.length === 0,
    errors,
    lines: normalizedLines,
    txnDate,
    totalDebit,
    totalCredit,
  };
}

function createLookup(items, keys) {
  const map = new Map();
  items.forEach((item) => {
    keys.forEach((key) => {
      const value = key(item);
      if (value) map.set(normalizeKey(value), item);
    });
  });
  return map;
}

function findRequired(lookup, value, label, rowNumber) {
  const item = lookup.get(normalizeKey(value));
  if (!item) throw new Error(`Row ${rowNumber}: ${label} "${value}" was not found in QuickBooks.`);
  return item;
}

function findOptional(lookup, value, label, rowNumber) {
  if (!normalizeText(value)) return null;
  return findRequired(lookup, value, label, rowNumber);
}

async function loadReferenceLookups(session) {
  // Accounts are required — propagate QBO errors so the handler can return 502
  const accountResult = await qboQuery(session, "select * from Account where Active = true maxresults 1000");
  const accounts = accountResult.Account || [];

  // Vendors and classes are optional — allow partial QBO failures
  const [vendorResult, classResult] = await Promise.allSettled([
    qboQuery(session, "select * from Vendor where Active = true maxresults 1000"),
    qboQuery(session, "select * from Class where Active = true maxresults 1000"),
  ]);

  const vendors = vendorResult.status === "fulfilled" ? vendorResult.value.Vendor || [] : [];
  const classes = classResult.status === "fulfilled" ? classResult.value.Class || [] : [];

  return {
    accounts: createLookup(accounts, [
      (item) => item.Id,
      (item) => item.Name,
      (item) => item.FullyQualifiedName,
      (item) => item.AcctNum,
    ]),
    vendors: createLookup(vendors, [
      (item) => item.Id,
      (item) => item.DisplayName,
      (item) => item.CompanyName,
    ]),
    classes: createLookup(classes, [
      (item) => item.Id,
      (item) => item.Name,
      (item) => item.FullyQualifiedName,
    ]),
  };
}

function buildJournalEntryPayload(validation, lookups) {
  const lines = validation.lines;
  const txnDate = validation.txnDate;
  const qboLines = lines.map((line, index) => {
    const rowNumber = line.rowNumber || index + 2;
    const debit = line.debit;
    const credit = line.credit;

    const amount = debit > 0 ? debit : credit;
    const postingType = debit > 0 ? "Debit" : "Credit";
    const account = findRequired(lookups.accounts, line.account, "account", rowNumber);
    const vendor = findOptional(lookups.vendors, line.vendor, "vendor", rowNumber);
    const classRef = findOptional(lookups.classes, line.className, "class", rowNumber);

    const detail = {
      PostingType: postingType,
      AccountRef: {
        value: account.Id,
        name: account.FullyQualifiedName || account.Name,
      },
    };

    if (vendor) {
      detail.Entity = {
        Type: "Vendor",
        EntityRef: {
          value: vendor.Id,
          name: vendor.DisplayName || vendor.CompanyName,
        },
      };
    }

    if (classRef) {
      detail.ClassRef = {
        value: classRef.Id,
        name: classRef.FullyQualifiedName || classRef.Name,
      };
    }

    return {
      DetailType: "JournalEntryLineDetail",
      Amount: amount,
      Description: normalizeText(line.description),
      JournalEntryLineDetail: detail,
    };
  });

  return {
    TxnDate: txnDate,
    PrivateNote: "Created from FinAccruals LedgerFlow Excel add-in.",
    Line: qboLines,
  };
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method not allowed." });
  }

  if (process.env.QBO_ENVIRONMENT === "production") {
    return res.status(403).json({ error: "Journal posting is currently enabled for sandbox only." });
  }

  try {
    const session = await activeSession(req, res);
    if (!session) return res.status(401).json({ error: "QuickBooks is not connected." });

    const { lines } = parseJsonBody(req);
    const validation = validateSupportedJournalLines(lines);
    if (!validation.valid) {
      return res.status(400).json({
        error: validation.errors.join(" "),
        errors: validation.errors,
      });
    }

    // 502 if QBO account lookup fails (network / auth issue)
    let lookups;
    try {
      lookups = await loadReferenceLookups(session);
    } catch (err) {
      return res.status(502).json({ error: `Failed to load accounts from QuickBooks: ${err.message}` });
    }

    // 400 if an account, vendor, or class name isn't found in QBO
    let payload;
    try {
      payload = buildJournalEntryPayload(validation, lookups);
    } catch (err) {
      return res.status(400).json({ error: err.message });
    }

    // 502 if the QBO create call itself fails
    let journalEntry;
    try {
      journalEntry = await qboCreate(session, "JournalEntry", payload);
    } catch (err) {
      return res.status(502).json({ error: err.message });
    }

    return res.status(200).json({
      success: true,
      id: journalEntry.Id,
      docNumber: journalEntry.DocNumber || "",
      date: journalEntry.TxnDate || payload.TxnDate,
      amount: journalEntry.TotalAmt || validation.totalDebit,
    });
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
}

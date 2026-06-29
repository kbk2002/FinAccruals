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
  if (!value) return new Date().toISOString().slice(0, 10);

  if (typeof value === "number") {
    const excelEpoch = Date.UTC(1899, 11, 30);
    return new Date(excelEpoch + value * 24 * 60 * 60 * 1000).toISOString().slice(0, 10);
  }

  const parsed = new Date(value);
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);

  return String(value).slice(0, 10);
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
  const [accountResult, vendorResult, classResult] = await Promise.allSettled([
    qboQuery(session, "select * from Account where Active = true maxresults 1000"),
    qboQuery(session, "select * from Vendor where Active = true maxresults 1000"),
    qboQuery(session, "select * from Class where Active = true maxresults 1000"),
  ]);

  const accounts = accountResult.status === "fulfilled" ? accountResult.value.Account || [] : [];
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

function buildJournalEntryPayload(lines, lookups) {
  if (!Array.isArray(lines) || lines.length === 0) {
    throw new Error("No journal lines were submitted.");
  }

  const txnDate = normalizeDate(lines[0].date);
  const qboLines = lines.map((line, index) => {
    const rowNumber = index + 2;
    const debit = normalizeAmount(line.debit);
    const credit = normalizeAmount(line.credit);

    if (!normalizeText(line.account)) {
      throw new Error(`Row ${rowNumber}: account is required.`);
    }

    if (debit === null || credit === null) {
      throw new Error(`Row ${rowNumber}: debit and credit must be numeric.`);
    }

    if (debit < 0 || credit < 0) {
      throw new Error(`Row ${rowNumber}: negative amounts are not allowed.`);
    }

    if ((debit > 0 && credit > 0) || (debit <= 0 && credit <= 0)) {
      throw new Error(`Row ${rowNumber}: enter an amount on exactly one side.`);
    }

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
    const lookups = await loadReferenceLookups(session);
    const payload = buildJournalEntryPayload(lines, lookups);
    const journalEntry = await qboCreate(session, "JournalEntry", payload);

    return res.status(200).json({
      success: true,
      id: journalEntry.Id,
      docNumber: journalEntry.DocNumber || "",
      date: journalEntry.TxnDate || payload.TxnDate,
      amount: journalEntry.TotalAmt || payload.Line.reduce((sum, line) => sum + line.Amount, 0) / 2,
    });
  } catch (error) {
    return res.status(400).json({ error: error.message });
  }
}

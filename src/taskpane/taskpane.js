const API_BASE = "https://fin-accruals.vercel.app/api";

Office.onReady(() => {
  document.getElementById("btnConnectQBO").onclick = handleConnectQBO;
  document.getElementById("btnPullAccounts").onclick = () => handlePullMasterData("accounts");
  document.getElementById("btnPullVendors").onclick = () => handlePullMasterData("vendors");

  const btnPullClasses = document.getElementById("btnPullClasses");
  if (btnPullClasses) {
    btnPullClasses.onclick = () => handlePullMasterData("classes");
  }

  document.getElementById("btnCreateTemplate").onclick = handleCreateTemplate;
  document.getElementById("btnValidateJE").onclick = handleValidateJE;
  document.getElementById("btnSubmitJE").onclick = handleSubmitJE;
  document.getElementById("btnLoadHistory").onclick = handleLoadHistory;
});

function setStatus(message, type = "info") {
  const bar = document.getElementById("statusBar");
  bar.textContent = message;
  bar.classList.remove("status-bar--info", "status-bar--success", "status-bar--error");
  bar.classList.add(`status-bar--${type}`);
}

function setAuthBadge(text, variant = "primary") {
  const badge = document.getElementById("authStatusBadge");
  badge.textContent = text;
  badge.className = `badge badge--${variant}`;
}

/* 1. Demo Connect to QBO */

async function handleConnectQBO() {
  setStatus("Demo QuickBooks connection active.", "success");
  setAuthBadge("Demo connected", "primary");
}

/* 2. Pull master data */

async function handlePullMasterData(kind) {
  try {
    const labelMap = {
      accounts: "Accounts",
      vendors: "Vendors",
      classes: "Classes",
    };

    const label = labelMap[kind];
    setStatus(`Pulling ${label.toLowerCase()} from deployed API...`, "info");

    const res = await fetch(`${API_BASE}/${kind}`);

    if (!res.ok) {
      throw new Error(`Failed to pull ${label.toLowerCase()}`);
    }

    const payload = await res.json();

    let data = [];

    if (kind === "accounts") data = payload.accounts || [];
    if (kind === "vendors") data = payload.vendors || [];
    if (kind === "classes") data = payload.classes || [];

    await writeMasterDataToSheet(kind, data);

    setStatus(`${label} updated in workbook.`, "success");
  } catch (err) {
    console.error(err);
    setStatus("Error pulling master data. See console.", "error");
  }
}

async function writeMasterDataToSheet(kind, data) {
  return Excel.run(async (context) => {
    const sheetNameMap = {
      accounts: "Accounts",
      vendors: "Vendors",
      classes: "Classes",
    };

    const sheetName = sheetNameMap[kind];

    const sheets = context.workbook.worksheets;
    sheets.load("items/name");

    await context.sync();

    let sheet = sheets.items.find((s) => s.name === sheetName);

    if (!sheet) {
      sheet = sheets.add(sheetName);
    }

    let headers = [];
    let rows = [];

    if (kind === "accounts") {
      headers = ["Code", "Name", "Type"];
      rows = data.map((item) => [item.code || item.id || "", item.name || "", item.type || ""]);
    }

    if (kind === "vendors") {
      headers = ["Name", "Email"];
      rows = data.map((item) => [item.name || "", item.email || ""]);
    }

    if (kind === "classes") {
      headers = ["Name"];
      rows = data.map((item) => [item.name || ""]);
    }

    const all = [headers, ...rows];

    sheet.getUsedRangeOrNullObject().clear();

    const range = sheet.getRangeByIndexes(0, 0, all.length, headers.length);
    range.values = all;
    range.format.font.name = "Aptos";
    range.format.autofitColumns();
    range.format.autofitRows();

    const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
    headerRange.format.font.bold = true;
    headerRange.format.fill.color = "#e5e7eb";

    sheet.activate();

    await context.sync();
  });
}

/* 3. Create / refresh JE template */

async function handleCreateTemplate() {
  try {
    setStatus("Creating / refreshing JE template...", "info");

    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");

      await context.sync();

      const sheetName = "JE_Template";
      let sheet = sheets.items.find((s) => s.name === sheetName);

      if (!sheet) {
        sheet = sheets.add(sheetName);
      }

      const headers = ["Line #", "Account", "Vendor", "Class", "Description", "Debit", "Credit", "Date"];

      const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
      headerRange.values = [headers];
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = "#e5e7eb";

      const dataRange = sheet.getRange("A2:H200");
      dataRange.clear();

      sheet.getRange("A1:H200").name = "JE_TABLE";

      sheet.activate();

      await context.sync();
    });

    setStatus("JE template ready. Fill in lines and validate.", "success");
  } catch (err) {
    console.error(err);
    setStatus("Error creating JE template.", "error");
  }
}

/* 4. Validate JE */

async function handleValidateJE() {
  try {
    const resultEl = document.getElementById("validationResult");
    resultEl.textContent = "Validating entry...";
    resultEl.className = "validation-result";

    const lines = await readJELinesFromSheet();

    if (lines.length === 0) {
      resultEl.textContent = "No lines found in JE_Template.";
      resultEl.classList.add("validation-result--error");
      return;
    }

    let totalDebit = 0;
    let totalCredit = 0;

    for (const line of lines) {
      totalDebit += Number(line.debit || 0);
      totalCredit += Number(line.credit || 0);
    }

    if (Math.abs(totalDebit - totalCredit) > 0.01) {
      resultEl.textContent = `Unbalanced entry. Debits: ${totalDebit.toFixed(
        2
      )}, Credits: ${totalCredit.toFixed(2)}.`;
      resultEl.classList.add("validation-result--error");
      setStatus("Validation failed: entry is not balanced.", "error");
      return;
    }

    resultEl.textContent = `Entry is balanced. Debits = Credits = ${totalDebit.toFixed(2)}.`;
    resultEl.classList.add("validation-result--ok");
    setStatus("Validation passed. Ready to submit.", "success");
  } catch (err) {
    console.error(err);
    setStatus("Error validating JE.", "error");
  }
}

async function readJELinesFromSheet() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("JE_Template");
    const range = sheet.getUsedRange();
    range.load("values");

    await context.sync();

    const values = range.values;

    if (values.length <= 1) return [];

    const lines = [];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      const [lineNo, account, vendor, className, description, debit, credit, date] = row;

      if (!account && !vendor && !className && !description && !debit && !credit) {
        continue;
      }

      lines.push({
        lineNo,
        account,
        vendor,
        className,
        description,
        debit,
        credit,
        date,
      });
    }

    return lines;
  });
}

/* 5. Submit JE demo */

async function handleSubmitJE() {
  try {
    setStatus("Reading JE and submitting to demo backend...", "info");

    const lines = await readJELinesFromSheet();

    if (lines.length === 0) {
      setStatus("No JE lines to submit.", "error");
      return;
    }

    const payload = {
      success: true,
      reference: `DEMO-JE-${Date.now()}`,
      lines,
    };

    setStatus("Journal entry submitted in demo mode.", "success");

    const resultEl = document.getElementById("validationResult");
    resultEl.textContent = `Submitted successfully. Ref: ${payload.reference}`;
    resultEl.className = "validation-result validation-result--ok";
  } catch (err) {
    console.error(err);
    setStatus("Error submitting JE. See console.", "error");

    const resultEl = document.getElementById("validationResult");
    resultEl.textContent = `Submission failed: ${err.message}`;
    resultEl.className = "validation-result validation-result--error";
  }
}

/* 6. History */

async function handleLoadHistory() {
  try {
    setStatus("Loading submission history...", "info");

    const res = await fetch(`${API_BASE}/history`);

    if (!res.ok) {
      throw new Error("Failed to load history");
    }

    const payload = await res.json();
    const items = payload.history || [];

    const container = document.getElementById("historyList");
    container.innerHTML = "";

    if (!items.length) {
      container.textContent = "No submissions found.";
      setStatus("History loaded empty.", "info");
      return;
    }

    items.forEach((item) => {
      const div = document.createElement("div");
      div.className = "history-item history-item--success";
      div.innerHTML = `
        <span>${item.id || ""}</span>
        <span>${item.date || ""}</span>
        <span>${item.status || ""}</span>
        <span>$${item.amount || 0}</span>
      `;
      container.appendChild(div);
    });

    setStatus("History loaded.", "success");
  } catch (err) {
    console.error(err);
    setStatus("Error loading history.", "error");
  }
}

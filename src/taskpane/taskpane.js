const API_BASE = "http://localhost:3001"; 

Office.onReady(() => {
  document.getElementById("btnConnectQBO").onclick = handleConnectQBO;
  document.getElementById("btnPullAccounts").onclick = () =>
    handlePullMasterData("accounts");
  document.getElementById("btnPullVendors").onclick = () =>
    handlePullMasterData("vendors");
  document.getElementById("btnCreateTemplate").onclick = handleCreateTemplate;
  document.getElementById("btnValidateJE").onclick = handleValidateJE;
  document.getElementById("btnSubmitJE").onclick = handleSubmitJE;
  document.getElementById("btnLoadHistory").onclick = handleLoadHistory;
});

function setStatus(message, type = "info") {
  const bar = document.getElementById("statusBar");
  bar.textContent = message;
  bar.classList.remove(
    "status-bar--info",
    "status-bar--success",
    "status-bar--error"
  );
  bar.classList.add(`status-bar--${type}`);
}

function setAuthBadge(text, variant = "primary") {
  const badge = document.getElementById("authStatusBadge");
  badge.textContent = text;
  badge.className = `badge badge--${variant}`;
}

/* 1. Connect to QBO */

async function handleConnectQBO() {
  try {
    setStatus("Opening QuickBooks login…", "info");
    const btn = document.getElementById("btnConnectQBO");
    btn.disabled = true;

    // Backend should return an OAuth URL to open
    const res = await fetch(`${API_BASE}/auth/qbo/start`, {
      credentials: "include",
    });
    if (!res.ok) throw new Error("Failed to start QBO auth");

    const { authUrl } = await res.json();
    // Open in external browser
    Office.context.ui.displayDialogAsync(
      authUrl,
      { height: 60, width: 30 },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          setStatus("Failed to open QuickBooks login.", "error");
          btn.disabled = false;
          return;
        }
        const dialog = result.value;
        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (arg) => {
            // Expect backend to send "success" or similar
            if (arg.message === "qbo-auth-success") {
              setStatus("Connected to QuickBooks Online.", "success");
              setAuthBadge("Connected", "primary");
              dialog.close();
              btn.disabled = false;
            }
          }
        );
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          btn.disabled = false;
        });
      }
    );
  } catch (err) {
    console.error(err);
    setStatus("Error connecting to QuickBooks. Check console.", "error");
    setAuthBadge("Not connected", "neutral");
    document.getElementById("btnConnectQBO").disabled = false;
  }
}

/* 2. Pull master data */

async function handlePullMasterData(kind) {
  try {
    const label = kind === "accounts" ? "Accounts" : "Vendors";
    setStatus(`Pulling ${label.toLowerCase()} from backend…`, "info");

    const endpoint =
      kind === "accounts" ? "/master/accounts" : "/master/vendors";

    const res = await fetch(`${API_BASE}${endpoint}`, {
      credentials: "include",
    });
    if (!res.ok) throw new Error(`Failed to pull ${label.toLowerCase()}`);

    const data = await res.json(); // expect array of objects
    await writeMasterDataToSheet(kind, data);

    setStatus(`${label} updated in workbook.`, "success");
  } catch (err) {
    console.error(err);
    setStatus("Error pulling master data. See console.", "error");
  }
}

async function writeMasterDataToSheet(kind, data) {
  return Excel.run(async (context) => {
    const sheetName = kind === "accounts" ? "Accounts" : "Vendors";
    let sheet;
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");

    await context.sync();

    const existing = sheets.items.find((s) => s.name === sheetName);
    sheet = existing || sheets.add(sheetName);

    const headers =
      kind === "accounts"
        ? ["Account ID", "Name", "Type"]
        : ["Vendor ID", "Name"];

    const rows = data.map((item) =>
      kind === "accounts"
        ? [item.id || item.AccountId, item.name || item.Name, item.type || item.AccountType]
        : [item.id || item.VendorId, item.name || item.DisplayName]
    );

    const all = [headers, ...rows];

    const range = sheet.getRangeByIndexes(0, 0, all.length, headers.length);
    range.values = all;
    range.format.autofitColumns();
    range.format.autofitRows();

    sheet.activate();
    await context.sync();
  });
}

/* 3. Create / refresh JE template */

async function handleCreateTemplate() {
  try {
    setStatus("Creating / refreshing JE template…", "info");
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const sheetName = "JE_Template";
      let sheet = sheets.items.find((s) => s.name === sheetName);
      if (!sheet) {
        sheet = sheets.add(sheetName);
      }

      const headers = [
        "Line #",
        "Account",
        "Vendor",
        "Description",
        "Debit",
        "Credit",
        "Date",
      ];
      const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
      headerRange.values = [headers];
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = "#e5e7eb";

      const dataRange = sheet.getRange("A2:G200");
      dataRange.clear();

      // Named ranges for later
      sheet.getRange("A1:G200").name = "JE_TABLE";

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
    resultEl.textContent = "Validating entry…";
    resultEl.className = "validation-result";

    const lines = await readJELinesFromSheet();
    if (lines.length === 0) {
      resultEl.textContent = "No lines found in JE_Template.";
      resultEl.classList.add("validation-result--error");
      return;
    }

    // Simple client-side validation: balanced debits/credits
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

    resultEl.textContent = `Entry is balanced. Debits = Credits = ${totalDebit.toFixed(
      2
    )}.`;
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
    range.load("values, rowCount, columnCount");
    await context.sync();

    const values = range.values;
    if (values.length <= 1) return [];

    const lines = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const [lineNo, account, vendor, description, debit, credit, date] = row;
      if (!account && !vendor && !description && !debit && !credit) continue;

      lines.push({
        lineNo,
        account,
        vendor,
        description,
        debit,
        credit,
        date,
      });
    }
    return lines;
  });
}

/* 5. Submit JE */

async function handleSubmitJE() {
  try {
    setStatus("Reading JE and submitting to backend…", "info");
    const lines = await readJELinesFromSheet();
    if (lines.length === 0) {
      setStatus("No JE lines to submit.", "error");
      return;
    }

    const res = await fetch(`${API_BASE}/journal/post`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      credentials: "include",
      body: JSON.stringify({ lines }),
    });

    if (!res.ok) {
      const text = await res.text();
      throw new Error(text || "Failed to post journal entry");
    }

    const payload = await res.json();
    setStatus("Journal entry submitted to QuickBooks.", "success");

    const resultEl = document.getElementById("validationResult");
    resultEl.textContent = `Submitted successfully. Ref: ${
      payload.reference || payload.id || "N/A"
    }`;
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
    setStatus("Loading submission history…", "info");
    const res = await fetch(`${API_BASE}/journal/history`, {
      credentials: "include",
    });
    if (!res.ok) throw new Error("Failed to load history");

    const items = await res.json(); // expect array
    const container = document.getElementById("historyList");
    container.innerHTML = "";

    if (!items.length) {
      container.textContent = "No submissions found for this workbook.";
      setStatus("History loaded (empty).", "info");
      return;
    }

    items.forEach((item) => {
      const div = document.createElement("div");
      div.className =
        "history-item " +
        (item.status === "success"
          ? "history-item--success"
          : "history-item--error");
      div.innerHTML = `
        <span>${item.timestamp || ""}</span>
        <span>${item.status || ""}</span>
      `;
      container.appendChild(div);
    });

    setStatus("History loaded.", "success");
  } catch (err) {
    console.error(err);
    setStatus("Error loading history.", "error");
  }
}

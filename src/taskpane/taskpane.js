/* global Office, Excel, document, fetch, console, window, setInterval, clearInterval */

const API_BASE = "https://fin-accruals.vercel.app/api";
let isQboConnected = false;
let validationPassed = false;
const demoSubmissionHistory = [];
const JE_TEMPLATE_HEADER_ROW_INDEX = 7;
const JE_TEMPLATE_DATA_START_ROW_INDEX = 8;
const JE_TEMPLATE_MAX_LINES = 200;
const JE_TEMPLATE_COLUMN_COUNT = 8;
const TRANSACTION_DATASETS = new Set([
  "invoices",
  "bills",
  "payments",
  "expenses",
  "deposits",
  "purchaseOrders",
  "journalEntries",
]);

Office.onReady(() => {
  initializeTabs();
  document.getElementById("btnConnectQBO").onclick = handleConnectQBO;
  document.getElementById("btnPullAccounts").onclick = () => handlePullMasterData("accounts");
  document.getElementById("btnPullVendors").onclick = () => handlePullMasterData("vendors");
  document.getElementById("btnPullCustomers").onclick = () => handlePullMasterData("customers");

  const btnPullClasses = document.getElementById("btnPullClasses");
  if (btnPullClasses) {
    btnPullClasses.onclick = () => handlePullMasterData("classes");
  }

  document.getElementById("moreDataSelect").onchange = updateMoreDataControls;
  document.getElementById("moreDataDateRange").onchange = updateMoreDataButton;
  document.getElementById("btnImportMoreData").onclick = handleImportMoreData;
  document.getElementById("btnCreateTemplate").onclick = handleCreateTemplate;
  document.getElementById("btnValidateJE").onclick = handleValidateJE;
  document.getElementById("btnSubmitJE").onclick = handleSubmitJE;
  document.getElementById("btnLoadHistory").onclick = handleLoadHistory;
  refreshQboStatus();
});

function initializeTabs() {
  const tabs = document.querySelectorAll(".workflow-tab");
  const panels = document.querySelectorAll(".panel");

  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      tabs.forEach((item) => item.classList.remove("active"));
      panels.forEach((panel) => panel.classList.remove("active"));

      tab.classList.add("active");
      document.getElementById(`panel-${tab.dataset.panel}`).classList.add("active");
    });
  });
}

function setStatus(message, type = "info") {
  const bar = document.getElementById("statusBar");
  const messageEl = bar.querySelector("span:last-child");
  messageEl.textContent = message;
  bar.classList.remove("status-bar--info", "status-bar--success", "status-bar--error");
  bar.classList.add(`status-bar--${type}`);
}

function setAuthBadge(text, variant = "primary", companyName = "") {
  const badge = document.getElementById("authStatusBadge");
  const statusText = document.getElementById("authStatusText");
  badge.textContent = text;
  badge.className = `status-badge badge--${variant}`;
  statusText.textContent =
    variant === "primary" ? companyName || "QuickBooks company" : "Not selected";
}

function setConnectionDependentActions(enabled) {
  ["btnPullAccounts", "btnPullVendors", "btnPullCustomers", "btnPullClasses"].forEach((id) => {
    document.getElementById(id).disabled = !enabled;
  });

  ["masterDataSurface", "moreDataSurface"].forEach((id) => {
    document.getElementById(id)?.classList.toggle("surface--locked", !enabled);
  });

  document.getElementById("dataLockNotice")?.classList.toggle("is-hidden", enabled);

  document.getElementById("moreDataSelect").disabled = !enabled;
  document.getElementById("moreDataDateRange").disabled = !enabled;
  document.getElementById("moreDataResult").textContent = enabled
    ? "Choose any supported table to create or refresh its worksheet."
    : "Connect QuickBooks to browse additional tables.";
  updateMoreDataControls();
}

function updateMoreDataButton() {
  const select = document.getElementById("moreDataSelect");
  document.getElementById("btnImportMoreData").disabled =
    !isQboConnected || select.disabled || !select.value;
}

function updateMoreDataControls() {
  const select = document.getElementById("moreDataSelect");
  const range = document.getElementById("moreDataDateRange");
  const isTransactionDataset = TRANSACTION_DATASETS.has(select.value);
  range.classList.toggle("is-hidden", !isTransactionDataset);
  range.disabled = !isQboConnected || !isTransactionDataset;
  updateMoreDataButton();
}

function updateSubmitAvailability() {
  document.getElementById("btnSubmitJE").disabled = !(isQboConnected && validationPassed);
}

/* 1. Connect to QBO */

function applyConnectionState(connected, companyName = "") {
  const button = document.getElementById("btnConnectQBO");
  const action = button.querySelector(".connector-cta");
  isQboConnected = connected;

  if (connected) {
    button.classList.add("is-connected");
    button.setAttribute("aria-pressed", "true");
    action.textContent = "Disconnect";
    setConnectionDependentActions(true);
    updateSubmitAvailability();
    setAuthBadge("Connected", "primary", companyName);
    return;
  }

  button.classList.remove("is-connected");
  button.setAttribute("aria-pressed", "false");
  action.textContent = "Connect";
  validationPassed = false;
  setConnectionDependentActions(false);
  updateSubmitAvailability();
  setAuthBadge("Not connected", "neutral");
}

async function apiFetch(path, options = {}) {
  return fetch(`${API_BASE}${path}`, {
    ...options,
    credentials: "include",
    headers: {
      Accept: "application/json",
      ...(options.headers || {}),
    },
  });
}

async function refreshQboStatus(showMessage = false) {
  try {
    const response = await apiFetch("/qbo/status");
    const payload = await response.json();
    applyConnectionState(Boolean(payload.connected), payload.companyName || "");

    if (showMessage) {
      setStatus(
        payload.connected
          ? `Connected to ${payload.companyName || "QuickBooks Online"}.`
          : "QuickBooks is not connected.",
        payload.connected ? "success" : "info"
      );
    }

    return Boolean(payload.connected);
  } catch (error) {
    console.error(error);
    applyConnectionState(false);
    if (showMessage) setStatus("Unable to check QuickBooks connection.", "error");
    return false;
  }
}

function openOAuthDialog(url) {
  return new Promise((resolve, reject) => {
    if (!Office.context?.ui?.displayDialogAsync) {
      const popup = window.open(url, "finaccruals-qbo", "width=620,height=720");
      if (!popup) {
        reject(new Error("The authentication window was blocked."));
        return;
      }

      const timer = setInterval(async () => {
        if (await refreshQboStatus()) {
          clearInterval(timer);
          popup.close();
          resolve();
        } else if (popup.closed) {
          clearInterval(timer);
          reject(new Error("QuickBooks connection was not completed."));
        }
      }, 1500);
      return;
    }

    Office.context.ui.displayDialogAsync(
      url,
      { height: 70, width: 45, displayInIframe: false },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(result.error?.message || "Unable to open QuickBooks authentication."));
          return;
        }

        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (event) => {
          dialog.close();
          try {
            const message = JSON.parse(event.message);
            message.success
              ? resolve()
              : reject(new Error(message.message || "Connection failed."));
          } catch {
            resolve();
          }
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (event) => {
          if (event.error) reject(new Error("QuickBooks authentication window was closed."));
        });
      }
    );
  });
}

async function handleConnectQBO() {
  const button = document.getElementById("btnConnectQBO");
  button.disabled = true;

  try {
    if (isQboConnected) {
      setStatus("Disconnecting QuickBooks...", "info");
      const response = await apiFetch("/qbo/disconnect", { method: "POST" });
      if (!response.ok) throw new Error("Unable to disconnect QuickBooks.");
      applyConnectionState(false);
      setStatus("QuickBooks disconnected.", "info");
      return;
    }

    setStatus("Opening secure QuickBooks authorization...", "info");
    await openOAuthDialog(`${API_BASE}/qbo/start`);
    const connected = await refreshQboStatus();
    if (!connected) throw new Error("QuickBooks did not return a valid connection.");
    setStatus("QuickBooks connected successfully.", "success");
  } catch (error) {
    console.error(error);
    setStatus(error.message || "QuickBooks connection failed.", "error");
  } finally {
    button.disabled = false;
  }
}

/* 2. Pull master data */

const MASTER_DATA_LABELS = {
  accounts: "Accounts",
  vendors: "Vendors",
  customers: "Customers",
  classes: "Classes",
};

function getMasterDataFromPayload(kind, payload) {
  if (kind === "accounts") return payload.accounts || [];
  if (kind === "vendors") return payload.vendors || [];
  if (kind === "customers") return payload.customers || [];
  if (kind === "classes") return payload.classes || [];
  return [];
}

async function fetchMasterData(kind) {
  const label = MASTER_DATA_LABELS[kind] || kind;
  const res = await apiFetch(`/${kind}`);

  if (!res.ok) {
    const errorPayload = await res.json().catch(() => ({}));
    if (res.status === 401) {
      applyConnectionState(false);
      throw new Error("Connect QuickBooks first, then try syncing again.");
    }
    throw new Error(errorPayload.error || `Failed to pull ${label.toLowerCase()}`);
  }

  const payload = await res.json();
  return getMasterDataFromPayload(kind, payload);
}

async function handlePullMasterData(kind) {
  try {
    if (!isQboConnected) {
      setStatus("Connect an accounting platform before syncing master data.", "error");
      return;
    }

    const label = MASTER_DATA_LABELS[kind];
    setStatus(`Pulling ${label.toLowerCase()} from deployed API...`, "info");

    const data = await fetchMasterData(kind);
    await writeMasterDataToSheet(kind, data);

    setStatus(`${label} updated in workbook.`, "success");
  } catch (err) {
    console.error(err);
    setStatus(err.message || "Error pulling master data.", "error");
  }
}

async function writeMasterDataToSheet(kind, data, options = {}) {
  const { activate = true, hidden = false } = options;

  return Excel.run(async (context) => {
    const sheetNameMap = {
      accounts: "Accounts",
      vendors: "Vendors",
      customers: "Customers",
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

    if (kind === "customers") {
      headers = ["QuickBooks ID", "Name", "Company", "Email", "Phone", "Balance"];
      rows = data.map((item) => [
        item.id || "",
        item.name || "",
        item.company || "",
        item.email || "",
        item.phone || "",
        Number(item.balance) || 0,
      ]);
    }

    if (kind === "classes") {
      headers = ["Name"];
      rows = data.map((item) => [item.name || ""]);
    }

    const all = [headers, ...rows];

    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load("address");

    await context.sync();

    if (!usedRange.isNullObject) {
      usedRange.clear();
    }

    const range = sheet.getRangeByIndexes(0, 0, all.length, headers.length);
    range.values = all;
    range.format.font.name = "Aptos";
    range.format.autofitColumns();
    range.format.autofitRows();

    const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
    headerRange.format.font.bold = true;
    headerRange.format.fill.color = "#e5e7eb";

    sheet.visibility = hidden ? Excel.SheetVisibility.hidden : Excel.SheetVisibility.visible;

    if (activate) {
      sheet.activate();
    }

    await context.sync();
  });
}

async function handleImportMoreData() {
  const select = document.getElementById("moreDataSelect");
  const range = document.getElementById("moreDataDateRange");
  const button = document.getElementById("btnImportMoreData");
  const result = document.getElementById("moreDataResult");
  const dataset = select.value;
  const dateRange = range.value;

  if (!isQboConnected) {
    setStatus("Connect QuickBooks before importing additional data.", "error");
    result.textContent = "Connect QuickBooks to browse additional tables.";
    return;
  }

  if (!dataset) return;

  button.disabled = true;
  select.disabled = true;
  range.disabled = true;
  const dateRangeSuffix = TRANSACTION_DATASETS.has(dataset)
    ? ` (${range.options[range.selectedIndex].text})`
    : "";
  result.textContent = `Importing ${select.options[select.selectedIndex].text}${dateRangeSuffix}...`;
  setStatus("Importing additional QuickBooks data...", "info");

  try {
    const query = new URLSearchParams({ dataset });
    if (TRANSACTION_DATASETS.has(dataset)) query.set("dateRange", dateRange);

    const response = await apiFetch(`/more-data?${query.toString()}`);
    const payload = await response.json();

    if (!response.ok) {
      if (response.status === 401) {
        applyConnectionState(false);
        throw new Error("Connect QuickBooks first, then try importing again.");
      }
      throw new Error(payload.error || "Unable to import the selected QuickBooks table.");
    }

    await writeTableToSheet(payload.sheetName, payload.headers, payload.rows);
    result.textContent = `${payload.label}: ${payload.count} record(s) imported to ${
      payload.sheetName
    }${payload.dateRangeLabel ? ` for ${payload.dateRangeLabel}.` : "."}`;
    setStatus(`${payload.label} updated in workbook.`, "success");
  } catch (error) {
    console.error(error);
    result.textContent = error.message || "Import failed.";
    setStatus(error.message || "Additional data import failed.", "error");
  } finally {
    select.disabled = !isQboConnected;
    updateMoreDataControls();
  }
}

async function writeTableToSheet(sheetName, headers, rows) {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    let sheet = sheets.items.find((item) => item.name === sheetName);
    if (!sheet) sheet = sheets.add(sheetName);

    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load("address");
    await context.sync();
    if (!usedRange.isNullObject) usedRange.clear();

    const values = [headers, ...(rows || [])];
    const range = sheet.getRangeByIndexes(0, 0, values.length, headers.length);
    range.values = values;
    range.format.font.name = "Aptos";
    range.format.autofitColumns();
    range.format.autofitRows();

    const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
    headerRange.format.font.bold = true;
    headerRange.format.fill.color = "#e5e7eb";
    headerRange.format.autofitColumns();

    sheet.activate();
    await context.sync();
  });
}

/* 3. Create / refresh JE template */

async function handleCreateTemplate() {
  const missingDropdownSources = [];
  const button = document.getElementById("btnCreateTemplate");
  const actionLabel = button?.querySelector(".sync-action");
  const originalActionText = actionLabel?.textContent || "Generate";

  try {
    if (!isQboConnected) {
      setStatus("Connect QuickBooks before creating the journal template.", "error");
      return;
    }

    if (button) button.disabled = true;
    if (actionLabel) actionLabel.textContent = "Preparing...";

    setStatus("Preparing journal template and syncing QuickBooks dropdown data...", "info");

    const requiredMasterData = ["accounts", "vendors", "classes"];
    const syncResults = [];

    for (const kind of requiredMasterData) {
      const label = MASTER_DATA_LABELS[kind];
      if (actionLabel) actionLabel.textContent = `Syncing ${label}...`;
      setStatus(`Syncing ${label.toLowerCase()} for JE dropdowns...`, "info");

      const data = await fetchMasterData(kind);
      await writeMasterDataToSheet(kind, data, { activate: false, hidden: true });
      syncResults.push({ kind, count: data.length });
    }

    if (actionLabel) actionLabel.textContent = "Creating...";
    setStatus("Creating JE template and applying dropdowns...", "info");

    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");

      await context.sync();

      const sheetName = "JE_Template";
      let sheet = sheets.items.find((s) => s.name === sheetName);

      if (!sheet) {
        sheet = sheets.add(sheetName);
      }

      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("address");
      await context.sync();
      if (!usedRange.isNullObject) {
        usedRange.unmerge();
        usedRange.clear();
      }

      const headers = [
        "Line #",
        "Account *",
        "Vendor",
        "Class",
        "Description",
        "Debit *",
        "Credit *",
        "Date",
      ];

      const titleRange = sheet.getRange("A1:H1");
      titleRange.values = [["FinAccruals Journal Entry Template", "", "", "", "", "", "", ""]];
      titleRange.merge(false);
      titleRange.format.font.bold = true;
      titleRange.format.font.size = 18;
      titleRange.format.font.color = "#0f172a";
      titleRange.format.wrapText = false;

      const subtitleRange = sheet.getRange("A2:H2");
      subtitleRange.values = [
        [
          "Enter the journal date once, complete the line details, validate, then post to the connected QuickBooks sandbox.",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
        ],
      ];
      subtitleRange.merge(false);
      subtitleRange.format.font.color = "#64748b";
      subtitleRange.format.wrapText = false;

      const quickGuideRange = sheet.getRange("A3:H3");
      quickGuideRange.values = [
        [
          "1  Set journal date",
          "",
          "2  Complete line details",
          "",
          "",
          "3  Validate, then post",
          "",
          "",
        ],
      ];
      sheet.getRange("A3:B3").merge(false);
      sheet.getRange("C3:E3").merge(false);
      sheet.getRange("F3:H3").merge(false);
      quickGuideRange.format.fill.color = "#eef4ff";
      quickGuideRange.format.font.color = "#1d4ed8";
      quickGuideRange.format.font.bold = true;
      quickGuideRange.format.wrapText = false;

      const headerCardRange = sheet.getRange("A4:E6");
      headerCardRange.format.fill.color = "#f8fafc";
      headerCardRange.format.font.color = "#334155";
      headerCardRange.format.wrapText = true;
      sheet.getRange("A4").values = [["Journal date *"]];
      sheet.getRange("B4").values = [[new Date().toISOString().slice(0, 10)]];
      sheet.getRange("C4").values = [["Memo / purpose"]];
      sheet.getRange("D4:E4").values = [["", ""]];
      sheet.getRange("D4:E4").merge(false);
      sheet.getRange("A5").values = [["Prepared by"]];
      sheet.getRange("B5").values = [[""]];
      sheet.getRange("C5").values = [["Posting mode"]];
      sheet.getRange("D5:E5").values = [["QuickBooks sandbox", ""]];
      sheet.getRange("D5:E5").merge(false);
      sheet.getRange("A6:E6").values = [
        [
          "Tip: choose Account, Vendor, and Class from dropdowns. Enter either Debit or Credit on each row, not both.",
          "",
          "",
          "",
          "",
        ],
      ];
      sheet.getRange("A6:E6").merge(false);
      sheet.getRange("A4:A5").format.font.bold = true;
      sheet.getRange("C4:C5").format.font.bold = true;
      sheet.getRange("A4:A5").format.wrapText = false;
      sheet.getRange("C4:C5").format.wrapText = false;
      sheet.getRange("B4").numberFormat = [["yyyy-mm-dd"]];
      sheet.getRange("B4").format.fill.color = "#fff7ed";
      sheet.getRange("D4:E4").format.fill.color = "#ffffff";
      sheet.getRange("B5").format.fill.color = "#ffffff";
      sheet.getRange("D5:E5").format.fill.color = "#ecfdf5";
      sheet.getRange("A6:E6").format.font.color = "#64748b";

      const totalsRange = sheet.getRange("F4:H6");
      totalsRange.values = [["Total debit", "Total credit", "Difference"], ["", "", ""], ["Status", "", ""]];
      totalsRange.format.fill.color = "#f8fafc";
      totalsRange.format.wrapText = true;
      sheet.getRange("F4:H4").format.font.bold = true;
      sheet.getRange("F5").formulas = [[`=SUM(F${JE_TEMPLATE_DATA_START_ROW_INDEX + 1}:F${JE_TEMPLATE_DATA_START_ROW_INDEX + JE_TEMPLATE_MAX_LINES})`]];
      sheet.getRange("G5").formulas = [[`=SUM(G${JE_TEMPLATE_DATA_START_ROW_INDEX + 1}:G${JE_TEMPLATE_DATA_START_ROW_INDEX + JE_TEMPLATE_MAX_LINES})`]];
      sheet.getRange("H5").formulas = [["=F5-G5"]];
      sheet.getRange("F5:H5").numberFormat = [["$#,##0.00", "$#,##0.00", "$#,##0.00"]];
      sheet.getRange("G6:H6").merge(false);
      sheet.getRange("F6").format.font.bold = true;
      sheet.getRange("G6").formulas = [["=IF(ABS(H5)<0.01,\"Balanced\",\"Needs review\")"]];
      sheet.getRange("G6").format.font.bold = true;

      const sectionRange = sheet.getRange("A7:H7");
      sectionRange.values = [["Journal lines", "", "", "", "", "", "", ""]];
      sectionRange.merge(false);
      sectionRange.format.font.bold = true;
      sectionRange.format.font.color = "#1d4ed8";

      const headerRange = sheet.getRangeByIndexes(
        JE_TEMPLATE_HEADER_ROW_INDEX,
        0,
        1,
        headers.length
      );
      headerRange.values = [headers];
      headerRange.format.font.bold = true;
      headerRange.format.fill.color = "#dbeafe";
      headerRange.format.font.color = "#0f172a";

      const dataRange = sheet.getRangeByIndexes(
        JE_TEMPLATE_DATA_START_ROW_INDEX,
        0,
        JE_TEMPLATE_MAX_LINES,
        JE_TEMPLATE_COLUMN_COUNT
      );
      dataRange.format.font.name = "Aptos";
      dataRange.format.fill.color = "#ffffff";

      const lineNumbers = Array.from({ length: JE_TEMPLATE_MAX_LINES }, (_, index) => [index + 1]);
      const lineNumberRange = sheet.getRangeByIndexes(
        JE_TEMPLATE_DATA_START_ROW_INDEX,
        0,
        JE_TEMPLATE_MAX_LINES,
        1
      );
      lineNumberRange.values = lineNumbers;
      lineNumberRange.format.font.color = "#64748b";
      lineNumberRange.format.fill.color = "#f8fafc";

      const sheetNames = new Set(sheets.items.map((s) => s.name));
      const dropdownColumns = sheet.getRangeByIndexes(
        JE_TEMPLATE_DATA_START_ROW_INDEX,
        1,
        JE_TEMPLATE_MAX_LINES,
        3
      );
      dropdownColumns.format.fill.color = "#eff6ff";

      if (sheetNames.has("Accounts")) {
        const accountRange = sheet.getRangeByIndexes(
          JE_TEMPLATE_DATA_START_ROW_INDEX,
          1,
          JE_TEMPLATE_MAX_LINES,
          1
        );
        accountRange.dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: "='Accounts'!$B$2:$B$1000",
          },
        };
      } else {
        missingDropdownSources.push("Accounts");
      }

      if (sheetNames.has("Vendors")) {
        const vendorRange = sheet.getRangeByIndexes(
          JE_TEMPLATE_DATA_START_ROW_INDEX,
          2,
          JE_TEMPLATE_MAX_LINES,
          1
        );
        vendorRange.dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: "='Vendors'!$A$2:$A$1000",
          },
        };
      } else {
        missingDropdownSources.push("Vendors");
      }

      if (sheetNames.has("Classes")) {
        const classRange = sheet.getRangeByIndexes(
          JE_TEMPLATE_DATA_START_ROW_INDEX,
          3,
          JE_TEMPLATE_MAX_LINES,
          1
        );
        classRange.dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: "='Classes'!$A$2:$A$1000",
          },
        };
      } else {
        missingDropdownSources.push("Classes");
      }

      const amountRange = sheet.getRangeByIndexes(
        JE_TEMPLATE_DATA_START_ROW_INDEX,
        5,
        JE_TEMPLATE_MAX_LINES,
        2
      );
      amountRange.numberFormat = Array.from({ length: JE_TEMPLATE_MAX_LINES }, () => [
        "$#,##0.00",
        "$#,##0.00",
      ]);

      const dateRange = sheet.getRangeByIndexes(
        JE_TEMPLATE_DATA_START_ROW_INDEX,
        7,
        JE_TEMPLATE_MAX_LINES,
        1
      );
      dateRange.formulas = Array.from({ length: JE_TEMPLATE_MAX_LINES }, () => [
        '=IF($B$4="","",$B$4)',
      ]);
      dateRange.numberFormat = Array.from({ length: JE_TEMPLATE_MAX_LINES }, () => ["yyyy-mm-dd"]);

      const requiredInputRange = sheet.getRangeByIndexes(
        JE_TEMPLATE_DATA_START_ROW_INDEX,
        1,
        JE_TEMPLATE_MAX_LINES,
        1
      );
      requiredInputRange.format.fill.color = "#eff6ff";
      amountRange.format.fill.color = "#fff7ed";
      dateRange.format.fill.color = "#f8fafc";

      const templateRange = sheet.getRangeByIndexes(0, 0, JE_TEMPLATE_DATA_START_ROW_INDEX + JE_TEMPLATE_MAX_LINES, JE_TEMPLATE_COLUMN_COUNT);
      const existingName = context.workbook.names.getItemOrNullObject("JE_TABLE");
      existingName.load("name");

      await context.sync();

      if (!existingName.isNullObject) {
        existingName.delete();
        await context.sync();
      }

      context.workbook.names.add("JE_TABLE", templateRange);

      sheet.getRange("A1:H211").format.font.name = "Aptos";
      sheet.getRange("A1:H211").format.autofitRows();
      sheet.getRange("A:A").format.columnWidth = 74;
      sheet.getRange("B:B").format.columnWidth = 185;
      sheet.getRange("C:C").format.columnWidth = 132;
      sheet.getRange("D:D").format.columnWidth = 125;
      sheet.getRange("E:E").format.columnWidth = 225;
      sheet.getRange("F:F").format.columnWidth = 82;
      sheet.getRange("G:G").format.columnWidth = 82;
      sheet.getRange("H:H").format.columnWidth = 72;
      sheet.getRange("A4:H6").format.rowHeight = 24;
      sheet.getRange("A:A").format.horizontalAlignment = Excel.HorizontalAlignment.center;
      sheet.getRange("B:D").format.horizontalAlignment = Excel.HorizontalAlignment.left;
      sheet.getRange("E:E").format.horizontalAlignment = Excel.HorizontalAlignment.left;
      sheet.getRange("E:E").format.wrapText = true;
      sheet.getRange("F:G").format.horizontalAlignment = Excel.HorizontalAlignment.right;
      sheet.getRange("H:H").format.horizontalAlignment = Excel.HorizontalAlignment.center;
      sheet.freezePanes.freezeRows(JE_TEMPLATE_HEADER_ROW_INDEX + 1);

      sheet.activate();

      await context.sync();
    });

    validationPassed = false;
    updateSubmitAvailability();

    if (missingDropdownSources.length) {
      setStatus(
        `JE template ready. Sync ${missingDropdownSources.join(", ")} to enable all dropdowns.`,
        "info"
      );
    } else {
      const syncSummary = syncResults
        .map((result) => `${MASTER_DATA_LABELS[result.kind]}: ${result.count}`)
        .join(", ");
      setStatus(`JE template ready with QuickBooks dropdowns (${syncSummary}).`, "success");
    }
  } catch (err) {
    console.error(err);
    setStatus(err.message || "Error creating JE template.", "error");
  } finally {
    if (button) button.disabled = !isQboConnected;
    if (actionLabel) actionLabel.textContent = originalActionText;
  }
}

/* 4. Validate JE */

async function handleValidateJE() {
  try {
    const resultEl = document.getElementById("validationResult");
    resultEl.textContent = "Validating entry...";
    resultEl.className = "validation-result";
    validationPassed = false;
    updateSubmitAvailability();

    const lines = await readJELinesFromSheet();
    const validation = validateJELines(lines);

    if (!validation.valid) {
      resultEl.textContent = validation.errors.join(" ");
      resultEl.classList.add("validation-result--error");
      setStatus("Validation failed. Review the journal entry.", "error");
      return;
    }

    validationPassed = true;
    updateSubmitAvailability();
    resultEl.textContent = `All controls passed. ${lines.length} line(s), total debits and credits ${formatCurrency(
      validation.totalDebit
    )}.`;
    resultEl.classList.add("validation-result--ok");
    setStatus("Validation passed. Ready to submit.", "success");
  } catch (err) {
    validationPassed = false;
    updateSubmitAvailability();
    console.error(err);
    const resultEl = document.getElementById("validationResult");
    resultEl.textContent =
      err.code === "ItemNotFound"
        ? "Create the JE_Template sheet before running validation."
        : `Validation failed: ${err.message}`;
    resultEl.className = "validation-result validation-result--error";
    setStatus("Error validating JE.", "error");
  }
}

function parseAmount(value) {
  if (value === "" || value === null || value === undefined) {
    return 0;
  }

  const normalized = typeof value === "string" ? value.replace(/[$,\s]/g, "") : value;
  const amount = Number(normalized);

  return Number.isFinite(amount) ? amount : null;
}

function validateJELines(lines) {
  const errors = [];
  let totalDebit = 0;
  let totalCredit = 0;

  if (lines.length === 0) {
    return {
      valid: false,
      errors: ["No journal lines were found in JE_Template."],
      totalDebit,
      totalCredit,
    };
  }

  lines.forEach((line, index) => {
    const rowNumber = line.rowNumber || index + JE_TEMPLATE_DATA_START_ROW_INDEX + 1;
    const debit = parseAmount(line.debit);
    const credit = parseAmount(line.credit);

    if (!line.account) {
      errors.push(`Row ${rowNumber}: account is required.`);
    }

    if (!line.date) {
      errors.push(`Row ${rowNumber}: date is required.`);
    }

    if (debit === null || credit === null) {
      errors.push(`Row ${rowNumber}: debit and credit must be numeric.`);
      return;
    }

    if (debit < 0 || credit < 0) {
      errors.push(`Row ${rowNumber}: negative amounts are not allowed.`);
    }

    if ((debit > 0 && credit > 0) || (debit === 0 && credit === 0)) {
      errors.push(`Row ${rowNumber}: enter an amount on exactly one side.`);
    }

    totalDebit += debit;
    totalCredit += credit;
  });

  if (totalDebit <= 0 && totalCredit <= 0) {
    errors.push("Journal total must be greater than zero.");
  }

  if (Math.abs(totalDebit - totalCredit) > 0.01) {
    errors.push(
      `Entry is unbalanced: debits ${formatCurrency(totalDebit)}, credits ${formatCurrency(
        totalCredit
      )}.`
    );
  }

  return {
    valid: errors.length === 0,
    errors,
    totalDebit,
    totalCredit,
  };
}

function formatCurrency(value) {
  return `$${value.toFixed(2)}`;
}

async function readJELinesFromSheet() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("JE_Template");
    const range = sheet.getRangeByIndexes(
      JE_TEMPLATE_DATA_START_ROW_INDEX,
      0,
      JE_TEMPLATE_MAX_LINES,
      JE_TEMPLATE_COLUMN_COUNT
    );
    range.load("values");

    await context.sync();

    const values = range.values;

    const lines = [];

    for (let i = 0; i < values.length; i++) {
      const row = values[i];

      const [lineNo, account, vendor, className, description, debit, credit, date] = row;

      const hasLineInput = Boolean(account || vendor || className || description || debit || credit);

      if (!hasLineInput) {
        continue;
      }

      lines.push({
        lineNo,
        rowNumber: JE_TEMPLATE_DATA_START_ROW_INDEX + i + 1,
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

/* 5. Submit JE to QuickBooks sandbox */

async function handleSubmitJE() {
  try {
    if (!isQboConnected) {
      setStatus("Connect QuickBooks before submitting a journal entry.", "error");
      return;
    }

    if (!validationPassed) {
      setStatus("Run validation successfully before submitting.", "error");
      return;
    }

    setStatus("Rechecking JE before sandbox posting...", "info");

    const lines = await readJELinesFromSheet();
    const validation = validateJELines(lines);

    if (!validation.valid) {
      validationPassed = false;
      updateSubmitAvailability();
      const resultEl = document.getElementById("validationResult");
      resultEl.textContent = validation.errors.join(" ");
      resultEl.className = "validation-result validation-result--error";
      setStatus("Journal changed after validation. Run validation again.", "error");
      return;
    }

    setStatus("Posting journal entry to QuickBooks sandbox...", "info");

    const response = await apiFetch("/journal-entries", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ lines }),
    });
    const payload = await response.json().catch(() => ({}));

    if (!response.ok) {
      if (response.status === 401) {
        applyConnectionState(false);
        throw new Error("Connect QuickBooks first, then try posting again.");
      }
      const detail = payload.error || payload.message || "QuickBooks sandbox posting failed.";
      throw new Error(detail);
    }

    const submission = {
      id: payload.id ? `QBO-JE-${payload.id}` : `QBO-JE-${Date.now()}`,
      date: payload.date || new Date().toISOString().slice(0, 10),
      status: "Posted to Sandbox",
      amount: Number(payload.amount) || validation.totalDebit,
    };

    demoSubmissionHistory.unshift(submission);
    renderHistory(demoSubmissionHistory);

    setStatus("Journal entry posted to QuickBooks sandbox.", "success");

    const resultEl = document.getElementById("validationResult");
    resultEl.textContent = `Posted successfully. QuickBooks Journal Entry ID: ${payload.id || "created"}`;
    resultEl.className = "validation-result validation-result--ok";
    validationPassed = false;
    updateSubmitAvailability();
  } catch (err) {
    console.error(err);
    const message = err.message || "QuickBooks sandbox posting failed.";
    const friendlyMessage = message.includes("was not found in QuickBooks")
      ? `${message} Sync QuickBooks master data and use the exact Account, Vendor, or Class name from the workbook.`
      : message;

    setStatus(friendlyMessage, "error");

    const resultEl = document.getElementById("validationResult");
    resultEl.textContent = `Submission failed: ${friendlyMessage}`;
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
    const items = [...demoSubmissionHistory, ...(payload.history || [])];

    if (!items.length) {
      const container = document.getElementById("historyList");
      container.replaceChildren();
      container.textContent = "No submissions found.";
      setStatus("History loaded empty.", "info");
      return;
    }

    renderHistory(items);
    setStatus("History loaded.", "success");
  } catch (err) {
    console.error(err);
    setStatus("Error loading history.", "error");
  }
}

function renderHistory(items) {
  const container = document.getElementById("historyList");
  container.replaceChildren();

  items.forEach((item) => {
    const div = document.createElement("div");
    div.className = "history-item history-item--success";
    [
      item.id || "",
      item.date || "",
      item.status || "",
      formatCurrency(Number(item.amount) || 0),
    ].forEach((value) => {
      const span = document.createElement("span");
      span.textContent = value;
      div.appendChild(span);
    });
    container.appendChild(div);
  });
}

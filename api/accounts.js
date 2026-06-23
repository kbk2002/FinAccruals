module.exports = function handler(req, res) {
  res.status(200).json({
    success: true,
    accounts: [
      { code: "1000", name: "Checking", type: "Asset" },
      { code: "1200", name: "Accounts Receivable", type: "Asset" },
      { code: "1400", name: "Prepaid Expenses", type: "Asset" },
      { code: "1500", name: "Fixed Assets", type: "Asset" },
      { code: "2000", name: "Accounts Payable", type: "Liability" },
      { code: "2100", name: "Accrued Liabilities", type: "Liability" },
      { code: "2400", name: "Payroll Liabilities", type: "Liability" },
      { code: "3000", name: "Owner's Equity", type: "Equity" },
      { code: "4000", name: "Sales Income", type: "Income" },
      { code: "4100", name: "Service Revenue", type: "Income" },
      { code: "5000", name: "Cost of Goods Sold", type: "Expense" },
      { code: "6000", name: "Rent Expense", type: "Expense" },
      { code: "6100", name: "Utilities Expense", type: "Expense" },
      { code: "6200", name: "Payroll Expense", type: "Expense" },
      { code: "6300", name: "Office Supplies", type: "Expense" },
      { code: "6400", name: "Depreciation Expense", type: "Expense" }
    ]
  });
};

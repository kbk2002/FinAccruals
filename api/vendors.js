module.exports = function handler(req, res) {
  res.status(200).json({
    success: true,
    vendors: [
      { name: "Acme Property Mgmt", email: "billing@acmeproperty.com" },
      { name: "Pacific Power Co.", email: "ar@pacificpower.com" },
      { name: "Staples Business", email: "orders@staples.com" },
      { name: "Northwind Payroll", email: "support@northwindpay.com" }
    ]
  });
};

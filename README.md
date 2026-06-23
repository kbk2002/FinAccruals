# FinAccruals / LedgerFlow Excel Add-in

## Overview

FinAccruals, presented in the Excel interface as **LedgerFlow for Excel**, is a Microsoft Excel Add-in built with Office.js to streamline accounting preparation workflows directly inside Excel.

The application connects an Excel workbook to QuickBooks Online through OAuth 2.0, imports accounting data from a connected QuickBooks sandbox company, creates dedicated worksheets, generates a standardized journal-entry template, validates journal entries, and provides a foundation for future posting and approval workflows.

The current repository and Intuit Developer application use the **FinAccruals** name. The task-pane interface uses the **LedgerFlow** product identity.

---

## Current Status

The project has completed the QuickBooks OAuth and accounting-data import phase of the MVP.

Users can currently:

- Connect a QuickBooks Online sandbox company.
- Disconnect the active QuickBooks connection.
- View the connected QuickBooks company name.
- Pull Accounts into Excel.
- Pull Vendors into Excel.
- Pull Customers into Excel.
- Pull Classes into Excel.
- Select and import additional QuickBooks tables.
- Create or refresh a journal-entry template.
- Validate journal-entry structure and balance.
- Submit demonstration journal entries locally.
- View demonstration submission history.
- Use the deployed application in Excel for the web.
- Use Vercel-hosted backend APIs.

All QuickBooks data currently comes from a QuickBooks **sandbox company**. Sandbox records are fictional test data supplied by Intuit or manually created during development.

---

## What Has Been Completed

### Excel Add-in Foundation

- Created the Office.js Excel Add-in project structure.
- Built the LedgerFlow task-pane interface.
- Added Excel ribbon integration through `manifest.xml`.
- Configured add-in branding, logos, and icons.
- Configured the Webpack build pipeline.
- Added production and development URL handling.
- Connected the GitHub repository to Vercel.
- Validated the Office Add-in manifest.
- Tested the task pane in Excel for the web.

### QuickBooks OAuth Integration

- Created an Intuit Developer workspace and QuickBooks application.
- Added QuickBooks OAuth 2.0 authorization.
- Added OAuth state validation.
- Added authorization-code exchange.
- Added secure access-token and refresh-token handling.
- Added automatic access-token refresh.
- Added QuickBooks connection-status checking.
- Added QuickBooks disconnect and token revocation.
- Added support for the QuickBooks sandbox API environment.
- Added secure encrypted HTTP-only session cookies.

### Excel Workbook Integration

- Added automated `Accounts` worksheet creation.
- Added automated `Vendors` worksheet creation.
- Added automated `Customers` worksheet creation.
- Added automated `Classes` worksheet creation.
- Added dynamic worksheet creation for additional QuickBooks tables.
- Added safe worksheet refresh behavior.
- Added automatic headers and worksheet formatting.
- Added column and row auto-sizing.
- Added workbook navigation to the newly imported worksheet.
- Added `JE_Template` worksheet creation.
- Added the `JE_TABLE` workbook name.

### Primary QuickBooks Data Synchronization

The task pane contains direct synchronization buttons for the most frequently used accounting reference data:

- Accounts
- Vendors
- Customers
- Classes

Each synchronization refreshes a dedicated worksheet using records returned by the connected QuickBooks company.

### More QuickBooks Data

The task pane includes a **More QuickBooks Data** selector for importing other supported QuickBooks API entities.

#### Reference Data

- Products & Services
- Employees
- Locations
- Payment Terms
- Payment Methods
- Tax Codes

#### Transaction Data

- Invoices
- Bills
- Customer Payments
- Expenses
- Deposits
- Purchase Orders
- Journal Entries

When a user selects a table and clicks **Import table**, the add-in:

1. Calls the deployed Vercel API.
2. Queries the connected QuickBooks sandbox company.
3. Maps the QuickBooks response into an Excel-friendly structure.
4. Creates or refreshes the corresponding worksheet.
5. Shows the number of imported records.

Transaction imports currently contain summary or header-level information. Detailed transaction-line worksheets are planned for a future phase.

### Journal Entry Workflow

- Added journal-entry template creation.
- Added required account validation.
- Added required date validation.
- Added numeric debit and credit validation.
- Prevented negative journal amounts.
- Required an amount on exactly one side of each line.
- Added debit and credit balance validation.
- Added validation result messaging.
- Added submit-button state management.
- Added demonstration journal submission references.
- Added demonstration submission history.

Journal entries are **not yet posted to QuickBooks**. The current submit workflow remains a controlled demonstration workflow.

### Backend API Integration

- Deployed backend API routes through Vercel Serverless Functions.
- Added QuickBooks OAuth routes.
- Added QuickBooks data-query helpers.
- Added secure session helpers.
- Added Accounts API integration.
- Added Vendors API integration.
- Added Customers API integration.
- Added Classes API integration.
- Added the More Data API integration.
- Added the API health endpoint.
- Added the demonstration history endpoint.
- Kept the deployment within the Vercel Hobby-plan serverless-function limit.

---

## Current Working Features

- ✅ QuickBooks Online OAuth 2.0 connection
- ✅ QuickBooks sandbox authorization
- ✅ Secure Connect and Disconnect
- ✅ Connected-company display
- ✅ Pull Accounts
- ✅ Pull Vendors
- ✅ Pull Customers
- ✅ Pull Classes
- ✅ Import Products & Services
- ✅ Import Employees
- ✅ Import Locations
- ✅ Import Payment Terms
- ✅ Import Payment Methods
- ✅ Import Tax Codes
- ✅ Import Invoices
- ✅ Import Bills
- ✅ Import Customer Payments
- ✅ Import Expenses
- ✅ Import Deposits
- ✅ Import Purchase Orders
- ✅ Import Journal Entries
- ✅ Create Journal Entry Template
- ✅ Validate Journal Entry Balance
- ✅ Submit Demo Journal Entry
- ✅ Load Demo Submission History
- ✅ Excel ↔ Vercel API Integration
- ✅ QuickBooks API ↔ Excel Integration
- ✅ Vercel Deployment
- ✅ Excel for the web testing

---

## How QuickBooks Data Works

The labels shown in the **More QuickBooks Data** selector are configured by FinAccruals based on supported QuickBooks Accounting API entities.

The actual imported rows are retrieved from the connected QuickBooks company.

```text
Dropdown label
    |
    v
FinAccruals dataset configuration
    |
    v
QuickBooks Accounting API query
    |
    v
Connected sandbox-company records
    |
    v
Excel worksheet
```

QuickBooks does not provide one API endpoint that dynamically lists every available accounting table. FinAccruals therefore maintains a controlled catalog of supported datasets.

Some QuickBooks API records may not appear in the current QuickBooks web interface. For example, a sandbox may contain legacy `Employee` API records while the newer **Team → Employees** screen appears empty. In that case, FinAccruals is still displaying genuine records returned by the QuickBooks Accounting API.

---

## Architecture

```text
Microsoft Excel
    |
    | Office.js
    v
LedgerFlow Task Pane
    |
    | HTTPS requests with secure session cookie
    v
Vercel Serverless API
    |
    | OAuth 2.0 and Accounting API requests
    v
QuickBooks Online Sandbox
```

### Frontend Responsibilities

- Render the Excel task pane.
- Manage workflow navigation.
- Open the QuickBooks authorization dialog.
- Display connection state.
- Request QuickBooks datasets.
- Create and refresh Excel worksheets.
- Create the journal-entry template.
- Validate journal-entry data.
- Display sync and validation results.

### Backend Responsibilities

- Generate QuickBooks authorization URLs.
- Validate OAuth state.
- Exchange authorization codes for tokens.
- Refresh expired access tokens.
- Encrypt session information.
- Revoke QuickBooks tokens.
- Query supported QuickBooks entities.
- Map QuickBooks responses into workbook-friendly data.
- Prevent secrets from being exposed to frontend code.

---

## Technology Stack

### Frontend

- JavaScript
- HTML
- CSS
- Microsoft Office.js

### Build and Development

- Node.js
- npm
- Webpack
- Babel
- Office Add-in debugging tools
- Office Add-in manifest validation
- ESLint and Office Add-in linting

### Backend and Deployment

- Vercel Serverless Functions
- QuickBooks Online Accounting API
- Intuit OAuth 2.0
- GitHub

### Planned

- Database persistence
- Multi-user authentication
- Audit logging
- QuickBooks journal posting
- Xero integration
- Production QuickBooks approval and compliance

---

## Project Structure

```text
FinAccruals/
|
|-- api/
|   |-- accounts.js
|   |-- classes.js
|   |-- customers.js
|   |-- health.js
|   |-- history.js
|   |-- more-data.js
|   |-- vendors.js
|   `-- qbo/
|       |-- callback.js
|       |-- disconnect.js
|       |-- start.js
|       `-- status.js
|
|-- server/
|   |-- qbo.js
|   `-- session.js
|
|-- src/
|   |-- taskpane/
|   |   |-- taskpane.html
|   |   |-- taskpane.css
|   |   `-- taskpane.js
|   |
|   `-- commands/
|       |-- commands.html
|       `-- commands.js
|
|-- assets/
|   |-- icon-16.png
|   |-- icon-32.png
|   |-- icon-64.png
|   |-- icon-80.png
|   |-- icon-128.png
|   |-- logo-filled.png
|   `-- additional brand assets
|
|-- manifest.xml
|-- package.json
|-- package-lock.json
|-- webpack.config.js
`-- README.md
```

The shared QuickBooks and session helpers are stored under `server/` instead of `api/`. This prevents Vercel from counting helper modules as separate Serverless Functions and keeps the deployment compatible with the Vercel Hobby plan.

---

## API Endpoints

### Application

```text
GET /api/health
GET /api/history
```

### Primary QuickBooks Data

```text
GET /api/accounts
GET /api/vendors
GET /api/customers
GET /api/classes
```

### Additional QuickBooks Data

```text
GET /api/more-data?dataset=<dataset-key>
```

Supported dataset keys:

```text
items
employees
locations
terms
paymentMethods
taxCodes
invoices
bills
payments
expenses
deposits
purchaseOrders
journalEntries
```

### QuickBooks OAuth

```text
GET  /api/qbo/start
GET  /api/qbo/callback
GET  /api/qbo/status
POST /api/qbo/disconnect
```

---

## Production Deployment

### Production URLs

```text
https://fin-accruals.vercel.app/taskpane.html
https://fin-accruals.vercel.app/commands.html
https://fin-accruals.vercel.app/manifest.xml
https://fin-accruals.vercel.app/api/health
https://fin-accruals.vercel.app/api/qbo/status
```

### Vercel Hobby Plan

The current deployment uses 11 Serverless Functions and remains within the Hobby-plan limit of 12.

Files under `server/` are shared modules and are not intended to become independent API routes.

---

## Intuit Developer Setup

### Requirements

- Intuit Developer account
- FinAccruals QuickBooks application
- QuickBooks sandbox company
- Development Client ID
- Development Client Secret

### Redirect URI

Add the following exact URI under the application's **Development Redirect URIs**:

```text
https://fin-accruals.vercel.app/api/qbo/callback
```

The redirect URI must match exactly. Do not add:

- A trailing slash
- Extra spaces
- A different protocol
- A Preview deployment domain

### QuickBooks Environment

The MVP uses:

```text
QBO_ENVIRONMENT=sandbox
```

Production credentials and real-company connections are not currently enabled.

---

## Vercel Environment Variables

Configure these variables under the Vercel project:

```text
QBO_CLIENT_ID
QBO_CLIENT_SECRET
QBO_REDIRECT_URI
QBO_ENVIRONMENT
SESSION_SECRET
```

Recommended values:

```text
QBO_REDIRECT_URI=https://fin-accruals.vercel.app/api/qbo/callback
QBO_ENVIRONMENT=sandbox
```

Apply the variables to:

- Production
- Preview
- Development

`SESSION_SECRET` should be a long, random value used to encrypt session data. If it is not configured, the current backend falls back to deriving its encryption key from `QBO_CLIENT_SECRET`.

Never commit environment variables, Client Secrets, OAuth tokens, or `.env` files to GitHub.

---

## Running the Excel Add-in

### Option 1: Excel for the Web

1. Open Excel for the web.
2. Open or create a workbook.
3. Open the Add-ins interface.
4. Select **Upload My Add-in**.
5. Upload `manifest.xml`.
6. Open FinAccruals from the Excel ribbon.
7. Connect QuickBooks.

Excel for the web is the recommended environment for testing the deployed Vercel application.

### Option 2: Excel Desktop Development

Open PowerShell in the repository:

```powershell
cd "C:\path\to\FinAccruals"
npm install
npm start
```

The Office Add-in debugging tools should:

1. Start the local development server.
2. Register `manifest.xml`.
3. Open Microsoft Excel.
4. Load the add-in for debugging.

If this message appears:

```text
'office-addin-debugging' is not recognized
```

Run:

```powershell
npm install
```

Make sure the command is executed inside the same project directory that contains `package.json`.

Excel desktop may not show the same **Upload My Add-in** option available in Excel for the web. Manual desktop sideloading may require a trusted add-in catalog.

---

## Local Development Setup

### Prerequisites

- Node.js
- npm
- Git
- Microsoft Excel
- Intuit Developer account
- QuickBooks sandbox company

### Installation

```bash
git clone https://github.com/kbk2002/FinAccruals.git
cd FinAccruals
npm install
```

### Development Server

```bash
npm start
```

### Stop Development Session

```bash
npm stop
```

### Run Webpack Development Server

```bash
npm run dev-server
```

### Production Build

```bash
npm run build
```

### Lint

```bash
npm run lint
```

### Validate Manifest

```bash
npm run validate
```

---

## Testing the Application

### Test QuickBooks Connection

1. Open the FinAccruals task pane.
2. Click **Connect** under QuickBooks Online.
3. Sign in to Intuit.
4. Select the sandbox company.
5. Click **Connect** on the Intuit authorization screen.

Expected result:

- The authorization window closes.
- The task pane shows **Connected**.
- The QuickBooks company name appears.
- Sync and import controls become enabled.
- The Connect action changes to Disconnect.

### Test Accounts

1. Click **Accounts → Sync**.
2. Open the `Accounts` worksheet.

Expected result:

- The worksheet is created or refreshed.
- QuickBooks account records are displayed.

### Test Vendors

1. Click **Vendors → Sync**.
2. Open the `Vendors` worksheet.

Expected result:

- The worksheet is created or refreshed.
- QuickBooks vendor records are displayed.

### Test Customers

1. Click **Customers → Sync**.
2. Open the `Customers` worksheet.

Expected result:

- The worksheet is created or refreshed.
- Customer identifiers, names, contact details, and balances are displayed.

### Test Classes

1. Enable class tracking in QuickBooks if necessary:

```text
Settings > Account and settings > Advanced > Categories > Track classes
```

2. Add classes through the QuickBooks Classes list.
3. Click **Classes → Sync**.

Expected result:

- The `Classes` worksheet is created or refreshed.
- QuickBooks class names are displayed.

### Test More QuickBooks Data

1. Open the **More QuickBooks Data** selector.
2. Select a dataset.
3. Click **Import table**.

Example:

```text
Products & Services
```

Expected worksheet:

```text
Products_Services
```

Another example:

```text
Customer Payments
```

Expected worksheet:

```text
Customer_Payments
```

If QuickBooks returns no records, the add-in creates a worksheet containing only the column headers. This is a valid empty result.

### Test Journal Entry Template

1. Click **Create template**.
2. Open the `JE_Template` worksheet.

Expected columns:

```text
Line #
Account
Vendor
Class
Description
Debit
Credit
Date
```

### Test Journal Entry Validation

Enter a balanced example:

| Line # | Account | Vendor | Class | Description | Debit | Credit | Date |
|---|---|---|---|---|---:|---:|---|
| 1 | Rent Expense | Sample Vendor | Operations | Rent expense | 500 | 0 | 6/23/2026 |
| 2 | Checking | Sample Vendor | Operations | Cash payment | 0 | 500 | 6/23/2026 |

Click **Validate Entry**.

Expected result:

```text
Validation passed.
Total Debits = Total Credits = $500.00
```

### Test Demo Submission

1. Complete successful validation.
2. Click **Submit demo journal**.

Expected result:

```text
Submitted successfully.
Ref: DEMO-JE-XXXXXXXX
```

No journal entry is posted to QuickBooks during this demonstration workflow.

---

## QuickBooks Sandbox Data

QuickBooks sandbox companies contain fictional sample records for development.

Imported data may include:

- Intuit-provided sample records
- Records created manually in the sandbox
- Legacy Accounting API records

Some records returned by the QuickBooks API may not be visible in the newest QuickBooks web-interface screens. This can occur when Intuit retains legacy sandbox records in the Accounting API while the modern UI uses a newer feature or workflow.

FinAccruals does not generate these imported records. It displays records returned by the connected QuickBooks API.

---

## Security

- QuickBooks Client Secrets remain on the backend.
- OAuth access tokens are never sent to frontend JavaScript.
- OAuth state validation protects the callback workflow.
- Session data is encrypted.
- Session cookies are HTTP-only and Secure.
- OAuth tokens and secrets are excluded from Git.
- The current integration uses sandbox credentials.
- Real businesses must explicitly authorize production access.
- Production rollout will require stronger multi-user token storage and compliance controls.

---

## Known Limitations

- QuickBooks does not provide one endpoint that lists every available table.
- Supported datasets must be configured in FinAccruals.
- The More Data menu is not yet dynamically filtered by company capability.
- QuickBooks API records may not always match the newest QuickBooks UI.
- Transaction imports currently include summary-level fields.
- Transaction line details are not yet imported into separate worksheets.
- QuickBooks queries currently request up to 1,000 records.
- Large-company pagination is not yet implemented.
- Journal entries are not yet posted to QuickBooks.
- Submission history is not persisted in a database.
- OAuth sessions are cookie-based and not yet designed for enterprise multi-user storage.
- Xero is not yet connected.

---

## What Is Not Completed Yet

- Real journal-entry posting to QuickBooks.
- Transaction line-detail imports.
- Pagination beyond 1,000 records.
- Date and status filters for transaction imports.
- Dynamic detection of supported QuickBooks datasets.
- QuickBooks financial reports such as Trial Balance and General Ledger.
- Persistent database storage.
- Persistent submission history.
- User authentication.
- Role-based access control.
- Multi-company and multi-tenant support.
- Approval workflows.
- Audit logging.
- Operational monitoring.
- Production QuickBooks compliance.
- Xero OAuth and API integration.

---

## Next Steps

1. Add transaction-line worksheets for Invoices, Bills, Purchases, and Journal Entries.
2. Add import filters for date range, status, and active/inactive records.
3. Add pagination for companies with more than 1,000 records.
4. Add a QuickBooks reports section.
5. Import Trial Balance, General Ledger, Profit and Loss, and Balance Sheet reports.
6. Add a centralized sync log.
7. Add persistent storage for submission and synchronization history.
8. Add real journal-entry posting to the QuickBooks sandbox.
9. Add user authentication and tenant isolation.
10. Add role-based approvals and audit logs.
11. Prepare the application for production QuickBooks review.
12. Add Xero as a separate integration.

---

## Summary

FinAccruals now includes a working Excel Add-in, an enterprise-style LedgerFlow task pane, real QuickBooks OAuth 2.0 authorization, secure sandbox connectivity, primary master-data synchronization, additional QuickBooks table imports, automatic Excel worksheet generation, journal-entry template creation, journal validation, demonstration submission workflows, and Vercel deployment.

The current phase proves the full connection:

```text
QuickBooks Sandbox
        |
        v
QuickBooks Accounting API
        |
        v
Vercel Backend
        |
        v
Office.js Task Pane
        |
        v
Excel Worksheets
```

The next phase will focus on deeper transaction detail, financial reports, persistence, journal posting, and enterprise controls.

---

## License

This project currently uses the MIT license inherited from the original Microsoft Office Add-in scaffold.

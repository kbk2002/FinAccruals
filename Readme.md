  FinAccruals / LedgerFlow Excel Add-in

   Overview

FinAccruals (LedgerFlow) is a Microsoft Excel Add-in built with Office.js to streamline accounting workflows directly within Excel. The current implementation provides a working Excel add-in with frontend functionality, workbook interaction, deployed API endpoints, and Excel-to-backend integration.

---

  Current Status

The project has completed the initial frontend and API integration phase.

Users can currently:

* Connect to the application in Demo Mode
* Pull Accounts into Excel
* Pull Vendors into Excel
* Pull Classes into Excel
* Create Journal Entry Templates
* Validate Journal Entries
* Submit Demo Journal Entries
* View Submission History
* Interact with deployed backend APIs hosted on Vercel

---

  What Has Been Completed

   Excel Add-in Foundation

* Created Office.js Excel Add-in project structure
* Built the LedgerFlow task pane interface
* Added Excel ribbon integration through `manifest.xml`
* Configured add-in branding and icons
* Configured Webpack build pipeline
* Connected GitHub repository to Vercel deployment
* Updated production manifest URLs

   Excel Workbook Integration

* Created automated `Accounts` worksheet generation
* Created automated `Vendors` worksheet generation
* Created automated `Classes` worksheet generation
* Created `JE_Template` worksheet generation
* Added workbook data writing using Office.js APIs

   Journal Entry Workflow

* Added journal entry template creation
* Implemented debit and credit balance validation
* Added journal entry submission workflow
* Added submission history display
* Added validation status messaging

   Backend API Integration

* Deployed backend API routes on Vercel
* Added API health endpoint
* Added Accounts API endpoint
* Added Vendors API endpoint
* Added Classes API endpoint
* Added Submission History API endpoint
* Connected Excel Add-in frontend to deployed API endpoints

   Master Data Functionality

* Pull Accounts from API into Excel
* Pull Vendors from API into Excel
* Pull Classes from API into Excel
* Populate dedicated worksheets automatically

---

  Current Working Features

✅ Connect to QuickBooks (Demo Mode)

✅ Pull Accounts

✅ Pull Vendors

✅ Pull Classes

✅ Create Journal Entry Template

✅ Validate Journal Entry Balance

✅ Submit Demo Journal Entry

✅ Load Submission History

✅ Excel ↔ API Integration

✅ Vercel Deployment

---

  Technology Stack

    Frontend

* JavaScript
* HTML
* CSS
* Office.js

    Build & Deployment

* Webpack
* Vercel Serverless Functions
* GitHub

    Planned

* Supabase
* QuickBooks Online API
* Xero API

---

  Project Structure

```text
src/
├── taskpane/
│   ├── taskpane.html
│   ├── taskpane.css
│   └── taskpane.js
│
├── commands/
│   ├── commands.html
│   └── commands.js

api/
├── health.js
├── accounts.js
├── vendors.js
├── classes.js
├── history.js

assets/
├── icon-16.png
├── icon-32.png
├── icon-64.png
├── icon-80.png
└── logo-filled.png

manifest.xml
webpack.config.js
package.json
```

---

  Production Deployment

   Production URLs

```text
https://fin-accruals.vercel.app/manifest.xml
https://fin-accruals.vercel.app/taskpane.html
https://fin-accruals.vercel.app/commands.html
https://fin-accruals.vercel.app/api/health
```

   API Endpoints

```text
GET /api/health
GET /api/accounts
GET /api/vendors
GET /api/classes
GET /api/history
```

---

  Running the Excel Add-in

   Option 1: Test Using the Deployed Manifest

    Excel Desktop

1. Open Microsoft Excel Desktop.
2. Open any workbook.
3. Go to:
   Insert → My Add-ins → Upload My Add-in
4. Upload the `manifest.xml` file.
5. Open the FinAccruals add-in from the Home ribbon.

    Excel Online

1. Open Excel Online.
2. Open any workbook.
3. Insert → Add-ins → Upload My Add-in.
4. Upload the `manifest.xml`.
5. Open the add-in.

---

  Testing the Application

   Pull Accounts

1. Click **Pull Accounts**
2. Verify an `Accounts` worksheet is created.
3. Verify sample account data is loaded.

   Pull Vendors

1. Click **Pull Vendors**
2. Verify a `Vendors` worksheet is created.
3. Verify vendor records are loaded.

   Pull Classes

1. Click **Pull Classes**
2. Verify a `Classes` worksheet is created.
3. Verify class records are loaded.

   Create Journal Entry Template

1. Click **Create / Refresh JE Template**
2. Verify a `JE_Template` worksheet is created.

   Validate Journal Entry

Enter sample data:

```text
Line   | Account       | Vendor              | Class      | Description | Debit | Credit | Date
1      | Rent Expense  | Acme Property Mgmt  | Operations | Rent        | 500   | 0      | 6/22/2026
2      | Checking      | Acme Property Mgmt  | Operations | Payment     | 0     | 500    | 6/22/2026
```

Click **Validate Entry**

Expected Result:

```text
Entry is balanced.
Debits = Credits = 500.00
```

   Submit Journal Entry

1. Click **Submit Journal Entry**
2. Demo reference number is generated.

Example:

```text
Submitted successfully.
Ref: DEMO-JE-XXXXXXXX
```

   Load History

1. Click **Load History**
2. Demo history records should appear.

Example:

```text
JE-001
2026-06-22
Demo Submitted
$500
```

---

  Local Development Setup

   Prerequisites

* Node.js 18+
* Git
* Microsoft Excel Desktop

   Installation

```bash
git clone <repository-url>
cd FinAccruals
npm install
```

   Run Development Server

```bash
npm start
```

or

```bash
npm run start
```

   Build Production Version

```bash
npm run build
```

---

  What Is Not Completed Yet

The following features are planned for future phases:

* Real QuickBooks Online OAuth authentication
* Real Xero integration
* Supabase database integration
* Persistent journal entry storage
* Persistent submission history
* Real account synchronization from QuickBooks
* Real vendor synchronization from QuickBooks
* Journal entry posting to QuickBooks
* User authentication
* Role-based access control
* Audit logging
* Production accounting workflow automation

---

  Next Steps

1. Integrate Supabase database.
2. Store submitted journal entries in the database.
3. Replace mock history with database-driven history.
4. Implement QuickBooks OAuth authentication.
5. Pull real accounts and vendors from QuickBooks.
6. Post journal entries to QuickBooks.
7. Add Xero integration.
8. Implement user authentication and audit logging.

---

  Summary

The project now includes a functioning Excel Add-in, deployed backend APIs, Excel workbook integration, master data synchronization, journal entry validation, demo submission workflows, and Vercel deployment.

The next phase focuses on database persistence, QuickBooks integration, and production accounting workflows.

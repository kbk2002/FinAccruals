# FinAccruals / LedgerFlow Excel Add-in

## Overview

FinAccruals, originally referred to as **LedgerFlow** in the project charter, is a Microsoft Excel add-in for accounting automation. The add-in lets accounting users stay inside Excel while connecting to QuickBooks Online, pulling accounting reference data, preparing journal entries, validating them, and posting approved entries back to QuickBooks.

The current implementation is an MVP focused on the **QuickBooks Online sandbox workflow**. Xero and production QuickBooks readiness are future phases.

## Current MVP Status

The QuickBooks sandbox MVP is working end to end.

Users can currently:

- Open the add-in inside Excel.
- Connect to a QuickBooks Online sandbox company through OAuth 2.0.
- Disconnect the active QuickBooks session.
- See connected company status.
- Pull QuickBooks master data into Excel:
  - Accounts
  - Vendors
  - Customers
  - Classes
- Import additional QuickBooks datasets through the "More QuickBooks data" dropdown.
- Generate a guided journal-entry worksheet.
- Use QuickBooks-powered dropdowns for journal entry fields.
- Validate journal entries before posting.
- Post a validated journal entry to QuickBooks sandbox.
- See success and failure messages.

The app has already successfully posted a journal entry to the QuickBooks sandbox during testing.

## Product Naming Note

The original project documents use the name **LedgerFlow**. During development, the product/repository name moved toward **FinAccruals**.

Current naming convention:

- Repository/internal project name: `FinAccruals`
- Earlier project/charter name: `LedgerFlow`
- UI may still contain both names while branding is finalized

A final branding cleanup can be done later once the organization confirms the official product name.

## What This Project Solves

Accounting teams often prepare data in Excel but still manually enter or upload that data into accounting platforms. This creates:

- Repetitive manual work
- Higher risk of data-entry errors
- Slow journal-entry posting
- Friction between Excel workflows and accounting systems

FinAccruals turns Excel into a connected accounting preparation layer. It allows users to pull accounting data from QuickBooks, prepare entries in Excel, validate them, and post them back through the API.

## Current Architecture

```text
Excel Add-in UI
src/taskpane/
        |
        v
Vercel Serverless APIs
api/
        |
        v
Shared QuickBooks helpers
server/
        |
        v
QuickBooks Online Sandbox API
```

Main parts:

- `src/taskpane/taskpane.html` - Excel task pane markup
- `src/taskpane/taskpane.css` - task pane styling
- `src/taskpane/taskpane.js` - Office.js workbook logic and UI behavior
- `manifest.xml` - Office add-in manifest
- `api/` - Vercel serverless API endpoints
- `server/qbo.js` - QuickBooks OAuth/API helper functions
- `server/session.js` - encrypted session cookie handling
- `assets/` - logo and add-in icon assets

## Key Features

### QuickBooks OAuth

The add-in connects to QuickBooks Online through OAuth 2.0.

Implemented endpoints:

- `GET /api/qbo/start`
- `GET /api/qbo/callback`
- `GET /api/qbo/status`
- `POST /api/qbo/disconnect`

The current environment is configured for QuickBooks sandbox testing.

### Master Data Sync

The add-in can create or refresh Excel worksheets from QuickBooks data:

- `Accounts`
- `Vendors`
- `Customers`
- `Classes`

These are used both for review and for journal-entry dropdowns.

### More QuickBooks Data

The add-in also supports importing additional QuickBooks datasets:

Reference data:

- Products & Services
- Employees
- Locations
- Payment Terms
- Payment Methods
- Tax Codes

Transaction data:

- Invoices
- Bills
- Customer Payments
- Expenses
- Deposits
- Purchase Orders
- Journal Entries

Transaction datasets support date-range filtering.

### Journal Entry Workspace

The add-in generates a `JE_Template` worksheet for journal-entry preparation.

The current template includes:

- One journal date at the top
- Memo/purpose field
- Posting mode indicator
- Line-entry table
- QuickBooks account dropdown
- QuickBooks vendor dropdown
- QuickBooks class dropdown
- Debit and credit columns
- Auto-filled journal date column
- Control totals
- Balance status

The supported journal-entry columns are:

```text
Line # | Account * | Vendor | Class | Description | Debit * | Credit * | Date *
```

### Journal Entry Validation

Validation exists in both the Excel add-in and backend API.

The supported JE validation checks include:

- At least one journal line exists
- Account is required
- Date is required
- All lines use the same journal date
- Debit and credit are numeric
- Negative amounts are not allowed
- Debit/credit values can have at most two decimals
- Each row must have either debit or credit, not both
- Total debits must equal total credits
- Referenced accounts/vendors/classes must exist in QuickBooks
- Unsupported submitted fields are rejected

### Journal Entry Posting

Validated journal entries can be posted to the connected QuickBooks sandbox company using:

- `POST /api/journal-entries`

The backend resolves Excel-entered account/vendor/class names to QuickBooks IDs before creating the QuickBooks `JournalEntry`.

## API Endpoints

Current API files:

```text
api/health.js
api/accounts.js
api/vendors.js
api/customers.js
api/classes.js
api/more-data.js
api/history.js
api/journal-entries.js
api/qbo/start.js
api/qbo/callback.js
api/qbo/status.js
api/qbo/disconnect.js
```

Important note: this project has been designed around Vercel Hobby limits. Adding too many separate API files may exceed the Hobby serverless function limit. If more endpoints are needed, consolidate routes or move to an organization-approved backend.

## Environment Variables

The deployed backend requires QuickBooks and session configuration.

Required variables:

```env
QBO_CLIENT_ID=
QBO_CLIENT_SECRET=
QBO_REDIRECT_URI=https://fin-accruals.vercel.app/api/qbo/callback
QBO_ENVIRONMENT=sandbox
SESSION_SECRET=
```

Do not commit real credentials to GitHub.

For local testing, use a local `.env` file or environment configuration supported by the hosting platform.

## QuickBooks Developer Setup

1. Create or open an Intuit Developer account.
2. Create a QuickBooks Online app.
3. Use sandbox mode for MVP testing.
4. Add the redirect URI:

```text
https://fin-accruals.vercel.app/api/qbo/callback
```

5. Copy the client ID and client secret into the deployment environment variables.
6. Confirm the app has access to Accounting APIs.

For local development, a localhost redirect URI may also be needed if testing OAuth locally.

## QuickBooks Sandbox Testing

This MVP currently uses QuickBooks sandbox data.

Sandbox data may include sample accounts such as:

- Checking
- Rent or Lease
- Services
- Accounts Payable
- Accounts Receivable
- Landscaping Services

That data comes from the connected QuickBooks sandbox company. It is not hardcoded in the add-in.

To let teammates test the same sandbox company:

1. Open the QuickBooks sandbox company.
2. Go to QuickBooks settings.
3. Open user management / manage users.
4. Invite teammates to the same QuickBooks company.
5. Ask them to accept the invite.
6. When connecting the add-in, they must select the same QuickBooks company.

Inviting teammates to the Intuit Developer workspace shares the app configuration, but it does not automatically make them use the same QuickBooks company data.

## Installation

Install dependencies:

```bash
npm install
```

Or, for a clean install from the lockfile:

```bash
npm ci
```

## Local Development

Start the development server:

```bash
npm run dev-server
```

Start Office add-in debugging:

```bash
npm start
```

If `office-addin-debugging` is not recognized, install dependencies first:

```bash
npm ci
```

Then run:

```bash
npm start
```

## Build

Create a production build:

```bash
npm run build
```

Create a development build:

```bash
npm run build:dev
```

Validate the Office manifest:

```bash
npm run validate
```

## Deployment

The current MVP is deployed through Vercel.

Deployment requirements:

- GitHub repository connected to Vercel
- Environment variables configured in Vercel
- `manifest.xml` points to the deployed task pane URLs
- Intuit Developer redirect URI matches the deployed callback URL exactly

Current production-style URL:

```text
https://fin-accruals.vercel.app
```

## How to Test the MVP

### 1. Open the add-in

Open Excel for the web and load the add-in using the deployed manifest.

### 2. Connect QuickBooks

Click the QuickBooks connect button and authorize the sandbox company.

Expected result:

- Add-in shows connected status
- Company name appears in the task pane

### 3. Sync master data

Use the sync buttons for:

- Accounts
- Vendors
- Customers
- Classes

Expected result:

- Matching Excel worksheets are created or refreshed
- Rows should match data from the connected QuickBooks sandbox company

### 4. Import more QuickBooks data

Use the "More QuickBooks data" dropdown.

Expected result:

- A dedicated worksheet is created for the selected dataset

### 5. Generate journal-entry workspace

Click the journal-entry workspace/template button.

Expected result:

- `JE_Template` worksheet is created
- Account/Vendor/Class dropdowns are available
- Helper sync data should not clutter the workbook unnecessarily

### 6. Prepare a balanced journal entry

Example test:

```text
Line 1: Account = Rent or Lease, Debit = 500
Line 2: Account = Checking, Credit = 500
```

Use account names that exist in the connected QuickBooks sandbox company.

### 7. Validate

Run the control review / validation step.

Expected result:

- Entry passes if balanced and all required fields exist
- Clear error appears if something is missing or invalid

### 8. Post to QuickBooks sandbox

Post the validated journal entry.

Expected result:

- Success message appears
- QuickBooks returns a journal entry ID/reference
- Entry can be found in QuickBooks sandbox reports/search

## Known Limitations

This is an MVP, not a complete enterprise release.

Current limitations:

- QuickBooks sandbox only
- Xero integration is not implemented yet
- No standalone FinAccruals user account system yet
- No enterprise database for multi-user token/session storage
- No production QuickBooks approval yet
- No Microsoft AppSource marketplace submission yet
- No formal subscription or billing module
- Audit logging is basic and not enterprise-grade
- Hosting is currently on Vercel and may need organization ownership
- Intuit Developer app ownership may need to move to the organization

## Project Charter Alignment

The project charter requested:

- Excel add-in
- Accounting platform connection
- Master-data pull
- Journal-entry template
- Journal-entry validation
- Journal-entry posting
- Success/failure status
- Backend/API infrastructure

The current MVP satisfies the core QuickBooks sandbox workflow:

```text
Excel -> Add-in -> QuickBooks OAuth -> Pull data -> Prepare JE -> Validate -> Post to QuickBooks sandbox
```

Not yet complete from the full 30-day application scope:

- Xero integration
- Production QuickBooks workflow
- Full app-level user authentication
- Enterprise audit logging
- Subscription/access-control foundation
- Formal QA report
- Full deployment and handoff documentation
- Marketplace readiness

## Recommended Next Steps

1. Finish team testing on the same QuickBooks sandbox company.
2. Confirm the final product name: `FinAccruals` or `LedgerFlow`.
3. Move GitHub, Vercel, and Intuit Developer ownership to organization-managed accounts.
4. Create a formal QA checklist/report.
5. Decide whether to keep Vercel or move APIs to an organization-approved backend.
6. Add durable database/session storage if multiple real users will use the app.
7. Begin Xero integration only after QuickBooks MVP is stable.
8. Prepare production QuickBooks review and marketplace strategy if required.

## Security Notes

- Never commit QuickBooks client secrets.
- Never commit `.env` files.
- Keep OAuth token handling server-side.
- Use HTTPS redirect URIs.
- Rotate secrets if they were ever exposed.
- Use sandbox companies for testing.
- Do not connect real client/company data until production security and ownership are approved.

## Repository Structure

```text
FinAccruals/
├── api/
│   ├── qbo/
│   │   ├── start.js
│   │   ├── callback.js
│   │   ├── status.js
│   │   └── disconnect.js
│   ├── accounts.js
│   ├── vendors.js
│   ├── customers.js
│   ├── classes.js
│   ├── more-data.js
│   ├── journal-entries.js
│   ├── history.js
│   └── health.js
├── assets/
├── server/
│   ├── qbo.js
│   └── session.js
├── src/
│   ├── commands/
│   └── taskpane/
│       ├── taskpane.html
│       ├── taskpane.css
│       └── taskpane.js
├── manifest.xml
├── package.json
├── package-lock.json
├── webpack.config.js
└── README.md
```

## License

This project is currently an MVP/prototype repository. Confirm final licensing and ownership with the organization before production release.

# FinAccruals

FinAccruals is a Microsoft Excel add-in for accounting preparation workflows. It connects Excel to QuickBooks Online, synchronizes reference data, creates a structured journal-entry template, and validates entries before submission.

The interface is branded as **LedgerFlow for Excel**, while the repository and QuickBooks application currently use the **FinAccruals** name.

## Current MVP status

The QuickBooks Online sandbox integration is working end to end:

- Connect and disconnect a QuickBooks sandbox company using OAuth 2.0.
- Display the connected QuickBooks company in the task pane.
- Synchronize the chart of accounts into an `Accounts` worksheet.
- Synchronize vendors into a `Vendors` worksheet.
- Synchronize classes into a `Classes` worksheet.
- Create or refresh a `JE_Template` worksheet.
- Validate journal lines, required fields, amounts, and debit/credit balance.
- Run the add-in in Excel for the web.
- Deploy the frontend and serverless API through Vercel.

Journal-entry submission and submission history are still demo workflows. Posting journal entries to QuickBooks and persistent database storage are the next major milestones.

## Architecture

```text
Excel task pane (Office.js)
        |
        v
Vercel serverless API
        |
        +-- QuickBooks OAuth 2.0
        +-- QuickBooks Online Accounting API
        |
        v
QuickBooks sandbox company
```

OAuth tokens are encrypted in secure HTTP-only cookies. QuickBooks credentials are stored as Vercel environment variables and must never be committed to Git.

## Technology

- JavaScript
- HTML and CSS
- Microsoft Office.js
- Webpack
- Vercel Serverless Functions
- QuickBooks Online Accounting API
- Intuit OAuth 2.0

## Project structure

```text
api/
|-- accounts.js
|-- classes.js
|-- health.js
|-- history.js
|-- vendors.js
|-- lib/
|   |-- qbo.js
|   `-- session.js
`-- qbo/
    |-- auth-url.js
    |-- callback.js
    |-- disconnect.js
    |-- start.js
    `-- status.js

src/
|-- commands/
|   |-- commands.html
|   `-- commands.js
`-- taskpane/
    |-- taskpane.css
    |-- taskpane.html
    `-- taskpane.js

assets/
manifest.xml
package.json
webpack.config.js
```

## Prerequisites

- Node.js and npm
- Microsoft Excel on the web or desktop Excel
- An Intuit Developer account
- A QuickBooks sandbox company
- A Vercel project connected to this repository

## Install

Clone the repository, open the project directory, and install its dependencies:

```bash
git clone https://github.com/kbk2002/FinAccruals.git
cd FinAccruals
npm install
```

If `office-addin-debugging` is not recognized, the dependencies have not been installed in the current project folder. Run `npm install` there before using `npm start`.

## QuickBooks developer configuration

Create a QuickBooks Online app in the Intuit Developer dashboard and use its **Development** credentials for sandbox testing.

Add this exact Development redirect URI:

```text
https://fin-accruals.vercel.app/api/qbo/callback
```

The URI must match exactly. Do not add a trailing slash.

## Vercel environment variables

Configure these variables in the Vercel project for Production, Preview, and Development:

```text
QBO_CLIENT_ID
QBO_CLIENT_SECRET
QBO_REDIRECT_URI=https://fin-accruals.vercel.app/api/qbo/callback
QBO_ENVIRONMENT=sandbox
```

An optional independent encryption secret can also be configured:

```text
SESSION_SECRET
```

Use a long random value for `SESSION_SECRET`. If it is omitted, the backend currently derives its encryption key from `QBO_CLIENT_SECRET`.

Never add credentials, tokens, `.env` files, or secrets to GitHub.

## Development commands

```bash
# Start desktop Excel debugging
npm start

# Stop the debugging session
npm stop

# Start only the local Webpack server
npm run dev-server

# Create a production build
npm run build

# Check code style
npm run lint

# Validate the Office add-in manifest
npm run validate
```

## Testing in Excel for the web

1. Push the latest code to GitHub.
2. Wait for the Vercel deployment to show **Ready**.
3. Open Excel for the web.
4. Sideload the deployed `manifest.xml`.
5. Open the FinAccruals task pane.
6. Select **Connect** under QuickBooks Online.
7. Authorize the sandbox company.
8. Confirm that the button changes to **Disconnect** and the company name appears.
9. Synchronize Accounts, Vendors, and Classes.

The status endpoint should return the current browser session state:

```text
https://fin-accruals.vercel.app/api/qbo/status
```

Before authorization:

```json
{"connected":false}
```

After authorization, the response includes the company name, realm ID, and sandbox environment.

## Testing in desktop Excel

From the project directory:

```bash
npm install
npm start
```

The Office add-in debugging tools should start the local development server, register `manifest.xml`, and open Excel.

Excel desktop does not always provide the same **Upload My Add-in** option shown in Excel for the web. Manual desktop sideloading may instead require a trusted add-in catalog.

## QuickBooks sandbox data

Accounts, vendors, and classes synchronized by the add-in come from the connected QuickBooks sandbox company. Intuit provides fictional sample data for development, so this is not real customer or business data.

To manage sandbox data:

1. Sign in to the Intuit Developer dashboard.
2. Open **My Hub**.
3. Open the relevant sandbox company.
4. Add or edit accounts, vendors, and classes in QuickBooks.
5. Return to Excel and synchronize again.

Class tracking may need to be enabled under:

```text
Settings > Account and settings > Advanced > Categories > Track classes
```

## Production URLs

```text
https://fin-accruals.vercel.app/taskpane.html
https://fin-accruals.vercel.app/manifest.xml
https://fin-accruals.vercel.app/api/health
https://fin-accruals.vercel.app/api/qbo/status
```

## Security notes

- Client secrets never belong in frontend JavaScript.
- OAuth tokens must never be committed to Git.
- The backend validates OAuth state to protect the callback flow.
- Session data is encrypted before being stored in an HTTP-only cookie.
- The current environment is intended for QuickBooks sandbox testing.
- Production QuickBooks access requires production credentials, compliance work, and authorization from each real business.

## Remaining work

- Post validated journal entries to the QuickBooks sandbox.
- Resolve account, vendor, and class names to QuickBooks IDs during submission.
- Add server-side journal validation.
- Add persistent storage for connections, submissions, and history.
- Add user authentication and tenant isolation.
- Add audit logs and operational monitoring.
- Improve token storage for multi-user production use.
- Add Xero as a separate future integration.
- Complete production security and compliance requirements.

## License

This project currently uses the MIT license inherited from the original Office add-in scaffold.

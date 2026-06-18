FinAccruals / LedgerFlow Excel Add-in
FinAccruals, also called LedgerFlow, is a Microsoft Excel Add-in built with Office.js. The current version focuses on the Excel add-in frontend structure, task pane interface, workbook interaction, and deployment setup.
Current Status
The project is currently in the frontend foundation phase. It includes the Excel add-in structure, UI screens, workbook template creation, basic journal entry validation, and Vercel deployment configuration.
What Has Been Completed
•	Created Office.js Excel Add-in project structure.
•	Built the LedgerFlow task pane UI.
•	Added Excel ribbon button support through manifest.xml.
•	Configured icons and add-in branding.
•	Added Webpack build setup.
•	Connected the deployment URL to Vercel.
•	Updated manifest URLs for production deployment.
•	Added workbook interaction for creating a JE_Template sheet.
•	Added basic client-side debit and credit balance validation.
•	Added frontend placeholders for:
o	QuickBooks connection
o	Pulling accounts
o	Pulling vendors
o	Submitting journal entries
o	Loading submission history
Tech Stack
•	JavaScript
•	HTML
•	CSS
•	Office.js
•	Webpack
•	Vercel
•	GitHub
Project Structure
src/
├── taskpane/
│   ├── taskpane.html
│   ├── taskpane.css
│   └── taskpane.js
│
├── commands/
│   ├── commands.html
│   └── commands.js

assets/
├── icon-16.png
├── icon-32.png
├── icon-64.png
├── icon-80.png
└── logo-filled.png

manifest.xml
webpack.config.js
package.json
Deployment
The project is deployed on Vercel.
Important URLs:
https://fin-accruals.vercel.app/manifest.xml
https://fin-accruals.vercel.app/taskpane.html
https://fin-accruals.vercel.app/commands.html
Note: The root URL may show a 404 because the project does not currently include a homepage index.html. This is expected for the current add-in setup.
What Is Not Completed Yet
The following functionality is not fully implemented yet:
•	Real QuickBooks Online OAuth connection
•	Real Xero integration
•	Backend API deployment
•	Database integration
•	Real account/vendor sync
•	Real journal entry posting
•	Authentication and authorization
•	Persistent audit logs
•	Production accounting workflow logic
Known Development Note
The frontend currently points API calls to:
const API_BASE = "http://localhost:3001";
This means backend-related buttons will only work if a backend server is running locally. For production, this needs to be changed to a deployed backend API URL.
Next Steps
•	Build and deploy the backend API.
•	Replace the local API URL with the production backend URL.
•	Implement QuickBooks OAuth.
•	Add real accounts and vendors sync.
•	Add journal entry posting.
•	Add database storage for logs and submission history.
•	Test the add-in inside Excel using the production manifest.
Summary
This repository currently provides the deployed frontend foundation for the FinAccruals / LedgerFlow Excel Add-in. The next phase is backend development and real accounting platform integration.


# FinAccruals

Excel Add-in MVP shared development environment for accounting automation and journal entry processing.

## MVP Architecture

- Office.js Excel Add-in
- GitHub for source control
- Vercel for shared deployment
- Supabase for MVP database

## Purpose

This repository provides the starter environment for the FinAccruals Excel Add-in so multiple developers can collaborate from one shared codebase and test using one deployed Vercel URL.

## Project Structure

```text
FinAccruals/
├── manifest.xml
├── index.html
├── package.json
├── vercel.json
├── api/
│   └── health.js
├── assets/
├── docs/
└── src/
    ├── commands/
    │   └── commands.html
    └── taskpane/
        ├── taskpane.html
        ├── taskpane.css
        └── taskpane.js
```

## Setup

1. Push this project to GitHub.
2. Import the GitHub repository into Vercel.
3. Deploy the project.
4. Replace all `YOUR_VERCEL_URL` placeholders in `manifest.xml` with the actual Vercel URL.
5. Sideload `manifest.xml` into Excel.

## Environment Variables

Add these in Vercel later when Supabase is connected:

```text
SUPABASE_URL=
SUPABASE_ANON_KEY=
SUPABASE_SERVICE_ROLE_KEY=
```

Do not commit secrets to GitHub.

## Status

Initial shared development environment created.

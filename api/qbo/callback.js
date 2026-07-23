import { exchangeCode, loadCompanyName } from "../../server/qbo.js";
import { consumeOAuthState, writeSession } from "../../server/session.js";

function callbackPage(success, message) {
  const safeMessage = JSON.stringify({ source: "finaccruals-qbo", success, message });
  const title = success ? "QuickBooks connected" : "Connection failed";
  return `<!doctype html>
<html>
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width,initial-scale=1">
    <title>${title}</title>
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  </head>
  <body style="font-family:Segoe UI,sans-serif;padding:32px;color:#172033">
    <h2>${title}</h2>
    <p>${success ? "You can return to FinAccruals." : "Please close this window and try again."}</p>
    <script>
      const payload = ${safeMessage};
      function notify() {
        if (window.Office && Office.context && Office.context.ui) {
          Office.context.ui.messageParent(JSON.stringify(payload));
        } else if (window.opener) {
          window.opener.postMessage(payload, window.location.origin);
          window.close();
        }
      }
      if (window.Office) Office.onReady(notify);
      setTimeout(notify, 1200);
    </script>
  </body>
</html>`;
}

export default async function handler(req, res) {
  res.setHeader("Content-Type", "text/html; charset=utf-8");

  try {
    const { code, realmId, state, error } = req.query;

    if (error) {
      return res.status(400).send(callbackPage(false, String(error)));
    }

    if (!code || !realmId || !consumeOAuthState(req, res, state)) {
      return res.status(400).send(callbackPage(false, "Invalid or expired OAuth response."));
    }

    const session = await exchangeCode(String(code), String(realmId));
    session.companyName = await loadCompanyName(session);
    await writeSession(res, session);
    return res.status(200).send(callbackPage(true, session.companyName || "QuickBooks"));
  } catch (error) {
    return res.status(500).send(callbackPage(false, error.message));
  }
}

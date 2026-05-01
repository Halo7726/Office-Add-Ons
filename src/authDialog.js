import { PublicClientApplication } from "@azure/msal-browser";

const SCOPES = ["User.Read", "Sites.ReadWrite.All", "Files.ReadWrite.All"];

function sendToParent(msg) {
  if (window.Office?.context?.ui) {
    Office.context.ui.messageParent(JSON.stringify(msg));
  }
}

async function run() {
  const params = new URLSearchParams(window.location.search);
  let clientId = params.get("clientId") || "";
  let tenantId = params.get("tenantId") || "common";

  // Persist across the MSAL redirect — after Microsoft redirects back to this page
  // the query params are gone. sessionStorage doesn't survive the cross-origin
  // round-trip through login.microsoftonline.com in Office's embedded browser,
  // so localStorage (same store MSAL uses) is used instead.
  if (clientId) {
    localStorage.setItem("auth_clientId", clientId);
    localStorage.setItem("auth_tenantId", tenantId);
  } else {
    clientId = localStorage.getItem("auth_clientId") || "";
    tenantId = localStorage.getItem("auth_tenantId") || "common";
  }

  if (!clientId) {
    sendToParent({ type: "error", message: "No clientId provided to auth dialog." });
    return;
  }

  // redirectUri must also be registered in your Azure AD app registration.
  const redirectUri = `${window.location.origin}/auth-dialog.html`;

  const pca = new PublicClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      redirectUri,
      navigateToLoginRequestUrl: false,
    },
    cache: { cacheLocation: "localStorage" },
  });

  await pca.initialize();

  const response = await pca.handleRedirectPromise().catch(() => null);

  if (response?.account) {
    pca.setActiveAccount(response.account);
    // Office dialogs run in a separate browser context that doesn't share
    // localStorage with the task pane. Serialize the MSAL cache so the parent
    // can import it before creating its own MSAL instance.
    const cache = {};
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key) cache[key] = localStorage.getItem(key);
    }
    sendToParent({ type: "done", cache });
    return;
  }

  // No existing redirect response — start the login flow.
  await pca.loginRedirect({ scopes: SCOPES });
}

function handleError(err) {
  sendToParent({ type: "error", message: err.message });
}

if (window.Office) {
  Office.onReady(() => run().catch(handleError));
} else {
  run().catch(handleError);
}

import { PublicClientApplication } from "@azure/msal-browser";

const SCOPES = ["User.Read", "Sites.ReadWrite.All", "Files.ReadWrite.All"];

function sendToParent(msg) {
  if (window.Office?.context?.ui) {
    Office.context.ui.messageParent(JSON.stringify(msg));
  }
}

async function run() {
  const params = new URLSearchParams(window.location.search);
  const clientId = params.get("clientId") || "";
  const tenantId = params.get("tenantId") || "common";

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
    sendToParent({ type: "done" });
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

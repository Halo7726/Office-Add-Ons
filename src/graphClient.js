import { PublicClientApplication } from "@azure/msal-browser";

const SCOPES = ["User.Read", "Sites.ReadWrite.All", "Files.ReadWrite.All"];

let msalClient;
let msalClientId;

function ensureClient(config) {
  if (msalClient && msalClientId === config.clientId) return msalClient;

  if (!config.clientId) {
    throw new Error("Client ID is missing. Add VITE_CLIENT_ID in .env.");
  }

  msalClient = new PublicClientApplication({
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId || "common"}`,
      // Must match a redirect URI registered in your Azure AD app registration.
      // For the Office dialog auth flow, also register: <origin>/auth-dialog.html
      redirectUri: window.location.origin,
    },
    cache: {
      cacheLocation: "localStorage",
    },
  });
  msalClientId = config.clientId;

  return msalClient;
}

// Opens a dedicated auth dialog using Office.context.ui.displayDialogAsync.
// The dialog runs the full MSAL redirect flow at /auth-dialog.html and writes
// tokens into the shared localStorage so the parent can acquire silently afterwards.
function loginWithOfficeDialog(config) {
  return new Promise((resolve, reject) => {
    const params = new URLSearchParams({
      clientId: config.clientId,
      tenantId: config.tenantId || "common",
    });
    const dialogUrl = `${window.location.origin}/auth-dialog.html?${params}`;

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 60, width: 30, promptBeforeOpen: false },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(`Auth dialog failed to open: ${asyncResult.error.message}`));
          return;
        }

        const dialog = asyncResult.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          dialog.close();
          try {
            const msg = JSON.parse(arg.message);
            if (msg.type === "done") {
              resolve();
            } else {
              reject(new Error(msg.message || "Authentication failed in dialog."));
            }
          } catch {
            reject(new Error("Unexpected response from auth dialog."));
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          if (arg.error === 12006) {
            reject(new Error("Sign-in cancelled."));
          }
        });
      }
    );
  });
}

async function interactiveLogin(config, client) {
  if (window.Office?.context?.ui) {
    await loginWithOfficeDialog(config);
    const accounts = client.getAllAccounts();
    if (accounts.length === 0) {
      throw new Error("Authentication completed but no account was found in cache.");
    }
    client.setActiveAccount(accounts[0]);
    return accounts[0];
  }

  const loginResult = await client.loginPopup({ scopes: SCOPES });
  client.setActiveAccount(loginResult.account);
  return loginResult.account;
}

export async function signIn(config) {
  const client = ensureClient(config);
  await client.initialize();

  const accounts = client.getAllAccounts();
  if (accounts.length > 0) {
    client.setActiveAccount(accounts[0]);
    return accounts[0];
  }

  return interactiveLogin(config, client);
}

async function getAccessToken(config) {
  const client = ensureClient(config);
  await client.initialize();

  let account = client.getActiveAccount();
  if (!account) {
    await interactiveLogin(config, client);
    account = client.getActiveAccount();
  }

  try {
    const token = await client.acquireTokenSilent({ account, scopes: SCOPES });
    return token.accessToken;
  } catch {
    await interactiveLogin(config, client);
    const freshAccount = client.getActiveAccount();
    const token = await client.acquireTokenSilent({ account: freshAccount, scopes: SCOPES });
    return token.accessToken;
  }
}

async function graphRequest(config, path, options = {}) {
  const accessToken = await getAccessToken(config);
  const hasBody = options.body !== undefined;
  const defaultHeaders = {
    Authorization: `Bearer ${accessToken}`,
  };

  if (hasBody && !options.headers?.["Content-Type"]) {
    defaultHeaders["Content-Type"] = "application/json";
  }

  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...options,
    headers: {
      ...defaultHeaders,
      ...(options.headers || {}),
    },
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Graph request failed (${response.status}): ${errorText}`);
  }

  if (response.status === 204) return null;
  return response.json();
}

export async function getListItems(config) {
  return getListItemsById(config, config.listId);
}

export async function getListItemsById(config, listId) {
  if (!config.siteId || !listId) {
    throw new Error("Both Site ID and List ID are required.");
  }

  return graphRequest(config, `/sites/${config.siteId}/lists/${listId}/items?expand=fields`, {
    method: "GET",
  });
}

export async function addListItem(config, fields) {
  if (!config.siteId || !config.listId) {
    throw new Error("Both Site ID and List ID are required.");
  }

  return graphRequest(config, `/sites/${config.siteId}/lists/${config.listId}/items`, {
    method: "POST",
    body: JSON.stringify({ fields }),
  });
}

export async function updateListItemFields(config, listId, itemId, fields) {
  if (!config.siteId || !listId || !itemId) {
    throw new Error("Site ID, List ID, and item ID are required.");
  }

  return graphRequest(config, `/sites/${config.siteId}/lists/${listId}/items/${itemId}/fields`, {
    method: "PATCH",
    body: JSON.stringify(fields),
  });
}

export async function uploadToLibrary(config, file) {
  if (!config.siteId || !config.driveId) {
    throw new Error("Site ID and Drive ID are required to upload files.");
  }

  const folder = (config.libraryFolder || "Shared Documents").replace(/^\/+|\/+$/g, "");
  return uploadBytesToLibrary(config, file.name, file, folder, file.type || "application/octet-stream");
}

async function getDriveRoot(config) {
  return graphRequest(config, `/sites/${config.siteId}/drives/${config.driveId}/root`, {
    method: "GET",
  });
}

async function ensureFolder(config, parentId, folderName) {
  const encodedName = encodeURIComponent(folderName);
  try {
    return await graphRequest(
      config,
      `/sites/${config.siteId}/drives/${config.driveId}/items/${parentId}:/${encodedName}`,
      { method: "GET" }
    );
  } catch {
    return graphRequest(config, `/sites/${config.siteId}/drives/${config.driveId}/items/${parentId}/children`, {
      method: "POST",
      body: JSON.stringify({
        name: folderName,
        folder: {},
        "@microsoft.graph.conflictBehavior": "fail",
      }),
    });
  }
}

export async function ensureLibraryPath(config, folderPath) {
  if (!folderPath) return;

  const cleaned = folderPath.replace(/^\/+|\/+$/g, "");
  if (!cleaned) return;

  const segments = cleaned.split("/").map((segment) => segment.trim()).filter(Boolean);
  if (segments.length === 0) return;

  let current = await getDriveRoot(config);

  for (const segment of segments) {
    current = await ensureFolder(config, current.id, segment);
  }
}

export async function uploadBytesToLibrary(
  config,
  fileName,
  data,
  folderPath,
  contentType = "application/octet-stream"
) {
  if (!config.siteId || !config.driveId) {
    throw new Error("Site ID and Drive ID are required to upload files.");
  }

  const folder = (folderPath || config.libraryFolder || "Shared Documents").replace(/^\/+|\/+$/g, "");
  await ensureLibraryPath(config, folder);

  const encodedFolder = folder.split("/").map(encodeURIComponent).join("/");
  const uploadPath = `/sites/${config.siteId}/drives/${config.driveId}/items/root:/${encodedFolder}/${encodeURIComponent(fileName)}:/content`;
  const accessToken = await getAccessToken(config);

  const response = await fetch(`https://graph.microsoft.com/v1.0${uploadPath}`, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": contentType,
    },
    body: data,
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`File upload failed (${response.status}): ${errorText}`);
  }

  return response.json();
}

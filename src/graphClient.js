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

function hasOfficeSso() {
  return typeof window !== "undefined" && typeof window.OfficeRuntime !== "undefined" && typeof window.OfficeRuntime.auth?.getAccessToken === "function";
}

function getSsoAccount() {
  if (typeof window === "undefined" || typeof window.Office === "undefined") return null;
  const profile = Office?.context?.mailbox?.userProfile;
  if (!profile) return null;
  return { username: profile.emailAddress || profile.displayName || "Outlook user" };
}

async function getSsoAccessToken(allowSignInPrompt = false) {
  if (!hasOfficeSso()) {
    throw new Error("Outlook SSO is not available in this host.");
  }

  try {
    return await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt });
  } catch (err) {
    throw err;
  }
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
              // Import the MSAL cache from the dialog's isolated browser context
              // into this window's localStorage so a fresh MSAL instance can find it.
              if (msg.cache) {
                for (const [key, value] of Object.entries(msg.cache)) {
                  localStorage.setItem(key, value);
                }
              }
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
    // The dialog wrote its MSAL cache into our localStorage above. Reset the
    // cached client so ensureClient creates a fresh instance that reads it.
    msalClient = null;
    const freshClient = ensureClient(config);
    await freshClient.initialize();
    const accounts = freshClient.getAllAccounts();
    if (accounts.length === 0) {
      throw new Error("Authentication completed but no account was found in cache.");
    }
    freshClient.setActiveAccount(accounts[0]);
    return accounts[0];
  }

  const loginResult = await client.loginPopup({ scopes: SCOPES });
  client.setActiveAccount(loginResult.account);
  return loginResult.account;
}

export async function signIn(config) {
  if (hasOfficeSso()) {
    try {
      await getSsoAccessToken(true);
      return getSsoAccount();
    } catch (err) {
      console.warn("Outlook SSO signin failed, falling back to MSAL:", err);
    }
  }

  const client = ensureClient(config);
  await client.initialize();

  const accounts = client.getAllAccounts();
  if (accounts.length > 0) {
    client.setActiveAccount(accounts[0]);
    return accounts[0];
  }

  return interactiveLogin(config, client);
}

export async function restoreSignIn(config) {
  if (hasOfficeSso()) {
    try {
      await getSsoAccessToken(false);
      return getSsoAccount();
    } catch {
      return null;
    }
  }

  const client = ensureClient(config);
  await client.initialize();

  const accounts = client.getAllAccounts();
  if (accounts.length > 0) {
    client.setActiveAccount(accounts[0]);
    return accounts[0];
  }
  return null;
}

async function getAccessToken(config) {
  if (hasOfficeSso()) {
    try {
      return await getSsoAccessToken(false);
    } catch (err) {
      console.warn("Outlook SSO token acquisition failed, falling back to MSAL:", err);
    }
  }

  const client = ensureClient(config);
  await client.initialize();

  let account = client.getActiveAccount();
  if (!account) {
    await interactiveLogin(config, client);
    account = ensureClient(config).getActiveAccount();
  }

  try {
    const token = await ensureClient(config).acquireTokenSilent({ account, scopes: SCOPES });
    return token.accessToken;
  } catch {
    await interactiveLogin(config, client);
    const freshAccount = ensureClient(config).getActiveAccount();
    const token = await ensureClient(config).acquireTokenSilent({ account: freshAccount, scopes: SCOPES });
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

  return uploadBytesToLibrary(config, file.name, file, "", file.type || "application/octet-stream");
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

// Returns the child folder of parentId whose name matches folderName (case-insensitive), or null.
async function getChildFolderByName(config, parentId, folderName) {
  try {
    const results = await graphRequest(
      config,
      `/sites/${config.siteId}/drives/${config.driveId}/items/${parentId}/children`,
      { method: "GET" }
    );
    const lower = folderName.toLowerCase();
    return (results?.value || []).find(
      (item) => item.folder !== undefined && item.name.toLowerCase() === lower
    ) || null;
  } catch {
    return null;
  }
}

// Searches the entire drive for a folder whose name exactly matches folderName.
async function searchDriveForFolder(config, folderName) {
  try {
    const q = encodeURIComponent(`"${folderName}"`);
    const results = await graphRequest(
      config,
      `/sites/${config.siteId}/drives/${config.driveId}/root/search(q=${q})`,
      { method: "GET" }
    );
    const lower = folderName.toLowerCase();
    return (results?.value || []).find(
      (item) => item.folder !== undefined && item.name.toLowerCase() === lower
    ) || null;
  } catch {
    return null;
  }
}

// Navigate to the project folder at projectFolderPath (drive-relative), then find
// the existing typeFolder (Vendors/Subcontractors) inside it, then find or create
// the companyName subfolder. Returns the company folder's item ID, or null if the
// project folder or type folder is not found.
export async function resolveUploadFolderByPath(config, projectFolderPath, typeFolder, companyName) {
  const encodedPath = projectFolderPath.split("/").map(encodeURIComponent).join("/");
  let projectItem;
  try {
    projectItem = await graphRequest(
      config,
      `/sites/${config.siteId}/drives/${config.driveId}/items/root:/${encodedPath}`,
      { method: "GET" }
    );
  } catch {
    return null;
  }

  const typeItem = await getChildFolderByName(config, projectItem.id, typeFolder);
  if (!typeItem) return null;

  const companyItem = await getChildFolderByName(config, typeItem.id, companyName);
  if (companyItem) return companyItem.id;

  const created = await graphRequest(
    config,
    `/sites/${config.siteId}/drives/${config.driveId}/items/${typeItem.id}/children`,
    {
      method: "POST",
      body: JSON.stringify({ name: companyName, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }),
    }
  );
  return created.id;
}

// Uses a direct SharePoint drive item ID (from the project list's FolderID field) to
// navigate into the type subfolder (must already exist), then find or create the company
// subfolder inside it. Returns the company folder's item ID.
export async function resolveUploadFolderByItemId(config, projectItemId, typeFolder, companyName) {
  const typeItem = await getChildFolderByName(config, projectItemId, typeFolder);
  if (!typeItem) {
    throw new Error(`"${typeFolder}" folder not found inside project folder.`);
  }

  const companyItem = await getChildFolderByName(config, typeItem.id, companyName);
  if (companyItem) return companyItem.id;

  const created = await graphRequest(
    config,
    `/sites/${config.siteId}/drives/${config.driveId}/items/${typeItem.id}/children`,
    {
      method: "POST",
      body: JSON.stringify({ name: companyName, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }),
    }
  );
  return created.id;
}

// Same as resolveUploadFolderByPath but locates the project folder via drive search
// instead of a known path. Used when the exact folder path is unknown.
export async function resolveUploadFolderBySearch(config, projectFolderName, typeFolder, companyName) {
  const projectItem = await searchDriveForFolder(config, projectFolderName);
  if (!projectItem) return null;

  const typeItem = await getChildFolderByName(config, projectItem.id, typeFolder);
  if (!typeItem) return null;

  const companyItem = await getChildFolderByName(config, typeItem.id, companyName);
  if (companyItem) return companyItem.id;

  const created = await graphRequest(
    config,
    `/sites/${config.siteId}/drives/${config.driveId}/items/${typeItem.id}/children`,
    {
      method: "POST",
      body: JSON.stringify({ name: companyName, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }),
    }
  );
  return created.id;
}

// Upload directly to a folder by its drive item ID instead of by path.
export async function uploadBytesToFolderById(
  config,
  fileName,
  data,
  folderId,
  contentType = "application/octet-stream"
) {
  if (!config.siteId || !config.driveId) {
    throw new Error("Site ID and Drive ID are required to upload files.");
  }

  const uploadPath = `/sites/${config.siteId}/drives/${config.driveId}/items/${folderId}:/${encodeURIComponent(fileName)}:/content`;
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

  const folder = (folderPath || "").replace(/^\/+|\/+$/g, "");
  await ensureLibraryPath(config, folder);

  const encodedFile = encodeURIComponent(fileName);
  const uploadPath = folder
    ? `/sites/${config.siteId}/drives/${config.driveId}/items/root:/${folder.split("/").map(encodeURIComponent).join("/")}/${encodedFile}:/content`
    : `/sites/${config.siteId}/drives/${config.driveId}/items/root:/${encodedFile}:/content`;
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

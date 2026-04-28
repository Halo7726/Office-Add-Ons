const fromEnv = {
  tenantId: import.meta.env.VITE_TENANT_ID || "common",
  clientId: import.meta.env.VITE_CLIENT_ID || "",
  siteId: import.meta.env.VITE_SITE_ID || "",
  listId: import.meta.env.VITE_LIST_ID || "",
  companyListId: import.meta.env.VITE_COMPANY_LIST_ID || "",
  responseListId: import.meta.env.VITE_RESPONSE_LIST_ID || "",
  renameUploadFiles: import.meta.env.VITE_RENAME_UPLOAD_FILES !== "false",
  driveId: import.meta.env.VITE_DRIVE_ID || "",
  folderTemplate:
    import.meta.env.VITE_FOLDER_TEMPLATE ||
    "Estimating Dashboard/Bids/Current/{project}/Subcontractors/{subcontractor}",
};

const STORAGE_KEY = "sp-outlook-addin-config";

export function loadConfig() {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return fromEnv;
    const parsed = JSON.parse(raw);
    return { ...fromEnv, ...parsed };
  } catch {
    return fromEnv;
  }
}

export function saveConfig(config) {
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(config));
}

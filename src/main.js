import "./styles.css";
import { loadConfig, saveConfig } from "./config";
import {
  addListItem,
  getListItemsById,
  signIn,
  updateListItemFields,
  uploadBytesToLibrary,
  uploadBytesToFolderById,
  resolveUploadFolderByItemId,
  resolveUploadFolderByPath,
} from "./graphClient";
import { getAttachmentBytes, getMessageContext, hasOutlookContext } from "./outlookContext";
import { resolveRoute, resolveItbMatch } from "./routingEngine";

const state = {
  config: loadConfig(),
  account: null,
  messageContext: null,
  projects: [],
  companies: [],
  itbItems: [],
  itbMatch: null,
  route: null,
  selectedCompanyId: "",
  selectedProjectId: "",
  selectedAttachmentIds: new Set(),
  uploadResults: [],
  showSettings: false,
  showNewCompanyForm: false,
  replaceContact1: false,
  setContact2: false,
  statusMessage: "Ready.",
  statusType: "info",
};

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function emailToDisplayName(email) {
  const localPart = String(email || "").split("@")[0] || "";
  const spaced = localPart.replace(/[._-]+/g, " ").trim();
  if (!spaced) return "Unknown Contact";
  return spaced
    .split(" ")
    .filter(Boolean)
    .map((part) => part[0].toUpperCase() + part.slice(1))
    .join(" ");
}

function isBlank(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

function normalizeFieldName(value) {
  return String(value || "")
    .replace(/_x[0-9a-fA-F]{4}_/g, " ")
    .replace(/[^a-zA-Z0-9]+/g, "")
    .toLowerCase();
}

function resolveFirstExistingFieldKey(fields, desiredNames = []) {
  const existingKeys = Object.keys(fields || {});
  const normalizedMap = new Map(
    existingKeys.map((key) => [normalizeFieldName(key), key])
  );

  for (const name of desiredNames) {
    const normalized = normalizeFieldName(name);
    if (normalizedMap.has(normalized)) {
      return normalizedMap.get(normalized);
    }
  }

  return null;
}

function getCompanyRecordById(companyId) {
  if (!companyId) return null;
  return state.companies.find((record) => record.id === companyId) || null;
}

function getProjectDisplayName(record) {
  return record?.fields?.Title || record?.fields?.Project_x0020_Name || "(Unnamed)";
}

function getProjectEstimateStatus(fields) {
  const key = resolveFirstExistingFieldKey(fields, ["EstimateStatus", "Estimate Status"]);
  return key ? String(fields[key] || "").trim() : "";
}

function buildProjectSelectOptions() {
  const activeId = state.selectedProjectId || state.route?.project?.id || "";

  const filtered = state.projects.filter((p) => {
    const status = getProjectEstimateStatus(p.fields || {}).toLowerCase();
    return status === "current" || status === "pending";
  });

  const sorted = filtered.sort((a, b) => {
    const sa = getProjectEstimateStatus(a.fields || {}).toLowerCase();
    const sb = getProjectEstimateStatus(b.fields || {}).toLowerCase();
    if (sa !== sb) return sa === "current" ? -1 : 1;
    return getProjectDisplayName(a).localeCompare(getProjectDisplayName(b));
  });

  return sorted
    .map((p) => {
      const name = getProjectDisplayName(p);
      const selected = p.id === activeId ? "selected" : "";
      return `<option value="${escapeHtml(p.id)}" ${selected}>${escapeHtml(name)}</option>`;
    })
    .join("");
}

function buildFolderPathFromRoute(projectRecord, companyName, typeFolder) {
  const fields = projectRecord?.fields || {};
  const pathKey = resolveFirstExistingFieldKey(fields, PROJECT_FOLDER_PATH_KEYS);
  const rawPath = pathKey ? fields[pathKey] : null;
  const cleanPath = extractDriveRelativePath(rawPath);
  const resolvedType = typeFolder || "Subcontractors";
  const sanitizedCompany = sanitizeFileName(String(companyName || "Unknown").trim());

  if (cleanPath) {
    return [cleanPath, resolvedType, sanitizedCompany].filter(Boolean).join("/");
  }

  // Fall back to template when no FilePath is stored on the project record
  const projectName = getProjectDisplayName(projectRecord);
  const template =
    state.config.folderTemplate || "Estimating Dashboard/Bids/Current/{project}/Subcontractors/{subcontractor}";
  const sanitize = (value, fallback) =>
    String(value || fallback || "Unknown")
      .replace(/[\\/:*?"<>|]/g, "-")
      .replace(/\s+/g, " ")
      .trim();
  return template
    .replaceAll("{project}", sanitize(projectName, "Unmapped Project"))
    .replaceAll("{subcontractor}", sanitize(companyName, "Unmapped Subcontractor"));
}

function sanitizeFileName(value) {
  return String(value || "")
    .trim()
    .replace(/[\\/:*?"<>|]+/g, "-")
    .replace(/\s+/g, " ")
    .replace(/^[-\s]+|[-\s]+$/g, "");
}

function buildUploadFileName(originalName, companyName, projectTitle, attachmentIndex = null) {
  const extensionMatch = originalName.match(/(\.[^.]*)$/);
  const extension = extensionMatch ? extensionMatch[1] : "";
  const dateSegment = new Date().toISOString().slice(0, 10);
  const indexSegment = attachmentIndex ? ` - ${attachmentIndex}` : "";

  return `Proposal - ${dateSegment}${indexSegment}${extension}`;
}

// Field names to check in the project list for the SharePoint drive item ID
const PROJECT_FOLDER_ID_KEYS = ["FolderID"];

// Field names to check in the project list for the SharePoint folder path
const PROJECT_FOLDER_PATH_KEYS = ["FilePath"];

// Field names to check in the ITB item for the owner / GC name
const ITB_OWNER_KEYS = [
  "Owner", "OwnerCompany", "GC", "GeneralContractor", "Client", "ClientName",
];

function normalizeTypeToFolder(typeValue) {
  const t = String(typeValue || "").trim().toLowerCase();
  if (t.startsWith("sub")) return "Subcontractors";
  if (t.startsWith("vend")) return "Vendors";
  // Unknown type: use the raw value or fall back to Subcontractors
  return String(typeValue || "").trim() || "Subcontractors";
}

// Converts a raw folder path value from SharePoint metadata into a drive-relative path.
// Handles full URLs, server-relative URLs, and plain relative paths.
function extractDriveRelativePath(rawPath) {
  if (!rawPath) return null;
  const path = String(rawPath).replace(/\\/g, "/").trim();

  if (path.startsWith("http")) {
    // Full URL — strip up through /{library}/ to get drive-relative portion
    // e.g. https://tenant.sharepoint.com/sites/SiteName/LibraryName/Folder/Sub
    const match = path.match(/\/sites\/[^/]+\/[^/]+\/(.*)/);
    return match ? decodeURIComponent(match[1]) : null;
  }

  if (path.startsWith("/sites/")) {
    const match = path.match(/\/sites\/[^/]+\/[^/]+\/(.*)/);
    return match ? decodeURIComponent(match[1]) : null;
  }

  // Already a relative path (no leading slash)
  if (!path.startsWith("/")) return path;

  // Server-relative path like /Estimating Dashboard/Bids/Current/Project Name — strip leading slash
  return path.slice(1);
}

// Tries up to four strategies to locate the existing project folder, then
// returns the company subfolder's drive item ID (creating it only if needed).
// Returns null if the folder cannot be found — caller falls back to path-based upload.
async function findUploadFolderId(typeFolder, companyName) {
  const projectListItemId = state.route?.project?.id || null;
  const projectRecord = projectListItemId
    ? state.projects.find((r) => r.id === projectListItemId)
    : null;

  const projectFields = projectRecord?.fields || {};

  // Strategy 0 (primary): FolderID field — direct SharePoint item ID, most reliable.
  const folderIdKey = resolveFirstExistingFieldKey(projectFields, PROJECT_FOLDER_ID_KEYS);
  const projectFolderItemId = folderIdKey ? String(projectFields[folderIdKey] || "").trim() : null;
  if (projectFolderItemId) {
    try {
      const id = await resolveUploadFolderByItemId(state.config, projectFolderItemId, typeFolder, companyName);
      if (id) return id;
    } catch { /* fall through */ }
  }

  // Strategy 1: FilePath field — drive-relative path stored in project list metadata.
  const pathKey = resolveFirstExistingFieldKey(projectFields, PROJECT_FOLDER_PATH_KEYS);
  const rawPath = pathKey ? projectFields[pathKey] : null;
  const cleanPath = extractDriveRelativePath(rawPath);
  if (cleanPath) {
    try {
      const id = await resolveUploadFolderByPath(state.config, cleanPath, typeFolder, companyName);
      if (id) return id;
    } catch { /* fall through */ }
  }

  return null;
}

function buildItbFolderPath(projectRecord, typeValue, companyName) {
  return buildFolderPathFromRoute(projectRecord, companyName, normalizeTypeToFolder(typeValue));
}

function buildCompanySelectOptions(route) {
  const candidates = route?.companyCandidates || [];
  const parts = [];

  // If a company was just created and selected but isn't in scored candidates yet, show it first
  if (
    state.selectedCompanyId &&
    state.selectedCompanyId !== "__new__" &&
    !candidates.some((c) => c.id === state.selectedCompanyId)
  ) {
    const rec = getCompanyRecordById(state.selectedCompanyId);
    if (rec) {
      const name = rec.fields?.Title || "New Company";
      parts.push(`<option value="${escapeHtml(state.selectedCompanyId)}" selected>${escapeHtml(name)}</option>`);
    }
  }

  if (candidates.length === 0 && parts.length === 0) {
    parts.push('<option value="">No confident company match</option>');
  }

  for (const candidate of candidates) {
    const selected = !state.showNewCompanyForm && candidate.id === state.selectedCompanyId ? "selected" : "";
    parts.push(
      `<option value="${escapeHtml(candidate.id)}" ${selected}>${escapeHtml(candidate.name)} (${candidate.score})</option>`
    );
  }

  parts.push(`<option value="__new__" ${state.showNewCompanyForm ? "selected" : ""}>— Create new company —</option>`);
  return parts.join("");
}

// Returns the actual SharePoint field key for the first candidate name found in any existing company record.
function inferCompanyFieldKey(desiredNames) {
  for (const company of state.companies) {
    const key = resolveFirstExistingFieldKey(company.fields || {}, desiredNames);
    if (key) return key;
  }
  return desiredNames[0];
}

function buildNewCompanyForm() {
  const fromEmail = state.messageContext?.from || "";
  const contactName = fromEmail ? emailToDisplayName(fromEmail) : "";
  const defaultType = state.itbMatch?.typeValue || "Subcontractor";
  return `
    <div class="new-company-form">
      <h3>New Company</h3>
      <div class="grid">
        <label>Company Name *
          <input id="newCompanyName" placeholder="Enter company name" />
        </label>
        <label>Company Type *
          <div class="multiselect-dropdown" id="companyTypeDropdown">
            <button type="button" class="multiselect-trigger" id="companyTypeBtn">
              <span id="companyTypeDisplay">${defaultType}</span>
              <span class="multiselect-arrow">&#9660;</span>
            </button>
            <div class="multiselect-menu" id="companyTypeMenu">
              ${["Subcontractor","Vendor","Competitor","Engineer/Architect","Customer/Owner"].map(type => `
                <label class="multiselect-item">
                  <input type="checkbox" name="newCompanyType" value="${type}" ${type === defaultType ? "checked" : ""} />
                  ${type}
                </label>`).join("")}
            </div>
          </div>
        </label>
        <label>Contact Name
          <input id="newContactName1" value="${escapeHtml(contactName)}" placeholder="Enter contact name" />
        </label>
        <label>Contact Title
          <input id="newContactTitle1" value="Estimator" />
        </label>
        <label>Email
          <input id="newEmail1" type="email" value="${escapeHtml(fromEmail)}" placeholder="Enter email" />
        </label>
        <label>Mobile
          <input id="newMobile1" placeholder="N/A" />
        </label>
      </div>
      <div class="actions">
        <button id="createCompany" class="btn-primary">Save to SharePoint</button>
        <button id="cancelNewCompany" class="btn-secondary">Cancel</button>
      </div>
    </div>
  `;
}

function buildCompanyContactPatch(companyRecord, fromEmail, opts = {}) {
  const { replaceContact1 = false, setContact2 = false } = opts;
  const fields = companyRecord?.fields || {};
  const patch = {};
  const senderEmail = String(fromEmail || "").trim();
  const senderName = emailToDisplayName(senderEmail);

  const contact1Filled = !isBlank(fields.Contact_x0020_Name_x0020_1);

  if (!contact1Filled || replaceContact1) {
    if (senderName) patch.Contact_x0020_Name_x0020_1 = senderName;
    if (senderEmail) patch.Email_x0020_1 = senderEmail;
    if (!contact1Filled) patch.Contact_x0020_Title = "Estimator Contact";
  }

  if (setContact2) {
    if (senderName) patch.Contact_x0020_Name_x0020_2 = senderName;
    if (senderEmail) patch.Email_x0020_2 = senderEmail;
    patch.Contact_x0020_Title_x0020_2 = "Estimator Contact";
  }

  return patch;
}

function resetContactUpdateFlags() {
  const company = getCompanyRecordById(state.selectedCompanyId);
  const fields = company?.fields || {};
  state.replaceContact1 = isBlank(fields.Contact_x0020_Name_x0020_1);
  state.setContact2 = false;
}

function buildContactUpdatePanel() {
  const companyId = state.selectedCompanyId;
  if (!companyId || companyId === "__new__" || state.showNewCompanyForm) return "";
  const company = getCompanyRecordById(companyId);
  if (!company || !state.messageContext?.from) return "";

  const f = company.fields || {};
  const c1Name  = String(f.Contact_x0020_Name_x0020_1 || "").trim();
  const c1Email = String(f.Email_x0020_1 || "").trim();
  const c2Name  = String(f.Contact_x0020_Name_x0020_2 || "").trim();
  const c2Email = String(f.Email_x0020_2 || "").trim();

  const senderName  = escapeHtml(emailToDisplayName(state.messageContext.from));
  const senderEmail = escapeHtml(state.messageContext.from);

  const c1Current = c1Name
    ? `${escapeHtml(c1Name)}${c1Email ? ` — ${escapeHtml(c1Email)}` : ""}`
    : "<em>empty</em>";
  const c2Current = c2Name
    ? `${escapeHtml(c2Name)}${c2Email ? ` — ${escapeHtml(c2Email)}` : ""}`
    : "<em>empty</em>";

  return `
    <div class="contact-update-panel">
      <span class="label">Update contacts from sender: ${senderName} — ${senderEmail}</span>
      <label class="contact-update-row">
        <input type="checkbox" id="replaceContact1" ${state.replaceContact1 ? "checked" : ""} />
        <span>Contact 1: ${c1Current}${c1Name ? " — <strong>replace</strong>" : " — fill"}</span>
      </label>
      <label class="contact-update-row">
        <input type="checkbox" id="setContact2" ${state.setContact2 ? "checked" : ""} />
        <span>Contact 2: ${c2Current}${c2Name ? " — <strong>replace</strong>" : " — set"}</span>
      </label>
    </div>`;
}

function normalizeText(value) {
  return String(value || "").trim().toLowerCase();
}

function parseProjectTitleFromSubject(subject) {
  if (!subject) return "";
  const cleaned = String(subject).replace(/[–—]/g, "-");
  const match = cleaned.match(/([A-Za-z0-9][^\-\r\n]+?)\s*-\s*([A-Za-z0-9][^\r\n]+)/);
  if (!match) return "";
  return `${match[1].trim()} - ${match[2].trim()}`;
}


function resolveResponsePatchField(fields) {
  return (
    resolveFirstExistingFieldKey(fields, [
      "Status",
      "ResponseStatus",
      "QuoteStatus",
      "ResponseType",
      "Quote",
      "Response",
    ]) || "Status"
  );
}

async function updateItbMatchStatus() {
  if (!state.config.responseListId) {
    return { updated: false, reason: "ITB/RFQ list is not configured." };
  }

  // Use the pre-matched ITB item from auto-match when available.
  // If auto-match wasn't run yet, search the pre-loaded items now.
  let item = state.itbMatch?.item || null;

  if (!item && state.itbItems.length > 0) {
    const match = resolveItbMatch(state.messageContext, state.itbItems, state.config);
    if (match) {
      state.itbMatch = match;
      item = match.item;
    }
  }

  if (!item) {
    const fromEmail = state.messageContext?.from || "";
    return {
      updated: false,
      reason: `No ITB/RFQ row found matching sender "${fromEmail}".`,
    };
  }

  const fields = item.fields || {};
  const statusKey = resolveResponsePatchField(fields);
  const currentValue = normalizeText(fields[statusKey]);

  if (currentValue === "quoted") {
    return {
      updated: false,
      reason: `ITB/RFQ item "${state.itbMatch?.projectName || item.id}" is already marked as Quoted.`,
    };
  }

  const patch = { [statusKey]: "Quoted" };
  await updateListItemFields(state.config, state.config.responseListId, item.id, patch);
  item.fields = { ...fields, ...patch };

  return {
    updated: true,
    reason: `Marked ITB/RFQ item "${state.itbMatch?.projectName || item.id}" as Quoted.`,
  };
}

function createApp() {
  const context = state.messageContext || {
    subject: "",
    from: "",
    attachments: [],
  };
  const route = state.route || {
    project: { name: "Not matched", score: 0 },
    subcontractor: { name: "Not matched", score: 0 },
    companyCandidates: [],
    confidence: 0,
    folderPath: state.config.folderTemplate || "",
    reason: "Run auto-match to generate route.",
  };

  const companyOptions = buildCompanySelectOptions(route);

  const attachmentRows = context.attachments.length
    ? context.attachments
        .map((attachment) => {
          const checked = state.selectedAttachmentIds.has(attachment.id) ? "checked" : "";
          const disabled = attachment.isInline ? "disabled" : "";
          return `
            <label class="attachment-row">
              <input type="checkbox" data-attachment-id="${attachment.id}" ${checked} ${disabled} />
              <span class="attachment-name">${attachment.name}</span>
              <span class="attachment-size">${Math.round((attachment.size || 0) / 1024)} KB</span>
            </label>
          `;
        })
        .join("")
    : '<p class="muted">No Outlook attachments detected. You can still upload a local file.</p>';

  return `
    <main class="app-shell fluent-root">
      <header class="topbar">
        <div>
          <h1>SHC Project Upload</h1>
          <p>Match this email to project routing and upload selected attachments.</p>
        </div>
        <div class="topbar-actions">
          <button id="toggleSettings" class="btn-secondary">${state.showSettings ? "Hide settings" : "Settings"}</button>
          <button id="refreshContext" class="btn-secondary">Refresh email context</button>
        </div>
      </header>
      <div id="statusBanner" class="status-banner ${state.statusType || "info"}">${escapeHtml(
        state.statusMessage || ""
      )}</div>

      <section class="panel context-panel">
        <h2>Email Context</h2>
        <div class="context-grid">
          <div>
            <span class="label">From</span>
            <p>${context.from || "Not available"}</p>
          </div>
          <div>
            <span class="label">Subject</span>
            <p>${context.subject || "Not available"}</p>
          </div>
          <div>
            <span class="label">Attachments</span>
            <p>${context.attachments.length}</p>
          </div>
        </div>
      </section>

      <section class="panel match-panel">
        <h2>Routing Match</h2>
        <div class="match-grid">
          <article class="match-card">
            <span class="label">Project</span>
            <strong>${route.project.name}</strong>
            <small>Score ${route.project.score}</small>
          </article>
          <article class="match-card">
            <span class="label">Subcontractor</span>
            <strong>${route.subcontractor.name}</strong>
            <small>Score ${route.subcontractor.score}</small>
          </article>
          <article class="match-card confidence ${route.confidence >= 70 ? "high" : "low"}">
            <span class="label">Confidence</span>
            <strong>${route.confidence}%</strong>
            <small>${route.reason}</small>
          </article>
        </div>
        ${(() => {
          const noMatch = state.projects.length > 0 && !route.project.id;
          if (noMatch) {
            const pathPreview = state.selectedProjectId ? (route.folderPath || "") : "";
            return `
              <label>Project
                <select id="projectSelect">
                  <option value="">— Select a project —</option>
                  ${buildProjectSelectOptions()}
                </select>
              </label>
              <label>Destination folder
                <input id="resolvedFolderPath" readonly value="${escapeHtml(pathPreview)}" placeholder="Select a project above" />
              </label>`;
          }
          return `
            <label>Destination folder (SharePoint path)
              <input id="resolvedFolderPath" value="${escapeHtml(route.folderPath || "")}" />
            </label>`;
        })()}
        <label>Subcontractor match options
          <select id="companySelect">${companyOptions}</select>
        </label>
        ${state.showNewCompanyForm ? buildNewCompanyForm() : buildContactUpdatePanel()}
        <div class="actions">
          <button id="signin" class="btn-secondary">Sign in</button>
          <button id="autoMatch" class="btn-primary">Auto-match</button>
        </div>
        <p id="accountLabel" class="hint">${
          state.account ? `Signed in as ${state.account.username}` : "Not signed in."
        }</p>
      </section>

      <section class="panel attachments-panel">
        <h2>Attachments</h2>
        <div class="attachment-list">${attachmentRows}</div>
        <div class="actions">
          <button id="uploadSelected" class="btn-primary">Upload selected</button>
        </div>
        <pre id="fileOutput" class="output hidden-panel" role="log" aria-live="polite">${JSON.stringify(
          state.uploadResults,
          null,
          2
        )}</pre>
      </section>

      <section class="panel slim ${state.showSettings ? "" : "hidden-panel"}" id="settingsPanel">
        <h2>Settings</h2>
        <div class="grid">
          <label>Tenant ID
            <input id="tenantId" value="${state.config.tenantId || "common"}" />
          </label>
          <label>Client ID
            <input id="clientId" value="${state.config.clientId || ""}" />
          </label>
          <label>Site ID
            <input id="siteId" value="${state.config.siteId || ""}" />
          </label>
          <label>List ID
            <input id="listId" value="${state.config.listId || ""}" />
          </label>
          <label>Company List ID
            <input id="companyListId" value="${state.config.companyListId || ""}" />
          </label>
          <label>Response list ID
            <input id="responseListId" value="${state.config.responseListId || ""}" />
          </label>
          <label>
            <span class="label">Rename upload files to shorter names</span>
            <input id="renameUploadFiles" type="checkbox" ${state.config.renameUploadFiles ? "checked" : ""} />
          </label>
          <label>Drive ID
            <input id="driveId" value="${state.config.driveId || ""}" />
          </label>
          <label>Folder template
            <input id="folderTemplate" value="${
              state.config.folderTemplate || "Estimating Dashboard/Bids/Current/{project}/Subcontractors/{subcontractor}"
            }" />
          </label>
        </div>
        <div class="actions">
          <button id="saveConfig" class="btn-secondary">Save settings</button>
        </div>
      </section>

      <section class="panel slim hidden-panel">
        <h2>Create / Update Record</h2>
        <label>List item fields (JSON)
          <textarea id="listFields" rows="5">{
  "Title": "Email Upload",
  "Status": "Received"
}</textarea>
        </label>
        <div class="actions">
          <button id="addListItem" class="btn-primary">Write list record</button>
        </div>
        <pre id="listOutput" class="output hidden-panel" role="log" aria-live="polite"></pre>
      </section>

      <section class="panel slim hidden-panel">
        <h2>Execution log</h2>
        <pre id="status" class="output" role="status" aria-live="polite"></pre>
      </section>
    </main>
  `;
}

function syncConfigFromUI() {
  state.config = {
    tenantId: document.getElementById("tenantId").value.trim() || "common",
    clientId: document.getElementById("clientId").value.trim(),
    siteId: document.getElementById("siteId").value.trim(),
    listId: document.getElementById("listId").value.trim(),
    companyListId: document.getElementById("companyListId").value.trim(),
    responseListId: document.getElementById("responseListId").value.trim(),
    renameUploadFiles: document.getElementById("renameUploadFiles").checked,
    driveId: document.getElementById("driveId").value.trim(),
    folderTemplate: document.getElementById("folderTemplate").value.trim(),
  };
}

function setStatus(message, type = "info") {
  state.statusMessage = message;
  state.statusType = type;
  const banner = document.getElementById("statusBanner");
  if (banner) {
    banner.textContent = message;
    banner.className = `status-banner ${type}`;
  }
}

function setOutput(elementId, payload) {
  document.getElementById(elementId).textContent =
    typeof payload === "string" ? payload : JSON.stringify(payload, null, 2);
}

function render() {
  const app = document.getElementById("app");
  app.innerHTML = createApp();
  wireEvents();
}

async function loadMessageContext() {
  state.messageContext = getMessageContext();
  state.selectedAttachmentIds = new Set(
    state.messageContext.attachments.filter((item) => !item.isInline).map((item) => item.id)
  );
}

async function loadRoutingLists() {
  const projects = await getListItemsById(state.config, state.config.listId);
  state.projects = projects.value || [];

  if (state.config.companyListId) {
    const companies = await getListItemsById(state.config, state.config.companyListId);
    state.companies = companies.value || [];
  } else {
    state.companies = [];
  }

  if (state.config.responseListId) {
    const itbItems = await getListItemsById(state.config, state.config.responseListId);
    state.itbItems = itbItems.value || [];
  } else {
    state.itbItems = [];
  }
}

async function runAutoMatch() {
  syncConfigFromUI();

  if (!state.messageContext) {
    await loadMessageContext();
  }

  await loadRoutingLists();

  // Find the ITB/RFQ row matching this email's sender (Recipient Email) + subject (Title)
  state.itbMatch = resolveItbMatch(state.messageContext, state.itbItems, state.config);

  // Standard routing for project list + company list (drives the subcontractor dropdown)
  state.route = resolveRoute(state.messageContext, state.projects, state.companies, state.config);
  state.selectedCompanyId =
    state.route.subcontractor.id || state.route.companyCandidates[0]?.id || "";
  resetContactUpdateFlags();

  if (state.selectedCompanyId) {
    state.route = resolveRoute(state.messageContext, state.projects, state.companies, state.config, {
      companyIdOverride: state.selectedCompanyId,
    });
  }

  const matchedProjectRec = state.route.project.id
    ? state.projects.find((r) => r.id === state.route.project.id) || null
    : null;

  if (state.itbMatch) {
    state.route.project = {
      id: state.route.project.id,
      name: state.itbMatch.projectName,
      score: state.itbMatch.emailScore + state.itbMatch.titleScore,
    };
    state.route.folderPath = buildItbFolderPath(
      matchedProjectRec,
      state.itbMatch.typeValue,
      state.route.subcontractor.name
    );
    state.route.confidence = state.itbMatch.confidence;
    state.route.reason = `ITB/RFQ matched on recipient email + title (${state.itbMatch.confidence}% confidence)`;
  } else if (matchedProjectRec) {
    const companyRec = getCompanyRecordById(state.selectedCompanyId);
    const typeRaw = companyRec?.fields?.Company_x0020_Type;
    const typeValue = (Array.isArray(typeRaw) ? typeRaw[0] : typeRaw) || "Subcontractor";
    state.route.folderPath = buildFolderPathFromRoute(
      matchedProjectRec,
      state.route.subcontractor.name,
      normalizeTypeToFolder(typeValue)
    );
  }
}

function applySelectedCompanyOverride() {
  if (!state.route) return;

  state.route = resolveRoute(
    state.messageContext,
    state.projects,
    state.companies,
    state.config,
    state.selectedCompanyId ? { companyIdOverride: state.selectedCompanyId } : undefined
  );

  const projectId = state.route.project.id || state.selectedProjectId || null;
  const projectRec = projectId ? state.projects.find((r) => r.id === projectId) || null : null;

  if (state.itbMatch) {
    state.route.project = {
      id: state.route.project.id,
      name: state.itbMatch.projectName,
      score: state.itbMatch.emailScore + state.itbMatch.titleScore,
    };
    state.route.folderPath = buildItbFolderPath(
      projectRec,
      state.itbMatch.typeValue,
      state.route.subcontractor.name
    );
    return;
  }

  if (projectRec) {
    const name = getProjectDisplayName(projectRec);
    state.route.project = { id: projectId, name, score: state.route.project.score };
    const companyRec = getCompanyRecordById(state.selectedCompanyId);
    const typeRaw = companyRec?.fields?.Company_x0020_Type;
    const typeValue = (Array.isArray(typeRaw) ? typeRaw[0] : typeRaw) || "Subcontractor";
    state.route.folderPath = buildFolderPathFromRoute(
      projectRec,
      state.route.subcontractor.name,
      normalizeTypeToFolder(typeValue)
    );
  }
}

async function patchSelectedCompanyContactsIfMissing() {
  if (!state.config.companyListId) {
    return { updated: false, reason: "Company List ID is not configured." };
  }

  const company = getCompanyRecordById(state.selectedCompanyId || state.route?.subcontractor?.id);
  if (!company) {
    return { updated: false, reason: "No matched company selected." };
  }

  const patch = buildCompanyContactPatch(company, state.messageContext?.from || "", {
    replaceContact1: state.replaceContact1,
    setContact2: state.setContact2,
  });
  if (Object.keys(patch).length === 0) {
    return { updated: false, reason: "Selected company already has contact fields populated." };
  }

  await updateListItemFields(state.config, state.config.companyListId, company.id, patch);
  company.fields = { ...company.fields, ...patch };

  return {
    updated: true,
    reason: `Updated missing contacts for ${company.fields?.Title || "selected company"}.`,
    patch,
  };
}

async function createNewCompany() {
  if (!state.config.companyListId) {
    throw new Error("Company List ID is not configured in settings.");
  }

  // Read all form values before any awaits — an async gap can trigger a render that
  // rebuilds the form from the original state, wiping any edits the user made.
  const name = document.getElementById("newCompanyName").value.trim();
  if (!name) throw new Error("Company name is required.");

  const companyTypeSelected = Array.from(document.querySelectorAll('input[name="newCompanyType"]:checked')).map(o => o.value);
  const patch = {
    "Company_x0020_Type@odata.type": "Collection(Edm.String)",
    Company_x0020_Type:          companyTypeSelected,
    Contact_x0020_Name_x0020_1:  document.getElementById("newContactName1").value.trim() || "N/A",
    Contact_x0020_Title:         document.getElementById("newContactTitle1").value.trim() || "N/A",
    Email_x0020_1:               document.getElementById("newEmail1").value.trim() || "N/A",
    Mobile_x0020_Number_x0020_1: document.getElementById("newMobile1").value.trim() || "N/A",
  };

  // POST with Title only — the /items endpoint is finicky with multi-select choice fields.
  // All other fields are set via a follow-up PATCH using the /fields endpoint.
  const result = await addListItem({ ...state.config, listId: state.config.companyListId }, { Title: name });
  await updateListItemFields(state.config, state.config.companyListId, result.id, patch);

  const newRecord = { id: result.id, fields: { Title: name, ...patch } };
  state.companies.push(newRecord);
  state.selectedCompanyId = result.id;
  state.showNewCompanyForm = false;

  if (state.route) {
    applySelectedCompanyOverride();
  }
}

async function uploadSelected() {
  syncConfigFromUI();

  if (!hasOutlookContext()) {
    throw new Error("Open this add-in from an Outlook message to access attachments.");
  }

  if (!state.messageContext) {
    await loadMessageContext();
  }

  if (!state.route) {
    await runAutoMatch();
  }

  const destination =
    document.getElementById("resolvedFolderPath").value.trim() ||
    state.route?.folderPath;

  if (!destination) {
    throw new Error("Destination folder is required. Run Auto-match or enter a folder path.");
  }

  const selectedIds = [...state.selectedAttachmentIds];
  const selected = state.messageContext.attachments.filter((item) => selectedIds.includes(item.id));

  if (selected.length === 0) {
    throw new Error("Select at least one attachment before uploading.");
  }

  const projectTitle =
    parseProjectTitleFromSubject(state.messageContext?.subject || "") ||
    state.route?.project?.name;
  const companyName = state.route?.subcontractor?.name || "Company";

  // Locate the existing project folder via the project record's FilePath or FolderID,
  // then navigate into the company type subfolder and find/create the company folder.
  let uploadFolderId = null;
  if (state.route?.project?.id) {
    const companyRecord = getCompanyRecordById(state.selectedCompanyId || state.route?.subcontractor?.id);
    const companyTypeRaw = companyRecord?.fields?.Company_x0020_Type;
    const typeValue = (Array.isArray(companyTypeRaw) ? companyTypeRaw[0] : companyTypeRaw)
      || state.itbMatch?.typeValue
      || "Subcontractor";
    const typeFolder = normalizeTypeToFolder(typeValue);
    const sanitizedCompany = sanitizeFileName(companyName);
    setStatus("Locating project folder in SharePoint…", "info");
    uploadFolderId = await findUploadFolderId(typeFolder, sanitizedCompany);
    if (!uploadFolderId) {
      setStatus("Existing folder not found — uploading to configured template path.", "info");
    }
  }

  // Upload files — primary action, must succeed before anything else runs
  const uploads = [];

  for (let index = 0; index < selected.length; index += 1) {
    const attachment = selected[index];
    const bytes = await getAttachmentBytes(attachment.id);
    const pathTooLong = destination.length + 1 + attachment.name.length > 260;
    const uploadName = pathTooLong
      ? buildUploadFileName(attachment.name, companyName, projectTitle, selected.length > 1 ? index + 1 : null)
      : attachment.name;

    const targetPath = uploadFolderId ? null : destination;
    const response = uploadFolderId
      ? await uploadBytesToFolderById(state.config, uploadName, bytes, uploadFolderId, "application/octet-stream")
      : await uploadBytesToLibrary(state.config, uploadName, bytes, targetPath, "application/octet-stream");

    uploads.push({
      source: "outlook",
      fileName: uploadName,
      webUrl: response.webUrl,
    });
  }

  state.uploadResults = uploads;

  const outputEl = document.getElementById("fileOutput");
  if (outputEl) {
    outputEl.textContent = JSON.stringify(uploads, null, 2);
    outputEl.classList.remove("hidden-panel");
  }

  // Side effects run after the upload so they never block it.
  // Failures are reported in the status banner but don't undo the upload.
  const notes = [];

  try {
    const contactResult = await patchSelectedCompanyContactsIfMissing();
    if (contactResult.updated) notes.push(contactResult.reason);
  } catch (err) {
    notes.push(`Contact backfill skipped: ${err.message}`);
  }

  try {
    const responseResult = await updateItbMatchStatus();
    if (responseResult.updated) notes.push(responseResult.reason);
  } catch (err) {
    notes.push(`ITB status update skipped: ${err.message}`);
  }

  setStatus(
    [`${uploads.length} file(s) uploaded to ${destination}.`, ...notes].join(" "),
    "success"
  );
}

async function withBusy(label, action) {
  try {
    setStatus(`${label}...`, "info");
    await action();
    setStatus(`${label} complete.`, "success");
  } catch (error) {
    setStatus(`${label} failed: ${error.message}`, "error");
  }
}

function wireEvents() {
  document.getElementById("saveConfig").addEventListener("click", () => {
    syncConfigFromUI();
    saveConfig(state.config);
    setStatus("Settings saved locally.");
  });

  document.getElementById("signin").addEventListener("click", async () => {
    await withBusy("Sign in", async () => {
      syncConfigFromUI();
      const account = await signIn(state.config);
      state.account = account;
      document.getElementById("accountLabel").textContent = `Signed in as ${account.username}`;
    });
  });

  document.getElementById("refreshContext").addEventListener("click", async () => {
    await withBusy("Refresh context", async () => {
      await loadMessageContext();
      render();
    });
  });

  document.getElementById("toggleSettings").addEventListener("click", () => {
    state.showSettings = !state.showSettings;
    render();
  });

  document.getElementById("autoMatch").addEventListener("click", async () => {
    await withBusy("Auto-match", async () => {
      syncConfigFromUI();
      await runAutoMatch();
      render();
    });
  });

  const projectSelectEl = document.getElementById("projectSelect");
  if (projectSelectEl) {
    projectSelectEl.addEventListener("change", (event) => {
      const projectId = event.target.value;
      state.selectedProjectId = projectId;

      if (!projectId || !state.route) return;

      const rec = state.projects.find((p) => p.id === projectId);
      if (!rec) return;

      const name = getProjectDisplayName(rec);
      state.route.project = { id: projectId, name, score: 0 };
      const companyRec = getCompanyRecordById(state.selectedCompanyId);
      const typeRaw = companyRec?.fields?.Company_x0020_Type;
      const typeValue = (Array.isArray(typeRaw) ? typeRaw[0] : typeRaw) || "Subcontractor";
      state.route.folderPath = buildFolderPathFromRoute(rec, state.route.subcontractor.name, normalizeTypeToFolder(typeValue));
      document.getElementById("resolvedFolderPath").value = state.route.folderPath;
    });
  }

  document.getElementById("companySelect").addEventListener("change", (event) => {
    const value = event.target.value;
    if (value === "__new__") {
      state.showNewCompanyForm = true;
      state.selectedCompanyId = "__new__";
      render();
      return;
    }
    state.showNewCompanyForm = false;
    state.selectedCompanyId = value;
    resetContactUpdateFlags();
    applySelectedCompanyOverride();
    document.getElementById("resolvedFolderPath").value = state.route?.folderPath || "";
  });

  const rc1 = document.getElementById("replaceContact1");
  if (rc1) rc1.addEventListener("change", (e) => { state.replaceContact1 = e.target.checked; });

  const sc2 = document.getElementById("setContact2");
  if (sc2) sc2.addEventListener("change", (e) => { state.setContact2 = e.target.checked; });

  if (state.showNewCompanyForm) {
    const companyTypeBtn = document.getElementById("companyTypeBtn");
    const companyTypeMenu = document.getElementById("companyTypeMenu");
    const companyTypeDisplay = document.getElementById("companyTypeDisplay");

    companyTypeBtn.addEventListener("click", () => {
      companyTypeMenu.classList.toggle("open");
    });

    document.addEventListener("click", (e) => {
      if (!document.getElementById("companyTypeDropdown")?.contains(e.target)) {
        companyTypeMenu.classList.remove("open");
      }
    }, { capture: true });

    companyTypeMenu.addEventListener("change", () => {
      const checked = Array.from(document.querySelectorAll('input[name="newCompanyType"]:checked')).map(o => o.value);
      companyTypeDisplay.textContent = checked.length ? checked.join(", ") : "Select...";
    });

    document.getElementById("createCompany").addEventListener("click", async () => {
      await withBusy("Create company", async () => {
        await createNewCompany();
        render();
      });
    });

    document.getElementById("cancelNewCompany").addEventListener("click", () => {
      state.showNewCompanyForm = false;
      state.selectedCompanyId = state.route?.subcontractor?.id || "";
      render();
    });
  }

  document.getElementById("addListItem").addEventListener("click", async () => {
    await withBusy("Write list record", async () => {
      syncConfigFromUI();
      const raw = document.getElementById("listFields").value;
      const fields = JSON.parse(raw);
      const result = await addListItem(state.config, fields);
      setOutput("listOutput", result);
    });
  });

  document.getElementById("uploadSelected").addEventListener("click", async () => {
    await withBusy("Upload selected", async () => {
      syncConfigFromUI();
      await uploadSelected();
    });
  });

  document.querySelectorAll("input[data-attachment-id]").forEach((input) => {
    input.addEventListener("change", (event) => {
      const id = event.target.getAttribute("data-attachment-id");
      if (event.target.checked) {
        state.selectedAttachmentIds.add(id);
      } else {
        state.selectedAttachmentIds.delete(id);
      }
    });
  });
}

async function bootstrap() {
  if (hasOutlookContext()) {
    await loadMessageContext();
  } else {
    state.messageContext = {
      subject: "Open this add-in from an Outlook message to read sender and attachments.",
      from: "",
      senderDomain: "",
      attachments: [],
    };
  }

  render();
}

if (window.Office) {
  Office.onReady(() => {
    bootstrap();
  });
} else {
  bootstrap();
}

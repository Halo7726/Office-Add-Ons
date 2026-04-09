import "./styles.css";
import { loadConfig, saveConfig } from "./config";
import {
  addListItem,
  getListItemsById,
  signIn,
  updateListItemFields,
  uploadBytesToLibrary,
  uploadToLibrary,
} from "./graphClient";
import { getAttachmentBytes, getMessageContext, hasOutlookContext } from "./outlookContext";
import { resolveRoute } from "./routingEngine";

const state = {
  config: loadConfig(),
  account: null,
  messageContext: null,
  projects: [],
  companies: [],
  route: null,
  selectedCompanyId: "",
  selectedAttachmentIds: new Set(),
  uploadResults: [],
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

function buildFolderPathFromRouteTemplate(projectName, companyName) {
  const template =
    state.config.folderTemplate || "Bids/Current/{project}/Subcontractors/{subcontractor}";

  const sanitize = (value, fallback) =>
    String(value || fallback || "Unknown")
      .replace(/[\\/:*?"<>|]/g, "-")
      .replace(/\s+/g, " ")
      .trim();

  return template
    .replaceAll("{project}", sanitize(projectName, "Unmapped Project"))
    .replaceAll("{subcontractor}", sanitize(companyName, "Unmapped Subcontractor"));
}

function buildCompanySelectOptions(route) {
  const candidates = route?.companyCandidates || [];

  if (candidates.length === 0) {
    return '<option value="">No confident company match</option>';
  }

  return candidates
    .map((candidate) => {
      const selected = candidate.id === state.selectedCompanyId ? "selected" : "";
      return `<option value="${escapeHtml(candidate.id)}" ${selected}>${escapeHtml(
        candidate.name
      )} (${candidate.score})</option>`;
    })
    .join("");
}

function buildCompanyContactPatch(companyRecord, fromEmail) {
  const fields = companyRecord?.fields || {};
  const patch = {};
  const senderEmail = String(fromEmail || "").trim();
  const senderName = emailToDisplayName(senderEmail);

  const fieldGroups = {
    contactName1: ["Contact Name 1", "ContactName1", "Primary Contact", "PrimaryContact"],
    contactTitle1: ["Contact Title 1", "ContactTitle1"],
    directLine1: ["Extension/Direct Line 1", "ExtensionDirectLine1"],
    mobile1: ["Mobile Number 1", "MobileNumber1"],
    email1: ["Email 1", "Email1", "Primary Contact Email", "PrimaryContactEmail"],
    contactName2: ["Contact Name 2", "ContactName2", "Secondary Contact", "SecondaryContact"],
    contactTitle2: ["Contact Title 2", "ContactTitle2"],
    directLine2: ["Extension/Direct Line 2", "ExtensionDirectLine2"],
    mobile2: ["Mobile Phone 2", "MobilePhone2", "Mobile Number 2", "MobileNumber2"],
    email2: ["Email 2", "Email2", "Secondary Contact Email", "SecondaryContactEmail"],
  };

  const setIfMissing = (groupName, value) => {
    const key = resolveFirstExistingFieldKey(fields, fieldGroups[groupName]);
    if (key && isBlank(fields[key])) {
      patch[key] = value;
      return true;
    }
    return false;
  };

  const filledPrimary = setIfMissing("contactName1", senderName);
  if (filledPrimary) {
    setIfMissing("contactTitle1", "Estimator Contact");
  }

  if (senderEmail) {
    setIfMissing("email1", senderEmail);
    setIfMissing("email2", senderEmail);
  }

  setIfMissing("contactName2", "Pending Contact");
  setIfMissing("contactTitle2", "Pending");
  setIfMissing("directLine1", "N/A");
  setIfMissing("directLine2", "N/A");
  setIfMissing("mobile1", "N/A");
  setIfMissing("mobile2", "N/A");

  return patch;
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
        <button id="refreshContext" class="btn-secondary">Refresh email context</button>
      </header>

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
        <label>Destination folder (SharePoint path)
          <input id="resolvedFolderPath" value="${route.folderPath || ""}" />
        </label>
        <label>Subcontractor match options
          <select id="companySelect">${companyOptions}</select>
        </label>
        <div class="actions">
          <button id="signin" class="btn-secondary">Sign in</button>
          <button id="autoMatch" class="btn-primary">Auto-match</button>
          <button id="patchContacts" class="btn-secondary">Patch missing contacts</button>
          <button id="saveConfig" class="btn-secondary">Save settings</button>
        </div>
        <p id="accountLabel" class="hint">${
          state.account ? `Signed in as ${state.account.username}` : "Not signed in."
        }</p>
      </section>

      <section class="panel attachments-panel">
        <h2>Attachments</h2>
        <div class="attachment-list">${attachmentRows}</div>
        <label>Fallback local file
          <input id="fileInput" type="file" />
        </label>
        <div class="actions">
          <button id="uploadSelected" class="btn-primary">Upload selected</button>
        </div>
        <pre id="fileOutput" class="output" role="log" aria-live="polite">${JSON.stringify(
          state.uploadResults,
          null,
          2
        )}</pre>
      </section>

      <section class="panel slim">
        <h2>Admin Settings</h2>
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
          <label>Drive ID
            <input id="driveId" value="${state.config.driveId || ""}" />
          </label>
          <label>Default library folder
            <input id="libraryFolder" value="${state.config.libraryFolder || "Shared Documents"}" />
          </label>
          <label>Folder template
            <input id="folderTemplate" value="${
              state.config.folderTemplate || "Bids/Current/{project}/Subcontractors/{subcontractor}"
            }" />
          </label>
        </div>
      </section>

      <section class="panel slim">
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
        <pre id="listOutput" class="output" role="log" aria-live="polite"></pre>
      </section>

      <section class="panel slim">
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
    driveId: document.getElementById("driveId").value.trim(),
    libraryFolder: document.getElementById("libraryFolder").value.trim(),
    folderTemplate: document.getElementById("folderTemplate").value.trim(),
  };
}

function setStatus(message) {
  document.getElementById("status").textContent = message;
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
}

async function runAutoMatch() {
  if (!state.messageContext) {
    await loadMessageContext();
  }

  await loadRoutingLists();
  state.route = resolveRoute(state.messageContext, state.projects, state.companies, state.config);
  state.selectedCompanyId =
    state.route.subcontractor.id || state.route.companyCandidates[0]?.id || "";

  if (state.selectedCompanyId) {
    state.route = resolveRoute(state.messageContext, state.projects, state.companies, state.config, {
      companyIdOverride: state.selectedCompanyId,
    });
  }
}

function applySelectedCompanyOverride() {
  if (!state.route) return;

  if (!state.selectedCompanyId) {
    state.route = resolveRoute(state.messageContext, state.projects, state.companies, state.config);
    return;
  }

  state.route = resolveRoute(state.messageContext, state.projects, state.companies, state.config, {
    companyIdOverride: state.selectedCompanyId,
  });
}

async function patchSelectedCompanyContactsIfMissing() {
  if (!state.config.companyListId) {
    return { updated: false, reason: "Company List ID is not configured." };
  }

  const company = getCompanyRecordById(state.selectedCompanyId || state.route?.subcontractor?.id);
  if (!company) {
    return { updated: false, reason: "No matched company selected." };
  }

  const patch = buildCompanyContactPatch(company, state.messageContext?.from || "");
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

async function uploadSelected() {
  const destination =
    document.getElementById("resolvedFolderPath").value.trim() ||
    state.route?.folderPath ||
    state.config.libraryFolder;

  if (!destination) {
    throw new Error("Destination folder is required.");
  }

  const patchResult = await patchSelectedCompanyContactsIfMissing();

  const uploads = [];

  if (hasOutlookContext()) {
    const selectedIds = [...state.selectedAttachmentIds];
    const selected = state.messageContext.attachments.filter((item) => selectedIds.includes(item.id));

    for (const attachment of selected) {
      const bytes = await getAttachmentBytes(attachment.id);
      const response = await uploadBytesToLibrary(
        state.config,
        attachment.name,
        bytes,
        destination,
        "application/octet-stream"
      );

      uploads.push({
        source: "outlook",
        fileName: attachment.name,
        webUrl: response.webUrl,
      });
    }
  }

  const fileInput = document.getElementById("fileInput");
  const localFile = fileInput.files[0];
  if (localFile) {
    const response = await uploadToLibrary({ ...state.config, libraryFolder: destination }, localFile);
    uploads.push({
      source: "local",
      fileName: localFile.name,
      webUrl: response.webUrl,
    });
  }

  if (uploads.length === 0) {
    throw new Error("Select at least one Outlook attachment or local file.");
  }

  if (patchResult.updated) {
    uploads.unshift({
      source: "company-list",
      fileName: "contact backfill",
      webUrl: patchResult.reason,
    });
  }

  state.uploadResults = uploads;
  setOutput("fileOutput", uploads);
}

async function withBusy(label, action) {
  try {
    setStatus(`${label}...`);
    await action();
    setStatus(`${label} complete.`);
  } catch (error) {
    setStatus(`${label} failed: ${error.message}`);
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

  document.getElementById("autoMatch").addEventListener("click", async () => {
    await withBusy("Auto-match", async () => {
      syncConfigFromUI();
      await runAutoMatch();
      render();
    });
  });

  document.getElementById("companySelect").addEventListener("change", (event) => {
    state.selectedCompanyId = event.target.value;
    applySelectedCompanyOverride();
    const folderInput = document.getElementById("resolvedFolderPath");
    folderInput.value = buildFolderPathFromRouteTemplate(
      state.route.project.name,
      state.route.subcontractor.name
    );
  });

  document.getElementById("patchContacts").addEventListener("click", async () => {
    await withBusy("Patch contacts", async () => {
      syncConfigFromUI();
      const result = await patchSelectedCompanyContactsIfMissing();
      setStatus(result.reason);
    });
  });

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

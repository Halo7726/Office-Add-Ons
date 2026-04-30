import "./styles.css";
import { loadConfig } from "./config";
import { signIn, getListItemsById } from "./graphClient";
import { loadTemplates, saveTemplates, generateKey } from "./template-store";

const state = {
  config: loadConfig(),
  account: null,
  projects: [],
  companies: [],
  templates: loadTemplates(),
  selectedProjectId: "",
  selectedCompanyId: "",
  recipientType: "",       // set to first template key on init
  showTemplateManager: false,
  editingTemplate: null,   // { key, label, subject, body, _isNew } or null
  statusMessage: "Sign in to load projects and companies.",
  statusType: "info",
};

// Initialize recipientType to the first loaded template.
if (state.templates.length > 0 && !state.recipientType) {
  state.recipientType = state.templates[0].key;
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

// Resolves a field value tolerating SharePoint _x-encoding, casing, and extra spaces.
function resolveField(fields, candidates) {
  const keys = Object.keys(fields || {});
  const map = new Map(
    keys.map((k) => [k.replace(/_x[0-9a-fA-F]{4}_/g, " ").replace(/[^a-zA-Z0-9]+/g, "").toLowerCase(), k])
  );
  for (const c of candidates) {
    const n = c.replace(/[^a-zA-Z0-9]+/g, "").toLowerCase();
    if (map.has(n)) return String(fields[map.get(n)] || "");
  }
  return "";
}

function projectDisplayName(fields) {
  return fields.Title || fields.ProjectName || fields.ProjectCode || fields.Name || "(Unnamed)";
}

function companyDisplayName(fields) {
  return fields.Title || fields.CompanyName || fields.Name || "(Unnamed)";
}

function buildTokenMap(project, company) {
  const pf = project?.fields || {};
  const cf = company?.fields || {};
  let senderName = "", senderEmail = "";
  try {
    const profile = Office.context.mailbox.userProfile;
    senderName = profile.displayName || "";
    senderEmail = profile.emailAddress || "";
  } catch { /* dev mode */ }
  return {
    project_name:   projectDisplayName(pf),
    project_number: pf.ProjectCode || pf.ProjectNumber || pf.ProjectNo || pf.Number || "",
    company_name:   companyDisplayName(cf),
    contact_name:   resolveField(cf, ["Contact Name 1", "ContactName1", "Primary Contact", "PrimaryContact"]) || companyDisplayName(cf),
    contact_email:  resolveField(cf, ["Email 1", "Email1", "Primary Contact Email", "PrimaryContactEmail"]),
    contact_title:  resolveField(cf, ["Contact Title 1", "ContactTitle1"]),
    today:          new Date().toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" }),
    sender_name:    senderName,
    sender_email:   senderEmail,
  };
}

function applyTokens(template, tokens) {
  return Object.entries(tokens).reduce(
    (result, [key, value]) => result.replaceAll(`{{${key}}}`, value || ""),
    template
  );
}

function buildPreview() {
  const tpl = state.templates.find((t) => t.key === state.recipientType);
  if (!tpl) return { subject: "", body: "" };
  const project = state.projects.find((p) => p.id === state.selectedProjectId) || null;
  const company = state.companies.find((c) => c.id === state.selectedCompanyId) || null;
  if (!project || !company) return { subject: tpl.subject, body: tpl.body };
  const tokens = buildTokenMap(project, company);
  return { subject: applyTokens(tpl.subject, tokens), body: applyTokens(tpl.body, tokens) };
}

// ─── Status ───────────────────────────────────────────────────────────────────

function setStatus(message, type = "info") {
  state.statusMessage = message;
  state.statusType = type;
  const el = document.getElementById("statusBanner");
  if (el) { el.textContent = message; el.className = `status-banner ${type}`; }
}

// ─── Template Manager ────────────────────────────────────────────────────────

const TOKEN_REFERENCE = [
  ["project_name",   "Project Title / Name from SharePoint"],
  ["project_number", "ProjectCode field"],
  ["company_name",   "Selected company name"],
  ["contact_name",   "Company primary contact name"],
  ["contact_email",  "Company primary contact email"],
  ["contact_title",  "Company primary contact title"],
  ["today",          "Current date"],
  ["sender_name",    "Your Outlook display name"],
  ["sender_email",   "Your Outlook email address"],
];

function createTokenRef() {
  const rows = TOKEN_REFERENCE.map(([token, desc]) =>
    `<li><code>{{${escapeHtml(token)}}}</code> — ${escapeHtml(desc)}</li>`
  ).join("");
  return `<details class="token-ref"><summary>Available tokens</summary><ul>${rows}</ul></details>`;
}

function createTemplateManager() {
  if (!state.showTemplateManager) return "";

  const et = state.editingTemplate;
  const editForm = et ? `
    <div class="template-edit-form">
      <h3>${et._isNew ? "New Template" : "Edit Template"}</h3>
      <label>Label
        <input id="tplLabel" value="${escapeHtml(et.label)}" placeholder="e.g. Subcontractor – Invitation to Bid" />
      </label>
      <label>Subject
        <input id="tplSubject" value="${escapeHtml(et.subject)}" />
      </label>
      <label>Body
        <textarea id="tplBody" rows="10">${escapeHtml(et.body)}</textarea>
      </label>
      ${createTokenRef()}
      <div class="actions">
        <button id="saveTemplate" class="btn-primary">Save template</button>
        <button id="cancelEdit" class="btn-secondary">Cancel</button>
      </div>
    </div>` : "";

  const list = state.templates.length
    ? state.templates.map((tpl) => `
        <div class="template-item">
          <div class="template-item-info">
            <strong>${escapeHtml(tpl.label)}</strong>
            <span>${escapeHtml(tpl.subject)}</span>
          </div>
          <div class="template-item-actions">
            <button class="btn-secondary btn-sm js-edit-tpl" data-key="${escapeHtml(tpl.key)}">Edit</button>
            <button class="btn-secondary btn-sm js-delete-tpl" data-key="${escapeHtml(tpl.key)}">Delete</button>
          </div>
        </div>`).join("")
    : `<p class="muted">No templates. Create one below.</p>`;

  return `
    <section class="panel">
      <div class="panel-header">
        <h2>Manage Templates</h2>
        <button id="newTemplate" class="btn-secondary btn-sm">+ New</button>
      </div>
      <div class="template-list">${list}</div>
      ${editForm}
    </section>`;
}

// ─── Render ───────────────────────────────────────────────────────────────────

function render() {
  const sortedProjects = [...state.projects].sort((a, b) =>
    projectDisplayName(a.fields || {}).localeCompare(projectDisplayName(b.fields || {}))
  );
  const sortedCompanies = [...state.companies].sort((a, b) =>
    companyDisplayName(a.fields || {}).localeCompare(companyDisplayName(b.fields || {}))
  );

  const typeOptions = state.templates
    .map((t) => `<option value="${escapeHtml(t.key)}" ${state.recipientType === t.key ? "selected" : ""}>${escapeHtml(t.label)}</option>`)
    .join("");
  const projectOptions = sortedProjects
    .map((p) => `<option value="${escapeHtml(p.id)}" ${p.id === state.selectedProjectId ? "selected" : ""}>${escapeHtml(projectDisplayName(p.fields || {}))}</option>`)
    .join("");
  const companyOptions = sortedCompanies
    .map((c) => `<option value="${escapeHtml(c.id)}" ${c.id === state.selectedCompanyId ? "selected" : ""}>${escapeHtml(companyDisplayName(c.fields || {}))}</option>`)
    .join("");

  const preview = buildPreview();
  const canApply = !!(state.selectedProjectId && state.selectedCompanyId);

  document.getElementById("app").innerHTML = `
    <main class="app-shell fluent-root">
      <header class="topbar">
        <div>
          <h1>Email Template</h1>
          <p>Select a project and company to generate an email.</p>
        </div>
        <div class="topbar-actions">
          <button id="toggleTemplateManager" class="btn-secondary">${state.showTemplateManager ? "Done" : "Manage templates"}</button>
          <button id="signin" class="btn-secondary">Sign in</button>
        </div>
      </header>
      <div id="statusBanner" class="status-banner ${escapeHtml(state.statusType)}">${escapeHtml(state.statusMessage)}</div>

      ${createTemplateManager()}

      <section class="panel">
        <h2>Compose</h2>
        <div class="grid">
          <label>Template
            <select id="recipientType">${typeOptions || '<option value="">No templates — create one above</option>'}</select>
          </label>
          <label>Project
            <select id="projectSelect">
              <option value="">— Select a project —</option>
              ${projectOptions}
            </select>
          </label>
          <label>Company
            <select id="companySelect">
              <option value="">— Select a company —</option>
              ${companyOptions}
            </select>
          </label>
        </div>
      </section>

      <section class="panel">
        <h2>Preview</h2>
        <label>Subject
          <input id="previewSubject" value="${escapeHtml(preview.subject)}" />
        </label>
        <label>Body
          <textarea id="previewBody" rows="14">${escapeHtml(preview.body)}</textarea>
        </label>
        <div class="actions">
          <button id="applyTemplate" class="btn-primary" ${canApply ? "" : "disabled"}>Apply to email</button>
          <button id="addRecipient" class="btn-secondary" ${canApply ? "" : "disabled"}>Add recipient</button>
        </div>
      </section>

      <p id="accountLabel" class="hint">${state.account ? `Signed in as ${escapeHtml(state.account.username)}` : "Not signed in."}</p>
    </main>`;

  wireEvents();
}

// Only refreshes the preview fields and button state without a full re-render.
function refreshPreview() {
  const preview = buildPreview();
  const canApply = !!(state.selectedProjectId && state.selectedCompanyId);
  const subjectEl = document.getElementById("previewSubject");
  const bodyEl    = document.getElementById("previewBody");
  const applyBtn  = document.getElementById("applyTemplate");
  const addBtn    = document.getElementById("addRecipient");
  if (subjectEl) subjectEl.value = preview.subject;
  if (bodyEl)    bodyEl.value    = preview.body;
  if (applyBtn)  applyBtn.disabled = !canApply;
  if (addBtn)    addBtn.disabled   = !canApply;
}

// ─── Events ───────────────────────────────────────────────────────────────────

function wireEvents() {
  document.getElementById("signin").addEventListener("click", () =>
    withBusy("Sign in", async () => {
      state.account = await signIn(state.config);
      document.getElementById("accountLabel").textContent = `Signed in as ${state.account.username}`;
      await loadData();
      render();
    })
  );

  document.getElementById("toggleTemplateManager").addEventListener("click", () => {
    state.showTemplateManager = !state.showTemplateManager;
    state.editingTemplate = null;
    render();
  });

  document.getElementById("recipientType").addEventListener("change", (e) => {
    state.recipientType = e.target.value;
    refreshPreview();
  });

  document.getElementById("projectSelect").addEventListener("change", (e) => {
    state.selectedProjectId = e.target.value;
    refreshPreview();
  });

  document.getElementById("companySelect").addEventListener("change", (e) => {
    state.selectedCompanyId = e.target.value;
    refreshPreview();
  });

  document.getElementById("applyTemplate").addEventListener("click", () =>
    withBusy("Apply template", async () => {
      const subject = document.getElementById("previewSubject").value;
      const body    = document.getElementById("previewBody").value;
      await officeAsync((cb) => Office.context.mailbox.item.subject.setAsync(subject, cb));
      await officeAsync((cb) => Office.context.mailbox.item.body.setAsync(body, { coercionType: Office.CoercionType.Text }, cb));
    })
  );

  document.getElementById("addRecipient").addEventListener("click", () =>
    withBusy("Add recipient", async () => {
      const company = state.companies.find((c) => c.id === state.selectedCompanyId);
      const email   = resolveField(company?.fields || {}, ["Email 1", "Email1", "Primary Contact Email", "PrimaryContactEmail"]);
      if (!email) throw new Error("No email address found for the selected company.");
      await officeAsync((cb) => Office.context.mailbox.item.to.addAsync([email], cb));
    })
  );

  // Template manager events (only wired when panel is visible)
  if (!state.showTemplateManager) return;

  document.getElementById("newTemplate").addEventListener("click", () => {
    state.editingTemplate = { key: "", label: "", subject: "", body: "", _isNew: true };
    render();
  });

  document.querySelectorAll(".js-edit-tpl").forEach((btn) => {
    btn.addEventListener("click", () => {
      const tpl = state.templates.find((t) => t.key === btn.dataset.key);
      if (tpl) { state.editingTemplate = { ...tpl, _isNew: false }; render(); }
    });
  });

  document.querySelectorAll(".js-delete-tpl").forEach((btn) => {
    btn.addEventListener("click", () => {
      if (!confirm(`Delete template "${btn.dataset.key}"?`)) return;
      state.templates = state.templates.filter((t) => t.key !== btn.dataset.key);
      saveTemplates(state.templates);
      if (state.recipientType === btn.dataset.key) {
        state.recipientType = state.templates[0]?.key || "";
      }
      render();
    });
  });

  const saveBtn   = document.getElementById("saveTemplate");
  const cancelBtn = document.getElementById("cancelEdit");

  if (saveBtn) {
    saveBtn.addEventListener("click", () => {
      const label   = document.getElementById("tplLabel").value.trim();
      const subject = document.getElementById("tplSubject").value.trim();
      const body    = document.getElementById("tplBody").value;

      if (!label)   { setStatus("Label is required.", "error"); return; }
      if (!subject) { setStatus("Subject is required.", "error"); return; }

      const key = state.editingTemplate._isNew
        ? generateKey(label, state.templates)
        : state.editingTemplate.key;

      const updated = { key, label, subject, body };
      const idx = state.templates.findIndex((t) => t.key === key);
      if (idx >= 0) {
        state.templates = state.templates.map((t, i) => (i === idx ? updated : t));
      } else {
        state.templates = [...state.templates, updated];
      }

      saveTemplates(state.templates);
      if (!state.recipientType) state.recipientType = key;
      state.editingTemplate = null;
      setStatus(`Template "${label}" saved.`, "success");
      render();
    });
  }

  if (cancelBtn) {
    cancelBtn.addEventListener("click", () => { state.editingTemplate = null; render(); });
  }
}

// ─── Data + Utilities ─────────────────────────────────────────────────────────

async function loadData() {
  const [projectsRes, companiesRes] = await Promise.all([
    getListItemsById(state.config, state.config.listId),
    state.config.companyListId
      ? getListItemsById(state.config, state.config.companyListId)
      : Promise.resolve({ value: [] }),
  ]);
  state.projects  = projectsRes.value  || [];
  state.companies = companiesRes.value || [];
}

async function withBusy(label, action) {
  try {
    setStatus(`${label}...`, "info");
    await action();
    setStatus(`${label} complete.`, "success");
  } catch (err) {
    setStatus(`${label} failed: ${err.message}`, "error");
  }
}

function officeAsync(fn) {
  return new Promise((resolve, reject) =>
    fn((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) reject(new Error(result.error.message));
      else resolve(result.value);
    })
  );
}

// ─── Bootstrap ────────────────────────────────────────────────────────────────

if (window.Office) {
  Office.onReady(() => render());
} else {
  render();
}

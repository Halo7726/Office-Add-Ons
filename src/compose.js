import "./styles.css";
import { loadConfig } from "./config";
import { signIn, restoreSignIn, getListItemsById } from "./graphClient";
import { loadTemplates, saveTemplates, generateKey } from "./template-store";

const state = {
  config: loadConfig(),
  account: null,
  projects: [],
  templates: loadTemplates(),
  selectedProjectId: "",
  selectedStatusFilter: "all",
  recipientType: "",
  showTemplateManager: false,
  editingTemplate: null,
  statusMessage: "Sign in to load projects.",
  statusType: "info",
};

if (state.templates.length > 0 && !state.recipientType) {
  state.recipientType = state.templates[0].key;
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function normalizeFieldValue(value) {
  if (value == null) return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return String(value);
  if (Array.isArray(value)) return value.map(normalizeFieldValue).filter(Boolean).join(", ");
  if (typeof value === "object") {
    if (typeof value.Url === "string") {
      return value.Description ? `${value.Description} (${value.Url})` : value.Url;
    }
    if (typeof value.url === "string") {
      return value.description ? `${value.description} (${value.url})` : value.url;
    }
    if (typeof value.Value !== "undefined") return String(value.Value || value.Label || "");
    if (typeof value.Label !== "undefined") return String(value.Label);
    if (typeof value.text === "string") return value.text;
    return JSON.stringify(value);
  }
  return String(value);
}

function resolveField(fields, candidates) {
  const keys = Object.keys(fields || {});
  const map = new Map(
    keys.map((k) => [k.replace(/_x[0-9a-fA-F]{4}_/g, " ").replace(/[^a-zA-Z0-9]+/g, "").toLowerCase(), k])
  );
  for (const c of candidates) {
    const n = c.replace(/[^a-zA-Z0-9]+/g, "").toLowerCase();
    if (map.has(n)) return normalizeFieldValue(fields[map.get(n)]);
  }
  return "";
}

function resolveHyperlinkField(fields, candidates) {
  const keys = Object.keys(fields || {});
  const map = new Map(
    keys.map((k) => [k.replace(/_x[0-9a-fA-F]{4}_/g, " ").replace(/[^a-zA-Z0-9]+/g, "").toLowerCase(), k])
  );
  for (const c of candidates) {
    const n = c.replace(/[^a-zA-Z0-9]+/g, "").toLowerCase();
    if (!map.has(n)) continue;
    const raw = fields[map.get(n)];
    if (!raw) return { url: "", label: "" };
    if (typeof raw === "string") return { url: raw, label: raw };
    if (typeof raw === "object") {
      const url = raw.Url || raw.url || "";
      const label = raw.Description || raw.description || url;
      return { url, label };
    }
    return { url: String(raw), label: String(raw) };
  }
  return { url: "", label: "" };
}

function projectDisplayName(fields) {
  return fields.Title || fields.Project_x0020_Name || "(Unnamed)";
}

function getProjectEstimateStatus(fields) {
  return resolveField(fields, ["EstimateStatus", "Estimate Status"]).trim();
}

function normalizeStatus(status) {
  return String(status || "").trim().toLowerCase();
}

function projectStatusLabel(status) {
  const normalized = normalizeStatus(status);
  if (!normalized) return "Unspecified";
  return status;
}

function getStatusOptions(projects) {
  const seen = new Set();
  const options = [];
  for (const project of projects) {
    const rawStatus = getProjectEstimateStatus(project.fields || {});
    const label = projectStatusLabel(rawStatus);
    if (!seen.has(label)) {
      seen.add(label);
      options.push(label);
    }
  }
  return options.sort((a, b) => {
    if (a === "Unspecified") return 1;
    if (b === "Unspecified") return -1;
    return a.localeCompare(b);
  });
}

function getFilteredProjects(projects) {
  return projects.filter((p) => {
    if (state.selectedStatusFilter === "all") return true;
    const status = projectStatusLabel(getProjectEstimateStatus(p.fields || {}));
    return normalizeStatus(status) === normalizeStatus(state.selectedStatusFilter);
  });
}

function parseDateValue(value) {
  if (!value) return null;
  const date = new Date(value);
  return Number.isFinite(date.valueOf()) ? date : null;
}

function formatDateValue(value) {
  const date = parseDateValue(value);
  if (!date) return "";
  return date.toLocaleString("en-US", {
    month: "long",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
    hour12: true,
  });
}

function buildTokenMap(project) {
  const pf = project?.fields || {};
  let senderName = "", senderEmail = "";
  try {
    const profile = Office.context.mailbox.userProfile;
    senderName = profile.displayName || "";
    senderEmail = profile.emailAddress || "";
  } catch { }

  const owner = resolveField(pf, ["Owner", "OWNER"]);
  const projectLocation = resolveField(pf, ["Project Location", "PROJECT_LOCATION", "Location"]);
  const bidDue = resolveField(pf, ["Bid Due Date", "BID_DUE_DATE", "BidDueDate", "BidDue"]);
  const prebidDate = resolveField(pf, ["Prebid Date", "PREBID_DATE", "PrebidDate", "Prebid"]);
  const questionDueDate = resolveField(pf, ["QuestionsDue", "Question Due Date", "QUESTION_DUE_DATE", "QuestionDueDate", "RFI Date", "RFI_Date"]);
  const planLinkField = resolveHyperlinkField(pf, ["SharedFolderLinkText", "SharedFolderLink", "Plan Link", "PLAN_LINK", "Plans"]);
  const takeoffLinkField = resolveHyperlinkField(pf, ["TakeoffLink", "Takeoff Link", "TAKEOFF_LINK"]);
  const bidDate = parseDateValue(bidDue);
  const deadlineDate = bidDate ? new Date(bidDate.valueOf() - 2 * 24 * 60 * 60 * 1000) : null;

  return {
    project_name:    projectDisplayName(pf),
    project_number:  String(pf.JobNumber || pf.EstimateNumber || ""),
    project_label:   [owner, projectDisplayName(pf)].filter(Boolean).join(" - "),
    owner,
    project_location: projectLocation,
    bid_due_date:     formatDateValue(bidDue),
    prebid_date:      formatDateValue(prebidDate),
    question_due_date: formatDateValue(questionDueDate),
    plan_link:        planLinkField.url,
    plan_link_label:  planLinkField.label,
    takeoff_link:     takeoffLinkField.url,
    takeoff_link_label: takeoffLinkField.label,
    deadline:         formatDateValue(deadlineDate),
    today:            new Date().toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" }),
    sender_name:      senderName,
    sender_email:     senderEmail,
  };
}

function applyTokens(template, tokens) {
  let result = template;
  for (const [key, value] of Object.entries(tokens)) {
    if (value) {
      result = result.replaceAll(`{{${key}}}`, value);
    } else {
      // Remove the entire line when the token is the only content after a label (e.g. "Pre-Bid:   {{prebid_date}}")
      result = result.replace(new RegExp(`^[^\n]*\\{\\{${key}\\}\\}[ \t]*\n?`, "gm"), "");
    }
  }
  // Collapse 3+ consecutive newlines down to 2 so removed lines don't leave gaps
  return result.replace(/\n{3,}/g, "\n\n").trim();
}

function buildBodyHtml(text, linkLabels = {}) {
  // Split on URLs first so we linkify before HTML-escaping (prevents & → &amp; from breaking href values).
  // linkLabels maps a raw URL to a friendly display label (e.g. { "https://...": "Plans" }).
  return String(text || "")
    .split(/(https?:\/\/[^\s]+)/g)
    .map((part, i) => {
      if (i % 2 === 1) {
        const url = escapeHtml(part);
        const label = escapeHtml(linkLabels[part] || part);
        return `<a href="${url}">${label}</a>`;
      }
      return escapeHtml(part).replace(/\r?\n/g, "<br />");
    })
    .join("");
}

function buildPreview() {
  const tpl = state.templates.find((t) => t.key === state.recipientType);
  if (!tpl) return { subject: "", body: "" };
  const project = state.projects.find((p) => p.id === state.selectedProjectId) || null;
  const tokens = buildTokenMap(project);
  return { subject: applyTokens(tpl.subject, tokens), body: applyTokens(tpl.body, tokens) };
}

function setStatus(message, type = "info") {
  state.statusMessage = message;
  state.statusType = type;
  const el = document.getElementById("statusBanner");
  if (el) { el.textContent = message; el.className = `status-banner ${type}`; }
}

const TOKEN_REFERENCE = [
  ["project_name",      "Project Title / Name from SharePoint"],
  ["project_number",    "JobNumber / EstimateNumber from SharePoint"],
  ["project_label",     "Combined owner and project label"],
  ["project_location",  "Project location"],
  ["owner",             "Project owner"],
  ["bid_due_date",      "Project bid due date formatted for email"],
  ["prebid_date",       "Project pre-bid date formatted for email"],
  ["question_due_date", "Questions due date (QuestionsDue)"],
  ["plan_link",         "Shared folder link (SharedFolderLinkText)"],
  ["takeoff_link",      "Project takeoff link"],
  ["deadline",          "Bid deadline two days before bid due date"],
  ["today",             "Current date"],
  ["sender_name",       "Your Outlook display name"],
  ["sender_email",      "Your Outlook email address"],
];

function createTokenRef() {
  const rows = TOKEN_REFERENCE.map(([token, desc]) =>
    `<li><button type="button" class="token-insert btn-secondary btn-sm" data-token="${escapeHtml(token)}"><code>{{${escapeHtml(token)}}}</code></button> — ${escapeHtml(desc)}</li>`
  ).join("");
  return `<details class="token-ref"><summary>Available tokens</summary><ul>${rows}</ul></details>`;
}

function getActiveTemplateField() {
  const active = document.activeElement;
  if (active?.id === "tplSubject" || active?.id === "tplBody") return active;
  return document.getElementById("tplBody") || document.getElementById("tplSubject");
}

function insertToken(token) {
  const field = getActiveTemplateField();
  if (!field) return;
  const placeholder = `{{${token}}}`;
  const start = field.selectionStart ?? field.value.length;
  const end = field.selectionEnd ?? start;
  field.value = field.value.slice(0, start) + placeholder + field.value.slice(end);
  field.selectionStart = field.selectionEnd = start + placeholder.length;
  field.focus();
}

function createTemplateManager() {
  if (!state.showTemplateManager) return "";

  const et = state.editingTemplate;
  const editForm = et ? `
    <div class="template-edit-form">
      <h3>${et._isNew ? "New Template" : "Edit Template"}</h3>
      <label>Label
        <input id="tplLabel" value="${escapeHtml(et.label)}" placeholder="e.g. Invitation to Bid" />
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

function render() {
  const sortedProjects = [...state.projects].sort((a, b) =>
    projectDisplayName(a.fields || {}).localeCompare(projectDisplayName(b.fields || {}))
  );

  const typeOptions = state.templates
    .map((t) => `<option value="${escapeHtml(t.key)}" ${state.recipientType === t.key ? "selected" : ""}>${escapeHtml(t.label)}</option>`)
    .join("");

  const statusOptions = ["All statuses", ...getStatusOptions(sortedProjects)];
  const statusFilterOptions = statusOptions
    .map((status) => {
      const value = status === "All statuses" ? "all" : escapeHtml(status);
      const selected = state.selectedStatusFilter === value ? "selected" : "";
      return `<option value="${value}" ${selected}>${escapeHtml(status)}</option>`;
    })
    .join("");

  const filteredProjects = getFilteredProjects(sortedProjects);

  const projectOptions = filteredProjects
    .map((p) => {
      const status = projectStatusLabel(getProjectEstimateStatus(p.fields || {}));
      const rawLabel = `${projectDisplayName(p.fields || {})} ${status && status !== "Unspecified" ? `— ${status}` : ""}`;
      return `<option value="${escapeHtml(p.id)}" ${p.id === state.selectedProjectId ? "selected" : ""}>${escapeHtml(rawLabel)}</option>`;
    })
    .join("");

  if (!state.recipientType && state.templates.length > 0) {
    state.recipientType = state.templates[0].key;
  }

  const preview = buildPreview();
  const canApply = !!state.recipientType;
  const app = document.getElementById("app");
  if (!app) {
    console.error("App root element not found: #app");
    return;
  }

  app.innerHTML = `
    <main class="app-shell fluent-root">
      <header class="topbar">
        <div>
          <h1>Email Template</h1>
          <p>Pick a project, choose a template, and generate clean project-focused email copy.</p>
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
          <label>Status
            <select id="statusFilter">
              ${statusFilterOptions}
            </select>
          </label>
          <label>Project
            <select id="projectSelect">
              <option value="">— Select a project —</option>
              ${projectOptions}
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
          <button id="copyBody" class="btn-secondary" ${preview.body ? "" : "disabled"}>Copy body</button>
          <button id="copySubject" class="btn-secondary" ${preview.subject ? "" : "disabled"}>Copy subject</button>
        </div>
      </section>

      <p id="accountLabel" class="hint">${state.account ? `Signed in as ${escapeHtml(state.account.username)}` : "Not signed in."}</p>
    </main>`;

  wireEvents();
}

function refreshPreview() {
  const preview = buildPreview();
  const canApply = !!state.recipientType;
  const subjectEl = document.getElementById("previewSubject");
  const bodyEl = document.getElementById("previewBody");
  const applyBtn = document.getElementById("applyTemplate");
  const copyBody = document.getElementById("copyBody");
  const copySubject = document.getElementById("copySubject");
  if (subjectEl) subjectEl.value = preview.subject;
  if (bodyEl) bodyEl.value = preview.body;
  if (applyBtn) applyBtn.disabled = !canApply;
  if (copyBody) copyBody.disabled = !preview.body;
  if (copySubject) copySubject.disabled = !preview.subject;
}

function wireEvents() {
  const signinBtn = document.getElementById("signin");
  const accountLabel = document.getElementById("accountLabel");
  const toggleTplBtn = document.getElementById("toggleTemplateManager");
  const recipientTypeSelect = document.getElementById("recipientType");
  const statusFilterSelect = document.getElementById("statusFilter");
  const projectSelect = document.getElementById("projectSelect");
  const applyBtn = document.getElementById("applyTemplate");
  const copyBodyBtn = document.getElementById("copyBody");
  const copySubjectBtn = document.getElementById("copySubject");

  if (signinBtn) {
    signinBtn.addEventListener("click", () =>
      withBusy("Sign in", async () => {
        state.account = await signIn(state.config);
        if (accountLabel) accountLabel.textContent = `Signed in as ${state.account.username}`;
        await loadData();
        render();
      })
    );
  }

  if (toggleTplBtn) {
    toggleTplBtn.addEventListener("click", () => {
      state.showTemplateManager = !state.showTemplateManager;
      state.editingTemplate = null;
      render();
    });
  }

  if (recipientTypeSelect) {
    recipientTypeSelect.addEventListener("change", (e) => {
      state.recipientType = e.target.value;
      refreshPreview();
    });
  }

  if (statusFilterSelect) {
    statusFilterSelect.addEventListener("change", (e) => {
      state.selectedStatusFilter = e.target.value;
      if (state.selectedProjectId) {
        const stillSelected = getFilteredProjects(state.projects).some((p) => p.id === state.selectedProjectId);
        if (!stillSelected) {
          state.selectedProjectId = "";
        }
      }
      render();
    });
  }

  if (projectSelect) {
    projectSelect.addEventListener("change", (e) => {
      state.selectedProjectId = e.target.value;
      refreshPreview();
      render();
    });
  }

  if (applyBtn) {
    applyBtn.addEventListener("click", () =>
      withBusy("Apply template", async () => {
        const subject = document.getElementById("previewSubject")?.value || "";
        const body = document.getElementById("previewBody")?.value || "";
        const project = state.projects.find((p) => p.id === state.selectedProjectId) || null;
        const tokens = buildTokenMap(project);
        const linkLabels = {};
        if (tokens.plan_link) linkLabels[tokens.plan_link] = tokens.plan_link_label || "Plans";
        if (tokens.takeoff_link) linkLabels[tokens.takeoff_link] = tokens.takeoff_link_label || "Takeoffs";
        await officeAsync((cb) => Office.context.mailbox.item.subject.setAsync(subject, cb));
        await officeAsync((cb) => Office.context.mailbox.item.body.setAsync(buildBodyHtml(body, linkLabels), { coercionType: Office.CoercionType.Html }, cb));
      })
    );
  }

  if (copyBodyBtn) {
    copyBodyBtn.addEventListener("click", async () => {
      const body = document.getElementById("previewBody")?.value || "";
      if (!body) return;
      if (!navigator.clipboard) throw new Error("Clipboard is unavailable.");
      await navigator.clipboard.writeText(body);
      setStatus("Email body copied to clipboard.", "success");
    });
  }

  if (copySubjectBtn) {
    copySubjectBtn.addEventListener("click", async () => {
      const subject = document.getElementById("previewSubject")?.value || "";
      if (!subject) return;
      if (!navigator.clipboard) throw new Error("Clipboard is unavailable.");
      await navigator.clipboard.writeText(subject);
      setStatus("Email subject copied to clipboard.", "success");
    });
  }

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

  const saveBtn = document.getElementById("saveTemplate");
  const cancelBtn = document.getElementById("cancelEdit");

  if (saveBtn) {
    saveBtn.addEventListener("click", () => {
      const label = document.getElementById("tplLabel").value.trim();
      const subject = document.getElementById("tplSubject").value.trim();
      const body = document.getElementById("tplBody").value;
      if (!label) { setStatus("Label is required.", "error"); return; }
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

  document.querySelectorAll(".token-insert").forEach((btn) => {
    btn.addEventListener("click", () => insertToken(btn.dataset.token));
  });
}

async function loadData() {
  const projectsRes = await getListItemsById(state.config, state.config.listId);
  state.projects = projectsRes.value || [];
}

async function tryAutoSignIn() {
  try {
    setStatus("Signing in...", "info");
    let account = await restoreSignIn(state.config);
    if (!account) {
      account = await signIn(state.config);
    }
    if (!account) return;
    state.account = account;
    setStatus("Loading projects...", "info");
    await loadData();
    setStatus(`Signed in as ${account.username}`, "success");
    render();
  } catch (err) {
    console.warn("Auto sign-in failed:", err);
    setStatus("Sign in to load projects.", "info");
  }
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

function showFatalError(error) {
  const app = document.getElementById("app");
  const message = error?.message || String(error);
  if (app) {
    app.innerHTML = `
      <main class="app-shell fluent-root">
        <section class="panel">
          <h1>Unable to load the email template tool</h1>
          <p>${escapeHtml(message)}</p>
          <p>Please refresh the taskpane or reload Outlook.</p>
        </section>
      </main>`;
  } else {
    document.body.textContent = `Error: ${message}`;
  }
  console.error(error);
}

function tryInitialize() {
  try {
    render();
  } catch (error) {
    showFatalError(error);
  }
}

if (window.Office && typeof Office.onReady === "function") {
  Office.onReady(async () => {
    tryInitialize();
    await tryAutoSignIn();
  });
} else if (window.Office) {
  // Some hosts expose a partial Office object before the library is fully initialized.
  setTimeout(() => {
    if (typeof Office.onReady === "function") {
      Office.onReady(async () => {
        tryInitialize();
        await tryAutoSignIn();
      });
    } else {
      tryInitialize();
    }
  }, 100);
} else {
  tryInitialize();
}

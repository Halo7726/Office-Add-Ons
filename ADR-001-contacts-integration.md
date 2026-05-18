# ADR-001: Outlook Contacts Integration

**Status:** Proposed  
**Date:** 2026-05-13  
**Deciders:** Michael Ferriss  

---

## Context

SHC Project Tools currently handles two workflows in Outlook: saving proposal attachments to SharePoint and applying email templates when composing. Both workflows already resolve the sender against a SharePoint **Company List** (subcontractors, vendors) using `routingEngine.js`, and extract sender info via `outlookContext.js`.

The gap is that contacts resolved during these workflows — subcontractors, owners, engineers, suppliers — are never persisted to Outlook Contacts. Repeatedly received emails from known subs like Ark Roofing or Empire Pipe have no associated contact card, making it impossible to look them up, call them from mobile, or see their history in the Outlook People pane.

The goal is to add a **contact creation and sync panel** to the existing read-mode taskpane (`main.js` / `index.html`) that:

1. Checks whether the email sender already exists in Outlook Contacts
2. If not, pre-fills a contact creation form from email sender data, signature parsing, and the matched Company List record
3. Saves the contact to Outlook via Graph API, filed into an appropriate contact folder (Subcontractors, Owners, Engineers, Suppliers)
4. Optionally syncs contact data back to the SharePoint Company List

### Existing assets that directly support this

| Asset | Relevance |
|---|---|
| `graphClient.js` → `graphRequest()` | Reusable authenticated Graph wrapper — contact endpoints plug straight in |
| `routingEngine.js` → `resolveRoute()` | Already scores sender against Company List — gives us company match for free |
| `outlookContext.js` → `getMessageContext()` | Already extracts `sender.email` and `sender.name` |
| `config.js` + `companyListId` | Company List already has `Email_x0020_1`, `Contact_x0020_Name_x0020_1`, `Contact_x0020_Title` fields |
| MSAL auth flow | Already established — just needs one additional scope |
| `styles.css` Fluent tokens | Contact card UI can reuse existing `.panel`, `.grid`, `.match-card` classes |

---

## Decision

**Extend the existing vanilla JS / native Graph API pattern** rather than introducing Microsoft Graph Toolkit (MGT) as a dependency.

Add a new `contactManager.js` module, extend `graphClient.js` with contacts endpoints, add `Contacts.ReadWrite` to the MSAL scope, and render a contact panel inside the existing `index.html` taskpane.

---

## Options Considered

### Option A: Microsoft Graph Toolkit (MGT) web components

Use `<mgt-person-card>`, `<mgt-people-picker>`, and `<mgt-person>` components from `@microsoft/mgt-element` and `@microsoft/mgt-msal2-provider`.

| Dimension | Assessment |
|---|---|
| Complexity | Medium — MGT auth provider must integrate with existing MSAL instance |
| Bundle size | High — MGT adds ~500KB+ to bundle |
| UI quality | High — native Microsoft card look out of the box |
| Codebase fit | Low — introduces a component framework into a vanilla JS codebase |
| Flexibility | Low — MGT card layout is opinionated; hard to embed company list data alongside |

**Pros:**
- Pre-built person card UI matches Outlook's native People pane look
- People picker handles search/autocomplete for free
- `Msal2Provider` can accept existing MSAL config

**Cons:**
- Significant bundle size increase for a single feature
- MGT's `Msal2Provider` initializes its own MSAL instance — needs careful wiring to avoid dual token caches
- Hard to inject SharePoint Company List fields (company type, ITB history) into MGT card layout
- All other add-in UI is vanilla JS; MGT adds a second paradigm to maintain

---

### Option B: Native Graph API + custom UI (vanilla JS) ✅ Recommended

Add contacts Graph endpoints to `graphClient.js`, build a lightweight contact card panel using existing Fluent CSS tokens, and add a new `contactManager.js` module.

| Dimension | Assessment |
|---|---|
| Complexity | Low — follows exact same pattern as existing list operations |
| Bundle size | Zero — no new dependencies |
| UI quality | Medium — consistent with existing add-in look; not the native Outlook card |
| Codebase fit | High — identical pattern to all existing Graph calls |
| Flexibility | High — full control to blend Outlook contact data with Company List data |

**Pros:**
- Zero new dependencies — `graphClient.js` already does everything needed
- Sender email is already extracted; company match is already scored — 80% of the data work is done before writing a line of contact code
- Contact card UI reuses `.panel`, `.match-card`, `.grid` CSS classes — consistent look with no extra styling work
- Full control to show Company List fields (type, ITB count, folder link) alongside contact data
- No MSAL wiring complexity

**Cons:**
- No auto-complete people picker (can implement basic email search against `/me/contacts` if needed)
- Contact card won't match Outlook's native People pane pixel-for-pixel

---

## Trade-off Analysis

MGT's value proposition is pre-built UI. But in this add-in, the UI is the smallest part of the problem — the sender's name, email, company, and title are already available from `outlookContext.js` and `routingEngine.js`. The contact card is just a confirmation form with a few fields. Pulling in 500KB of web components to render a form that takes ~40 lines of vanilla JS is not a good trade.

The bigger unlock here is the **bidirectional sync between Outlook Contacts and the SharePoint Company List** — specifically, keeping contact details current in both places without manual data entry. That logic lives in `contactManager.js` and has nothing to do with which UI library renders the card.

---

## Implementation Plan

### Phase 1 — Auth + Graph endpoints (1–2 hours)

**1. Add `Contacts.ReadWrite` scope**

In `graphClient.js`, update the scopes array:
```js
// graphClient.js
const SCOPES = [
  'User.Read',
  'Sites.ReadWrite.All',
  'Files.ReadWrite.All',
  'Contacts.ReadWrite'   // ← add this
];
```

**2. Add contact functions to `graphClient.js`**

```js
// Find contact by email address
export async function findContactByEmail(config, email) {
  const encoded = encodeURIComponent(`emailAddresses/any(a:a/address eq '${email}')`);
  return graphRequest(config, `/me/contacts?$filter=${encoded}&$top=1`);
}

// Create a new contact
export async function createContact(config, fields) {
  return graphRequest(config, '/me/contacts', 'POST', fields);
}

// Update existing contact
export async function updateContact(config, contactId, fields) {
  return graphRequest(config, `/me/contacts/${contactId}`, 'PATCH', fields);
}

// List contact folders
export async function getContactFolders(config) {
  return graphRequest(config, '/me/contactFolders');
}

// Create a contact folder
export async function createContactFolder(config, displayName) {
  return graphRequest(config, '/me/contactFolders', 'POST', { displayName });
}

// Ensure standard folders exist, return map of name → folderId
export async function ensureContactFolders(config) {
  const folders = ['Subcontractors', 'Owners', 'Engineers & Architects', 'Suppliers', 'Government'];
  const existing = await getContactFolders(config);
  const map = {};
  for (const f of existing.value) map[f.displayName] = f.id;
  for (const name of folders) {
    if (!map[name]) {
      const created = await createContactFolder(config, name);
      map[name] = created.id;
    }
  }
  return map;
}
```

---

### Phase 2 — `contactManager.js` module (2–3 hours)

New file: `src/contactManager.js`

**Responsibilities:**
- Parse signature from email body to extract phone number and title
- Match sender against existing Outlook contacts
- Build a pre-filled contact object from sender + company match + signature
- Create or update contact via Graph
- Sync contact data back to SharePoint Company List item

```js
// src/contactManager.js

export function parseSignature(bodyHtml) {
  const text = bodyHtml.replace(/<[^>]+>/g, ' ');
  const phoneMatch = text.match(/(\+?1?\s?)?(\(?\d{3}\)?[\s.\-]?\d{3}[\s.\-]?\d{4})/);
  const phone = phoneMatch ? phoneMatch[0].trim() : null;
  // Title heuristic: line containing common title words near sender name
  const titleMatch = text.match(/\b(Project Manager|Estimator|President|Owner|Director|VP|Engineer|Architect|Sales|Associate|Superintendent)\b/i);
  const title = titleMatch ? titleMatch[0] : null;
  return { phone, title };
}

export function buildContactPayload(sender, companyMatch, signatureData, folderMap) {
  const nameParts = (sender.name || '').trim().split(' ');
  const folderCategory = companyMatch?.companyType ?? 'Subcontractors';
  return {
    givenName: nameParts[0] ?? '',
    surname: nameParts.slice(1).join(' ') ?? '',
    emailAddresses: [{ address: sender.email, name: sender.name }],
    companyName: companyMatch?.companyName ?? '',
    jobTitle: signatureData.title ?? '',
    businessPhones: signatureData.phone ? [signatureData.phone] : [],
    parentFolderId: folderMap[folderCategory] ?? null,
  };
}

export async function resolveContact(config, sender, companyMatch, bodyHtml) {
  // 1. Check if contact already exists
  const existing = await findContactByEmail(config, sender.email);
  if (existing.value.length > 0) return { status: 'exists', contact: existing.value[0] };

  // 2. Parse signature
  const sig = parseSignature(bodyHtml);

  // 3. Ensure folders exist
  const folderMap = await ensureContactFolders(config);

  // 4. Build payload
  const payload = buildContactPayload(sender, companyMatch, sig, folderMap);

  return { status: 'new', payload, folderMap };
}

export async function saveContact(config, payload) {
  return createContact(config, payload);
}

export async function syncContactToCompanyList(config, companyItemId, contactPayload) {
  const fields = {
    Contact_x0020_Name_x0020_1: `${contactPayload.givenName} ${contactPayload.surname}`.trim(),
    Email_x0020_1: contactPayload.emailAddresses[0]?.address ?? '',
    Contact_x0020_Title: contactPayload.jobTitle ?? '',
  };
  return updateListItemFields(config, config.companyListId, companyItemId, fields);
}
```

---

### Phase 3 — Contact panel UI in `main.js` (2–3 hours)

Add a contact panel section that renders after the existing project match section. The panel has three states:

**State 1 — Contact exists**
```
┌─────────────────────────────────┐
│ ✓  Noah Denney                  │
│    Ark Roofing, LLC             │
│    arkroofsga@gmail.com         │
│    478-444-5009                 │
│    [Subcontractors]             │
└─────────────────────────────────┘
```

**State 2 — New contact, pre-filled form**
```
┌─────────────────────────────────┐
│ New contact detected            │
│                                 │
│ Name    [Noah Denney          ] │
│ Email   [arkroofsga@gmail.com ] │
│ Company [Ark Roofing, LLC     ] │
│ Title   [Owner                ] │
│ Phone   [478-444-5009         ] │
│ Type    [Subcontractors  ▾    ] │
│                                 │
│ [Save Contact]  [Skip]          │
└─────────────────────────────────┘
```

**State 3 — Loading / no sender**
```
┌─────────────────────────────────┐
│ Checking contacts...            │
└─────────────────────────────────┘
```

**Rendering logic in `main.js`:**
```js
async function renderContactPanel(sender, companyMatch, bodyHtml) {
  const panel = document.getElementById('contact-panel');
  panel.innerHTML = '<p class="muted">Checking contacts…</p>';

  const result = await resolveContact(config, sender, companyMatch, bodyHtml);

  if (result.status === 'exists') {
    const c = result.contact;
    panel.innerHTML = `
      <div class="panel">
        <div class="panel-header">Contact on file</div>
        <div class="match-card match-high">
          <strong>${c.displayName}</strong>
          <span>${c.companyName}</span>
          <span>${c.emailAddresses?.[0]?.address ?? ''}</span>
          <span>${c.businessPhones?.[0] ?? ''}</span>
        </div>
      </div>`;
  } else {
    const p = result.payload;
    panel.innerHTML = `
      <div class="panel">
        <div class="panel-header">Save new contact</div>
        <div class="grid">
          <label>Name<input id="c-name" value="${p.givenName} ${p.surname}"></label>
          <label>Email<input id="c-email" value="${p.emailAddresses[0].address}"></label>
          <label>Company<input id="c-company" value="${p.companyName}"></label>
          <label>Title<input id="c-title" value="${p.jobTitle}"></label>
          <label>Phone<input id="c-phone" value="${p.businessPhones[0] ?? ''}"></label>
          <label>Type
            <select id="c-type">
              ${['Subcontractors','Owners','Engineers & Architects','Suppliers','Government']
                .map(t => `<option ${t === categoryFromMatch(companyMatch) ? 'selected' : ''}>${t}</option>`)
                .join('')}
            </select>
          </label>
        </div>
        <div class="action-row">
          <button id="btn-save-contact" class="btn-primary">Save Contact</button>
          <button id="btn-skip-contact" class="btn-ghost">Skip</button>
        </div>
      </div>`;

    document.getElementById('btn-save-contact').onclick = async () => {
      const payload = buildPayloadFromForm(result.folderMap);
      await saveContact(config, payload);
      if (companyMatch?.itemId) {
        await syncContactToCompanyList(config, companyMatch.itemId, payload);
      }
      renderContactPanel(sender, companyMatch, bodyHtml); // re-render as "exists"
    };
  }
}
```

---

### Phase 4 — Wire into existing `main.js` flow (30 min)

After `runAutoMatch()` resolves, call `renderContactPanel()` in parallel:

```js
// In main.js, after existing auto-match logic
const [routeResult] = await Promise.all([
  runAutoMatch(),
  renderContactPanel(sender, topCompanyMatch, messageBody)
]);
```

The contact panel runs alongside the existing project/attachment UI — no changes to the existing flow.

---

### Phase 5 — Manifest update (15 min)

No new taskpane or command button needed. The contact panel lives inside the existing `index.html` read-mode taskpane. No manifest changes required.

---

## File Change Summary

| File | Change |
|---|---|
| `src/graphClient.js` | Add `findContactByEmail`, `createContact`, `updateContact`, `getContactFolders`, `createContactFolder`, `ensureContactFolders`. Update MSAL scopes to include `Contacts.ReadWrite`. |
| `src/contactManager.js` | **New file.** Signature parser, contact payload builder, resolveContact, saveContact, syncContactToCompanyList. |
| `src/main.js` | Add `renderContactPanel()`. Call it after `runAutoMatch()`. Add `contact-panel` div to HTML string. |
| `src/styles.css` | Minor additions: `.contact-exists` success state, `.action-row` button row. Reuse existing `.panel`, `.grid`, `.match-card` classes. |
| `manifest.xml` | No changes required. |

---

## Consequences

**What becomes easier:**
- Every email from a subcontractor, owner, or engineer becomes a one-click contact save
- Outlook Contacts and SharePoint Company List stay in sync automatically
- Estimators can call/text subs directly from mobile using Outlook People
- Future features (Planner task creation, contact-linked proposals) have a clean contact record to reference

**What becomes harder:**
- MSAL scope change requires users to re-consent on next sign-in (one-time prompt)
- Signature parsing is heuristic — phone and title extraction will miss edge cases (international formats, non-standard signatures)

**What to revisit later:**
- If contact volume grows, add a full people picker / search UI against `/me/contacts`
- Consider Microsoft Graph Toolkit's `<mgt-person-card>` if the native card look becomes a priority after the core feature ships
- Evaluate whether contact folders should mirror SharePoint Company Types exactly (currently `Company_x0020_Type` field values are not yet known)

---

## Action Items

1. [ ] Add `Contacts.ReadWrite` to MSAL scopes in `graphClient.js`
2. [ ] Add contact Graph functions to `graphClient.js` (findContactByEmail, createContact, updateContact, getContactFolders, createContactFolder, ensureContactFolders)
3. [ ] Create `src/contactManager.js` (parseSignature, buildContactPayload, resolveContact, saveContact, syncContactToCompanyList)
4. [ ] Add `renderContactPanel()` to `main.js` and wire into existing flow
5. [ ] Add `contact-panel` div to taskpane HTML in `main.js`
6. [ ] Add minor CSS for contact panel states to `styles.css`
7. [ ] Test re-consent flow on next sign-in after scope change
8. [ ] Verify `Company_x0020_Type` field values in SharePoint Company List to align contact folder names

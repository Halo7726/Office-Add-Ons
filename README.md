# Office-Add-Ons

Outlook task pane add-in starter that authenticates with Microsoft Entra ID and calls Microsoft Graph to:

- Read SharePoint list items
- Create SharePoint list items
- Upload files to a SharePoint document library

Primary use case for this starter: add proposal records to project-related SharePoint lists and upload proposal documents into a project library folder.

This repo gives you a local Outlook task pane app that runs at `https://localhost:3000`, signs the user in with Microsoft Entra ID, and then uses Microsoft Graph with delegated permissions. The add-in does not use a backend service or client secret.

## What This Includes

- `manifest.xml` for Outlook desktop/web sideloading
- Vite HTTPS local dev app at `https://localhost:3000`
- MSAL browser auth (popup) with Graph delegated permissions
- Graph helpers for list and library operations

## Prerequisites

Before you start, make sure you have:

- Node.js 18 or newer
- An M365 account that can use Outlook add-ins
- Access to the target SharePoint site, list, and document library
- Permission to register an app in Entra ID, or help from your tenant admin

## 1. Register an App in Entra ID

1. Open Azure Portal -> Entra ID -> App registrations -> New registration.
2. Give the app a name such as `Office Add-Ons Local Dev`.
3. Choose who can sign in:
	- Single tenant: choose this if only your organization will use the add-in.
	- Multi-tenant: choose this only if you understand the broader consent and support implications.
4. Do not create a client secret. This project is a browser-based add-in using delegated sign-in.
5. After the app is created, open Authentication.
6. Add a platform of type `Single-page application`.
7. Add this redirect URI exactly: `https://localhost:3000`.
8. Save the Application (client) ID and Directory (tenant) ID.
9. Open API permissions and add Microsoft Graph delegated permissions:
	- `User.Read`
	- `Sites.ReadWrite.All`
	- `Files.ReadWrite.All`
10. Grant admin consent if your tenant requires it.

Notes:

- `Sites.ReadWrite.All` covers SharePoint list read and write operations.
- `Files.ReadWrite.All` covers uploading into the document library.
- If your admin will not allow these broad permissions, you will need a tighter production design, usually with a backend and app-specific authorization rules.

## 2. Configure This Project

1. Copy `.env.example` to `.env`.
2. Fill in these values:
	- `VITE_TENANT_ID`
	- `VITE_CLIENT_ID`
	- `VITE_SITE_ID`
	- `VITE_LIST_ID`
	- `VITE_DRIVE_ID`
	- `VITE_LIBRARY_FOLDER`

Example:

```env
VITE_TENANT_ID=your-tenant-id-or-common
VITE_CLIENT_ID=your-app-registration-client-id
VITE_SITE_ID=contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222
VITE_LIST_ID=33333333-3333-3333-3333-333333333333
VITE_DRIVE_ID=b!abcdefghijklmnopqrstuvwxyz1234567890
VITE_LIBRARY_FOLDER=Shared Documents/Incoming
```

What each setting means:

- `VITE_TENANT_ID`: your Entra tenant ID. Use `common` only if your app registration and tenant policy support it.
- `VITE_CLIENT_ID`: the Application (client) ID from the app registration.
- `VITE_SITE_ID`: the target SharePoint site ID.
- `VITE_LIST_ID`: the list you want to read from and write to.
- `VITE_DRIVE_ID`: the document library drive ID.
- `VITE_LIBRARY_FOLDER`: the folder path inside the library where uploads should land.

The task pane also lets you edit these values and click `Save config`. Those values are stored in local browser storage and override `.env` for that browser profile.

## 3. Find the SharePoint IDs

If you do not already know the site, list, and drive IDs, Graph Explorer is the fastest path.

### Find the site ID

1. Open Graph Explorer.
2. Sign in with an account that can access the SharePoint site.
3. Run:

```http
GET https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site-path}
```

Example:

```http
GET https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/Operations
```

4. Copy the `id` value from the response into `VITE_SITE_ID`.

### Find the list ID

Run:

```http
GET https://graph.microsoft.com/v1.0/sites/{site-id}/lists
```

Find the list you want by `displayName`, then copy its `id` into `VITE_LIST_ID`.

### Find the document library drive ID

Run:

```http
GET https://graph.microsoft.com/v1.0/sites/{site-id}/drives
```

Find the document library you want, then copy its `id` into `VITE_DRIVE_ID`.

### Choose the upload folder path

Use the path relative to the library root, for example:

- `Shared Documents`
- `Shared Documents/Incoming`
- `General/Attachments`

If the folder does not exist, the current starter does not create it for you. Create the folder first in SharePoint, or point the config at an existing folder.

## 4. Install and Run

```bash
npm install
npm run dev
```

What to expect:

- Vite serves the add-in UI over HTTPS at `https://localhost:3000`.
- Keep the dev server running while Outlook is using the add-in.
- The manifest is already pointed at `https://localhost:3000/index.html`.

If your browser or Outlook complains about the local HTTPS certificate, you need to trust the local dev certificate on your machine before sideloading the add-in.

## 5. Sideload in Outlook

1. Open Outlook on the web or new Outlook desktop.
2. Go to Get Add-ins -> My add-ins -> Add a custom add-in -> Add from file.
3. Select `manifest.xml` from this repo.
4. Open an email message, then click `SharePoint Bridge` on the ribbon.

If the add-in does not appear:

- Make sure the message is opened in read mode.
- Make sure the dev server is still running.
- Confirm the manifest still references `https://localhost:3000`.

## 6. First Run Inside Outlook

1. Open an email and launch `SharePoint Bridge`.
2. Verify the connection fields are populated.
3. Click `Save config` if you changed any values in the task pane.
4. Click `Sign in` and complete the Microsoft sign-in popup.
5. Click `Get list items` to verify Graph can read from the SharePoint list.
6. Enter JSON in the `New item fields` box and click `Add list item`.
7. Choose a file and click `Upload file` to send it into the configured document library folder.

Example list item payload:

```json
{
	"Title": "Proposal - ACME Network Upgrade",
	"Status": "Draft",
	"ProjectCode": "PRJ-1042"
}
```

The property names in the JSON must match the SharePoint internal field names, not always the labels you see in the SharePoint UI.

## 6.1 Proposal Workflow Setup (Recommended)

If your main scenario is "add proposals to projects", define your SharePoint list with proposal-focused columns and then map JSON keys to those internal names.

Suggested list columns:

- `Title` (Single line of text)
- `ProjectCode` (Single line of text)
- `ClientName` (Single line of text)
- `ProposalAmount` (Currency or Number)
- `DueDate` (Date)
- `Status` (Choice: Draft, Submitted, Approved, Rejected)
- `OwnerEmail` (Single line of text or Person)

Suggested document library folder pattern:

- `Shared Documents/Proposals/{ProjectCode}`

Example proposal payload:

```json
{
	"Title": "Proposal - South Campus Cabling",
	"ProjectCode": "PRJ-2078",
	"ClientName": "South Campus District",
	"ProposalAmount": 125000,
	"DueDate": "2026-04-20",
	"Status": "Draft",
	"OwnerEmail": "owner@contoso.com"
}
```

If you use a Person column for `OwnerEmail`, you may need to send the field as the internal People field format in Graph (for example `OwnerEmailLookupId`) depending on your list configuration.

## 7. Validate the Manifest

Run this any time you change the add-in manifest:

```bash
npm run validate-manifest
```

## Troubleshooting

### Sign-in popup fails or closes immediately

- Recheck the SPA redirect URI in Entra ID. It must be exactly `https://localhost:3000`.
- Make sure popup sign-in is allowed in your tenant and browser.
- Verify `VITE_CLIENT_ID` and `VITE_TENANT_ID` are correct.

### List reads fail with 403 or 401

- Confirm the signed-in user has access to the SharePoint site.
- Confirm delegated Graph permissions were added and consented.
- Confirm the `VITE_SITE_ID` and `VITE_LIST_ID` values are correct.

### File upload fails

- Confirm `VITE_DRIVE_ID` points to the correct document library.
- Confirm `VITE_LIBRARY_FOLDER` exists in that library.
- Confirm the user has upload permissions in SharePoint.

### The add-in opens but shows a blank pane

- Make sure `npm run dev` is still running.
- Open `https://localhost:3000` in a browser and verify the page loads.
- Re-sideload the manifest if Outlook cached an older copy.

## Important Security Notes

- This starter uses delegated user permissions via popup auth.
- Restrict permissions in production to least privilege.
- If your org blocks popup auth in add-ins, switch to a backend token exchange flow.
// Email template configuration.
//
// Each key is a recipient type (must match the <select> values in compose.js).
// Templates support the following tokens — all replaced at send time:
//
//   {{project_name}}   — project Title / ProjectName from SharePoint
//   {{project_number}} — JobNumber / EstimateNumber from SharePoint
//   {{company_name}}   — selected company Title / CompanyName
//   {{contact_name}}   — company "Contact Name 1" field (falls back to company name)
//   {{contact_email}}  — company "Email 1" field
//   {{contact_title}}  — company "Contact Title 1" field
//   {{today}}          — current date, e.g. "January 1, 2026"
//   {{sender_name}}    — signed-in user's display name (from Outlook profile)
//   {{sender_email}}   — signed-in user's email address

export const emailTemplates = {
  sub: {
    label: "Subcontractor – Invitation to Bid",
    subject: "Invitation to Bid – {{project_name}}",
    body: `Dear {{contact_name}},

We would like to invite {{company_name}} to submit a bid for the following project:

Project:    {{project_name}}
Project No: {{project_number}}

Please review the attached bid documents and return your proposal at your earliest convenience. Feel free to reach out with any questions.

We look forward to working with you.

Best regards,
{{sender_name}}
{{sender_email}}`,
  },

  vendor: {
    label: "Vendor – Request for Quote",
    subject: "Request for Quote – {{project_name}}",
    body: `Dear {{contact_name}},

We are requesting a quote from {{company_name}} for materials/services on the following project:

Project:    {{project_name}}
Project No: {{project_number}}

Please review the attached scope and provide your best pricing at your earliest convenience. Feel free to reach out with any questions.

We look forward to your response.

Best regards,
{{sender_name}}
{{sender_email}}`,
  },
};

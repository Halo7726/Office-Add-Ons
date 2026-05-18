export const emailTemplates = {
  invitationToBid: {
    label: "Invitation to Bid",
    subject: "Invitation to Bid – {{project_name}}",
    body: `Good morning,

We would like to invite you to bid on the following project:

Project:      {{project_name}}
Owner:        {{owner}}
Location:     {{project_location}}
Bid Due:      {{bid_due_date}}
Pre-Bid:      {{prebid_date}}
RFI Cut-off:  {{question_due_date}}
Plans:        {{plan_link}}
Takeoff:      {{takeoff_link}}

Please submit your bid by {{deadline}} and reach out with any questions.

Thanks,
{{sender_name}}
{{sender_email}}`,
  },
  requestForQuote: {
    label: "Request for Quote",
    subject: "Request for Quote – {{project_name}}",
    body: `Hello,

We are requesting a quote for the following project:

Project:      {{project_name}}
Owner:        {{owner}}
Location:     {{project_location}}
Bid Due:      {{bid_due_date}}
Pre-Bid:      {{prebid_date}}
RFI Cut-off:  {{question_due_date}}
Plans:        {{plan_link}}
Takeoff:      {{takeoff_link}}

Please review the attached scope and return your pricing by {{deadline}}.

Best,
{{sender_name}}
{{sender_email}}`,
  },
  proposalSummary: {
    label: "Project Proposal Summary",
    subject: "Proposal Summary – {{project_name}}",
    body: `Hi,

Here is a summary for the following project:

Project:      {{project_name}}
Owner:        {{owner}}
Location:     {{project_location}}
Bid Due:      {{bid_due_date}}
Pre-Bid:      {{prebid_date}}
RFI Cut-off:  {{question_due_date}}
Plans:        {{plan_link}}
Takeoff:      {{takeoff_link}}

We are targeting a bid submission by {{deadline}}.

If anything needs clarification, I can follow up directly.

Regards,
{{sender_name}}
{{sender_email}}`,
  },
  prebidReminder: {
    label: "Pre-Bid Meeting Reminder",
    subject: "Pre-Bid Meeting – {{project_name}}",
    body: `Hello,

This is a reminder about the upcoming pre-bid meeting for the following project:

Project:      {{project_name}}
Owner:        {{owner}}
Location:     {{project_location}}
Pre-Bid:      {{prebid_date}}
RFI Cut-off:  {{question_due_date}}
Bid Due:      {{bid_due_date}}
Plans:        {{plan_link}}
Takeoff:      {{takeoff_link}}

Please plan to attend and submit any questions before the RFI cut-off.

Thanks,
{{sender_name}}
{{sender_email}}`,
  },
};

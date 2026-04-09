function asPromise(asyncCall) {
  return new Promise((resolve, reject) => {
    asyncCall((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error?.message || "Office async call failed."));
      }
    });
  });
}

function extractAddress(fromValue) {
  if (!fromValue) return "";
  if (typeof fromValue === "string") return fromValue;
  return fromValue.emailAddress || fromValue.address || "";
}

export function hasOutlookContext() {
  return Boolean(window.Office?.context?.mailbox?.item);
}

export function getFallbackContext() {
  return {
    subject: "",
    from: "",
    senderDomain: "",
    attachments: [],
  };
}

export function getMessageContext() {
  if (!hasOutlookContext()) {
    return getFallbackContext();
  }

  const item = Office.context.mailbox.item;
  const from = extractAddress(item.from || item.sender);
  const senderDomain = from.includes("@") ? from.split("@")[1].toLowerCase() : "";

  const attachments = (item.attachments || []).map((attachment) => ({
    id: attachment.id,
    name: attachment.name,
    size: attachment.size,
    isInline: Boolean(attachment.isInline),
  }));

  return {
    subject: item.subject || "",
    from,
    senderDomain,
    attachments,
  };
}

function decodeBase64ToBytes(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

export async function getAttachmentBytes(attachmentId) {
  if (!hasOutlookContext()) {
    throw new Error("Attachment access requires Outlook context.");
  }

  const item = Office.context.mailbox.item;

  const content = await asPromise((callback) => {
    item.getAttachmentContentAsync(attachmentId, callback);
  });

  if (content.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
    return decodeBase64ToBytes(content.content);
  }

  throw new Error("Unsupported attachment format from Outlook. Expected Base64.");
}

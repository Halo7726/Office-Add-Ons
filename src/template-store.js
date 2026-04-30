import { emailTemplates } from "./email-templates";

const STORAGE_KEY = "sp-email-templates";

// Converts the static emailTemplates object into the array format used at runtime.
function seed() {
  return Object.entries(emailTemplates).map(([key, tpl]) => ({ key, ...tpl }));
}

export function loadTemplates() {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return seed();
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) && parsed.length > 0 ? parsed : seed();
  } catch {
    return seed();
  }
}

export function saveTemplates(templates) {
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(templates));
}

// Generates a unique key from the template label, avoiding collisions with existing keys.
export function generateKey(label, existing) {
  const base =
    String(label || "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "")
      .slice(0, 32) || "template";

  const keys = new Set(existing.map((t) => t.key));
  if (!keys.has(base)) return base;

  let n = 2;
  while (keys.has(`${base}_${n}`)) n++;
  return `${base}_${n}`;
}

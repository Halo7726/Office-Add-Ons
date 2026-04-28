function normalize(value) {
  return String(value || "").trim().toLowerCase();
}

function allStringValues(record) {
  const fields = record?.fields || {};
  return Object.values(fields)
    .filter((value) => typeof value === "string")
    .map((value) => normalize(value));
}

function chooseDisplayName(record) {
  const fields = record?.fields || {};
  return (
    fields.Title ||
    fields.ProjectName ||
    fields.ProjectCode ||
    fields.CompanyName ||
    fields.Name ||
    "(Unnamed)"
  );
}

function scoreRecord(searchText, record, preferredKeys = []) {
  const fields = record?.fields || {};
  const normalizedSearch = normalize(searchText);
  if (!normalizedSearch) return 0;

  let score = 0;

  for (const key of preferredKeys) {
    const value = normalize(fields[key]);
    if (!value) continue;

    if (normalizedSearch.includes(value)) {
      score += value.length > 5 ? 60 : 30;
    }
  }

  for (const value of allStringValues(record)) {
    if (!value || value.length < 4) continue;
    if (normalizedSearch.includes(value)) score += 10;
  }

  return score;
}

function topMatch(records, scoreFn) {
  let best = null;

  for (const record of records) {
    const score = scoreFn(record);
    if (!best || score > best.score) {
      best = { record, score };
    }
  }

  if (!best || best.score <= 0) return null;
  return best;
}

function rankMatches(records, scoreFn) {
  return records
    .map((record) => ({ record, score: scoreFn(record) }))
    .filter((item) => item.score > 0)
    .sort((left, right) => right.score - left.score);
}

function confidenceFromScores(projectScore, companyScore) {
  const bounded = Math.min(95, Math.max(5, projectScore + companyScore));
  return Math.round(bounded);
}

function sanitizeSegment(value, fallback) {
  return String(value || fallback || "Unknown")
    .replace(/[\\/:*?"<>|]/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}

// Resolves the actual SharePoint field key for a given set of candidate names,
// tolerating spaces, casing, and SharePoint's _x-encoded characters.
function findFieldKey(fields, candidates) {
  const keys = Object.keys(fields || {});
  const normalizedMap = new Map(
    keys.map((k) => [normalize(k).replace(/[^a-z0-9]/g, ""), k])
  );
  for (const candidate of candidates) {
    const normalized = normalize(candidate).replace(/[^a-z0-9]/g, "");
    if (normalizedMap.has(normalized)) return normalizedMap.get(normalized);
  }
  return null;
}

const ITB_RECIPIENT_EMAIL_KEYS = [
  "Recipient Email",
  "RecipientEmail",
  "Email",
  "VendorEmail",
  "SubEmail",
  "ContactEmail",
];

const ITB_EMAIL_SUBJECT_KEYS = [
  "Email Subject",
  "EmailSubject",
  "Subject",
  "ITBSubject",
];

export function resolveRoute(messageContext, projects, companies, config, options = {}) {
  const subject = messageContext?.subject || "";
  const from = messageContext?.from || "";
  const senderDomain = messageContext?.senderDomain || "";
  const companyIdOverride = options.companyIdOverride || null;

  const projectMatch = topMatch(projects, (record) =>
    scoreRecord(subject, record, ["ProjectCode", "Title", "ProjectName"])
  );

  const rankedCompanies = rankMatches(companies, (record) => {
    const domainScore = scoreRecord(senderDomain, record, ["Email", "Domain", "Website"]);
    const subjectScore = scoreRecord(subject, record, ["Title", "CompanyName"]);
    const fromScore = scoreRecord(from, record, ["Email"]);
    return domainScore + subjectScore + fromScore;
  });

  const defaultCompanyMatch = rankedCompanies[0] || null;

  const overrideRecord = companyIdOverride
    ? companies.find((record) => record.id === companyIdOverride) || null
    : null;

  const overrideRank = companyIdOverride
    ? rankedCompanies.find((item) => item.record.id === companyIdOverride) || null
    : null;

  const companyMatch =
    overrideRecord != null
      ? { record: overrideRecord, score: overrideRank?.score || 0 }
      : defaultCompanyMatch;

  const projectName = projectMatch ? chooseDisplayName(projectMatch.record) : "Unmapped Project";
  const subcontractorName = companyMatch ? chooseDisplayName(companyMatch.record) : "Unmapped Subcontractor";

  const template =
    config.folderTemplate || "Bids/Current/{project}/Subcontractors/{subcontractor}";

  const folderPath = template
    .replaceAll("{project}", sanitizeSegment(projectName, "Unmapped Project"))
    .replaceAll("{subcontractor}", sanitizeSegment(subcontractorName, "Unmapped Subcontractor"));

  const projectScore = projectMatch?.score || 0;
  const companyScore = companyMatch?.score || 0;
  const companyCandidates = rankedCompanies.slice(0, 12).map((item) => ({
    id: item.record.id,
    name: chooseDisplayName(item.record),
    score: item.score,
  }));

  if (companyMatch && !companyCandidates.some((item) => item.id === companyMatch.record.id)) {
    companyCandidates.unshift({
      id: companyMatch.record.id,
      name: chooseDisplayName(companyMatch.record),
      score: companyScore,
    });
  }

  return {
    project: {
      id: projectMatch?.record?.id || null,
      name: projectName,
      score: projectScore,
    },
    subcontractor: {
      id: companyMatch?.record?.id || null,
      name: subcontractorName,
      score: companyScore,
    },
    companyCandidates,
    confidence: confidenceFromScores(projectScore, companyScore),
    folderPath,
    reason: `${projectMatch ? "project matched" : "project fallback"}; ${
      companyMatch ? "subcontractor matched" : "subcontractor fallback"
    }`,
  };
}

// Searches the ITB/RFQ tracking list for a row where the Recipient Email matches
// the incoming email sender AND the Title/Email Subject fuzzy-matches the subject.
// Returns the matched item and its project name, or null if no confident match found.
export function resolveItbMatch(messageContext, itbItems, config) {
  if (!itbItems?.length) return null;

  const incomingSubject = normalize(messageContext?.subject || "");
  const senderEmail = normalize(messageContext?.from || "");

  if (!senderEmail) return null;

  let best = null;

  for (const item of itbItems) {
    const fields = item.fields || {};

    // Score: Recipient Email vs sender email
    const emailKey = findFieldKey(fields, ITB_RECIPIENT_EMAIL_KEYS);
    const recipientEmail = normalize(fields[emailKey] || "");

    let emailScore = 0;
    if (recipientEmail && senderEmail) {
      if (recipientEmail === senderEmail) {
        emailScore = 100;
      } else if (recipientEmail.includes(senderEmail) || senderEmail.includes(recipientEmail)) {
        emailScore = 60;
      }
    }

    // Require at least a partial email match — without it we can't confirm identity
    if (emailScore === 0) continue;

    // Score: ITB Email Subject or Title vs incoming email subject
    const subjectKey = findFieldKey(fields, ITB_EMAIL_SUBJECT_KEYS);
    const itbSubject = normalize(fields[subjectKey] || "");
    const itbTitle = normalize(fields.Title || fields.ProjectName || "");

    let titleScore = 0;
    if (itbTitle.length >= 4 && incomingSubject.includes(itbTitle)) {
      titleScore += itbTitle.length > 8 ? 60 : 30;
    }
    if (itbSubject.length >= 4 && incomingSubject.includes(itbSubject.slice(0, 40))) {
      titleScore += 30;
    }
    // Reverse: does the ITB subject contain significant words from the incoming subject?
    if (itbSubject.length >= 4 && itbSubject.includes(incomingSubject.slice(0, 20))) {
      titleScore += 10;
    }

    const total = emailScore + titleScore;
    if (!best || total > best.total || (total === best.total && emailScore > best.emailScore)) {
      best = { item, emailScore, titleScore, total };
    }
  }

  if (!best) return null;

  const fields = best.item.fields || {};
  const projectName = fields.Title || fields.ProjectName || fields.ProjectCode || "(Unnamed Project)";

  // confidence weighs email match heavily since it's the strong identifier
  const confidence = Math.min(95, Math.round(best.emailScore * 0.65 + best.titleScore * 0.35));

  return {
    item: best.item,
    projectName,
    emailScore: best.emailScore,
    titleScore: best.titleScore,
    confidence,
  };
}

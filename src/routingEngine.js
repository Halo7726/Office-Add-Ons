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

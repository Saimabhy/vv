import { ExtractionResult } from "../models/intermediate";
import { Contact, Org, Project, TargetModel } from "../models/target";
import { createIssueCollector, Issue } from "../models/report";

export interface NormalizeOutput {
  model: TargetModel;
  issues: Issue[];
}

// Merges raw per-source records into a deduplicated target model.
// Orgs are deduped by domain, Contacts by email, Projects by key.
export function normalize(extraction: ExtractionResult): NormalizeOutput {
  const issueCollector = createIssueCollector();

  const orgsByDomain = new Map<string, Org>();
  for (const raw of extraction.orgs) {
    const existing = orgsByDomain.get(raw.domain);
    if (existing) {
      if (!existing.sources.includes(raw.source)) existing.sources.push(raw.source);
    } else {
      orgsByDomain.set(raw.domain, {
        domain: raw.domain,
        displayName: raw.displayName,
        sources: [raw.source],
      });
    }
  }

  const contactsByEmail = new Map<string, Contact>();
  for (const raw of extraction.contacts) {
    const existing = contactsByEmail.get(raw.email);
    if (existing) {
      if (!existing.sources.includes(raw.source)) existing.sources.push(raw.source);
      if (!existing.displayName && raw.displayName) existing.displayName = raw.displayName;
    } else {
      if (!raw.displayName) {
        issueCollector.add("warning", `Contact ${raw.email} has no display name`, {
          email: raw.email,
          source: raw.source,
        });
      }
      contactsByEmail.set(raw.email, {
        email: raw.email,
        displayName: raw.displayName ?? raw.email,
        sources: [raw.source],
      });
    }
  }

  const projectsByKey = new Map<string, Project>();
  for (const raw of extraction.projects) {
    const existing = projectsByKey.get(raw.key);
    if (existing) {
      if (!existing.sources.includes(raw.source)) existing.sources.push(raw.source);
      if (!existing.sourceRefs.includes(raw.sourceRef)) existing.sourceRefs.push(raw.sourceRef);
    } else {
      projectsByKey.set(raw.key, {
        key: raw.key,
        name: raw.name,
        sources: [raw.source],
        sourceRefs: [raw.sourceRef],
      });
    }
  }

  if (extraction.assets.length === 0) {
    issueCollector.add("manual", "No assets extracted from OneDrive for this run");
  }

  return {
    model: {
      orgs: [...orgsByDomain.values()],
      contacts: [...contactsByEmail.values()],
      projects: [...projectsByKey.values()],
    },
    issues: issueCollector.all(),
  };
}

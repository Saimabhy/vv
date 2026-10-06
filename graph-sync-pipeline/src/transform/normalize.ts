import { ExtractionResult, RawContact } from "../models/intermediate";
import { Contact, Org, Project, TargetModel } from "../models/target";
import { createIssueCollector, IssueCollector, Issue } from "../models/report";

export interface NormalizeOutput {
  model: TargetModel;
  issues: Issue[];
}

// Merges raw per-source records into a deduplicated target model.
// Orgs are deduped by domain, Contacts by person, Projects by key.
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

  const contacts = mergeContacts(extraction.contacts, issueCollector);

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
      contacts,
      projects: [...projectsByKey.values()],
    },
    issues: issueCollector.all(),
  };
}

// The directory (Azure AD) is authoritative: each personId is one contact,
// however many emails it owns. Contacts from other sources (e.g. message
// senders) are matched against those known emails and merged in; only an
// email that matches nobody in the directory becomes its own external
// contact (e.g. a person outside the company).
function mergeContacts(rawContacts: RawContact[], issueCollector: IssueCollector): Contact[] {
  const contactsByPersonId = new Map<string, Contact>();
  const personIdByEmail = new Map<string, string>();

  for (const raw of rawContacts) {
    if (!raw.personId) continue;

    const existing = contactsByPersonId.get(raw.personId);
    if (existing) {
      if (!existing.sources.includes(raw.source)) existing.sources.push(raw.source);
      for (const email of raw.emails) {
        if (!existing.emails.includes(email)) existing.emails.push(email);
      }
    } else {
      contactsByPersonId.set(raw.personId, {
        id: raw.personId,
        displayName: raw.displayName ?? raw.emails[0],
        emails: [...raw.emails],
        sources: [raw.source],
      });
    }

    for (const email of raw.emails) personIdByEmail.set(email, raw.personId);
  }

  const externalContactsByEmail = new Map<string, Contact>();
  for (const raw of rawContacts) {
    if (raw.personId) continue;

    for (const email of raw.emails) {
      const personId = personIdByEmail.get(email);
      if (personId) {
        const directoryContact = contactsByPersonId.get(personId)!;
        if (!directoryContact.sources.includes(raw.source)) directoryContact.sources.push(raw.source);
        continue;
      }

      const existing = externalContactsByEmail.get(email);
      if (existing) {
        if (!existing.sources.includes(raw.source)) existing.sources.push(raw.source);
        continue;
      }

      if (!raw.displayName) {
        issueCollector.add("warning", `Contact ${email} has no display name`, {
          email,
          source: raw.source,
        });
      }
      externalContactsByEmail.set(email, {
        id: `email:${email}`,
        displayName: raw.displayName ?? email,
        emails: [email],
        sources: [raw.source],
      });
    }
  }

  return [...contactsByPersonId.values(), ...externalContactsByEmail.values()];
}

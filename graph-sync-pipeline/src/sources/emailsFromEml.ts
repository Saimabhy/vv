import { readdirSync, readFileSync } from "fs";
import { join } from "path";
import { simpleParser, AddressObject, EmailAddress } from "mailparser";
import { emptyExtractionResult, ExtractionResult } from "../models/intermediate";

function collectAddresses(field?: AddressObject | AddressObject[]): EmailAddress[] {
  if (!field) return [];
  const objects = Array.isArray(field) ? field : [field];
  return objects.flatMap((o) => o.value);
}

// Extracts Contacts (message participants) and Org (their domains) from a
// folder of manually exported .eml files — no Microsoft Graph API/credentials
// needed. Each .eml is a single email, exported via "Save as" from
// Outlook/Gmail/etc.
export async function extractEmailsFromEmlFolder(folderPath: string): Promise<ExtractionResult> {
  const result = emptyExtractionResult();
  const seenEmails = new Set<string>();
  const seenDomains = new Set<string>();

  const files = readdirSync(folderPath).filter((f) => f.toLowerCase().endsWith(".eml"));
  if (files.length === 0) {
    throw new Error(`No .eml files found in ${folderPath}`);
  }

  for (const file of files) {
    const raw = readFileSync(join(folderPath, file));
    const parsed = await simpleParser(raw);

    const addresses = [...collectAddresses(parsed.from), ...collectAddresses(parsed.to)];

    for (const addr of addresses) {
      if (!addr.address) continue;
      const email = addr.address.toLowerCase();

      if (!seenEmails.has(email)) {
        seenEmails.add(email);
        result.contacts.push({
          emails: [email],
          displayName: addr.name || undefined,
          source: "emails",
        });
      }

      const domain = email.split("@")[1];
      if (domain && !seenDomains.has(domain)) {
        seenDomains.add(domain);
        result.orgs.push({ domain, displayName: domain, source: "emails" });
      }
    }
  }

  return result;
}

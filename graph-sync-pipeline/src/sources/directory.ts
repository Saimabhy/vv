import { Client } from "@microsoft/microsoft-graph-client";
import { emptyExtractionResult, ExtractionResult } from "../models/intermediate";

interface DirectoryUser {
  id: string;
  displayName?: string;
  mail?: string;
  otherMails?: string[];
  proxyAddresses?: string[];
}

const SMTP_PROXY = /^smtp:(.+)$/i;

// Extracts the company directory from Azure AD. This is the authoritative
// source for Contacts: each user has one stable id, but can own several
// email addresses (primary mail, aliases, proxy addresses) that all belong
// to the same person.
export async function extractDirectory(client: Client): Promise<ExtractionResult> {
  const result = emptyExtractionResult();

  const response = await client
    .api("/users")
    .select("id,displayName,mail,otherMails,proxyAddresses")
    .get();

  const users: DirectoryUser[] = response.value ?? [];
  for (const user of users) {
    const emails = new Set<string>();
    if (user.mail) emails.add(user.mail.toLowerCase());
    for (const mail of user.otherMails ?? []) emails.add(mail.toLowerCase());
    for (const proxy of user.proxyAddresses ?? []) {
      const match = SMTP_PROXY.exec(proxy);
      if (match) emails.add(match[1].toLowerCase());
    }

    if (emails.size === 0) continue;

    result.contacts.push({
      personId: user.id,
      emails: [...emails],
      displayName: user.displayName,
      source: "directory",
    });
  }

  return result;
}

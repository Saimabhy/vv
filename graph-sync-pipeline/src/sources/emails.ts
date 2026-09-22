import { Client } from "@microsoft/microsoft-graph-client";
import { emptyExtractionResult, ExtractionResult } from "../models/intermediate";

interface Recipient {
  emailAddress?: { address?: string; name?: string };
}

interface Message {
  id: string;
  subject: string;
  from?: Recipient;
  toRecipients?: Recipient[];
}

// Extracts Contacts (message participants) and Org (their domains) from a
// user's mailbox. Projects are not derived from emails yet — the whiteboard
// shows this feeding the intermediate model only.
export async function extractEmails(
  client: Client,
  userId: string
): Promise<ExtractionResult> {
  const result = emptyExtractionResult();
  const seenEmails = new Set<string>();
  const seenDomains = new Set<string>();

  const response = await client
    .api(`/users/${userId}/messages`)
    .select("id,subject,from,toRecipients")
    .top(50)
    .get();

  const messages: Message[] = response.value ?? [];

  const addRecipient = (recipient?: Recipient) => {
    const address = recipient?.emailAddress?.address;
    if (!address) return;

    const email = address.toLowerCase();
    if (!seenEmails.has(email)) {
      seenEmails.add(email);
      result.contacts.push({
        email,
        displayName: recipient?.emailAddress?.name,
        source: "emails",
      });
    }

    const domain = email.split("@")[1];
    if (domain && !seenDomains.has(domain)) {
      seenDomains.add(domain);
      result.orgs.push({ domain, displayName: domain, source: "emails" });
    }
  };

  for (const message of messages) {
    addRecipient(message.from);
    (message.toRecipients ?? []).forEach(addRecipient);
  }

  return result;
}

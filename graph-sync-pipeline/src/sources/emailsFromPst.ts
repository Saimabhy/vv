import { PSTFile, PSTFolder, PSTMessage } from "pst-extractor";
import { emptyExtractionResult, ExtractionResult } from "../models/intermediate";

interface FoundAddress {
  email: string;
  displayName?: string;
}

function collectFolder(folder: PSTFolder, found: FoundAddress[]): void {
  if (folder.hasSubfolders) {
    for (const child of folder.getSubFolders()) {
      collectFolder(child, found);
    }
  }

  if (folder.contentCount > 0) {
    let email = folder.getNextChild() as PSTMessage | null;
    while (email != null) {
      if (email.senderEmailAddress) {
        found.push({ email: email.senderEmailAddress, displayName: email.senderName || undefined });
      }
      for (let i = 0; i < email.numberOfRecipients; i++) {
        const recipient = email.getRecipient(i);
        const address = recipient?.smtpAddress || recipient?.emailAddress;
        if (address) {
          found.push({ email: address, displayName: recipient?.displayName || undefined });
        }
      }
      email = folder.getNextChild() as PSTMessage | null;
    }
  }
}

// Extracts Contacts (message participants) and Org (their domains) from a
// .pst file exported via Outlook's "Import/Export > Export to a file"
// wizard. No Microsoft Graph API/credentials needed — covers every message
// in the exported folder(s) in one pass, unlike one .eml file per email.
export async function extractEmailsFromPst(pstFilePath: string): Promise<ExtractionResult> {
  const result = emptyExtractionResult();
  const seenEmails = new Set<string>();
  const seenDomains = new Set<string>();

  const pstFile = new PSTFile(pstFilePath);
  const found: FoundAddress[] = [];
  collectFolder(pstFile.getRootFolder(), found);

  for (const { email: rawEmail, displayName } of found) {
    const email = rawEmail.toLowerCase();
    if (!email.includes("@")) continue;

    if (!seenEmails.has(email)) {
      seenEmails.add(email);
      result.contacts.push({ emails: [email], displayName, source: "emails" });
    }

    const domain = email.split("@")[1];
    if (domain && !seenDomains.has(domain)) {
      seenDomains.add(domain);
      result.orgs.push({ domain, displayName: domain, source: "emails" });
    }
  }

  return result;
}

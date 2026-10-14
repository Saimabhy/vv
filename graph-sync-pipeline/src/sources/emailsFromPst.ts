import { PSTFile, PSTFolder, PSTMessage } from "pst-extractor";
import { emptyExtractionResult, ExtractionResult } from "../models/intermediate";

interface CollectState {
  result: ExtractionResult;
  seenEmails: Set<string>;
  seenDomains: Set<string>;
  processed: number;
}

function addAddress(state: CollectState, rawEmail: string, displayName?: string): void {
  const email = rawEmail.toLowerCase();
  if (!email.includes("@")) return;

  if (!state.seenEmails.has(email)) {
    state.seenEmails.add(email);
    state.result.contacts.push({ emails: [email], displayName, source: "emails" });
  }

  const domain = email.split("@")[1];
  if (domain && !state.seenDomains.has(domain)) {
    state.seenDomains.add(domain);
    state.result.orgs.push({ domain, displayName: domain, source: "emails" });
  }
}

function collectFolder(folder: PSTFolder, state: CollectState): void {
  if (folder.hasSubfolders) {
    for (const child of folder.getSubFolders()) {
      collectFolder(child, state);
    }
  }

  if (folder.contentCount > 0) {
    console.log(`  [pst] scanning folder "${folder.displayName}" (${folder.contentCount} items)...`);
    let email = folder.getNextChild() as PSTMessage | null;
    while (email != null) {
      if (email.senderEmailAddress) {
        addAddress(state, email.senderEmailAddress, email.senderName || undefined);
      }
      for (let i = 0; i < email.numberOfRecipients; i++) {
        const recipient = email.getRecipient(i);
        const address = recipient?.smtpAddress || recipient?.emailAddress;
        if (address) {
          addAddress(state, address, recipient?.displayName || undefined);
        }
      }

      state.processed++;
      if (state.processed % 200 === 0) {
        console.log(
          `  [pst] ${state.processed} messages processed, ${state.result.contacts.length} unique contacts so far...`
        );
      }

      email = folder.getNextChild() as PSTMessage | null;
    }
  }
}

// Extracts Contacts (message participants) and Org (their domains) from a
// .pst file exported via Outlook's "Import/Export > Export to a file"
// wizard. No Microsoft Graph API/credentials needed — covers every message
// in the exported folder(s) in one pass, unlike one .eml file per email.
// Dedupes as it scans rather than collecting every raw occurrence first, to
// keep memory use down on large mailboxes.
export async function extractEmailsFromPst(pstFilePath: string): Promise<ExtractionResult> {
  const state: CollectState = {
    result: emptyExtractionResult(),
    seenEmails: new Set<string>(),
    seenDomains: new Set<string>(),
    processed: 0,
  };

  console.log(`[pst] opening ${pstFilePath}...`);
  const pstFile = new PSTFile(pstFilePath);
  collectFolder(pstFile.getRootFolder(), state);
  console.log(
    `[pst] done — ${state.processed} messages, ${state.result.contacts.length} unique contacts, ${state.result.orgs.length} orgs`
  );

  return state.result;
}

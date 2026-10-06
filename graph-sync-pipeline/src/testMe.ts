import { createGraphClientDelegated } from "./graph/graphClient";
import { extractOneDrive } from "./sources/oneDrive";
import { extractEmails } from "./sources/emails";
import { mergeExtractionResults } from "./models/intermediate";
import { normalize } from "./transform/normalize";

// Standalone entry point to test the pipeline against just your own mailbox,
// using delegated (device code) auth. No admin consent or client secret
// needed. Does not include the company directory or SharePoint — see
// runPipeline (index.ts) for the full app-only pipeline.
async function main() {
  const client = createGraphClientDelegated();

  console.log("Signing in — follow the instructions above to authenticate...\n");

  const [oneDriveResult, emailsResult] = await Promise.all([
    extractOneDrive(client, "me"),
    extractEmails(client, "me"),
  ]);

  const merged = mergeExtractionResults([oneDriveResult, emailsResult]);
  const { model, issues } = normalize(merged);

  console.log(`\nAssets found: ${oneDriveResult.assets.length}`);
  for (const asset of oneDriveResult.assets) console.log(`  - ${asset.name}`);

  console.log(`\nContacts found: ${model.contacts.length}`);
  for (const contact of model.contacts) {
    console.log(`  - ${contact.displayName} (${contact.emails.join(", ")})`);
  }

  console.log(`\nOrgs found: ${model.orgs.length}`);
  for (const org of model.orgs) console.log(`  - ${org.displayName}`);

  if (issues.length > 0) {
    console.log(`\n${issues.length} issue(s):`);
    for (const issue of issues) console.log(`  [${issue.level}] ${issue.message}`);
  }
}

main().catch((error) => {
  console.error("Test run failed:", error);
  process.exitCode = 1;
});

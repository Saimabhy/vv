import { extractEmailsFromPst } from "./sources/emailsFromPst";
import { normalize } from "./transform/normalize";
import { exportToCsv } from "./export/csv";

// Standalone entry point: builds the directory from a single .pst file
// exported via Outlook's Import/Export wizard. No Azure/Graph credentials
// needed at all, and covers an entire mailbox folder in one export.
async function main() {
  const pstPath = process.argv[2];
  if (!pstPath) {
    console.error("Usage: npm run test-pst -- <path-to-file.pst>");
    process.exitCode = 1;
    return;
  }

  const extraction = await extractEmailsFromPst(pstPath);
  const { model, issues } = normalize(extraction);

  console.log(`Orgs: ${model.orgs.length}`);
  for (const o of model.orgs) console.log(`  - ${o.displayName}`);

  console.log(`\nContacts: ${model.contacts.length}`);
  for (const c of model.contacts) {
    console.log(`  - ${c.displayName} (${c.emails.join(", ")})`);
  }

  if (issues.length > 0) {
    console.log(`\n${issues.length} issue(s):`);
    for (const issue of issues) console.log(`  [${issue.level}] ${issue.message}`);
  }

  exportToCsv(model, "output");
  console.log("\nExported to output/contacts.csv, orgs.csv, projects.csv");
}

main().catch((error) => {
  console.error("Test run failed:", error);
  process.exitCode = 1;
});

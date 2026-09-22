import { runPipeline } from "./pipeline";

async function main() {
  const { model, issues } = await runPipeline();

  console.log(`Orgs: ${model.orgs.length}`);
  console.log(`Contacts: ${model.contacts.length}`);
  console.log(`Projects: ${model.projects.length}`);

  if (issues.length > 0) {
    console.log(`\n${issues.length} issue(s):`);
    for (const issue of issues) {
      console.log(`  [${issue.level}] ${issue.message}`);
    }
  }
}

main().catch((error) => {
  console.error("Pipeline failed:", error);
  process.exitCode = 1;
});

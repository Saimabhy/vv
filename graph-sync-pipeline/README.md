# graph-sync-pipeline

Extracts data from Microsoft 365 (OneDrive, Emails, SharePoint) via Microsoft
Graph, and normalizes it into a deduplicated target model.

## How it maps to the whiteboard sketch

- **Sources**: OneDrive, Emails, SharePoint.
- **Intermediate model** (`src/models/intermediate.ts`): raw, per-source
  `Org` / `Contacts` / `Projects` / `Assets` records, produced by OneDrive +
  Emails extraction (`src/sources/oneDrive.ts`, `src/sources/emails.ts`).
- **"xN" loop**: the pipeline runs OneDrive + Emails extraction once per
  entry in `GRAPH_TARGET_USERS` (`src/pipeline.ts`).
- **SharePoint path**: bypasses the intermediate merge and feeds `Projects`
  directly into the final model, as drawn on the whiteboard
  (`src/sources/sharePoint.ts`).
- **Target model** (`src/models/target.ts`): deduplicated `Org` / `Contacts`
  / `Projects` — Orgs deduped by domain, Contacts by email, Projects by key
  (`src/transform/normalize.ts`).
- **Red markers (circled "M", "!" triangle, "!" circle)**: not yet defined.
  `src/models/report.ts` adds a placeholder `manual` / `warning` / `error`
  issue log so normalization has somewhere to record problems instead of
  silently dropping data — revisit once the intended meaning is settled.

## Setup

```bash
npm install
cp .env.example .env
# fill in AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET,
# GRAPH_TARGET_USERS and GRAPH_SHAREPOINT_SITE
```

The Azure AD app registration needs application permissions (client
credentials flow) for at least: `User.Read.All`, `Mail.Read`, `Files.Read.All`,
`Sites.Read.All`, granted with admin consent.

## Run

```bash
npm start
```

Prints the counts of orgs/contacts/projects in the final model and any
issues collected during normalization.

## Open questions

- Exact semantics of the red markers on the whiteboard (M / warning / error).
- Whether Projects should also be derived from OneDrive/Emails, or only from
  SharePoint as currently implemented.

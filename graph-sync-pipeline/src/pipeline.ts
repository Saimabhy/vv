import { Client } from "@microsoft/microsoft-graph-client";
import { createGraphClient } from "./graph/graphClient";
import { config } from "./config";
import { mergeExtractionResults } from "./models/intermediate";
import { extractDirectory } from "./sources/directory";
import { extractOneDrive } from "./sources/oneDrive";
import { extractEmails } from "./sources/emails";
import { extractSharePoint, resolveDefaultListId, resolveSiteId } from "./sources/sharePoint";
import { normalize, NormalizeOutput } from "./transform/normalize";

// Pulls the company directory once (authoritative source for Contacts), then
// runs OneDrive + Emails extraction once per target user (the "xN" loop on
// the whiteboard), and merges in Projects pulled directly from SharePoint.
export async function runPipeline(): Promise<NormalizeOutput> {
  if (!config.sharePointSite) {
    throw new Error("GRAPH_SHAREPOINT_SITE is required for the full pipeline (runPipeline)");
  }

  const client: Client = createGraphClient();

  const directoryResult = await extractDirectory(client);

  const perUserResults = [directoryResult];
  for (const userId of config.targetUsers) {
    const [oneDriveResult, emailsResult] = await Promise.all([
      extractOneDrive(client, userId),
      extractEmails(client, userId),
    ]);
    perUserResults.push(oneDriveResult, emailsResult);
  }

  const siteId = await resolveSiteId(client, config.sharePointSite);
  const listId = await resolveDefaultListId(client, siteId);
  const sharePointResult = await extractSharePoint(client, siteId, listId);
  perUserResults.push(sharePointResult);

  const merged = mergeExtractionResults(perUserResults);
  return normalize(merged);
}

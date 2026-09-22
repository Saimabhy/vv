import { Client } from "@microsoft/microsoft-graph-client";
import { createGraphClient } from "./graph/graphClient";
import { config } from "./config";
import { mergeExtractionResults } from "./models/intermediate";
import { extractOneDrive } from "./sources/oneDrive";
import { extractEmails } from "./sources/emails";
import { extractSharePoint, resolveDefaultListId, resolveSiteId } from "./sources/sharePoint";
import { normalize, NormalizeOutput } from "./transform/normalize";

// Runs the OneDrive + Emails extraction once per target user (the "xN" loop
// on the whiteboard), then merges in Projects pulled directly from SharePoint.
export async function runPipeline(): Promise<NormalizeOutput> {
  const client: Client = createGraphClient();

  const perUserResults = [];
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

import { Client } from "@microsoft/microsoft-graph-client";
import { emptyExtractionResult, ExtractionResult } from "../models/intermediate";

interface DriveItem {
  id: string;
  name: string;
  webUrl: string;
}

// Extracts Assets for a given user's OneDrive. Pass userId "me" to read the
// signed-in user's own OneDrive under delegated (device code) auth.
export async function extractOneDrive(
  client: Client,
  userId: string
): Promise<ExtractionResult> {
  const result = emptyExtractionResult();

  const basePath = userId === "me" ? "/me" : `/users/${userId}`;
  const response = await client
    .api(`${basePath}/drive/root/children`)
    .select("id,name,webUrl")
    .get();

  const items: DriveItem[] = response.value ?? [];
  for (const item of items) {
    result.assets.push({
      id: item.id,
      name: item.name,
      webUrl: item.webUrl,
      source: "oneDrive",
    });
  }

  return result;
}

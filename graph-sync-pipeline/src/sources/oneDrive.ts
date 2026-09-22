import { Client } from "@microsoft/microsoft-graph-client";
import { emptyExtractionResult, ExtractionResult } from "../models/intermediate";

interface DriveItem {
  id: string;
  name: string;
  webUrl: string;
}

// Extracts Assets for a given user's OneDrive.
export async function extractOneDrive(
  client: Client,
  userId: string
): Promise<ExtractionResult> {
  const result = emptyExtractionResult();

  const response = await client
    .api(`/users/${userId}/drive/root/children`)
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

import { Client } from "@microsoft/microsoft-graph-client";
import { emptyExtractionResult, ExtractionResult } from "../models/intermediate";

interface ListItemFields {
  Title?: string;
}

interface ListItem {
  id: string;
  webUrl: string;
  fields?: ListItemFields;
}

// Extracts Projects directly from a SharePoint site's default list.
// On the whiteboard, SharePoint feeds the target model directly rather than
// going through the OneDrive/Emails intermediate merge step.
export async function extractSharePoint(
  client: Client,
  siteId: string,
  listId: string
): Promise<ExtractionResult> {
  const result = emptyExtractionResult();

  const response = await client
    .api(`/sites/${siteId}/lists/${listId}/items`)
    .expand("fields(select=Title)")
    .get();

  const items: ListItem[] = response.value ?? [];
  for (const item of items) {
    const name = item.fields?.Title ?? item.id;
    result.projects.push({
      key: `sharepoint:${item.id}`,
      name,
      source: "sharePoint",
      sourceRef: item.webUrl,
    });
  }

  return result;
}

export async function resolveSiteId(client: Client, siteHostAndPath: string): Promise<string> {
  const site = await client.api(`/sites/${siteHostAndPath}`).get();
  return site.id;
}

export async function resolveDefaultListId(client: Client, siteId: string): Promise<string> {
  const response = await client
    .api(`/sites/${siteId}/lists`)
    .filter("displayName eq 'Projects'")
    .get();

  const list = (response.value ?? [])[0];
  if (!list) {
    throw new Error(`No 'Projects' list found on site ${siteId}`);
  }
  return list.id;
}

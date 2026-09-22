export type SourceSystem = "oneDrive" | "emails" | "sharePoint";

export interface RawOrg {
  domain: string;
  displayName: string;
  source: SourceSystem;
}

export interface RawContact {
  email: string;
  displayName?: string;
  source: SourceSystem;
}

export interface RawProject {
  key: string;
  name: string;
  source: SourceSystem;
  sourceRef: string;
}

export interface RawAsset {
  id: string;
  name: string;
  webUrl: string;
  source: SourceSystem;
}

export interface ExtractionResult {
  orgs: RawOrg[];
  contacts: RawContact[];
  projects: RawProject[];
  assets: RawAsset[];
}

export function emptyExtractionResult(): ExtractionResult {
  return { orgs: [], contacts: [], projects: [], assets: [] };
}

export function mergeExtractionResults(results: ExtractionResult[]): ExtractionResult {
  return results.reduce<ExtractionResult>((acc, result) => {
    acc.orgs.push(...result.orgs);
    acc.contacts.push(...result.contacts);
    acc.projects.push(...result.projects);
    acc.assets.push(...result.assets);
    return acc;
  }, emptyExtractionResult());
}

import { SourceSystem } from "./intermediate";

export interface Org {
  domain: string;
  displayName: string;
  sources: SourceSystem[];
}

export interface Contact {
  email: string;
  displayName: string;
  sources: SourceSystem[];
}

export interface Project {
  key: string;
  name: string;
  sources: SourceSystem[];
  sourceRefs: string[];
}

export interface TargetModel {
  orgs: Org[];
  contacts: Contact[];
  projects: Project[];
}

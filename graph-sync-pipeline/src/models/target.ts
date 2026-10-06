import { SourceSystem } from "./intermediate";

export interface Org {
  domain: string;
  displayName: string;
  sources: SourceSystem[];
}

export interface Contact {
  // Azure AD user id for a directory-matched person, or "email:<address>"
  // for someone only seen in message traffic (e.g. an external contact).
  id: string;
  displayName: string;
  emails: string[];
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

import { writeFileSync, mkdirSync } from "fs";
import { join } from "path";
import { TargetModel } from "../models/target";

function escapeCsvField(value: string): string {
  if (/[",\n]/.test(value)) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}

function toCsv(headers: string[], rows: string[][]): string {
  const lines = [headers, ...rows].map((row) => row.map(escapeCsvField).join(","));
  return lines.join("\n") + "\n";
}

// Writes the directory as CSV files, ready to share (e.g. import into Excel).
export function exportToCsv(model: TargetModel, outDir: string): void {
  mkdirSync(outDir, { recursive: true });

  const contactsCsv = toCsv(
    ["id", "displayName", "emails", "sources"],
    model.contacts.map((c) => [c.id, c.displayName, c.emails.join("; "), c.sources.join("; ")])
  );
  writeFileSync(join(outDir, "contacts.csv"), contactsCsv);

  const orgsCsv = toCsv(
    ["domain", "displayName", "sources"],
    model.orgs.map((o) => [o.domain, o.displayName, o.sources.join("; ")])
  );
  writeFileSync(join(outDir, "orgs.csv"), orgsCsv);

  const projectsCsv = toCsv(
    ["key", "name", "sources", "sourceRefs"],
    model.projects.map((p) => [p.key, p.name, p.sources.join("; "), p.sourceRefs.join("; ")])
  );
  writeFileSync(join(outDir, "projects.csv"), projectsCsv);
}

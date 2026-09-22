// Placeholder for the red markers on the whiteboard (circled "M", triangle "!",
// circled "!"). Their exact meaning wasn't defined yet, so this just gives the
// normalize step somewhere to record issues instead of silently dropping data.
// Revisit the level names/semantics once that's settled.
export type IssueLevel = "manual" | "warning" | "error";

export interface Issue {
  level: IssueLevel;
  message: string;
  context?: Record<string, unknown>;
}

export function createIssueCollector() {
  const issues: Issue[] = [];
  return {
    add(level: IssueLevel, message: string, context?: Record<string, unknown>) {
      issues.push({ level, message, context });
    },
    all(): Issue[] {
      return issues;
    },
  };
}

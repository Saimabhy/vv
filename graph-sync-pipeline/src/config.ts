import "dotenv/config";

function required(name: string): string {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

// Only tenantId + clientId are always required (needed by both the delegated
// "test my own account" flow and the full app-only pipeline). clientSecret,
// targetUsers and sharePointSite are only needed for the full app-only
// pipeline (runPipeline), which validates them itself when it runs.
export const config = {
  azure: {
    tenantId: required("AZURE_TENANT_ID"),
    clientId: required("AZURE_CLIENT_ID"),
    clientSecret: process.env.AZURE_CLIENT_SECRET,
  },
  targetUsers: (process.env.GRAPH_TARGET_USERS ?? "")
    .split(",")
    .map((u) => u.trim())
    .filter(Boolean),
  sharePointSite: process.env.GRAPH_SHAREPOINT_SITE,
};

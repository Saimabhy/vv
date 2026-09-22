import "dotenv/config";

function required(name: string): string {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

export const config = {
  azure: {
    tenantId: required("AZURE_TENANT_ID"),
    clientId: required("AZURE_CLIENT_ID"),
    clientSecret: required("AZURE_CLIENT_SECRET"),
  },
  targetUsers: required("GRAPH_TARGET_USERS")
    .split(",")
    .map((u) => u.trim())
    .filter(Boolean),
  sharePointSite: required("GRAPH_SHAREPOINT_SITE"),
};

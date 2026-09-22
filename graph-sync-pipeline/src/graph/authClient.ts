import { ConfidentialClientApplication } from "@azure/msal-node";
import { config } from "../config";

const msalApp = new ConfidentialClientApplication({
  auth: {
    clientId: config.azure.clientId,
    authority: `https://login.microsoftonline.com/${config.azure.tenantId}`,
    clientSecret: config.azure.clientSecret,
  },
});

const GRAPH_SCOPE = "https://graph.microsoft.com/.default";

export async function getGraphAccessToken(): Promise<string> {
  const result = await msalApp.acquireTokenByClientCredential({
    scopes: [GRAPH_SCOPE],
  });
  if (!result?.accessToken) {
    throw new Error("Failed to acquire Microsoft Graph access token");
  }
  return result.accessToken;
}

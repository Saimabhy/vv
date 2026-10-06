import { ConfidentialClientApplication, PublicClientApplication } from "@azure/msal-node";
import { config } from "../config";

const GRAPH_SCOPE = "https://graph.microsoft.com/.default";

// App-only auth (client credentials). Needs a client secret and admin
// consent for the application permissions granted to the app registration.
// Used by the full pipeline (runPipeline), which can read any user's mailbox
// and the whole company directory.
export async function getGraphAccessToken(): Promise<string> {
  if (!config.azure.clientSecret) {
    throw new Error("AZURE_CLIENT_SECRET is required for app-only auth (runPipeline)");
  }

  const msalApp = new ConfidentialClientApplication({
    auth: {
      clientId: config.azure.clientId,
      authority: `https://login.microsoftonline.com/${config.azure.tenantId}`,
      clientSecret: config.azure.clientSecret,
    },
  });

  const result = await msalApp.acquireTokenByClientCredential({
    scopes: [GRAPH_SCOPE],
  });
  if (!result?.accessToken) {
    throw new Error("Failed to acquire Microsoft Graph access token");
  }
  return result.accessToken;
}

// Delegated auth (device code flow): the caller signs in with their own
// Microsoft account and consents for themselves, so no admin consent or
// client secret is needed in tenants that allow user consent. Only grants
// access to the signed-in user's own mailbox/OneDrive (via /me), not other
// users or the company directory.
const DELEGATED_SCOPES = ["Mail.Read", "Files.Read", "User.Read"];

export async function getGraphAccessTokenDelegated(): Promise<string> {
  const pca = new PublicClientApplication({
    auth: {
      clientId: config.azure.clientId,
      authority: `https://login.microsoftonline.com/${config.azure.tenantId}`,
    },
  });

  const result = await pca.acquireTokenByDeviceCode({
    scopes: DELEGATED_SCOPES,
    deviceCodeCallback: (response) => {
      console.log(response.message);
    },
  });
  if (!result?.accessToken) {
    throw new Error("Failed to acquire delegated Microsoft Graph access token");
  }
  return result.accessToken;
}

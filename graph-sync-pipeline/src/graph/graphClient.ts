import "isomorphic-fetch";
import { Client } from "@microsoft/microsoft-graph-client";
import { getGraphAccessToken, getGraphAccessTokenDelegated } from "./authClient";

export function createGraphClient(): Client {
  return Client.init({
    authProvider: async (done) => {
      try {
        const token = await getGraphAccessToken();
        done(null, token);
      } catch (error) {
        done(error as Error, null);
      }
    },
  });
}

// Same as createGraphClient, but delegated (device code) auth. Caches the
// token in memory for the life of the process so the device-code sign-in
// prompt only happens once per run, not once per API call.
export function createGraphClientDelegated(): Client {
  let cachedToken: string | undefined;

  return Client.init({
    authProvider: async (done) => {
      try {
        if (!cachedToken) {
          cachedToken = await getGraphAccessTokenDelegated();
        }
        done(null, cachedToken);
      } catch (error) {
        done(error as Error, null);
      }
    },
  });
}

import "isomorphic-fetch";
import { Client } from "@microsoft/microsoft-graph-client";
import { getGraphAccessToken } from "./authClient";

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

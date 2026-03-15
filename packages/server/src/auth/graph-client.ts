import { Client } from "@microsoft/microsoft-graph-client";
import type { ConfidentialClientApplication } from "@azure/msal-node";
import { getAccessToken } from "./msal-client.js";

export function createGraphClient(msal: ConfidentialClientApplication): Client {
  return Client.initWithMiddleware({
    authProvider: {
      getAccessToken: () => getAccessToken(msal),
    },
  });
}

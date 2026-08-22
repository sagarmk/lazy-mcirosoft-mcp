import { ConfidentialClientApplication } from "@azure/msal-node";
import type { AuthConfig } from "lazy-ms-graph-mcp-shared";
import { GRAPH_SCOPES } from "lazy-ms-graph-mcp-shared";

export function createMsalClient(config: AuthConfig): ConfidentialClientApplication {
  return new ConfidentialClientApplication({
    auth: {
      clientId: config.clientId,
      clientSecret: config.clientSecret,
      authority: `https://login.microsoftonline.com/${config.tenantId}`,
    },
  });
}

export async function getAccessToken(msal: ConfidentialClientApplication): Promise<string> {
  const result = await msal.acquireTokenByClientCredential({
    scopes: GRAPH_SCOPES,
  });
  if (!result || !result.accessToken) {
    throw new Error("Failed to acquire access token from MSAL");
  }
  return result.accessToken;
}

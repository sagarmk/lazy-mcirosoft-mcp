#!/usr/bin/env node

import { startServer } from "./server.js";
import { logger } from "./utils/logger.js";

function getConfig() {
  const clientId = process.env.MSGRAPH_CLIENT_ID || process.env.MICROSOFT_MCP_CLIENT_ID || "";
  const clientSecret = process.env.MSGRAPH_CLIENT_SECRET || process.env.MICROSOFT_MCP_CLIENT_SECRET || "";
  const tenantId = process.env.MSGRAPH_TENANT_ID || process.env.MICROSOFT_MCP_TENANT_ID || "";
  const port = parseInt(process.env.MSGRAPH_MCP_PORT || process.env.MICROSOFT_MCP_PORT || "3100", 10);
  const transport = (process.env.MSGRAPH_MCP_TRANSPORT || process.env.MICROSOFT_MCP_TRANSPORT || "stdio") as "stdio" | "sse";

  if (!clientId || !clientSecret || !tenantId) {
    logger.warn(
      "Missing auth credentials. Set MSGRAPH_CLIENT_ID, MSGRAPH_CLIENT_SECRET, and MSGRAPH_TENANT_ID environment variables."
    );
  }

  return {
    auth: { clientId, clientSecret, tenantId },
    transport,
    port,
  };
}

async function main() {
  const config = getConfig();
  logger.info(`Starting Lazy MS Graph MCP Server (transport: ${config.transport})`);

  try {
    await startServer(config);
  } catch (error) {
    logger.error(`Failed to start server: ${error instanceof Error ? error.message : String(error)}`);
    process.exit(1);
  }
}

main().catch((error) => {
  console.error("Fatal error:", error);
  process.exit(1);
});

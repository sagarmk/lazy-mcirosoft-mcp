#!/usr/bin/env node

import { startServer } from "./server.js";
import { resolveTenants } from "./config/tenant-config.js";
import { logger } from "./utils/logger.js";

function getConfig() {
  const port = parseInt(process.env.MSGRAPH_MCP_PORT || process.env.MICROSOFT_MCP_PORT || "3100", 10);
  const transport = (process.env.MSGRAPH_MCP_TRANSPORT || process.env.MICROSOFT_MCP_TRANSPORT || "stdio") as "stdio" | "sse";

  const { tenants, defaultTenant } = resolveTenants(process.env, { warn: logger.warn });

  if (tenants.length === 0) {
    logger.warn(
      "Missing auth credentials. Set MSGRAPH_CLIENT_ID, MSGRAPH_CLIENT_SECRET, and MSGRAPH_TENANT_ID (single tenant), or MSGRAPH_TENANTS / MSGRAPH_TENANT_<NAME>_* variables (multiple tenants)."
    );
  }

  return {
    tenants,
    defaultTenant,
    transport,
    port,
  };
}

async function main() {
  try {
    const config = getConfig();
    logger.info(`Starting Lazy MS Graph MCP Server (transport: ${config.transport})`);
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

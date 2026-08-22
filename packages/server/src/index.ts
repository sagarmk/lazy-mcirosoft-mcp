#!/usr/bin/env node

import { startServer } from "./server.js";
import { resolveTenants } from "./config/tenant-config.js";
import { logger } from "./utils/logger.js";

// --print-tools: dump the tool catalog as JSON and exit (used by the macOS
// app's Tools page so it never goes stale). Needs no credentials.
async function printTools(): Promise<void> {
  const { ToolRegistry } = await import("./registry/tool-registry.js");
  const { registerAllTools } = await import("./tools/index.js");
  const { registerTenantTools } = await import("./tools/tenants.js");
  const { TenantManager } = await import("./auth/tenant-manager.js");

  const registry = new ToolRegistry();
  registerAllTools(registry);
  registerTenantTools(registry, new TenantManager([]));

  const tools = registry.getAll().map((tool) => ({
    name: tool.name,
    category: tool.category,
    description: tool.description,
  }));
  process.stdout.write(JSON.stringify({ total: tools.length, tools }, null, 2));
}

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
  if (process.argv.includes("--print-tools")) {
    await printTools();
    return;
  }
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

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";
import { z } from "zod";
import { createServer } from "http";
import type { TenantConfig } from "lazy-ms-graph-mcp-shared";
import { TenantManager } from "./auth/tenant-manager.js";
import { ToolRegistry } from "./registry/tool-registry.js";
import { toolToInfo } from "./registry/tool-definition.js";
import { registerAllTools } from "./tools/index.js";
import { registerTenantTools } from "./tools/tenants.js";
import { logger } from "./utils/logger.js";
import { formatGraphError } from "./utils/graph-helpers.js";

export interface StartOptions {
  tenants: TenantConfig[];
  defaultTenant?: string;
  transport: "stdio" | "sse";
  port: number;
}

export async function startServer(options: StartOptions): Promise<void> {
  const { tenants, defaultTenant, transport, port } = options;

  // One MSAL + Graph client pair per tenant, created lazily on first use
  const tenantManager = new TenantManager(tenants, defaultTenant);

  if (tenantManager.size() === 0) {
    logger.warn("No credentials configured. Tool search will work but execution will fail until credentials are provided.");
  } else {
    logger.info(
      `Configured ${tenantManager.size()} Entra ID tenant(s): ${tenantManager.names().join(", ")} (default: ${tenantManager.getDefaultName()})`
    );
  }

  // Create tool registry and register all tools
  const registry = new ToolRegistry();
  registerAllTools(registry);
  registerTenantTools(registry, tenantManager);
  logger.info(`Registered ${registry.size()} tools across ${registry.getCategories().length} categories`);

  // Create MCP server
  const server = new McpServer({
    name: "lazy-ms-graph",
    version: "1.0.0",
  });

  // Build category summary for the tool description
  const categoryList = registry.getCategories().map((cat) => {
    const tools = registry.getAll().filter((t) => t.category === cat);
    const names = tools.map((t) => t.name).join(", ");
    return `  - ${cat} (${tools.length}): ${names}`;
  }).join("\n");

  // Register the search_tools MCP tool
  server.tool(
    "search_tools",
    `Search available Microsoft Graph tools. Use category filter for best results.

Available categories and tools:
${categoryList}

Tip: Search by category first (e.g. category="mail") to see all tools in that area, then pick the right one to execute.`,
    {
      query: z.string().describe("Search query (e.g., 'send email', 'list users'). Use broad terms or just pass '*' with a category filter to list all tools in a category."),
      category: z.string().optional().describe("Filter by category for targeted results. One of: " + registry.getCategories().join(", ")),
      limit: z.number().optional().describe("Maximum results to return (default: 10, use higher to see more)"),
    },
    async ({ query, category, limit }) => {
      // If query is '*' or empty-ish with a category, return all tools in that category
      const isWildcard = query === "*" || query === "all" || query === "list";
      const effectiveLimit = limit ?? (isWildcard && category ? 50 : 10);

      const results = registry.search(query, {
        category: category as any,
        limit: effectiveLimit,
      });

      // If wildcard with category and no search results, fall back to listing all in category
      let toolInfos = results.map(toolToInfo);
      if (isWildcard && category && toolInfos.length === 0) {
        const allTools = registry.getAll().filter((t) => t.category === category);
        toolInfos = allTools.map(toolToInfo);
      }

      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify(
              {
                tools: toolInfos,
                total: toolInfos.length,
                query,
                category: category || "all",
                available_categories: registry.getCategories(),
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  // Register the execute_tool MCP tool
  server.tool(
    "execute_tool",
    "Execute a Microsoft Graph tool by name with the given parameters. Use search_tools first to discover available tools and their required parameters. When multiple Entra ID tenants are configured, pass 'tenant' to target a specific one (see tenants_list); otherwise the default tenant is used.",
    {
      tool_name: z.string().describe("The exact name of the tool to execute (from search_tools results)"),
      parameters: z.record(z.unknown()).describe("Parameters for the tool as a JSON object"),
      tenant: z.string().optional().describe("Tenant profile name to execute against (from tenants_list). Defaults to the current default tenant."),
    },
    async ({ tool_name, parameters, tenant }) => {
      const tool = registry.get(tool_name);
      if (!tool) {
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify({
                error: `Tool '${tool_name}' not found. Use search_tools to find available tools.`,
              }),
            },
          ],
          isError: true,
        };
      }

      // Tenant-management tools don't call the Graph API, so they run even
      // with no credentials configured; everything else needs a client.
      let graphClient: import("@microsoft/microsoft-graph-client").Client | null = null;
      try {
        graphClient = tenantManager.getClient(tenant);
      } catch (error) {
        if (tool.category !== "tenants") {
          return {
            content: [
              {
                type: "text" as const,
                text: JSON.stringify({
                  error: error instanceof Error ? error.message : String(error),
                  available_tenants: tenantManager.names(),
                }),
              },
            ],
            isError: true,
          };
        }
      }

      try {
        const validated = tool.parameters.parse(parameters);
        const result = await tool.handler(graphClient!, validated);
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      } catch (error: unknown) {
        const message =
          error instanceof z.ZodError
            ? `Validation error: ${error.errors.map((e) => `${e.path.join(".")}: ${e.message}`).join(", ")}`
            : formatGraphError(error);

        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify({ error: message }),
            },
          ],
          isError: true,
        };
      }
    }
  );

  // Connect transport
  if (transport === "stdio") {
    logger.info("Starting MCP server with stdio transport...");
    const stdioTransport = new StdioServerTransport();
    await server.connect(stdioTransport);
    logger.info("MCP server running on stdio");
  } else {
    logger.info(`Starting MCP server with SSE transport on port ${port}...`);

    let sseTransport: SSEServerTransport | null = null;

    const httpServer = createServer(async (req, res) => {
      res.setHeader("Access-Control-Allow-Origin", "*");
      res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
      res.setHeader("Access-Control-Allow-Headers", "Content-Type");

      if (req.method === "OPTIONS") {
        res.writeHead(204);
        res.end();
        return;
      }

      if (req.url === "/sse" && req.method === "GET") {
        sseTransport = new SSEServerTransport("/messages", res);
        await server.connect(sseTransport);
        return;
      }

      if (req.url === "/messages" && req.method === "POST") {
        if (sseTransport) {
          await sseTransport.handlePostMessage(req, res);
        } else {
          res.writeHead(400);
          res.end("No SSE connection established");
        }
        return;
      }

      // Health check
      if (req.url === "/health") {
        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(
          JSON.stringify({
            status: "ok",
            tools: registry.size(),
            tenants: tenantManager.names(),
            default_tenant: tenantManager.getDefaultName(),
          })
        );
        return;
      }

      res.writeHead(404);
      res.end("Not found");
    });

    httpServer.listen(port, () => {
      logger.info(`MCP server listening on http://localhost:${port}`);
      logger.info(`SSE endpoint: http://localhost:${port}/sse`);
      logger.info(`Health check: http://localhost:${port}/health`);
    });
  }

  // Signal ready to parent process (Electron app)
  if (process.send) {
    process.send({ type: "ready", tools: registry.size() });
  }
}

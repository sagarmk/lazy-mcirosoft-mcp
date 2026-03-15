import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { SSEServerTransport } from "@modelcontextprotocol/sdk/server/sse.js";
import { z } from "zod";
import { createServer } from "http";
import type { AuthConfig } from "lazy-ms-graph-mcp-shared";
import { createMsalClient } from "./auth/msal-client.js";
import { createGraphClient } from "./auth/graph-client.js";
import { ToolRegistry } from "./registry/tool-registry.js";
import { toolToInfo } from "./registry/tool-definition.js";
import { registerAllTools } from "./tools/index.js";
import { logger } from "./utils/logger.js";
import { formatGraphError } from "./utils/graph-helpers.js";

export interface StartOptions {
  auth: AuthConfig;
  transport: "stdio" | "sse";
  port: number;
}

export async function startServer(options: StartOptions): Promise<void> {
  const { auth, transport, port } = options;

  // Create auth + graph client
  const hasCredentials = auth.clientId && auth.clientSecret && auth.tenantId;
  let graphClient: import("@microsoft/microsoft-graph-client").Client | null = null;

  if (hasCredentials) {
    logger.info("Initializing MSAL client...");
    const msal = createMsalClient(auth);
    graphClient = createGraphClient(msal);
  } else {
    logger.warn("No credentials configured. Tool search will work but execution will fail until credentials are provided.");
  }

  // Create tool registry and register all tools
  const registry = new ToolRegistry();
  registerAllTools(registry);
  logger.info(`Registered ${registry.size()} tools across ${registry.getCategories().length} categories`);

  // Create MCP server
  const server = new McpServer({
    name: "lazy-ms-graph",
    version: "1.0.0",
  });

  // Register the search_tools MCP tool
  server.tool(
    "search_tools",
    "Search available Microsoft Graph tools by query. Returns tool names, descriptions, categories, and parameter schemas. Use this to discover what operations are available before executing them.",
    {
      query: z.string().describe("Search query to find relevant tools (e.g., 'send email', 'list users', 'create event')"),
      category: z.string().optional().describe("Filter by category: users, mail, calendar, contacts, files, teams, sharepoint, planner, onenote, groups"),
      limit: z.number().optional().describe("Maximum number of results to return (default: 10)"),
    },
    async ({ query, category, limit }) => {
      const results = registry.search(query, {
        category: category as any,
        limit: limit ?? 10,
      });

      const toolInfos = results.map(toolToInfo);

      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify(
              {
                tools: toolInfos,
                total: toolInfos.length,
                query,
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
    "Execute a Microsoft Graph tool by name with the given parameters. Use search_tools first to discover available tools and their required parameters.",
    {
      tool_name: z.string().describe("The exact name of the tool to execute (from search_tools results)"),
      parameters: z.record(z.unknown()).describe("Parameters for the tool as a JSON object"),
    },
    async ({ tool_name, parameters }) => {
      if (!graphClient) {
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify({
                error: "No credentials configured. Set MSGRAPH_CLIENT_ID, MSGRAPH_CLIENT_SECRET, and MSGRAPH_TENANT_ID.",
              }),
            },
          ],
          isError: true,
        };
      }

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
        res.end(JSON.stringify({ status: "ok", tools: registry.size() }));
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

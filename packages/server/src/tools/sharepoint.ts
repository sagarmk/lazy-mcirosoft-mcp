import { z } from "zod";
import type { Client } from "@microsoft/microsoft-graph-client";
import type { ToolRegistry } from "../registry/tool-registry.js";

export function registerSharePointTools(registry: ToolRegistry): void {
  registry.register({
    name: "sharepoint_list_sites",
    category: "sharepoint",
    description: "List SharePoint sites across the organization. Returns sites matching a broad search.",
    keywords: ["sharepoint", "sites", "list", "search", "intranet"],
    parameters: z.object({}),
    handler: async (client: Client) => {
      return await client.api("/sites?search=*").get();
    },
  });

  registry.register({
    name: "sharepoint_get_site",
    category: "sharepoint",
    description: "Get a specific SharePoint site by its ID.",
    keywords: ["sharepoint", "site", "get", "details"],
    parameters: z.object({
      siteId: z.string().describe("The unique identifier of the SharePoint site."),
    }),
    handler: async (client: Client, params: { siteId: string }) => {
      return await client.api(`/sites/${params.siteId}`).get();
    },
  });

  registry.register({
    name: "sharepoint_list_lists",
    category: "sharepoint",
    description: "List all lists on a SharePoint site.",
    keywords: ["sharepoint", "lists", "site", "document library", "custom list"],
    parameters: z.object({
      siteId: z.string().describe("The unique identifier of the SharePoint site."),
    }),
    handler: async (client: Client, params: { siteId: string }) => {
      return await client.api(`/sites/${params.siteId}/lists`).get();
    },
  });

  registry.register({
    name: "sharepoint_get_list_items",
    category: "sharepoint",
    description: "Get items from a SharePoint list with their field values expanded.",
    keywords: ["sharepoint", "list", "items", "data", "rows", "fields"],
    parameters: z.object({
      siteId: z.string().describe("The unique identifier of the SharePoint site."),
      listId: z.string().describe("The unique identifier of the list."),
      $top: z.number().optional().describe("Number of items to return."),
    }),
    handler: async (client: Client, params: {
      siteId: string;
      listId: string;
      $top?: number;
    }) => {
      let request = client.api(`/sites/${params.siteId}/lists/${params.listId}/items`).expand("fields");

      if (params.$top) {
        request = request.top(params.$top);
      }

      return await request.get();
    },
  });
}

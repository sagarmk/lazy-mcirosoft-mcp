import { z } from "zod";
import type { Client } from "@microsoft/microsoft-graph-client";
import type { ToolRegistry } from "../registry/tool-registry.js";

export function registerFileTools(registry: ToolRegistry): void {
  registry.register({
    name: "files_list",
    category: "files",
    description: "List files and folders in a user's OneDrive. Defaults to the root folder if no path is specified.",
    keywords: ["files", "folders", "list", "onedrive", "drive", "directory", "browse"],
    parameters: z.object({
      userId: z.string().describe("The user's ID (GUID) or user principal name (e.g. user@example.com)."),
      path: z.string().optional().describe("The folder path relative to the drive root (e.g. \"Documents/Reports\"). Defaults to root if not specified."),
      $top: z.number().optional().describe("Number of items to return."),
    }),
    handler: async (client: Client, params: {
      userId: string;
      path?: string;
      $top?: number;
    }) => {
      const endpoint = params.path
        ? `/users/${params.userId}/drive/root:/${params.path}:/children`
        : `/users/${params.userId}/drive/root/children`;

      let request = client.api(endpoint);

      if (params.$top) {
        request = request.top(params.$top);
      }

      return await request.get();
    },
  });

  registry.register({
    name: "files_get",
    category: "files",
    description: "Get metadata for a specific file or folder in a user's OneDrive by item ID.",
    keywords: ["file", "get", "metadata", "details", "info", "onedrive", "drive"],
    parameters: z.object({
      userId: z.string().describe("The user's ID (GUID) or user principal name (e.g. user@example.com)."),
      itemId: z.string().describe("The unique ID of the drive item."),
    }),
    handler: async (client: Client, params: {
      userId: string;
      itemId: string;
    }) => {
      return await client.api(`/users/${params.userId}/drive/items/${params.itemId}`).get();
    },
  });

  registry.register({
    name: "files_upload",
    category: "files",
    description: "Upload a small file (up to 4 MB) to a user's OneDrive by providing the file content as a string.",
    keywords: ["file", "upload", "create", "write", "onedrive", "drive", "content"],
    parameters: z.object({
      userId: z.string().describe("The user's ID (GUID) or user principal name (e.g. user@example.com)."),
      parentPath: z.string().describe("The parent folder path relative to the drive root (e.g. \"Documents/Reports\")."),
      fileName: z.string().describe("The name of the file to create or overwrite (e.g. \"report.txt\")."),
      content: z.string().describe("The file content as a string."),
    }),
    handler: async (client: Client, params: {
      userId: string;
      parentPath: string;
      fileName: string;
      content: string;
    }) => {
      return await client
        .api(`/users/${params.userId}/drive/root:/${params.parentPath}/${params.fileName}:/content`)
        .put(params.content);
    },
  });

  registry.register({
    name: "files_download",
    category: "files",
    description: "Get the download URL for a file in a user's OneDrive. Returns the pre-authenticated download URL.",
    keywords: ["file", "download", "url", "onedrive", "drive", "content", "link"],
    parameters: z.object({
      userId: z.string().describe("The user's ID (GUID) or user principal name (e.g. user@example.com)."),
      itemId: z.string().describe("The unique ID of the drive item to download."),
    }),
    handler: async (client: Client, params: {
      userId: string;
      itemId: string;
    }) => {
      const item = await client.api(`/users/${params.userId}/drive/items/${params.itemId}`).get();
      return {
        name: item.name,
        downloadUrl: item["@microsoft.graph.downloadUrl"],
      };
    },
  });

  registry.register({
    name: "files_search",
    category: "files",
    description: "Search for files and folders in a user's OneDrive by query string. Searches file names, metadata, and file content.",
    keywords: ["files", "search", "find", "query", "onedrive", "drive", "lookup"],
    parameters: z.object({
      userId: z.string().describe("The user's ID (GUID) or user principal name (e.g. user@example.com)."),
      query: z.string().describe("The search query text to find matching files."),
    }),
    handler: async (client: Client, params: {
      userId: string;
      query: string;
    }) => {
      return await client
        .api(`/users/${params.userId}/drive/root/search(q='${params.query}')`)
        .get();
    },
  });

  registry.register({
    name: "files_share",
    category: "files",
    description: "Create a sharing link for a file or folder in a user's OneDrive.",
    keywords: ["file", "share", "link", "permission", "access", "onedrive", "drive"],
    parameters: z.object({
      userId: z.string().describe("The user's ID (GUID) or user principal name (e.g. user@example.com)."),
      itemId: z.string().describe("The unique ID of the drive item to share."),
      type: z.enum(["view", "edit"]).describe("The type of sharing link: \"view\" for read-only or \"edit\" for read-write."),
      scope: z.enum(["anonymous", "organization"]).describe("The scope of the link: \"anonymous\" for anyone with the link, or \"organization\" for people in the same organization."),
    }),
    handler: async (client: Client, params: {
      userId: string;
      itemId: string;
      type: "view" | "edit";
      scope: "anonymous" | "organization";
    }) => {
      return await client
        .api(`/users/${params.userId}/drive/items/${params.itemId}/createLink`)
        .post({
          type: params.type,
          scope: params.scope,
        });
    },
  });
}

import { z } from "zod";
import type { Client } from "@microsoft/microsoft-graph-client";
import type { ToolRegistry } from "../registry/tool-registry.js";

export function registerOneNoteTools(registry: ToolRegistry): void {
  registry.register({
    name: "onenote_list_notebooks",
    category: "onenote",
    description: "List all OneNote notebooks for a specific user.",
    keywords: ["onenote", "notebooks", "list", "notes"],
    parameters: z.object({
      userId: z.string().describe("The user's ID (GUID) or user principal name."),
    }),
    handler: async (client: Client, params: { userId: string }) => {
      return await client.api(`/users/${params.userId}/onenote/notebooks`).get();
    },
  });

  registry.register({
    name: "onenote_list_sections",
    category: "onenote",
    description: "List all sections in a specific OneNote notebook.",
    keywords: ["onenote", "sections", "list", "notebook", "organize"],
    parameters: z.object({
      userId: z.string().describe("The user's ID (GUID) or user principal name."),
      notebookId: z.string().describe("The unique identifier of the notebook."),
    }),
    handler: async (client: Client, params: { userId: string; notebookId: string }) => {
      return await client.api(`/users/${params.userId}/onenote/notebooks/${params.notebookId}/sections`).get();
    },
  });

  registry.register({
    name: "onenote_list_pages",
    category: "onenote",
    description: "List all pages in a specific OneNote section.",
    keywords: ["onenote", "pages", "list", "section", "notes"],
    parameters: z.object({
      userId: z.string().describe("The user's ID (GUID) or user principal name."),
      sectionId: z.string().describe("The unique identifier of the section."),
    }),
    handler: async (client: Client, params: { userId: string; sectionId: string }) => {
      return await client.api(`/users/${params.userId}/onenote/sections/${params.sectionId}/pages`).get();
    },
  });

  registry.register({
    name: "onenote_get_page_content",
    category: "onenote",
    description: "Get the HTML content of a specific OneNote page.",
    keywords: ["onenote", "page", "content", "html", "read", "body"],
    parameters: z.object({
      userId: z.string().describe("The user's ID (GUID) or user principal name."),
      pageId: z.string().describe("The unique identifier of the page."),
    }),
    handler: async (client: Client, params: { userId: string; pageId: string }) => {
      return await client.api(`/users/${params.userId}/onenote/pages/${params.pageId}/content`).get();
    },
  });
}

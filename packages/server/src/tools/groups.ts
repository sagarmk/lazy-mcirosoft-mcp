import { z } from "zod";
import type { Client } from "@microsoft/microsoft-graph-client";
import type { ToolRegistry } from "../registry/tool-registry.js";

export function registerGroupTools(registry: ToolRegistry): void {
  registry.register({
    name: "groups_list",
    category: "groups",
    description: "List groups in the organization with optional filtering and pagination.",
    keywords: ["groups", "list", "directory", "teams", "security", "distribution"],
    parameters: z.object({
      $top: z.number().optional().describe("Number of groups to return."),
      $filter: z.string().optional().describe("OData filter expression (e.g. \"displayName eq 'Engineering'\")."),
    }),
    handler: async (client: Client, params: {
      $top?: number;
      $filter?: string;
    }) => {
      let request = client.api("/groups");

      if (params.$top) {
        request = request.top(params.$top);
      }
      if (params.$filter) {
        request = request.filter(params.$filter);
      }

      return await request.get();
    },
  });

  registry.register({
    name: "groups_get",
    category: "groups",
    description: "Get a specific group by its ID.",
    keywords: ["group", "get", "details", "info"],
    parameters: z.object({
      groupId: z.string().describe("The unique identifier of the group."),
    }),
    handler: async (client: Client, params: { groupId: string }) => {
      return await client.api(`/groups/${params.groupId}`).get();
    },
  });

  registry.register({
    name: "groups_create",
    category: "groups",
    description: "Create a new group in the directory.",
    keywords: ["group", "create", "new", "add", "team", "security"],
    parameters: z.object({
      displayName: z.string().describe("The display name for the group."),
      mailEnabled: z.boolean().describe("Whether the group is mail-enabled."),
      mailNickname: z.string().describe("The mail alias for the group (no spaces)."),
      securityEnabled: z.boolean().describe("Whether the group is a security group."),
      groupTypes: z.array(z.string()).optional().describe("Group types (e.g. [\"Unified\"] for Microsoft 365 groups)."),
    }),
    handler: async (client: Client, params: {
      displayName: string;
      mailEnabled: boolean;
      mailNickname: string;
      securityEnabled: boolean;
      groupTypes?: string[];
    }) => {
      const body: Record<string, unknown> = {
        displayName: params.displayName,
        mailEnabled: params.mailEnabled,
        mailNickname: params.mailNickname,
        securityEnabled: params.securityEnabled,
      };

      if (params.groupTypes) {
        body.groupTypes = params.groupTypes;
      }

      return await client.api("/groups").post(body);
    },
  });

  registry.register({
    name: "groups_list_members",
    category: "groups",
    description: "List members of a specific group.",
    keywords: ["group", "members", "list", "users", "membership"],
    parameters: z.object({
      groupId: z.string().describe("The unique identifier of the group."),
    }),
    handler: async (client: Client, params: { groupId: string }) => {
      return await client.api(`/groups/${params.groupId}/members`).get();
    },
  });

  registry.register({
    name: "groups_manage_members",
    category: "groups",
    description: "Add or remove a member from a group.",
    keywords: ["group", "member", "add", "remove", "manage", "membership"],
    parameters: z.object({
      groupId: z.string().describe("The unique identifier of the group."),
      userId: z.string().describe("The unique identifier of the user to add or remove."),
      action: z.enum(["add", "remove"]).describe("The action to perform: \"add\" to add a member or \"remove\" to remove a member."),
    }),
    handler: async (client: Client, params: {
      groupId: string;
      userId: string;
      action: "add" | "remove";
    }) => {
      if (params.action === "add") {
        const body = {
          "@odata.id": `https://graph.microsoft.com/v1.0/users/${params.userId}`,
        };
        await client.api(`/groups/${params.groupId}/members/$ref`).post(body);
        return { success: true, message: `User ${params.userId} added to group ${params.groupId}.` };
      } else {
        await client.api(`/groups/${params.groupId}/members/${params.userId}/$ref`).delete();
        return { success: true, message: `User ${params.userId} removed from group ${params.groupId}.` };
      }
    },
  });
}

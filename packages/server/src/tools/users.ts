import { z } from "zod";
import type { Client } from "@microsoft/microsoft-graph-client";
import type { ToolRegistry } from "../registry/tool-registry.js";

export function registerUserTools(registry: ToolRegistry): void {
  registry.register({
    name: "users_list",
    category: "users",
    description: "List users in the organization with optional OData query parameters for filtering, selecting fields, searching, and pagination.",
    keywords: ["users", "list", "directory", "people", "organization", "filter", "search"],
    parameters: z.object({
      $top: z.number().optional().describe("Number of users to return (max 999)."),
      $filter: z.string().optional().describe("OData filter expression (e.g. \"department eq 'Engineering'\")."),
      $select: z.string().optional().describe("Comma-separated list of properties to include (e.g. \"displayName,mail,jobTitle\")."),
      $search: z.string().optional().describe("Search expression for searching across displayName, mail, etc. Requires ConsistencyLevel: eventual header."),
    }),
    handler: async (client: Client, params: {
      $top?: number;
      $filter?: string;
      $select?: string;
      $search?: string;
    }) => {
      let request = client.api("/users");

      if (params.$top) {
        request = request.top(params.$top);
      }
      if (params.$filter) {
        request = request.filter(params.$filter);
      }
      if (params.$select) {
        request = request.select(params.$select);
      }
      if (params.$search) {
        request = request.header("ConsistencyLevel", "eventual").query({ $search: params.$search });
      }

      return await request.get();
    },
  });

  registry.register({
    name: "users_get",
    category: "users",
    description: "Get a specific user by their ID or user principal name (UPN).",
    keywords: ["user", "get", "profile", "details", "identity", "upn"],
    parameters: z.object({
      userId: z.string().describe("The user's ID (GUID) or user principal name (e.g. user@example.com)."),
    }),
    handler: async (client: Client, params: { userId: string }) => {
      return await client.api(`/users/${params.userId}`).get();
    },
  });

  registry.register({
    name: "users_create",
    category: "users",
    description: "Create a new user in the directory. Requires User.ReadWrite.All permission.",
    keywords: ["user", "create", "new", "add", "provision", "account"],
    parameters: z.object({
      displayName: z.string().describe("The name displayed in the address book for the user (e.g. \"Alex Johnson\")."),
      mailNickname: z.string().describe("The mail alias for the user (e.g. \"alexj\")."),
      userPrincipalName: z.string().describe("The user principal name (e.g. \"alexj@contoso.com\")."),
      password: z.string().describe("The initial password for the user."),
      accountEnabled: z.boolean().describe("Whether the account is enabled."),
    }),
    handler: async (client: Client, params: {
      displayName: string;
      mailNickname: string;
      userPrincipalName: string;
      password: string;
      accountEnabled: boolean;
    }) => {
      const body = {
        displayName: params.displayName,
        mailNickname: params.mailNickname,
        userPrincipalName: params.userPrincipalName,
        accountEnabled: params.accountEnabled,
        passwordProfile: {
          password: params.password,
          forceChangePasswordNextSignIn: true,
        },
      };

      return await client.api("/users").post(body);
    },
  });

  registry.register({
    name: "users_update",
    category: "users",
    description: "Update properties of an existing user. Pass any valid user properties to update.",
    keywords: ["user", "update", "modify", "edit", "patch", "change"],
    parameters: z.object({
      userId: z.string().describe("The user's ID (GUID) or user principal name."),
      properties: z.record(z.unknown()).describe("An object containing the user properties to update (e.g. { \"jobTitle\": \"Senior Dev\", \"department\": \"Engineering\" })."),
    }),
    handler: async (client: Client, params: {
      userId: string;
      properties: Record<string, unknown>;
    }) => {
      await client.api(`/users/${params.userId}`).patch(params.properties);
      return { success: true, message: `User ${params.userId} updated successfully.` };
    },
  });

  registry.register({
    name: "users_delete",
    category: "users",
    description: "Delete a user from the directory. This moves the user to the deleted items container and can be restored within 30 days.",
    keywords: ["user", "delete", "remove", "deprovision"],
    parameters: z.object({
      userId: z.string().describe("The user's ID (GUID) or user principal name to delete."),
    }),
    handler: async (client: Client, params: { userId: string }) => {
      await client.api(`/users/${params.userId}`).delete();
      return { success: true, message: `User ${params.userId} deleted successfully.` };
    },
  });
}

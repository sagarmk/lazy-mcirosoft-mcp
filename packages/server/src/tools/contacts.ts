import { z } from "zod";
import type { Client } from "@microsoft/microsoft-graph-client";
import type { ToolRegistry } from "../registry/tool-registry.js";

export function registerContactTools(registry: ToolRegistry): void {
  registry.register({
    name: "contacts_list",
    category: "contacts",
    description: "List contacts for a user.",
    keywords: ["contacts", "list", "people", "address book"],
    parameters: z.object({
      userId: z.string().describe("The ID or UPN of the user"),
      $top: z
        .number()
        .optional()
        .describe("Maximum number of contacts to return"),
      $filter: z
        .string()
        .optional()
        .describe("OData filter expression to filter contacts"),
      $select: z
        .string()
        .optional()
        .describe("Comma-separated list of properties to select"),
    }),
    handler: async (client: Client, params: Record<string, unknown>) => {
      const { userId, $top, $filter, $select } = params as {
        userId: string;
        $top?: number;
        $filter?: string;
        $select?: string;
      };

      let request = client.api(`/users/${userId}/contacts`);

      if ($top !== undefined) {
        request = request.top($top);
      }
      if ($filter) {
        request = request.filter($filter);
      }
      if ($select) {
        request = request.select($select);
      }

      return await request.get();
    },
  });

  registry.register({
    name: "contacts_get",
    category: "contacts",
    description: "Get a specific contact by ID for a given user.",
    keywords: ["contact", "get", "details", "person"],
    parameters: z.object({
      userId: z.string().describe("The ID or UPN of the user"),
      contactId: z.string().describe("The ID of the contact to retrieve"),
    }),
    handler: async (client: Client, params: Record<string, unknown>) => {
      const { userId, contactId } = params as {
        userId: string;
        contactId: string;
      };

      return await client
        .api(`/users/${userId}/contacts/${contactId}`)
        .get();
    },
  });

  registry.register({
    name: "contacts_create",
    category: "contacts",
    description: "Create a new contact for a user.",
    keywords: ["contact", "create", "new", "add", "person"],
    parameters: z.object({
      userId: z.string().describe("The ID or UPN of the user"),
      givenName: z.string().describe("The contact's first/given name"),
      surname: z.string().describe("The contact's last name/surname"),
      emailAddresses: z
        .array(
          z.object({
            address: z.string().describe("The email address"),
            name: z
              .string()
              .optional()
              .describe("Display name for the email address"),
          }),
        )
        .optional()
        .describe("List of email addresses for the contact"),
      businessPhones: z
        .array(z.string())
        .optional()
        .describe("List of business phone numbers"),
      jobTitle: z
        .string()
        .optional()
        .describe("The contact's job title"),
      companyName: z
        .string()
        .optional()
        .describe("The name of the contact's company"),
    }),
    handler: async (client: Client, params: Record<string, unknown>) => {
      const { userId, ...contactData } = params as {
        userId: string;
        givenName: string;
        surname: string;
        emailAddresses?: Array<{ address: string; name?: string }>;
        businessPhones?: string[];
        jobTitle?: string;
        companyName?: string;
      };

      return await client
        .api(`/users/${userId}/contacts`)
        .post(contactData);
    },
  });

  registry.register({
    name: "contacts_update",
    category: "contacts",
    description: "Update an existing contact for a user.",
    keywords: ["contact", "update", "edit", "modify", "person"],
    parameters: z.object({
      userId: z.string().describe("The ID or UPN of the user"),
      contactId: z.string().describe("The ID of the contact to update"),
      givenName: z
        .string()
        .optional()
        .describe("Updated first/given name"),
      surname: z
        .string()
        .optional()
        .describe("Updated last name/surname"),
      emailAddresses: z
        .array(
          z.object({
            address: z.string().describe("The email address"),
            name: z
              .string()
              .optional()
              .describe("Display name for the email address"),
          }),
        )
        .optional()
        .describe("Updated list of email addresses"),
      businessPhones: z
        .array(z.string())
        .optional()
        .describe("Updated list of business phone numbers"),
      jobTitle: z
        .string()
        .optional()
        .describe("Updated job title"),
      companyName: z
        .string()
        .optional()
        .describe("Updated company name"),
    }),
    handler: async (client: Client, params: Record<string, unknown>) => {
      const { userId, contactId, ...updateData } = params as {
        userId: string;
        contactId: string;
        [key: string]: unknown;
      };

      return await client
        .api(`/users/${userId}/contacts/${contactId}`)
        .patch(updateData);
    },
  });

  registry.register({
    name: "contacts_delete",
    category: "contacts",
    description: "Delete a contact for a user.",
    keywords: ["contact", "delete", "remove", "person"],
    parameters: z.object({
      userId: z.string().describe("The ID or UPN of the user"),
      contactId: z.string().describe("The ID of the contact to delete"),
    }),
    handler: async (client: Client, params: Record<string, unknown>) => {
      const { userId, contactId } = params as {
        userId: string;
        contactId: string;
      };

      await client
        .api(`/users/${userId}/contacts/${contactId}`)
        .delete();

      return {
        success: true,
        message: `Contact ${contactId} deleted successfully`,
      };
    },
  });
}

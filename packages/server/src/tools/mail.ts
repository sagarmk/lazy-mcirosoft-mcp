import { z } from "zod";
import type { Client } from "@microsoft/microsoft-graph-client";
import type { ToolRegistry } from "../registry/tool-registry.js";

const recipientSchema = z.object({
  emailAddress: z.object({
    address: z.string().describe("The email address of the recipient."),
    name: z.string().optional().describe("The display name of the recipient."),
  }),
});

export function registerMailTools(registry: ToolRegistry): void {
  registry.register({
    name: "mail_list_messages",
    category: "mail",
    description: "List mail messages for a user. Optionally filter by folder, select specific fields, and paginate results.",
    keywords: ["mail", "email", "messages", "list", "inbox", "folder"],
    parameters: z.object({
      userId: z.string().describe("The user's ID or user principal name."),
      folderId: z.string().optional().describe("The mail folder ID to list messages from (e.g. 'Inbox', 'SentItems', 'Drafts', or a folder ID). If omitted, lists messages from all folders."),
      $top: z.number().optional().describe("Number of messages to return (max 999)."),
      $filter: z.string().optional().describe("OData filter expression (e.g. \"isRead eq false\")."),
      $select: z.string().optional().describe("Comma-separated list of properties to include (e.g. \"subject,from,receivedDateTime\")."),
    }),
    handler: async (client: Client, params: {
      userId: string;
      folderId?: string;
      $top?: number;
      $filter?: string;
      $select?: string;
    }) => {
      const endpoint = params.folderId
        ? `/users/${params.userId}/mailFolders/${params.folderId}/messages`
        : `/users/${params.userId}/messages`;

      let request = client.api(endpoint);

      if (params.$top) {
        request = request.top(params.$top);
      }
      if (params.$filter) {
        request = request.filter(params.$filter);
      }
      if (params.$select) {
        request = request.select(params.$select);
      }

      return await request.get();
    },
  });

  registry.register({
    name: "mail_get_message",
    category: "mail",
    description: "Get a specific mail message by ID, including its full body content and metadata.",
    keywords: ["mail", "email", "message", "get", "read", "details"],
    parameters: z.object({
      userId: z.string().describe("The user's ID or user principal name."),
      messageId: z.string().describe("The ID of the message to retrieve."),
    }),
    handler: async (client: Client, params: {
      userId: string;
      messageId: string;
    }) => {
      return await client.api(`/users/${params.userId}/messages/${params.messageId}`).get();
    },
  });

  registry.register({
    name: "mail_send_message",
    category: "mail",
    description: "Send an email message on behalf of a user. Supports HTML and text content types, with optional CC recipients.",
    keywords: ["mail", "email", "send", "compose", "new", "write"],
    parameters: z.object({
      userId: z.string().describe("The user's ID or user principal name."),
      subject: z.string().describe("The subject line of the email."),
      body: z.string().describe("The body content of the email."),
      contentType: z.enum(["Text", "HTML"]).optional().describe("The content type of the body. Defaults to 'Text'."),
      toRecipients: z.array(recipientSchema).describe("Array of recipients for the To line."),
      ccRecipients: z.array(recipientSchema).optional().describe("Array of recipients for the CC line."),
    }),
    handler: async (client: Client, params: {
      userId: string;
      subject: string;
      body: string;
      contentType?: "Text" | "HTML";
      toRecipients: Array<{ emailAddress: { address: string; name?: string } }>;
      ccRecipients?: Array<{ emailAddress: { address: string; name?: string } }>;
    }) => {
      const message = {
        message: {
          subject: params.subject,
          body: {
            contentType: params.contentType || "Text",
            content: params.body,
          },
          toRecipients: params.toRecipients,
          ...(params.ccRecipients && { ccRecipients: params.ccRecipients }),
        },
      };

      await client.api(`/users/${params.userId}/sendMail`).post(message);
      return { success: true, message: "Email sent successfully." };
    },
  });

  registry.register({
    name: "mail_reply_message",
    category: "mail",
    description: "Reply to an existing email message. Can reply to just the sender or reply all to include all original recipients.",
    keywords: ["mail", "email", "reply", "respond", "answer"],
    parameters: z.object({
      userId: z.string().describe("The user's ID or user principal name."),
      messageId: z.string().describe("The ID of the message to reply to."),
      comment: z.string().describe("The reply comment/body text."),
      replyAll: z.boolean().optional().describe("If true, replies to all recipients. Defaults to false (reply to sender only)."),
    }),
    handler: async (client: Client, params: {
      userId: string;
      messageId: string;
      comment: string;
      replyAll?: boolean;
    }) => {
      const action = params.replyAll ? "replyAll" : "reply";
      const body = { comment: params.comment };

      await client.api(`/users/${params.userId}/messages/${params.messageId}/${action}`).post(body);
      return { success: true, message: `Reply sent successfully.` };
    },
  });

  registry.register({
    name: "mail_forward_message",
    category: "mail",
    description: "Forward an existing email message to one or more recipients with an optional comment.",
    keywords: ["mail", "email", "forward", "share", "send"],
    parameters: z.object({
      userId: z.string().describe("The user's ID or user principal name."),
      messageId: z.string().describe("The ID of the message to forward."),
      comment: z.string().describe("A comment to include with the forwarded message."),
      toRecipients: z.array(recipientSchema).describe("Array of recipients to forward the message to."),
    }),
    handler: async (client: Client, params: {
      userId: string;
      messageId: string;
      comment: string;
      toRecipients: Array<{ emailAddress: { address: string; name?: string } }>;
    }) => {
      const body = {
        comment: params.comment,
        toRecipients: params.toRecipients,
      };

      await client.api(`/users/${params.userId}/messages/${params.messageId}/forward`).post(body);
      return { success: true, message: "Message forwarded successfully." };
    },
  });

  registry.register({
    name: "mail_search_messages",
    category: "mail",
    description: "Search a user's mail messages using a keyword query. Searches across subject, body, and other message fields.",
    keywords: ["mail", "email", "search", "find", "query", "lookup"],
    parameters: z.object({
      userId: z.string().describe("The user's ID or user principal name."),
      query: z.string().describe("The search query string (e.g. \"budget report\", \"from:john@example.com\")."),
      $top: z.number().optional().describe("Number of results to return (max 999)."),
    }),
    handler: async (client: Client, params: {
      userId: string;
      query: string;
      $top?: number;
    }) => {
      let request = client.api(`/users/${params.userId}/messages`)
        .query({ $search: `"${params.query}"` });

      if (params.$top) {
        request = request.top(params.$top);
      }

      return await request.get();
    },
  });

  registry.register({
    name: "mail_list_folders",
    category: "mail",
    description: "List all mail folders for a user, including built-in folders (Inbox, Sent Items, Drafts, etc.) and custom folders.",
    keywords: ["mail", "email", "folders", "list", "mailbox", "organize"],
    parameters: z.object({
      userId: z.string().describe("The user's ID or user principal name."),
    }),
    handler: async (client: Client, params: { userId: string }) => {
      return await client.api(`/users/${params.userId}/mailFolders`).get();
    },
  });

  registry.register({
    name: "mail_create_folder",
    category: "mail",
    description: "Create a new mail folder for a user. Can be created as a top-level folder or as a child of an existing folder.",
    keywords: ["mail", "email", "folder", "create", "new", "organize"],
    parameters: z.object({
      userId: z.string().describe("The user's ID or user principal name."),
      displayName: z.string().describe("The name of the new mail folder."),
      parentFolderId: z.string().optional().describe("The ID of the parent folder. If omitted, the folder is created at the top level."),
    }),
    handler: async (client: Client, params: {
      userId: string;
      displayName: string;
      parentFolderId?: string;
    }) => {
      const endpoint = params.parentFolderId
        ? `/users/${params.userId}/mailFolders/${params.parentFolderId}/childFolders`
        : `/users/${params.userId}/mailFolders`;

      const body = { displayName: params.displayName };

      return await client.api(endpoint).post(body);
    },
  });
}

import { z } from "zod";
import type { Client } from "@microsoft/microsoft-graph-client";
import type { ToolRegistry } from "../registry/tool-registry.js";

export function registerTeamsTools(registry: ToolRegistry): void {
  registry.register({
    name: "teams_list",
    category: "teams",
    description: "List all teams in the organization. Retrieves Microsoft 365 groups that have been provisioned as teams.",
    keywords: ["teams", "list", "groups", "organization", "all"],
    parameters: z.object({}),
    handler: async (client: Client) => {
      return await client
        .api("/groups")
        .filter("resourceProvisioningOptions/Any(x:x eq 'Team')")
        .get();
    },
  });

  registry.register({
    name: "teams_list_channels",
    category: "teams",
    description: "List all channels in a specific Microsoft Teams team.",
    keywords: ["teams", "channels", "list", "conversations"],
    parameters: z.object({
      teamId: z.string().describe("The unique ID of the team."),
    }),
    handler: async (client: Client, params: { teamId: string }) => {
      return await client.api(`/teams/${params.teamId}/channels`).get();
    },
  });

  registry.register({
    name: "teams_list_messages",
    category: "teams",
    description: "List messages in a Microsoft Teams channel. Returns the most recent messages.",
    keywords: ["teams", "messages", "list", "channel", "chat", "conversation", "history"],
    parameters: z.object({
      teamId: z.string().describe("The unique ID of the team."),
      channelId: z.string().describe("The unique ID of the channel."),
      $top: z.number().optional().describe("Number of messages to return."),
    }),
    handler: async (client: Client, params: {
      teamId: string;
      channelId: string;
      $top?: number;
    }) => {
      let request = client.api(`/teams/${params.teamId}/channels/${params.channelId}/messages`);

      if (params.$top) {
        request = request.top(params.$top);
      }

      return await request.get();
    },
  });

  registry.register({
    name: "teams_send_message",
    category: "teams",
    description: "Send a message to a Microsoft Teams channel.",
    keywords: ["teams", "message", "send", "post", "channel", "chat", "write"],
    parameters: z.object({
      teamId: z.string().describe("The unique ID of the team."),
      channelId: z.string().describe("The unique ID of the channel."),
      content: z.string().describe("The message content to send."),
      contentType: z.enum(["text", "html"]).optional().describe("The content type of the message body. Defaults to \"text\"."),
    }),
    handler: async (client: Client, params: {
      teamId: string;
      channelId: string;
      content: string;
      contentType?: "text" | "html";
    }) => {
      return await client
        .api(`/teams/${params.teamId}/channels/${params.channelId}/messages`)
        .post({
          body: {
            content: params.content,
            contentType: params.contentType || "text",
          },
        });
    },
  });

  registry.register({
    name: "teams_get_message",
    category: "teams",
    description: "Get a specific message from a Microsoft Teams channel by its message ID.",
    keywords: ["teams", "message", "get", "read", "channel", "details"],
    parameters: z.object({
      teamId: z.string().describe("The unique ID of the team."),
      channelId: z.string().describe("The unique ID of the channel."),
      messageId: z.string().describe("The unique ID of the message."),
    }),
    handler: async (client: Client, params: {
      teamId: string;
      channelId: string;
      messageId: string;
    }) => {
      return await client
        .api(`/teams/${params.teamId}/channels/${params.channelId}/messages/${params.messageId}`)
        .get();
    },
  });
}

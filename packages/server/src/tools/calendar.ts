import { z } from "zod";
import type { Client } from "@microsoft/microsoft-graph-client";
import type { ToolRegistry } from "../registry/tool-registry.js";

export function registerCalendarTools(registry: ToolRegistry): void {
  registry.register({
    name: "calendar_list_events",
    category: "calendar",
    description:
      "List calendar events for a user within a specified date range using the calendarView endpoint.",
    keywords: [
      "calendar",
      "events",
      "list",
      "schedule",
      "meetings",
      "appointments",
    ],
    parameters: z.object({
      userId: z.string().describe("The ID or UPN of the user"),
      startDateTime: z
        .string()
        .describe(
          "Start of the date range in ISO 8601 format (e.g. 2024-01-01T00:00:00Z)",
        ),
      endDateTime: z
        .string()
        .describe(
          "End of the date range in ISO 8601 format (e.g. 2024-01-31T23:59:59Z)",
        ),
      $top: z
        .number()
        .optional()
        .describe("Maximum number of events to return"),
      $select: z
        .string()
        .optional()
        .describe("Comma-separated list of properties to select"),
    }),
    handler: async (client: Client, params: Record<string, unknown>) => {
      const { userId, startDateTime, endDateTime, $top, $select } = params as {
        userId: string;
        startDateTime: string;
        endDateTime: string;
        $top?: number;
        $select?: string;
      };

      let request = client
        .api(`/users/${userId}/calendarView`)
        .query({ startDateTime, endDateTime });

      if ($top !== undefined) {
        request = request.top($top);
      }
      if ($select) {
        request = request.select($select);
      }

      return await request.get();
    },
  });

  registry.register({
    name: "calendar_get_event",
    category: "calendar",
    description: "Get a specific calendar event by its ID for a given user.",
    keywords: ["calendar", "event", "get", "details", "meeting"],
    parameters: z.object({
      userId: z.string().describe("The ID or UPN of the user"),
      eventId: z.string().describe("The ID of the event to retrieve"),
    }),
    handler: async (client: Client, params: Record<string, unknown>) => {
      const { userId, eventId } = params as {
        userId: string;
        eventId: string;
      };

      return await client.api(`/users/${userId}/events/${eventId}`).get();
    },
  });

  registry.register({
    name: "calendar_create_event",
    category: "calendar",
    description: "Create a new calendar event for a user.",
    keywords: [
      "calendar",
      "event",
      "create",
      "new",
      "meeting",
      "schedule",
      "appointment",
    ],
    parameters: z.object({
      userId: z.string().describe("The ID or UPN of the user"),
      subject: z.string().describe("The subject/title of the event"),
      start: z
        .object({
          dateTime: z
            .string()
            .describe("Start date and time in ISO 8601 format"),
          timeZone: z
            .string()
            .describe("Time zone for the start time (e.g. UTC, Pacific Standard Time)"),
        })
        .describe("The start date, time, and time zone of the event"),
      end: z
        .object({
          dateTime: z
            .string()
            .describe("End date and time in ISO 8601 format"),
          timeZone: z
            .string()
            .describe("Time zone for the end time (e.g. UTC, Pacific Standard Time)"),
        })
        .describe("The end date, time, and time zone of the event"),
      body: z
        .object({
          contentType: z
            .enum(["text", "html"])
            .optional()
            .describe("The type of the body content"),
          content: z.string().describe("The content of the event body"),
        })
        .optional()
        .describe("The body/description of the event"),
      attendees: z
        .array(
          z.object({
            emailAddress: z.object({
              address: z.string().describe("Email address of the attendee"),
              name: z.string().optional().describe("Display name of the attendee"),
            }),
            type: z
              .enum(["required", "optional", "resource"])
              .optional()
              .describe("The type of attendee"),
          }),
        )
        .optional()
        .describe("List of attendees for the event"),
      location: z
        .object({
          displayName: z.string().describe("The name of the location"),
        })
        .optional()
        .describe("The location of the event"),
      isOnlineMeeting: z
        .boolean()
        .optional()
        .describe("Whether the event is an online meeting (e.g. Teams)"),
    }),
    handler: async (client: Client, params: Record<string, unknown>) => {
      const { userId, ...eventData } = params as {
        userId: string;
        subject: string;
        start: { dateTime: string; timeZone: string };
        end: { dateTime: string; timeZone: string };
        body?: { contentType?: string; content: string };
        attendees?: Array<{
          emailAddress: { address: string; name?: string };
          type?: string;
        }>;
        location?: { displayName: string };
        isOnlineMeeting?: boolean;
      };

      return await client.api(`/users/${userId}/events`).post(eventData);
    },
  });

  registry.register({
    name: "calendar_update_event",
    category: "calendar",
    description: "Update an existing calendar event for a user.",
    keywords: [
      "calendar",
      "event",
      "update",
      "edit",
      "modify",
      "meeting",
      "reschedule",
    ],
    parameters: z.object({
      userId: z.string().describe("The ID or UPN of the user"),
      eventId: z.string().describe("The ID of the event to update"),
      subject: z.string().optional().describe("Updated subject/title"),
      start: z
        .object({
          dateTime: z
            .string()
            .describe("Start date and time in ISO 8601 format"),
          timeZone: z
            .string()
            .describe("Time zone for the start time"),
        })
        .optional()
        .describe("Updated start date, time, and time zone"),
      end: z
        .object({
          dateTime: z
            .string()
            .describe("End date and time in ISO 8601 format"),
          timeZone: z
            .string()
            .describe("Time zone for the end time"),
        })
        .optional()
        .describe("Updated end date, time, and time zone"),
      body: z
        .object({
          contentType: z
            .enum(["text", "html"])
            .optional()
            .describe("The type of the body content"),
          content: z.string().describe("The content of the event body"),
        })
        .optional()
        .describe("Updated body/description"),
      attendees: z
        .array(
          z.object({
            emailAddress: z.object({
              address: z.string().describe("Email address of the attendee"),
              name: z.string().optional().describe("Display name of the attendee"),
            }),
            type: z
              .enum(["required", "optional", "resource"])
              .optional()
              .describe("The type of attendee"),
          }),
        )
        .optional()
        .describe("Updated list of attendees"),
      location: z
        .object({
          displayName: z.string().describe("The name of the location"),
        })
        .optional()
        .describe("Updated location"),
      isOnlineMeeting: z
        .boolean()
        .optional()
        .describe("Whether the event is an online meeting"),
    }),
    handler: async (client: Client, params: Record<string, unknown>) => {
      const { userId, eventId, ...updateData } = params as {
        userId: string;
        eventId: string;
        [key: string]: unknown;
      };

      return await client
        .api(`/users/${userId}/events/${eventId}`)
        .patch(updateData);
    },
  });

  registry.register({
    name: "calendar_delete_event",
    category: "calendar",
    description: "Delete a calendar event for a user.",
    keywords: ["calendar", "event", "delete", "remove", "cancel", "meeting"],
    parameters: z.object({
      userId: z.string().describe("The ID or UPN of the user"),
      eventId: z.string().describe("The ID of the event to delete"),
    }),
    handler: async (client: Client, params: Record<string, unknown>) => {
      const { userId, eventId } = params as {
        userId: string;
        eventId: string;
      };

      await client.api(`/users/${userId}/events/${eventId}`).delete();
      return { success: true, message: `Event ${eventId} deleted successfully` };
    },
  });

  registry.register({
    name: "calendar_find_free_busy",
    category: "calendar",
    description:
      "Get the free/busy availability schedule for one or more users.",
    keywords: [
      "calendar",
      "free",
      "busy",
      "availability",
      "schedule",
      "free/busy",
    ],
    parameters: z.object({
      schedules: z
        .array(z.string())
        .describe("Array of email addresses to check availability for"),
      startTime: z
        .object({
          dateTime: z
            .string()
            .describe("Start date and time in ISO 8601 format"),
          timeZone: z.string().describe("Time zone (e.g. UTC, Pacific Standard Time)"),
        })
        .describe("Start of the time range to check"),
      endTime: z
        .object({
          dateTime: z
            .string()
            .describe("End date and time in ISO 8601 format"),
          timeZone: z.string().describe("Time zone (e.g. UTC, Pacific Standard Time)"),
        })
        .describe("End of the time range to check"),
      availabilityViewInterval: z
        .number()
        .optional()
        .describe(
          "Duration of each time slot in the availability view, in minutes (default is 30)",
        ),
    }),
    handler: async (client: Client, params: Record<string, unknown>) => {
      const { schedules, startTime, endTime, availabilityViewInterval } =
        params as {
          schedules: string[];
          startTime: { dateTime: string; timeZone: string };
          endTime: { dateTime: string; timeZone: string };
          availabilityViewInterval?: number;
        };

      const requestBody: Record<string, unknown> = {
        schedules,
        startTime,
        endTime,
      };

      if (availabilityViewInterval !== undefined) {
        requestBody.availabilityViewInterval = availabilityViewInterval;
      }

      return await client
        .api("/me/calendar/getSchedule")
        .post(requestBody);
    },
  });
}

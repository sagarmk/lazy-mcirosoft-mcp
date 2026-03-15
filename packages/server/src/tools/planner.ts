import { z } from "zod";
import type { Client } from "@microsoft/microsoft-graph-client";
import type { ToolRegistry } from "../registry/tool-registry.js";

export function registerPlannerTools(registry: ToolRegistry): void {
  registry.register({
    name: "planner_list_plans",
    category: "planner",
    description: "List all Planner plans for a specific Microsoft 365 group.",
    keywords: ["planner", "plans", "list", "group", "project"],
    parameters: z.object({
      groupId: z.string().describe("The unique identifier of the Microsoft 365 group."),
    }),
    handler: async (client: Client, params: { groupId: string }) => {
      return await client.api(`/groups/${params.groupId}/planner/plans`).get();
    },
  });

  registry.register({
    name: "planner_get_plan",
    category: "planner",
    description: "Get details of a specific Planner plan.",
    keywords: ["planner", "plan", "get", "details"],
    parameters: z.object({
      planId: z.string().describe("The unique identifier of the Planner plan."),
    }),
    handler: async (client: Client, params: { planId: string }) => {
      return await client.api(`/planner/plans/${params.planId}`).get();
    },
  });

  registry.register({
    name: "planner_list_tasks",
    category: "planner",
    description: "List all tasks in a specific Planner plan.",
    keywords: ["planner", "tasks", "list", "plan", "todo", "work items"],
    parameters: z.object({
      planId: z.string().describe("The unique identifier of the Planner plan."),
    }),
    handler: async (client: Client, params: { planId: string }) => {
      return await client.api(`/planner/plans/${params.planId}/tasks`).get();
    },
  });

  registry.register({
    name: "planner_create_task",
    category: "planner",
    description: "Create a new task in a Planner plan.",
    keywords: ["planner", "task", "create", "new", "add", "assign"],
    parameters: z.object({
      planId: z.string().describe("The unique identifier of the Planner plan to create the task in."),
      title: z.string().describe("The title of the task."),
      bucketId: z.string().optional().describe("The bucket ID to place the task in."),
      assignedTo: z.string().optional().describe("The user ID to assign the task to."),
      dueDateTime: z.string().optional().describe("The due date and time for the task in ISO 8601 format (e.g. \"2025-12-31T00:00:00Z\")."),
    }),
    handler: async (client: Client, params: {
      planId: string;
      title: string;
      bucketId?: string;
      assignedTo?: string;
      dueDateTime?: string;
    }) => {
      const body: Record<string, unknown> = {
        planId: params.planId,
        title: params.title,
      };

      if (params.bucketId) {
        body.bucketId = params.bucketId;
      }

      if (params.assignedTo) {
        body.assignments = {
          [params.assignedTo]: {
            "@odata.type": "#microsoft.graph.plannerAssignment",
            orderHint: " !",
          },
        };
      }

      if (params.dueDateTime) {
        body.dueDateTime = params.dueDateTime;
      }

      return await client.api("/planner/tasks").post(body);
    },
  });

  registry.register({
    name: "planner_update_task",
    category: "planner",
    description: "Update an existing Planner task. Requires the task's current etag for concurrency control.",
    keywords: ["planner", "task", "update", "modify", "edit", "progress", "complete"],
    parameters: z.object({
      taskId: z.string().describe("The unique identifier of the task to update."),
      title: z.string().optional().describe("The updated title of the task."),
      percentComplete: z.number().optional().describe("The percentage of task completion (0, 25, 50, 75, or 100)."),
      dueDateTime: z.string().optional().describe("The updated due date and time in ISO 8601 format (e.g. \"2025-12-31T00:00:00Z\")."),
    }),
    handler: async (client: Client, params: {
      taskId: string;
      title?: string;
      percentComplete?: number;
      dueDateTime?: string;
    }) => {
      // Fetch the current task to get its etag for the If-Match header
      const currentTask = await client.api(`/planner/tasks/${params.taskId}`).get();
      const etag = currentTask["@odata.etag"];

      const body: Record<string, unknown> = {};

      if (params.title !== undefined) {
        body.title = params.title;
      }
      if (params.percentComplete !== undefined) {
        body.percentComplete = params.percentComplete;
      }
      if (params.dueDateTime !== undefined) {
        body.dueDateTime = params.dueDateTime;
      }

      await client
        .api(`/planner/tasks/${params.taskId}`)
        .header("If-Match", etag)
        .patch(body);

      return { success: true, message: `Task ${params.taskId} updated successfully.` };
    },
  });
}

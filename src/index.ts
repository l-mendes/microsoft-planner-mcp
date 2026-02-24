import "dotenv/config";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { createGraphClient } from "./graph-client.js";
import { PlannerService } from "./planner-service.js";
import { ok, err } from "./utils.js";

// ── Bootstrap ───────────────────────────────────────────────────────────────

const graphClient = createGraphClient();
const planner = new PlannerService(graphClient);

const server = new McpServer({
  name: "microsoft-planner-mcp",
  version: "1.0.2",
});

// ── Tools ───────────────────────────────────────────────────────────────────

// 0. List Microsoft 365 groups
server.registerTool(
  "list_groups",
  {
    description:
      "List Microsoft 365 groups (teams) that have Planner plans. Use this to discover the groupId needed for list_plans.",
    inputSchema: {
      search: z
        .string()
        .optional()
        .describe("Filter groups by name prefix (e.g. 'Marketing'). Omit to list all."),
    },
  },
  async ({ search }) => {
    try {
      const groups = await planner.listGroups(search);
      return ok(groups);
    } catch (e) {
      return err(e);
    }
  }
);

// 1. List plans for a group
server.registerTool(
  "list_plans",
  {
    description: "List all Planner plans for a Microsoft 365 group (team)",
    inputSchema: { groupId: z.string().describe("The ID of the Microsoft 365 group") },
  },
  async ({ groupId }) => {
    try {
      const plans = await planner.listPlans(groupId);
      return ok(plans);
    } catch (e) {
      return err(e);
    }
  }
);

// 2. List buckets in a plan
server.registerTool(
  "list_buckets",
  {
    description: "List all buckets in a Planner plan",
    inputSchema: { planId: z.string().optional().describe("The ID of the plan. Defaults to PLAN_ID env var if omitted.") },
  },
  async ({ planId }) => {
    try {
      const resolvedPlanId = planId ?? process.env.PLAN_ID;
      if (!resolvedPlanId) {
        return err(new Error("planId is required (or set PLAN_ID in .env)"));
      }

      const buckets = await planner.listBuckets(resolvedPlanId);
      return ok(buckets);
    } catch (e) {
      return err(e);
    }
  }
);

// 3. List tasks in a plan
server.registerTool(
  "list_tasks",
  {
    description: "List all tasks in a Planner plan",
    inputSchema: { planId: z.string().optional().describe("The ID of the plan. Defaults to PLAN_ID env var if omitted.") },
  },
  async ({ planId }) => {
    try {
      const resolvedPlanId = planId ?? process.env.PLAN_ID;
      if (!resolvedPlanId) {
        return err(new Error("planId is required (or set PLAN_ID in .env)"));
      }

      const tasks = await planner.listTasks(resolvedPlanId);
      return ok(tasks);
    } catch (e) {
      return err(e);
    }
  }
);

// 4. Get a single task
server.registerTool(
  "get_task",
  {
    description: "Get details of a specific Planner task by its ID",
    inputSchema: { taskId: z.string().describe("The ID of the task") },
  },
  async ({ taskId }) => {
    try {
      const [task, details] = await Promise.all([
        planner.getTask(taskId),
        planner.getTaskDetails(taskId),
      ]);
      return ok({ ...task, details });
    } catch (e) {
      return err(e);
    }
  }
);

// 5. Create a task
server.registerTool(
  "create_task",
  {
    description: "Create a new task in Microsoft Planner",
    inputSchema: {
      planId: z.string().optional().describe("The ID of the plan to create the task in. Defaults to PLAN_ID env var if omitted."),
      title: z.string().describe("Title of the task"),
      bucketId: z.string().optional().describe("The bucket ID to place the task in"),
      startDateTime: z
        .string()
        .optional()
        .describe("Start date in ISO 8601 format (e.g. 2025-03-01T00:00:00Z)"),
      dueDateTime: z
        .string()
        .optional()
        .describe("Due date in ISO 8601 format (e.g. 2025-03-15T00:00:00Z)"),
      percentComplete: z
        .number()
        .min(0)
        .max(100)
        .optional()
        .describe("Progress percentage (0, 25, 50, 75, 100)"),
      priority: z
        .number()
        .min(0)
        .max(10)
        .optional()
        .describe("Priority: 1 = Urgent, 3 = Important, 5 = Medium, 9 = Low"),
      assigneeIds: z
        .array(z.string())
        .optional()
        .describe("Array of user IDs to assign the task to"),
      categoryIds: z
        .array(z.string())
        .optional()
        .describe("Category keys to apply (e.g. ['category1', 'category3']). Use list_categories to discover available keys."),
    },
  },
  async ({ planId, title, bucketId, startDateTime, dueDateTime, percentComplete, priority, assigneeIds, categoryIds }) => {
    try {
      const resolvedPlanId = planId ?? process.env.PLAN_ID;
      if (!resolvedPlanId) {
        return err(new Error("planId is required (or set PLAN_ID in .env)"));
      }

      const assignments: Record<string, { "@odata.type": string; orderHint: string }> = {};
      if (assigneeIds) {
        for (const userId of assigneeIds) {
          assignments[userId] = {
            "@odata.type": "#microsoft.graph.plannerAssignment",
            orderHint: " !",
          };
        }
      }

      const appliedCategories: Record<string, boolean> = {};
      if (categoryIds) {
        for (const cat of categoryIds) appliedCategories[cat] = true;
      }

      const task = await planner.createTask({
        planId: resolvedPlanId,
        title,
        bucketId,
        startDateTime,
        dueDateTime,
        percentComplete,
        priority,
        assignments: Object.keys(assignments).length > 0 ? assignments : undefined,
        appliedCategories: Object.keys(appliedCategories).length > 0 ? appliedCategories : undefined,
      });
      return ok(task);
    } catch (e) {
      return err(e);
    }
  }
);

// 6. Update a task
server.registerTool(
  "update_task",
  {
    description: "Update an existing Planner task (title, dates, progress, priority, bucket, assignments)",
    inputSchema: {
      taskId: z.string().describe("The ID of the task to update"),
      title: z.string().optional().describe("New title for the task"),
      bucketId: z.string().optional().describe("Move the task to a different bucket"),
      startDateTime: z
        .string()
        .nullable()
        .optional()
        .describe("New start date (ISO 8601) or null to clear"),
      dueDateTime: z
        .string()
        .nullable()
        .optional()
        .describe("New due date (ISO 8601) or null to clear"),
      percentComplete: z
        .number()
        .min(0)
        .max(100)
        .optional()
        .describe("Progress percentage (0, 25, 50, 75, 100)"),
      priority: z
        .number()
        .min(0)
        .max(10)
        .optional()
        .describe("Priority: 1 = Urgent, 3 = Important, 5 = Medium, 9 = Low"),
      assigneeIds: z
        .array(z.string())
        .optional()
        .describe("Array of user IDs to assign (replaces current assignments)"),
      categoryId: z
        .string()
        .optional()
        .describe("Single category key to apply (e.g. 'category1'). Due to API limitations only one category can be updated at a time. Use list_categories to discover available keys."),
      description: z.string().optional().describe("Task description / notes"),
    },
  },
  async ({
    taskId,
    title,
    bucketId,
    startDateTime,
    dueDateTime,
    percentComplete,
    priority,
    assigneeIds,
    categoryId,
    description,
  }) => {
    try {
      // Build task-level updates
      const updates: Record<string, unknown> = {};
      if (title !== undefined) updates.title = title;
      if (bucketId !== undefined) updates.bucketId = bucketId;
      if (startDateTime !== undefined) updates.startDateTime = startDateTime;
      if (dueDateTime !== undefined) updates.dueDateTime = dueDateTime;
      if (percentComplete !== undefined) updates.percentComplete = percentComplete;
      if (priority !== undefined) updates.priority = priority;
      if (assigneeIds !== undefined) {
        const assignments: Record<string, { "@odata.type": string; orderHint: string }> = {};
        for (const userId of assigneeIds) {
          assignments[userId] = {
            "@odata.type": "#microsoft.graph.plannerAssignment",
            orderHint: " !",
          };
        }
        updates.assignments = assignments;
      }
      if (categoryId !== undefined) {
        updates.appliedCategories = { [categoryId]: true };
      }

      const results: unknown[] = [];

      // Update task properties
      if (Object.keys(updates).length > 0) {
        const task = await planner.getTask(taskId);
        const etag = task["@odata.etag"];
        await planner.updateTask(taskId, etag, updates);
        results.push({ taskUpdated: true });
      }

      // Update task details (description)
      if (description !== undefined) {
        const details = await planner.getTaskDetails(taskId);
        const detailsEtag = details["@odata.etag"];
        await planner.updateTaskDetails(taskId, detailsEtag, {
          description,
          previewType: "description",
        });
        results.push({ detailsUpdated: true });
      }

      // Return the fresh task
      const updatedTask = await planner.getTask(taskId);
      return ok(updatedTask);
    } catch (e) {
      return err(e);
    }
  }
);

// 7. Complete a task
server.registerTool(
  "complete_task",
  {
    description: "Mark a Planner task as complete (sets percentComplete to 100)",
    inputSchema: { taskId: z.string().describe("The ID of the task to complete") },
  },
  async ({ taskId }) => {
    try {
      await planner.completeTask(taskId);
      const task = await planner.getTask(taskId);
      return ok(task);
    } catch (e) {
      return err(e);
    }
  }
);

// 8. List categories of a plan
server.registerTool(
  "list_categories",
  {
    description:
      "List the category labels (coloured tags) defined in a Planner plan. Returns an object with keys category1–category25 and their display names.",
    inputSchema: {
      planId: z
        .string()
        .optional()
        .describe("The ID of the plan. Defaults to PLAN_ID env var if omitted."),
    },
  },
  async ({ planId }) => {
    try {
      const resolvedPlanId = planId ?? process.env.PLAN_ID;
      if (!resolvedPlanId) {
        return err(new Error("planId is required (or set PLAN_ID in .env)"));
      }

      const details = await planner.getPlanDetails(resolvedPlanId);
      const categories = details.categoryDescriptions ?? {};
      // Return only entries that have a non-empty label
      const named = Object.fromEntries(
        Object.entries(categories).filter(([, v]) => v)
      );
      return ok(named);
    } catch (e) {
      return err(e);
    }
  }
);

// 9. Delete a task
server.registerTool(
  "delete_task",
  {
    description: "Delete a Planner task permanently",
    inputSchema: { taskId: z.string().describe("The ID of the task to delete") },
  },
  async ({ taskId }) => {
    try {
      await planner.deleteTask(taskId);
      return ok({ deleted: true, taskId });
    } catch (e) {
      return err(e);
    }
  }
);

// 10. List users
server.registerTool(
  "list_users",
  {
    description:
      "List users in the Microsoft 365 directory. Use this to discover user IDs for task assignment.",
    inputSchema: {
      search: z
        .string()
        .optional()
        .describe("Filter users by display name prefix (e.g. 'John'). Omit to list all."),
    },
  },
  async ({ search }) => {
    try {
      const users = await planner.listUsers(search);
      return ok(users);
    } catch (e) {
      return err(e);
    }
  }
);

// ── Start ───────────────────────────────────────────────────────────────────

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Planner MCP Server running on stdio");
}

main().catch((e) => {
  console.error("Fatal error:", e);
  process.exit(1);
});

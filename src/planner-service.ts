import { Client } from "@microsoft/microsoft-graph-client";

// ─── Types ──────────────────────────────────────────────────────────────────

export interface PlannerTaskInput {
  planId: string;
  bucketId?: string;
  title: string;
  assignments?: Record<string, { "@odata.type": string; orderHint: string }>;
  startDateTime?: string;
  dueDateTime?: string;
  percentComplete?: number;
  priority?: number;
  appliedCategories?: Record<string, boolean>;
}

export interface PlannerTaskUpdate {
  title?: string;
  bucketId?: string;
  startDateTime?: string | null;
  dueDateTime?: string | null;
  percentComplete?: number;
  priority?: number;
  assignments?: Record<string, { "@odata.type": string; orderHint: string } | null>;
  appliedCategories?: Record<string, boolean>;
}

export interface PlannerTaskDetails {
  description?: string;
  previewType?: string;
}

// ─── Service ────────────────────────────────────────────────────────────────

export class PlannerService {
  constructor(private client: Client) {}

  // ── Users ──────────────────────────────────────────────────────────

  /** List users in the directory. */
  async listUsers(search?: string) {
    let req = this.client
      .api("/users")
      .select("id,displayName,mail,userPrincipalName,jobTitle,department")
      .top(100);

    if (search) {
      req = req
        .header("ConsistencyLevel", "eventual")
        .search(`"displayName:${search}"`);
    }

    const res = await req.get();
    return res.value;
  }

  // ── Groups ─────────────────────────────────────────────────────────

  /** List all Microsoft 365 groups the app has access to. */
  async listGroups(search?: string) {
    let req = this.client
      .api("/groups")
      .select("id,displayName,description,mail,groupTypes")
      .top(100);

    if (search) {
      req = req
        .header("ConsistencyLevel", "eventual")
        .search(`"displayName:${search}"`);
    }

    const res = await req.get();
    return res.value;
  }

  // ── Plans ──────────────────────────────────────────────────────────

  /** List all plans for a given group (team). */
  async listPlans(groupId: string) {
    const res = await this.client.api(`/groups/${groupId}/planner/plans`).get();
    return res.value;
  }

  /** Get a single plan by ID. */
  async getPlan(planId: string) {
    return this.client.api(`/planner/plans/${planId}`).get();
  }

  /** Get plan details, including categoryDescriptions (the coloured labels). */
  async getPlanDetails(planId: string) {
    return this.client.api(`/planner/plans/${planId}/details`).get();
  }

  // ── Buckets ────────────────────────────────────────────────────────

  /** List buckets inside a plan. */
  async listBuckets(planId: string) {
    const res = await this.client.api(`/planner/plans/${planId}/buckets`).get();
    return res.value;
  }

  // ── Tasks ──────────────────────────────────────────────────────────

  /** List all tasks for a plan. */
  async listTasks(planId: string) {
    const res = await this.client.api(`/planner/plans/${planId}/tasks`).get();
    return res.value;
  }

  /** Get a single task by ID. */
  async getTask(taskId: string) {
    return this.client.api(`/planner/tasks/${taskId}`).get();
  }

  /** Get task details (description, checklist, references). */
  async getTaskDetails(taskId: string) {
    return this.client.api(`/planner/tasks/${taskId}/details`).get();
  }

  /** Create a new task. */
  async createTask(input: PlannerTaskInput) {
    return this.client.api("/planner/tasks").post(input);
  }

  /**
   * Update an existing task.
   * Requires the current ETag for optimistic concurrency.
   */
  async updateTask(taskId: string, etag: string, updates: PlannerTaskUpdate) {
    return this.client
      .api(`/planner/tasks/${taskId}`)
      .header("If-Match", etag)
      .patch(updates);
  }

  /**
   * Update task details (e.g. description).
   * Requires the current ETag of the task details object.
   */
  async updateTaskDetails(taskId: string, etag: string, details: PlannerTaskDetails) {
    return this.client
      .api(`/planner/tasks/${taskId}/details`)
      .header("If-Match", etag)
      .patch(details);
  }

  /**
   * Mark a task as complete (percentComplete = 100).
   * Fetches the current ETag automatically.
   */
  async completeTask(taskId: string) {
    const task = await this.getTask(taskId);
    const etag = task["@odata.etag"];
    return this.updateTask(taskId, etag, { percentComplete: 100 });
  }

  /** Delete a task. Fetches the current ETag automatically. */
  async deleteTask(taskId: string) {
    const task = await this.getTask(taskId);
    const etag = task["@odata.etag"];
    return this.client
      .api(`/planner/tasks/${taskId}`)
      .header("If-Match", etag)
      .delete();
  }
}

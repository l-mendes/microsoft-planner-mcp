import { describe, it, expect, vi, beforeEach } from "vitest";
import type { Mock } from "vitest";
import { PlannerService } from "../planner-service.js";
import type { Client } from "@microsoft/microsoft-graph-client";

// ── Mock factory ─────────────────────────────────────────────────────────────

/**
 * Creates a chainable mock that mimics the fluent Graph client request builder.
 * Methods like .select(), .top(), .header(), .search() return `this` so they
 * can be chained.  Terminal methods (.get, .post, .patch, .delete) are plain
 * vi.fn() that tests configure with mockResolvedValue / mockResolvedValueOnce.
 */
function createMockChain() {
  const chain = {
    select: vi.fn(),
    top: vi.fn(),
    header: vi.fn(),
    search: vi.fn(),
    get: vi.fn(),
    post: vi.fn(),
    patch: vi.fn(),
    delete: vi.fn(),
  };

  chain.select.mockReturnValue(chain);
  chain.top.mockReturnValue(chain);
  chain.header.mockReturnValue(chain);
  chain.search.mockReturnValue(chain);

  return chain;
}

type MockChain = ReturnType<typeof createMockChain>;

function createMockClient() {
  const chain = createMockChain();
  const client = { api: vi.fn().mockReturnValue(chain) } as unknown as Client;
  return { client, chain, api: (client as unknown as { api: Mock }).api };
}

// ── Fixtures ─────────────────────────────────────────────────────────────────

const ETAG = 'W/"some-etag-value"';

const fakeTask = {
  id: "task-1",
  title: "Write tests",
  percentComplete: 0,
  "@odata.etag": ETAG,
};

const fakeTaskDetails = {
  id: "task-1",
  description: "Write comprehensive tests",
  "@odata.etag": 'W/"details-etag"',
};

// ── Tests ────────────────────────────────────────────────────────────────────

describe("PlannerService", () => {
  let service: PlannerService;
  let chain: MockChain;
  let api: Mock;

  beforeEach(() => {
    const mock = createMockClient();
    service = new PlannerService(mock.client);
    chain = mock.chain;
    api = mock.api;
  });

  // ── listUsers ──────────────────────────────────────────────────────

  describe("listUsers", () => {
    it("calls the correct endpoint and returns values", async () => {
      const users = [{ id: "u1", displayName: "Alice" }];
      chain.get.mockResolvedValue({ value: users });

      const result = await service.listUsers();

      expect(api).toHaveBeenCalledWith("/users");
      expect(chain.select).toHaveBeenCalledWith(
        "id,displayName,mail,userPrincipalName,jobTitle,department"
      );
      expect(chain.top).toHaveBeenCalledWith(100);
      expect(result).toEqual(users);
    });

    it("adds ConsistencyLevel header and search param when search is provided", async () => {
      chain.get.mockResolvedValue({ value: [] });

      await service.listUsers("Alice");

      expect(chain.header).toHaveBeenCalledWith("ConsistencyLevel", "eventual");
      expect(chain.search).toHaveBeenCalledWith('"displayName:Alice"');
    });

    it("does NOT add search param when search is not provided", async () => {
      chain.get.mockResolvedValue({ value: [] });

      await service.listUsers();

      expect(chain.search).not.toHaveBeenCalled();
      expect(chain.header).not.toHaveBeenCalled();
    });
  });

  // ── listGroups ─────────────────────────────────────────────────────

  describe("listGroups", () => {
    it("calls the correct endpoint and returns values", async () => {
      const groups = [{ id: "g1", displayName: "Engineering" }];
      chain.get.mockResolvedValue({ value: groups });

      const result = await service.listGroups();

      expect(api).toHaveBeenCalledWith("/groups");
      expect(chain.select).toHaveBeenCalledWith(
        "id,displayName,description,mail,groupTypes"
      );
      expect(result).toEqual(groups);
    });

    it("adds ConsistencyLevel header and search param when search is provided", async () => {
      chain.get.mockResolvedValue({ value: [] });

      await service.listGroups("Engineering");

      expect(chain.header).toHaveBeenCalledWith("ConsistencyLevel", "eventual");
      expect(chain.search).toHaveBeenCalledWith('"displayName:Engineering"');
    });
  });

  // ── listPlans ──────────────────────────────────────────────────────

  describe("listPlans", () => {
    it("calls the correct endpoint with the given groupId", async () => {
      const plans = [{ id: "plan-1", title: "Q1 Planning" }];
      chain.get.mockResolvedValue({ value: plans });

      const result = await service.listPlans("group-abc");

      expect(api).toHaveBeenCalledWith("/groups/group-abc/planner/plans");
      expect(result).toEqual(plans);
    });
  });

  // ── getPlan ────────────────────────────────────────────────────────

  describe("getPlan", () => {
    it("calls the correct endpoint and returns the plan", async () => {
      const plan = { id: "plan-1", title: "Q1 Planning" };
      chain.get.mockResolvedValue(plan);

      const result = await service.getPlan("plan-1");

      expect(api).toHaveBeenCalledWith("/planner/plans/plan-1");
      expect(result).toEqual(plan);
    });
  });

  // ── getPlanDetails ─────────────────────────────────────────────────

  describe("getPlanDetails", () => {
    it("calls the correct endpoint and returns the plan details", async () => {
      const details = { categoryDescriptions: { category1: "Bug", category2: "" } };
      chain.get.mockResolvedValue(details);

      const result = await service.getPlanDetails("plan-1");

      expect(api).toHaveBeenCalledWith("/planner/plans/plan-1/details");
      expect(result).toEqual(details);
    });
  });

  // ── listBuckets ────────────────────────────────────────────────────

  describe("listBuckets", () => {
    it("calls the correct endpoint and returns buckets", async () => {
      const buckets = [{ id: "bucket-1", name: "Backlog" }];
      chain.get.mockResolvedValue({ value: buckets });

      const result = await service.listBuckets("plan-1");

      expect(api).toHaveBeenCalledWith("/planner/plans/plan-1/buckets");
      expect(result).toEqual(buckets);
    });
  });

  // ── listTasks ──────────────────────────────────────────────────────

  describe("listTasks", () => {
    it("calls the correct endpoint and returns tasks", async () => {
      const tasks = [fakeTask];
      chain.get.mockResolvedValue({ value: tasks });

      const result = await service.listTasks("plan-1");

      expect(api).toHaveBeenCalledWith("/planner/plans/plan-1/tasks");
      expect(result).toEqual(tasks);
    });
  });

  // ── getTask ────────────────────────────────────────────────────────

  describe("getTask", () => {
    it("calls the correct endpoint and returns the task", async () => {
      chain.get.mockResolvedValue(fakeTask);

      const result = await service.getTask("task-1");

      expect(api).toHaveBeenCalledWith("/planner/tasks/task-1");
      expect(result).toEqual(fakeTask);
    });
  });

  // ── getTaskDetails ─────────────────────────────────────────────────

  describe("getTaskDetails", () => {
    it("calls the correct endpoint and returns task details", async () => {
      chain.get.mockResolvedValue(fakeTaskDetails);

      const result = await service.getTaskDetails("task-1");

      expect(api).toHaveBeenCalledWith("/planner/tasks/task-1/details");
      expect(result).toEqual(fakeTaskDetails);
    });
  });

  // ── createTask ─────────────────────────────────────────────────────

  describe("createTask", () => {
    it("posts to /planner/tasks with the correct payload", async () => {
      const newTask = { ...fakeTask, id: "task-99" };
      chain.post.mockResolvedValue(newTask);

      const input = {
        planId: "plan-1",
        title: "Write tests",
        bucketId: "bucket-1",
        dueDateTime: "2026-03-01T00:00:00Z",
        percentComplete: 0,
        priority: 5,
      };

      const result = await service.createTask(input);

      expect(api).toHaveBeenCalledWith("/planner/tasks");
      expect(chain.post).toHaveBeenCalledWith(input);
      expect(result).toEqual(newTask);
    });
  });

  // ── updateTask ─────────────────────────────────────────────────────

  describe("updateTask", () => {
    it("patches the task with If-Match header and updates", async () => {
      chain.patch.mockResolvedValue({});

      const updates = { title: "Updated title", percentComplete: 50 };
      await service.updateTask("task-1", ETAG, updates);

      expect(api).toHaveBeenCalledWith("/planner/tasks/task-1");
      expect(chain.header).toHaveBeenCalledWith("If-Match", ETAG);
      expect(chain.patch).toHaveBeenCalledWith(updates);
    });
  });

  // ── updateTaskDetails ──────────────────────────────────────────────

  describe("updateTaskDetails", () => {
    it("patches task details with If-Match header", async () => {
      chain.patch.mockResolvedValue({});

      const detailsEtag = 'W/"details-etag"';
      const details = { description: "New description", previewType: "description" as const };

      await service.updateTaskDetails("task-1", detailsEtag, details);

      expect(api).toHaveBeenCalledWith("/planner/tasks/task-1/details");
      expect(chain.header).toHaveBeenCalledWith("If-Match", detailsEtag);
      expect(chain.patch).toHaveBeenCalledWith(details);
    });
  });

  // ── completeTask ───────────────────────────────────────────────────

  describe("completeTask", () => {
    it("fetches the task ETag then patches percentComplete to 100", async () => {
      // First call: getTask returns the task with an etag
      chain.get.mockResolvedValue(fakeTask);
      chain.patch.mockResolvedValue({});

      await service.completeTask("task-1");

      // api should have been called twice: once for getTask, once for updateTask
      expect(api).toHaveBeenNthCalledWith(1, "/planner/tasks/task-1");
      expect(api).toHaveBeenNthCalledWith(2, "/planner/tasks/task-1");

      expect(chain.header).toHaveBeenCalledWith("If-Match", ETAG);
      expect(chain.patch).toHaveBeenCalledWith({ percentComplete: 100 });
    });
  });

  // ── deleteTask ─────────────────────────────────────────────────────

  describe("deleteTask", () => {
    it("fetches the task ETag then sends DELETE with If-Match header", async () => {
      chain.get.mockResolvedValue(fakeTask);
      chain.delete.mockResolvedValue(undefined);

      await service.deleteTask("task-1");

      expect(api).toHaveBeenNthCalledWith(1, "/planner/tasks/task-1");
      expect(api).toHaveBeenNthCalledWith(2, "/planner/tasks/task-1");

      expect(chain.header).toHaveBeenCalledWith("If-Match", ETAG);
      expect(chain.delete).toHaveBeenCalled();
    });
  });
});

import type { PluginContext } from "@paperclipai/plugin-sdk";
import type { Issue } from "@paperclipai/shared";
import { ENTITY_TYPES, type M365Config } from "../constants.js";
import type { GraphClient } from "../graph/client.js";
import type {
  PlannerBucket,
  PlannerTask,
} from "../graph/types.js";
import { toPlannerStatus } from "../sync/status-map.js";
import type { PaperclipIssueStatus } from "../constants.js";

/**
 * Manages Planner task CRUD and entity tracking for issue sync.
 */
export class PlannerService {
  private bucketCache = new Map<string, PlannerBucket>();

  constructor(
    private readonly ctx: PluginContext,
    private readonly graph: GraphClient,
    private readonly config: M365Config,
  ) {}

  /**
   * Create a Planner task linked to a Paperclip issue.
   */
  async createTask(issue: Issue): Promise<PlannerTask> {
    const status = toPlannerStatus(issue.status as PaperclipIssueStatus);
    const bucketId = await this.resolveBucketId(status.bucketName);

    const task = await this.graph.post<PlannerTask>("/planner/tasks", {
      planId: this.config.plannerPlanId,
      bucketId,
      title: issue.title,
      percentComplete: status.percentComplete,
      dueDateTime: (issue as unknown as Record<string, unknown>).dueDate ?? null,
    });

    await this.ctx.entities.upsert({
      entityType: ENTITY_TYPES.plannerTask,
      scopeKind: "issue",
      scopeId: issue.id,
      externalId: task.id,
      title: task.title,
      status: "synced",
      data: {
        plannerTaskId: task.id,
        etag: task["@odata.etag"] ?? null,
        lastSyncedAt: new Date().toISOString(),
        bucketId: task.bucketId,
      },
    });

    this.ctx.logger.info("Created Planner task", {
      issueId: issue.id,
      taskId: task.id,
    });

    return task;
  }

  /**
   * Update an existing Planner task from Paperclip issue changes.
   */
  async updateTask(
    issueId: string,
    plannerTaskId: string,
    etag: string | null,
    patch: {
      title?: string;
      status?: PaperclipIssueStatus;
      description?: string | null;
    },
  ): Promise<PlannerTask> {
    const body: Record<string, unknown> = {};
    if (patch.title !== undefined) body.title = patch.title;
    if (patch.status !== undefined) {
      const mapped = toPlannerStatus(patch.status);
      body.percentComplete = mapped.percentComplete;
      body.bucketId = await this.resolveBucketId(mapped.bucketName);
    }

    // Planner PATCH requires If-Match. Fetch current ETag if we don't have one.
    let currentEtag = etag;
    if (!currentEtag) {
      const current = await this.getTask(plannerTaskId);
      currentEtag = current["@odata.etag"] ?? null;
    }

    const headers: Record<string, string> = {};
    if (currentEtag) headers["If-Match"] = currentEtag;

    const task = await this.graph.patch<PlannerTask>(
      `/planner/tasks/${plannerTaskId}`,
      body,
      { headers },
    );

    await this.ctx.entities.upsert({
      entityType: ENTITY_TYPES.plannerTask,
      scopeKind: "issue",
      scopeId: issueId,
      externalId: plannerTaskId,
      title: patch.title ?? undefined,
      status: "synced",
      data: {
        plannerTaskId,
        etag: task?.["@odata.etag"] ?? null,
        lastSyncedAt: new Date().toISOString(),
        bucketId: task?.bucketId ?? null,
      },
    });

    this.ctx.logger.info("Updated Planner task", {
      issueId,
      taskId: plannerTaskId,
    });

    return task;
  }

  /**
   * Fetch a single Planner task by ID.
   */
  async getTask(taskId: string): Promise<PlannerTask> {
    return this.graph.get<PlannerTask>(`/planner/tasks/${taskId}`);
  }

  /**
   * List all tasks in the configured plan.
   */
  async listTasks(): Promise<PlannerTask[]> {
    return this.graph.listAll<PlannerTask>(
      `/planner/plans/${this.config.plannerPlanId}/tasks`,
    );
  }

  /**
   * List all buckets in the configured plan.
   */
  async listBuckets(): Promise<PlannerBucket[]> {
    return this.graph.listAll<PlannerBucket>(
      `/planner/plans/${this.config.plannerPlanId}/buckets`,
    );
  }

  /**
   * Resolve a bucket name to its ID, creating the bucket if it doesn't exist.
   */
  private async resolveBucketId(bucketName: string): Promise<string> {
    // Check cache
    const cached = this.bucketCache.get(bucketName);
    if (cached) return cached.id;

    // Refresh cache from API
    const buckets = await this.listBuckets();
    this.bucketCache.clear();
    for (const bucket of buckets) {
      this.bucketCache.set(bucket.name, bucket);
    }

    const found = this.bucketCache.get(bucketName);
    if (found) return found.id;

    // Create bucket if it doesn't exist
    const newBucket = await this.graph.post<PlannerBucket>("/planner/buckets", {
      name: bucketName,
      planId: this.config.plannerPlanId,
    });
    this.bucketCache.set(newBucket.name, newBucket);
    return newBucket.id;
  }
}

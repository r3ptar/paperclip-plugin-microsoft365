import type { PluginContext } from "@paperclipai/plugin-sdk";
import type { Issue } from "@paperclipai/shared";
import { ENTITY_TYPES, STATE_KEYS, type M365Config, type PaperclipIssueStatus } from "../constants.js";
import type { PlannerService } from "../services/planner.js";
import type { PlannerBucket, PlannerTask } from "../graph/types.js";
import { isStatusInSync, toPaperclipStatus, toPlannerStatus } from "./status-map.js";
import { resolveConflict } from "./conflict.js";

interface EntityData {
  plannerTaskId: string;
  etag: string | null;
  lastSyncedAt: string;
  bucketId: string | null;
}

/**
 * Full reconciliation job: compares all Planner tasks against tracked entities
 * and syncs any that have drifted.
 */
export async function reconcile(
  ctx: PluginContext,
  planner: PlannerService,
  config: M365Config,
): Promise<{ synced: number; conflicts: number; errors: number }> {
  const stats = { synced: 0, conflicts: 0, errors: 0 };

  try {
    // Paginate through all tracked entities
    const allEntities = [];
    const PAGE_SIZE = 200;
    let offset = 0;
    let page;
    do {
      page = await ctx.entities.list({
        entityType: ENTITY_TYPES.plannerTask,
        limit: PAGE_SIZE,
        offset,
      });
      allEntities.push(...page);
      offset += PAGE_SIZE;
    } while (page.length === PAGE_SIZE);

    const [tasks, buckets] = await Promise.all([
      planner.listTasks(),
      planner.listBuckets(),
    ]);
    const trackedEntities = allEntities;

    const bucketMap = new Map(buckets.map((b) => [b.id, b]));
    const taskMap = new Map(tasks.map((t) => [t.id, t]));

    for (const entity of trackedEntities) {
      try {
        const data = entity.data as unknown as EntityData;
        if (!data?.plannerTaskId || !entity.scopeId) continue;

        const task = taskMap.get(data.plannerTaskId);
        if (!task) {
          ctx.logger.warn("Tracked Planner task not found — may have been deleted", {
            issueId: entity.scopeId,
            taskId: data.plannerTaskId,
          });
          continue;
        }

        const issue = await ctx.issues.get(entity.scopeId, "");
        if (!issue) continue;

        const bucket = bucketMap.get(task.bucketId);
        const bucketName = bucket?.name ?? "";

        if (isStatusInSync(issue.status as PaperclipIssueStatus, task.percentComplete, bucketName)) {
          continue; // Already in sync
        }

        // Conflict detected
        stats.conflicts += 1;
        const winner = resolveConflict({
          paperclipUpdatedAt: issue.updatedAt instanceof Date
            ? issue.updatedAt.toISOString()
            : String(issue.updatedAt),
          plannerUpdatedAt: task.lastModifiedDateTime ?? task.createdDateTime,
          strategy: config.conflictStrategy,
        });

        await ctx.activity.log({
          companyId: issue.companyId,
          entityType: "issue",
          entityId: issue.id,
          message: `Planner sync conflict resolved: ${winner} wins`,
          metadata: {
            issueStatus: issue.status,
            plannerPercent: task.percentComplete,
            plannerBucket: bucketName,
            winner,
          },
        });

        if (winner === "paperclip") {
          await planner.updateTask(issue.id, task.id, task["@odata.etag"] ?? null, {
            title: issue.title,
            status: issue.status as PaperclipIssueStatus,
          });
        } else {
          const newStatus = toPaperclipStatus(task.percentComplete, bucketName);
          await ctx.issues.update(issue.id, { status: newStatus }, issue.companyId);
        }

        stats.synced += 1;
      } catch (err) {
        stats.errors += 1;
        ctx.logger.error("Reconciliation error for entity", {
          entityId: entity.id,
          error: err instanceof Error ? err.message : String(err),
        });
      }
    }

    await ctx.state.set(
      { scopeKind: "instance", stateKey: STATE_KEYS.lastReconcileAt },
      new Date().toISOString(),
    );

    ctx.logger.info("Planner reconciliation complete", stats);
  } catch (err) {
    ctx.logger.error("Reconciliation failed", {
      error: err instanceof Error ? err.message : String(err),
    });
    stats.errors += 1;
  }

  return stats;
}

import type { PluginContext, ToolResult, ToolRunContext } from "@paperclipai/plugin-sdk";
import { ENTITY_TYPES } from "../constants.js";
import type { PlannerService } from "../services/planner.js";

export interface PlannerStatusParams {
  issueId: string;
}

export async function handlePlannerStatus(
  params: unknown,
  runCtx: ToolRunContext,
  ctx: PluginContext,
  planner: PlannerService,
): Promise<ToolResult> {
  const { issueId } = params as PlannerStatusParams;
  if (!issueId) {
    return { error: "issueId is required" };
  }

  const entities = await ctx.entities.list({
    entityType: ENTITY_TYPES.plannerTask,
    scopeKind: "issue",
    scopeId: issueId,
    limit: 1,
    offset: 0,
  });

  if (entities.length === 0) {
    return {
      content: `No Planner task linked to issue ${issueId}`,
      data: { linked: false },
    };
  }

  const entity = entities[0]!;
  const data = entity.data as { plannerTaskId?: string; lastSyncedAt?: string } | undefined;
  if (!data?.plannerTaskId) {
    return { content: "Entity found but missing task ID", data: { linked: false } };
  }

  const task = await planner.getTask(data.plannerTaskId);

  return {
    content: `Planner task "${task.title}" — ${task.percentComplete}% complete (last synced: ${data.lastSyncedAt ?? "unknown"})`,
    data: {
      linked: true,
      taskId: task.id,
      title: task.title,
      percentComplete: task.percentComplete,
      lastSyncedAt: data.lastSyncedAt,
    },
  };
}

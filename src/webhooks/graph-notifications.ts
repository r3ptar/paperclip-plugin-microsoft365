import type { PluginContext, PluginWebhookInput } from "@paperclipai/plugin-sdk";
import { ENTITY_TYPES, WEBHOOK_KEYS, type M365Config, type PaperclipIssueStatus } from "../constants.js";
import type { PlannerService } from "../services/planner.js";
import type { GraphChangeNotification, PlannerBucket } from "../graph/types.js";
import { toPaperclipStatus } from "../sync/status-map.js";
import type { GraphClient } from "../graph/client.js";

/**
 * Handle incoming Microsoft Graph change notifications.
 * Verifies the clientState secret before processing.
 */
export async function handleGraphNotification(
  ctx: PluginContext,
  input: PluginWebhookInput,
  config: M365Config,
  planner: PlannerService,
  graph: GraphClient,
): Promise<void> {
  // Graph subscription validation: Microsoft sends a POST with a `validationToken`
  // query parameter. The endpoint must respond with HTTP 200 and the token as
  // plain-text body. Since the plugin SDK's onWebhook returns void, the host
  // server must intercept validation requests at the route level before
  // dispatching to the plugin worker. This handler logs the event so operators
  // can confirm the request arrived, but the actual token echo must happen
  // in the host's webhook route handler.
  // See: https://learn.microsoft.com/en-us/graph/webhooks#notification-endpoint-validation
  const body = input.parsedBody as Record<string, unknown> | undefined;
  if (body && typeof body === "object" && "validationToken" in body) {
    ctx.logger.info("Graph webhook validation request received — host must echo validationToken in response");
    return;
  }

  const notification = input.parsedBody as GraphChangeNotification;
  if (!notification?.value?.length) {
    ctx.logger.warn("Empty Graph notification received");
    return;
  }

  // Verify clientState
  if (!config.webhookClientStateRef) {
    ctx.logger.warn(
      "webhookClientStateRef is not configured — accepting notification without verification. " +
      "Set a webhook client state secret in plugin settings to enable signature verification.",
    );
  } else {
    const expectedState = await ctx.secrets.resolve(config.webhookClientStateRef);
    for (const item of notification.value) {
      if (item.clientState !== expectedState) {
        ctx.logger.error("Graph notification clientState mismatch — rejecting", {
          subscriptionId: item.subscriptionId,
        });
        return;
      }
    }
  }

  for (const item of notification.value) {
    try {
      // Skip activity logging here — companyId is not available until after
      // entity lookup. The ctx.logger call below provides observability.
      ctx.logger.info("Processing Graph notification", {
        changeType: item.changeType,
        resource: item.resource,
        subscriptionId: item.subscriptionId,
      });

      const taskId = item.resourceData?.id;
      if (!taskId) continue;

      // Fetch the updated task
      const task = await planner.getTask(taskId);

      // Find the entity mapping
      const entities = await ctx.entities.list({
        entityType: ENTITY_TYPES.plannerTask,
        limit: 1,
        offset: 0,
        externalId: taskId,
      });

      if (entities.length === 0) {
        ctx.logger.debug("No tracked entity for Planner task — skipping", { taskId });
        continue;
      }

      const entity = entities[0]!;
      if (!entity.scopeId) continue;

      // Resolve bucket name
      const buckets = await planner.listBuckets();
      const bucket = buckets.find((b) => b.id === task.bucketId);
      const bucketName = bucket?.name ?? "";

      const newStatus = toPaperclipStatus(task.percentComplete, bucketName);

      // Update the Paperclip issue
      const issue = await ctx.issues.get(entity.scopeId, "");
      if (!issue) continue;

      if (issue.status !== newStatus || issue.title !== task.title) {
        const patch: Record<string, unknown> = {};
        if (issue.status !== newStatus) patch.status = newStatus;
        if (issue.title !== task.title) patch.title = task.title;

        await ctx.issues.update(issue.id, patch as { status?: PaperclipIssueStatus; title?: string }, issue.companyId);

        ctx.logger.info("Updated issue from Planner notification", {
          issueId: issue.id,
          taskId,
          newStatus,
        });
      }

      // Update entity tracking
      await ctx.entities.upsert({
        entityType: ENTITY_TYPES.plannerTask,
        scopeKind: "issue",
        scopeId: entity.scopeId,
        externalId: taskId,
        title: task.title,
        status: "synced",
        data: {
          plannerTaskId: taskId,
          etag: task["@odata.etag"] ?? null,
          lastSyncedAt: new Date().toISOString(),
          bucketId: task.bucketId,
        },
      });
    } catch (err) {
      ctx.logger.error("Error processing Graph notification", {
        error: err instanceof Error ? err.message : String(err),
        resource: item.resource,
      });
    }
  }
}

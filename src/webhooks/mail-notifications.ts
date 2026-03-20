import type { PluginContext, PluginWebhookInput } from "@paperclipai/plugin-sdk";
import { ENTITY_TYPES, type M365Config } from "../constants.js";
import type { GraphClient } from "../graph/client.js";
import type { GraphChangeNotification, GraphMailMessage } from "../graph/types.js";
import { parseInboundEmail } from "../services/email-parser.js";

/**
 * Validates that an ID is safe for Graph API URL path interpolation.
 */
const SAFE_GRAPH_ID_RE = /^[a-zA-Z0-9._:,@-]+$/;
function isValidGraphId(id: string): boolean {
  return SAFE_GRAPH_ID_RE.test(id) && !id.includes("..");
}

/**
 * Handle incoming Microsoft Graph mail change notifications.
 * Fetches the full message, parses it for Paperclip actions,
 * and updates issues or logs comments accordingly.
 */
export async function handleMailNotification(
  ctx: PluginContext,
  input: PluginWebhookInput,
  config: M365Config,
  outlookClient: GraphClient,
): Promise<void> {
  // Graph subscription validation: the host server must echo the
  // validationToken before dispatching to the plugin worker.
  const body = input.parsedBody as Record<string, unknown> | undefined;
  if (body && typeof body === "object" && "validationToken" in body) {
    ctx.logger.info("Mail webhook validation request received — host must echo validationToken in response");
    return;
  }

  const notification = input.parsedBody as GraphChangeNotification;
  if (!notification?.value?.length) {
    ctx.logger.warn("Empty mail notification received");
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
        ctx.logger.error("Mail notification clientState mismatch — rejecting", {
          subscriptionId: item.subscriptionId,
        });
        return;
      }
    }
  }

  for (const item of notification.value) {
    try {
      // Skip activity logging here — companyId is not available until after
      // entity/issue lookup. The ctx.logger call below provides observability.
      ctx.logger.info("Processing mail notification", {
        changeType: item.changeType,
        resource: item.resource,
        subscriptionId: item.subscriptionId,
      });

      const messageId = item.resourceData?.id;
      if (!messageId) continue;

      // Validate IDs before interpolating into URL paths
      if (!isValidGraphId(config.outlookMailboxUserId)) {
        ctx.logger.error("Invalid outlookMailboxUserId format — skipping", {
          outlookMailboxUserId: config.outlookMailboxUserId,
        });
        continue;
      }
      if (!isValidGraphId(messageId)) {
        ctx.logger.error("Invalid message ID format — skipping", { messageId });
        continue;
      }

      // Fetch the full message from Graph API
      const message = await outlookClient.get<GraphMailMessage>(
        `/users/${config.outlookMailboxUserId}/messages/${messageId}` +
        `?$select=id,subject,body,from,toRecipients,internetMessageHeaders,internetMessageId,conversationId,receivedDateTime`,
      );

      const action = parseInboundEmail(message);

      switch (action.kind) {
        case "status_change": {
          // Look up the issue — use empty companyId to match the pattern
          // used by graph-notifications.ts (entity-scoped lookup).
          const issue = await ctx.issues.get(action.issueId, "");
          if (!issue) {
            ctx.logger.warn("Issue not found for status change from email", {
              issueId: action.issueId,
              messageId,
            });
            break;
          }

          await ctx.issues.update(
            action.issueId,
            { status: action.newStatus },
            issue.companyId,
          );

          await ctx.activity.log({
            companyId: issue.companyId,
            message: `Status changed to "${action.newStatus}" via email reply from ${message.from.emailAddress.address}`,
            metadata: {
              issueId: action.issueId,
              newStatus: action.newStatus,
              sender: message.from.emailAddress.address,
              messageId: message.id,
            },
          });

          // Track the email as an entity
          await ctx.entities.upsert({
            entityType: ENTITY_TYPES.taskEmail,
            scopeKind: "issue",
            scopeId: action.issueId,
            externalId: message.id,
            title: message.subject,
            status: "processed",
            data: {
              messageId: message.id,
              conversationId: message.conversationId ?? null,
              action: "status_change",
              newStatus: action.newStatus,
              processedAt: new Date().toISOString(),
            },
          });

          ctx.logger.info("Processed status change from inbound email", {
            issueId: action.issueId,
            newStatus: action.newStatus,
            sender: message.from.emailAddress.address,
          });
          break;
        }

        case "comment": {
          const issue = await ctx.issues.get(action.issueId, "");
          if (!issue) {
            ctx.logger.warn("Issue not found for comment from email", {
              issueId: action.issueId,
              messageId,
            });
            break;
          }

          await ctx.activity.log({
            companyId: issue.companyId,
            message: `Email comment from ${message.from.emailAddress.name} <${message.from.emailAddress.address}>: ${action.body}`,
            metadata: {
              issueId: action.issueId,
              sender: message.from.emailAddress.address,
              senderName: message.from.emailAddress.name,
              messageId: message.id,
              body: action.body,
            },
          });

          // Track the email as an entity
          await ctx.entities.upsert({
            entityType: ENTITY_TYPES.taskEmail,
            scopeKind: "issue",
            scopeId: action.issueId,
            externalId: message.id,
            title: message.subject,
            status: "processed",
            data: {
              messageId: message.id,
              conversationId: message.conversationId ?? null,
              action: "comment",
              processedAt: new Date().toISOString(),
            },
          });

          ctx.logger.info("Logged comment from inbound email", {
            issueId: action.issueId,
            sender: message.from.emailAddress.address,
          });
          break;
        }

        case "unrecognized": {
          ctx.logger.warn("Unrecognized inbound email — skipping", {
            reason: action.reason,
            messageId,
            subject: message.subject,
          });
          break;
        }
      }
    } catch (err) {
      ctx.logger.error("Error processing mail notification", {
        error: err instanceof Error ? err.message : String(err),
        resource: item.resource,
      });
    }
  }
}

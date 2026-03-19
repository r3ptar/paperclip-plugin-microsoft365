import {
  definePlugin,
  runWorker,
  type PaperclipPlugin,
  type PluginContext,
  type PluginEvent,
  type PluginHealthDiagnostics,
  type PluginJobContext,
  type PluginWebhookInput,
  type ToolResult,
} from "@paperclipai/plugin-sdk";
import type { Issue } from "@paperclipai/shared";
import {
  DEFAULT_CONFIG,
  ENTITY_TYPES,
  JOB_KEYS,
  PLUGIN_ID,
  STATE_KEYS,
  TOOL_NAMES,
  WEBHOOK_KEYS,
  type M365Config,
  type PaperclipIssueStatus,
} from "./constants.js";
import { TokenManager } from "./graph/auth.js";
import { GraphClient } from "./graph/client.js";
import { PlannerService } from "./services/planner.js";
import { SharePointService } from "./services/sharepoint.js";
import { OutlookService } from "./services/outlook.js";
import { reconcile } from "./sync/reconcile.js";
import { handleGraphNotification } from "./webhooks/graph-notifications.js";
import { handleMailNotification } from "./webhooks/mail-notifications.js";
import { handleSharePointSearch } from "./tools/sharepoint-search.js";
import { handleSharePointRead } from "./tools/sharepoint-read.js";
import { handleSharePointUpload } from "./tools/sharepoint-upload.js";
import { handlePlannerStatus } from "./tools/planner-status.js";
import { handleOutlookSendTaskEmail } from "./tools/outlook-send-task-email.js";
import { validateConfig } from "./validation.js";
import type {
  GraphListResponse,
  GraphGroup,
  GraphSite,
  GraphDrive,
  GraphCalendar,
  PlannerPlan,
  DriveItem,
} from "./graph/types.js";

/**
 * Validates that an ID is safe for Graph API URL path interpolation.
 * Allows GUIDs, opaque Graph IDs (alphanumeric, dots, hyphens, colons),
 * and SharePoint site IDs (which contain commas, e.g. "contoso.sharepoint.com,guid,guid").
 * Rejects path traversal characters (/, \, ..) and whitespace.
 */
const SAFE_GRAPH_ID_RE = /^[a-zA-Z0-9._:,@-]+$/;
function isValidGraphId(id: string): boolean {
  return SAFE_GRAPH_ID_RE.test(id) && !id.includes("..");
}

let pluginCtx: PluginContext | null = null;
let tokenManager: TokenManager | null = null;
let configClient: GraphClient | null = null;
let plannerClient: GraphClient | null = null;
let sharepointClient: GraphClient | null = null;
let outlookClient: GraphClient | null = null;
let plannerService: PlannerService | null = null;
let sharepointService: SharePointService | null = null;
let outlookService: OutlookService | null = null;

async function getConfig(ctx: PluginContext): Promise<M365Config> {
  const raw = await ctx.config.get();
  // save-config persists to ctx.state (SDK lacks ctx.config.set()),
  // so state overrides the initial config from the host.
  const stateConfig = await ctx.state.get({
    scopeKind: "instance",
    stateKey: "plugin-config",
  }) as Partial<M365Config> | null;
  return { ...DEFAULT_CONFIG, ...(raw as Partial<M365Config>), ...(stateConfig ?? {}) };
}

function initServices(ctx: PluginContext, config: M365Config): void {
  // Reset all service references to prevent stale instances after config change
  tokenManager = null;
  configClient = null;
  plannerClient = null;
  sharepointClient = null;
  outlookClient = null;
  plannerService = null;
  sharepointService = null;
  outlookService = null;

  if (!config.tenantId || !config.clientId || !config.clientSecretRef) {
    ctx.logger.warn("M365 plugin: Azure AD credentials not configured");
    return;
  }

  tokenManager = new TokenManager(ctx, config.tenantId, config.clientId, config.clientSecretRef);
  configClient = new GraphClient(ctx, tokenManager, "config");

  if (config.enablePlanner) {
    plannerClient = new GraphClient(ctx, tokenManager, "planner");
    plannerService = new PlannerService(ctx, plannerClient, config);
  }

  if (config.enableSharePoint) {
    sharepointClient = new GraphClient(ctx, tokenManager, "sharepoint");
    sharepointService = new SharePointService(ctx, sharepointClient, config);
  }

  if (config.enableOutlook) {
    outlookClient = new GraphClient(ctx, tokenManager, "outlook");
    outlookService = new OutlookService(ctx, outlookClient, config);
  }
}

// ── Event Handlers ──────────────────────────────────────────────────────────

async function registerEventHandlers(ctx: PluginContext): Promise<void> {
  ctx.events.on("issue.created", async (event: PluginEvent) => {
    if (!event.companyId) return;
    const config = await getConfig(ctx);

    const payload = event.payload as { issueId?: string; companyId?: string };
    if (!payload.issueId) return;

    const issue = await ctx.issues.get(payload.issueId, event.companyId);
    if (!issue) return;

    // Planner sync
    if (config.enablePlanner && plannerService) {
      try {
        await plannerService.createTask(issue);
        await ctx.metrics.write("m365.planner.task_created", 1);
      } catch (err) {
        ctx.logger.error("Failed to create Planner task for new issue", {
          issueId: issue.id,
          error: err instanceof Error ? err.message : String(err),
        });
      }
    }

    // Create calendar event if issue has a due date and Outlook is enabled
    if (config.enableOutlook && outlookService) {
      const dueDate = (issue as unknown as Record<string, unknown>).dueDate as string | undefined;
      if (dueDate) {
        try {
          await outlookService.createDeadlineEvent(issue, dueDate);
          await ctx.metrics.write("m365.outlook.event_created", 1);
        } catch (err) {
          ctx.logger.error("Failed to create calendar event", {
            issueId: issue.id,
            error: err instanceof Error ? err.message : String(err),
          });
        }
      }
    }
  });

  ctx.events.on("issue.updated", async (event: PluginEvent) => {
    if (!event.companyId) return;
    const config = await getConfig(ctx);
    const payload = event.payload as {
      issueId?: string;
      companyId?: string;
      changes?: Record<string, unknown>;
    };
    if (!payload.issueId) return;

    const issue = await ctx.issues.get(payload.issueId, event.companyId);
    if (!issue) return;

    // Sync status/title to Planner
    if (config.enablePlanner && plannerService) {
      const entities = await ctx.entities.list({
        entityType: ENTITY_TYPES.plannerTask,
        scopeKind: "issue",
        scopeId: issue.id,
        limit: 1,
        offset: 0,
      });

      if (entities.length > 0) {
        const entityData = entities[0]!.data as {
          plannerTaskId?: string;
          etag?: string | null;
        };
        if (entityData?.plannerTaskId) {
          try {
            await plannerService.updateTask(
              issue.id,
              entityData.plannerTaskId,
              entityData.etag ?? null,
              {
                title: issue.title,
                status: issue.status as PaperclipIssueStatus,
              },
            );
            await ctx.metrics.write("m365.planner.task_updated", 1);
          } catch (err) {
            ctx.logger.error("Failed to update Planner task", {
              issueId: issue.id,
              error: err instanceof Error ? err.message : String(err),
            });
          }
        }
      }
    }

    // Update calendar event if deadline changed
    if (config.enableOutlook && outlookService) {
      const entities = await ctx.entities.list({
        entityType: ENTITY_TYPES.calendarEvent,
        scopeKind: "issue",
        scopeId: issue.id,
        limit: 1,
        offset: 0,
      });

      const dueDate = (issue as unknown as Record<string, unknown>).dueDate as string | undefined;
      const isDoneOrCancelled = issue.status === "done" || issue.status === "cancelled";

      if (entities.length > 0) {
        const eventData = entities[0]!.data as { eventId?: string };
        if (eventData?.eventId) {
          try {
            await outlookService.updateDeadlineEvent(
              issue.id,
              eventData.eventId,
              isDoneOrCancelled ? null : (dueDate ?? null),
            );
          } catch (err) {
            ctx.logger.error("Failed to update calendar event", {
              issueId: issue.id,
              error: err instanceof Error ? err.message : String(err),
            });
          }
        }
      } else if (dueDate && !isDoneOrCancelled) {
        try {
          await outlookService.createDeadlineEvent(issue, dueDate);
        } catch (err) {
          ctx.logger.error("Failed to create calendar event on update", {
            issueId: issue.id,
            error: err instanceof Error ? err.message : String(err),
          });
        }
      }
    }
  });
}

// ── Job Handlers ────────────────────────────────────────────────────────────

async function registerJobs(ctx: PluginContext): Promise<void> {
  ctx.jobs.register(JOB_KEYS.plannerReconcile, async (job: PluginJobContext) => {
    const config = await getConfig(ctx);
    if (!config.enablePlanner || !plannerService) {
      ctx.logger.info("Planner reconciliation skipped — not enabled");
      return;
    }

    const stats = await reconcile(ctx, plannerService, config);
    await ctx.metrics.write("m365.planner.reconcile.synced", stats.synced);
    await ctx.metrics.write("m365.planner.reconcile.conflicts", stats.conflicts);
    await ctx.metrics.write("m365.planner.reconcile.errors", stats.errors);
  });

  ctx.jobs.register(JOB_KEYS.graphSubscriptionRenew, async (job: PluginJobContext) => {
    const config = await getConfig(ctx);
    if (!config.enablePlanner || !plannerClient) return;

    const subId = await ctx.state.get({
      scopeKind: "instance",
      stateKey: STATE_KEYS.subscriptionId,
    }) as string | null;

    if (!subId) {
      ctx.logger.info("No Graph subscription to renew");
      return;
    }

    try {
      const expiry = new Date(Date.now() + 48 * 60 * 60 * 1000).toISOString();
      await plannerClient.patch(`/subscriptions/${subId}`, {
        expirationDateTime: expiry,
      });
      await ctx.state.set(
        { scopeKind: "instance", stateKey: STATE_KEYS.subscriptionExpiry },
        expiry,
      );
      ctx.logger.info("Renewed Graph subscription", { subId, expiry });
    } catch (err) {
      ctx.logger.error("Failed to renew Graph subscription", {
        error: err instanceof Error ? err.message : String(err),
      });
    }

    // Renew mail subscription if one exists
    if (outlookClient) {
      const mailSubId = await ctx.state.get({
        scopeKind: "instance",
        stateKey: STATE_KEYS.mailSubscriptionId,
      }) as string | null;

      if (mailSubId) {
        try {
          const mailExpiry = new Date(Date.now() + 48 * 60 * 60 * 1000).toISOString();
          await outlookClient.patch(`/subscriptions/${mailSubId}`, {
            expirationDateTime: mailExpiry,
          });
          await ctx.state.set(
            { scopeKind: "instance", stateKey: STATE_KEYS.mailSubscriptionExpiry },
            mailExpiry,
          );
          ctx.logger.info("Renewed mail subscription", { mailSubId, expiry: mailExpiry });
        } catch (err) {
          ctx.logger.error("Failed to renew mail subscription", {
            error: err instanceof Error ? err.message : String(err),
          });
        }
      }
    }
  });

  ctx.jobs.register(JOB_KEYS.outlookDigest, async (job: PluginJobContext) => {
    const config = await getConfig(ctx);
    if (!config.enableOutlook || !outlookService) return;

    try {
      const companies = await ctx.companies.list({ limit: 10, offset: 0 });
      const allIssues: Array<{ id: string; title: string; status: string; updatedAt: string }> = [];

      for (const company of companies) {
        const issues = await ctx.issues.list({ companyId: company.id, limit: 50, offset: 0 });
        const yesterday = new Date(Date.now() - 24 * 60 * 60 * 1000);

        for (const issue of issues) {
          const updatedAt = issue.updatedAt instanceof Date ? issue.updatedAt : new Date(String(issue.updatedAt));
          if (updatedAt > yesterday) {
            allIssues.push({
              id: issue.id,
              title: issue.title,
              status: issue.status,
              updatedAt: updatedAt.toISOString(),
            });
          }
        }
      }

      const html = outlookService.buildDigestHtml(allIssues);
      const date = new Date().toLocaleDateString("en-US", {
        weekday: "long",
        year: "numeric",
        month: "long",
        day: "numeric",
      });
      await outlookService.sendDigest(`Paperclip Daily Digest — ${date}`, html);
      await ctx.metrics.write("m365.outlook.digest_sent", 1);
    } catch (err) {
      ctx.logger.error("Failed to send digest", {
        error: err instanceof Error ? err.message : String(err),
      });
    }
  });

  ctx.jobs.register(JOB_KEYS.tokenHealthCheck, async (job: PluginJobContext) => {
    if (!tokenManager) return;

    const result = await tokenManager.healthCheck();
    await ctx.state.set(
      { scopeKind: "instance", stateKey: STATE_KEYS.syncHealth },
      { tokenHealthy: result.ok, checkedAt: new Date().toISOString() },
    );
    await ctx.metrics.write("m365.token.health", result.ok ? 1 : 0);

    if (!result.ok) {
      ctx.logger.error("Token health check failed — OAuth credentials may be invalid", {
        error: result.error,
      });
    }
  });
}

// ── Data & Action Handlers ──────────────────────────────────────────────────

async function registerDataHandlers(ctx: PluginContext): Promise<void> {
  ctx.data.register("sync-health", async () => {
    const config = await getConfig(ctx);
    const health = await ctx.state.get({
      scopeKind: "instance",
      stateKey: STATE_KEYS.syncHealth,
    });
    const lastReconcile = await ctx.state.get({
      scopeKind: "instance",
      stateKey: STATE_KEYS.lastReconcileAt,
    });
    const subExpiry = await ctx.state.get({
      scopeKind: "instance",
      stateKey: STATE_KEYS.subscriptionExpiry,
    });

    const taskCount = (await ctx.entities.list({
      entityType: ENTITY_TYPES.plannerTask,
      limit: 500,
      offset: 0,
    })).length;

    return {
      configured: Boolean(config.tenantId && config.clientId),
      enablePlanner: config.enablePlanner,
      enableSharePoint: config.enableSharePoint,
      enableOutlook: config.enableOutlook,
      health,
      lastReconcile,
      subscriptionExpiry: subExpiry,
      trackedTasks: taskCount,
    };
  });

  ctx.data.register("issue-m365", async (params) => {
    const issueId = typeof params.issueId === "string" ? params.issueId : "";
    if (!issueId) return { plannerTask: null, calendarEvent: null };

    const [plannerEntities, calendarEntities] = await Promise.all([
      ctx.entities.list({
        entityType: ENTITY_TYPES.plannerTask,
        scopeKind: "issue",
        scopeId: issueId,
        limit: 1,
        offset: 0,
      }),
      ctx.entities.list({
        entityType: ENTITY_TYPES.calendarEvent,
        scopeKind: "issue",
        scopeId: issueId,
        limit: 1,
        offset: 0,
      }),
    ]);

    return {
      plannerTask: plannerEntities[0] ?? null,
      calendarEvent: calendarEntities[0] ?? null,
    };
  });

  ctx.data.register("plugin-config", async () => {
    const config = await getConfig(ctx);
    // Return config data for the settings UI.
    // clientSecretRef is a reference identifier, not the raw secret, so it is
    // safe to send to the browser for round-tripping through the form.
    return {
      tenantId: config.tenantId,
      clientId: config.clientId,
      clientSecretRef: config.clientSecretRef,
      hasClientSecret: Boolean(config.clientSecretRef),
      enablePlanner: config.enablePlanner,
      enableSharePoint: config.enableSharePoint,
      enableOutlook: config.enableOutlook,
      plannerPlanId: config.plannerPlanId,
      plannerGroupId: config.plannerGroupId,
      conflictStrategy: config.conflictStrategy,
      sharepointSiteId: config.sharepointSiteId,
      sharepointDriveId: config.sharepointDriveId,
      sharepointUploadFolderId: config.sharepointUploadFolderId,
      maxDocSizeBytes: config.maxDocSizeBytes,
      outlookCalendarId: config.outlookCalendarId,
      digestRecipients: config.digestRecipients,
      digestSenderUserId: config.digestSenderUserId,
      hasWebhookClientState: Boolean(config.webhookClientStateRef),
    };
  });

  // ── Setup Wizard data handlers ─────────────────────────────────────────────

  /**
   * Returns the configClient if available, or creates a temporary GraphClient
   * from wizard-provided credentials (tenantId, clientId, clientSecret passed
   * as data handler params before config is saved).
   */
  function getWizardClient(params: Record<string, unknown>): GraphClient | null {
    if (configClient) return configClient;
    const tenantId = typeof params.tenantId === "string" ? params.tenantId : "";
    const clientId = typeof params.clientId === "string" ? params.clientId : "";
    const clientSecret = typeof params.clientSecret === "string" ? params.clientSecret : "";
    if (!tenantId || !clientId || !clientSecret) return null;
    const tm = new TokenManager(ctx, tenantId, clientId, "", clientSecret);
    return new GraphClient(ctx, tm, "wizard");
  }

  ctx.data.register("m365-groups", async (params) => {
    const client = getWizardClient(params);
    if (!client) return { error: "Azure AD credentials not configured" };
    try {
      const res = await client.get<GraphListResponse<GraphGroup>>(
        "/groups?$filter=groupTypes/any(c:c eq 'Unified')&$select=id,displayName&$top=100",
      );
      return {
        items: res.value.map((g) => ({ id: g.id, name: g.displayName })),
      };
    } catch (err) {
      return { error: err instanceof Error ? err.message : String(err) };
    }
  });

  ctx.data.register("m365-plans", async (params) => {
    const client = getWizardClient(params);
    if (!client) return { error: "Azure AD credentials not configured" };
    const groupId = typeof params.groupId === "string" ? params.groupId : "";
    if (!groupId) return { error: "groupId is required" };
    if (!isValidGraphId(groupId)) return { error: "Invalid groupId format" };
    try {
      const res = await client.get<GraphListResponse<PlannerPlan>>(
        `/groups/${groupId}/planner/plans?$select=id,title`,
      );
      return {
        items: res.value.map((p) => ({ id: p.id, name: p.title })),
      };
    } catch (err) {
      return { error: err instanceof Error ? err.message : String(err) };
    }
  });

  ctx.data.register("m365-sites", async (params) => {
    const client = getWizardClient(params);
    if (!client) return { error: "Azure AD credentials not configured" };
    try {
      const res = await client.get<GraphListResponse<GraphSite>>(
        "/sites?search=*&$select=id,displayName,webUrl&$top=100",
      );
      return {
        items: res.value.map((s) => ({ id: s.id, name: s.displayName })),
      };
    } catch (err) {
      return { error: err instanceof Error ? err.message : String(err) };
    }
  });

  ctx.data.register("m365-drives", async (params) => {
    const client = getWizardClient(params);
    if (!client) return { error: "Azure AD credentials not configured" };
    const siteId = typeof params.siteId === "string" ? params.siteId : "";
    if (!siteId) return { error: "siteId is required" };
    if (!isValidGraphId(siteId)) return { error: "Invalid siteId format" };
    try {
      const res = await client.get<GraphListResponse<GraphDrive>>(
        `/sites/${siteId}/drives?$select=id,name,driveType`,
      );
      return {
        items: res.value.map((d) => ({ id: d.id, name: d.name })),
      };
    } catch (err) {
      return { error: err instanceof Error ? err.message : String(err) };
    }
  });

  ctx.data.register("m365-folders", async (params) => {
    const client = getWizardClient(params);
    if (!client) return { error: "Azure AD credentials not configured" };
    const driveId = typeof params.driveId === "string" ? params.driveId : "";
    if (!driveId) return { error: "driveId is required" };
    if (!isValidGraphId(driveId)) return { error: "Invalid driveId format" };
    try {
      const res = await client.get<GraphListResponse<DriveItem>>(
        `/drives/${driveId}/root/children?$select=id,name,folder`,
      );
      return {
        items: res.value
          .filter((item) => item.folder !== undefined)
          .map((item) => ({ id: item.id, name: item.name })),
      };
    } catch (err) {
      return { error: err instanceof Error ? err.message : String(err) };
    }
  });

  ctx.data.register("m365-calendars", async (params) => {
    const client = getWizardClient(params);
    if (!client) return { error: "Azure AD credentials not configured" };
    const userId = typeof params.userId === "string" ? params.userId : "";
    if (!userId) return { error: "userId is required" };
    if (!isValidGraphId(userId)) return { error: "Invalid userId format" };
    try {
      const res = await client.get<GraphListResponse<GraphCalendar>>(
        `/users/${userId}/calendars?$select=id,name,isDefaultCalendar`,
      );
      return {
        items: res.value.map((c) => ({ id: c.id, name: c.name })),
      };
    } catch (err) {
      return { error: err instanceof Error ? err.message : String(err) };
    }
  });
}

async function registerActionHandlers(ctx: PluginContext): Promise<void> {
  ctx.actions.register("test-connection", async (params) => {
    const { tenantId, clientId, clientSecretRef, clientSecret } = (params ?? {}) as {
      tenantId?: string;
      clientId?: string;
      clientSecretRef?: string;
      clientSecret?: string;
    };

    let tm: TokenManager | null;

    if (tenantId && clientId && (clientSecret || clientSecretRef)) {
      // Use provided credentials (e.g. from Setup Wizard before config is saved)
      // Pass raw secret directly to avoid ctx.secrets.resolve() scope issues
      tm = new TokenManager(ctx, tenantId, clientId, clientSecretRef ?? "", clientSecret);
    } else {
      // Fall back to module-level instance (post-setup)
      tm = tokenManager;
    }

    if (!tm) {
      return { ok: false, error: "Azure AD credentials not configured" };
    }
    const result = await tm.healthCheck();
    return { ok: result.ok, error: result.ok ? null : result.error ?? "Failed to acquire OAuth token" };
  });

  ctx.actions.register("trigger-reconcile", async () => {
    const config = await getConfig(ctx);
    if (!config.enablePlanner || !plannerService) {
      return { ok: false, error: "Planner sync not enabled" };
    }
    const stats = await reconcile(ctx, plannerService, config);
    return { ok: true, stats };
  });

  ctx.actions.register("save-config", async (params) => {
    const incoming = params as Partial<M365Config>;

    // Merge with defaults for fields not provided
    const merged = { ...DEFAULT_CONFIG, ...incoming };

    // Validate
    const validation = validateConfig(merged);
    if (!validation.ok) {
      return { ok: false, errors: validation.errors, warnings: validation.warnings };
    }

    // Persist config
    // NOTE: ctx.config.set() does not exist on PluginConfigClient — fall back
    // to ctx.state.set() with instance-scoped "plugin-config" key.
    await ctx.state.set({ scopeKind: "instance", stateKey: "plugin-config" }, merged);

    // Reinitialize services with new config
    initServices(ctx, merged);

    return { ok: true, warnings: validation.warnings };
  });

  ctx.actions.register("create-mail-subscription", async (params) => {
    const config = await getConfig(ctx);

    if (!outlookClient) {
      return { ok: false, error: "Outlook integration is not enabled" };
    }
    if (!config.enableInboundEmail) {
      return { ok: false, error: "Inbound email processing is not enabled" };
    }
    if (!config.outlookMailboxUserId) {
      return { ok: false, error: "Outlook mailbox user ID is not configured" };
    }

    const { notificationUrl } = (params ?? {}) as { notificationUrl?: string };
    if (!notificationUrl) {
      return { ok: false, error: "notificationUrl parameter is required" };
    }

    let clientState: string | undefined;
    if (config.webhookClientStateRef) {
      clientState = await ctx.secrets.resolve(config.webhookClientStateRef);
    }

    try {
      const expirationDateTime = new Date(Date.now() + 48 * 60 * 60 * 1000).toISOString();

      const subscription = await outlookClient.post<{ id: string; expirationDateTime: string }>(
        "/subscriptions",
        {
          changeType: "created",
          notificationUrl,
          resource: `/users/${config.outlookMailboxUserId}/mailFolders/inbox/messages`,
          expirationDateTime,
          ...(clientState ? { clientState } : {}),
        },
      );

      await ctx.state.set(
        { scopeKind: "instance", stateKey: STATE_KEYS.mailSubscriptionId },
        subscription.id,
      );
      await ctx.state.set(
        { scopeKind: "instance", stateKey: STATE_KEYS.mailSubscriptionExpiry },
        subscription.expirationDateTime,
      );

      ctx.logger.info("Created mail subscription", {
        subscriptionId: subscription.id,
        expirationDateTime: subscription.expirationDateTime,
      });

      return { ok: true, subscriptionId: subscription.id };
    } catch (err) {
      ctx.logger.error("Failed to create mail subscription", {
        error: err instanceof Error ? err.message : String(err),
      });
      return {
        ok: false,
        error: `Failed to create mail subscription: ${err instanceof Error ? err.message : String(err)}`,
      };
    }
  });
}

// ── Tool Handlers ───────────────────────────────────────────────────────────

async function registerToolHandlers(ctx: PluginContext): Promise<void> {
  ctx.tools.register(
    TOOL_NAMES.sharepointSearch,
    {
      displayName: "SharePoint Search",
      description: "Search documents in the configured SharePoint site.",
      parametersSchema: {
        type: "object",
        properties: {
          query: { type: "string" },
          maxResults: { type: "number" },
        },
        required: ["query"],
      },
    },
    async (params, runCtx): Promise<ToolResult> => {
      if (!sharepointService) {
        return { error: "SharePoint integration is not enabled" };
      }
      return handleSharePointSearch(params, runCtx, sharepointService);
    },
  );

  ctx.tools.register(
    TOOL_NAMES.sharepointRead,
    {
      displayName: "SharePoint Read Document",
      description: "Read text content of a SharePoint document.",
      parametersSchema: {
        type: "object",
        properties: {
          driveId: { type: "string" },
          itemId: { type: "string" },
        },
        required: ["driveId", "itemId"],
      },
    },
    async (params, runCtx): Promise<ToolResult> => {
      if (!sharepointService) {
        return { error: "SharePoint integration is not enabled" };
      }
      return handleSharePointRead(params, runCtx, sharepointService);
    },
  );

  ctx.tools.register(
    TOOL_NAMES.sharepointUpload,
    {
      displayName: "SharePoint Upload",
      description: "Upload a file to SharePoint.",
      parametersSchema: {
        type: "object",
        properties: {
          fileName: { type: "string" },
          content: { type: "string" },
          contentType: { type: "string" },
        },
        required: ["fileName", "content"],
      },
    },
    async (params, runCtx): Promise<ToolResult> => {
      if (!sharepointService) {
        return { error: "SharePoint integration is not enabled" };
      }
      return handleSharePointUpload(params, runCtx, sharepointService);
    },
  );

  ctx.tools.register(
    TOOL_NAMES.plannerStatus,
    {
      displayName: "Planner Task Status",
      description: "Check the linked Planner task status for a Paperclip issue.",
      parametersSchema: {
        type: "object",
        properties: {
          issueId: { type: "string" },
        },
        required: ["issueId"],
      },
    },
    async (params, runCtx): Promise<ToolResult> => {
      if (!plannerService) {
        return { error: "Planner integration is not enabled" };
      }
      return handlePlannerStatus(params, runCtx, ctx, plannerService);
    },
  );

  ctx.tools.register(
    TOOL_NAMES.outlookSendTaskEmail,
    {
      displayName: "Send Task Email",
      description: "Send an email to someone about a specific Paperclip issue.",
      parametersSchema: {
        type: "object",
        properties: {
          issueId: { type: "string", description: "The Paperclip issue ID" },
          recipientEmail: { type: "string", description: "Email address of the recipient" },
          emailType: {
            type: "string",
            enum: ["assignment", "status_change", "blocked", "request", "custom"],
            description: "Type of email notification",
          },
          customMessage: { type: "string", description: "Custom message to include in the email" },
        },
        required: ["issueId", "recipientEmail"],
      },
    },
    async (params, runCtx): Promise<ToolResult> => {
      if (!outlookService) {
        return { error: "Outlook integration is not enabled" };
      }
      return handleOutlookSendTaskEmail(params, runCtx, ctx, outlookService);
    },
  );
}

// ── Plugin Definition ───────────────────────────────────────────────────────

const plugin: PaperclipPlugin = definePlugin({
  async setup(ctx) {
    pluginCtx = ctx;
    const config = await getConfig(ctx);
    initServices(ctx, config);

    await registerEventHandlers(ctx);
    await registerJobs(ctx);
    await registerDataHandlers(ctx);
    await registerActionHandlers(ctx);
    await registerToolHandlers(ctx);

    ctx.logger.info("Microsoft 365 plugin setup complete", {
      planner: config.enablePlanner,
      sharepoint: config.enableSharePoint,
      outlook: config.enableOutlook,
    });
  },

  async onHealth(): Promise<PluginHealthDiagnostics> {
    const ctx = pluginCtx;
    if (!ctx) return { status: "error", message: "Plugin not initialized" };

    const config = await getConfig(ctx);
    if (!config.tenantId || !config.clientId) {
      return { status: "degraded", message: "Azure AD credentials not configured" };
    }

    const health = await ctx.state.get({
      scopeKind: "instance",
      stateKey: STATE_KEYS.syncHealth,
    }) as { tokenHealthy?: boolean } | null;

    if (health && !health.tokenHealthy) {
      return { status: "error", message: "OAuth token health check failed" };
    }

    return {
      status: "ok",
      message: "Microsoft 365 plugin ready",
      details: {
        plannerEnabled: config.enablePlanner,
        sharepointEnabled: config.enableSharePoint,
        outlookEnabled: config.enableOutlook,
      },
    };
  },

  async onConfigChanged(newConfig) {
    if (!pluginCtx) return;
    const config = { ...DEFAULT_CONFIG, ...(newConfig as Partial<M365Config>) };
    initServices(pluginCtx, config);
    pluginCtx.logger.info("M365 config updated — services reinitialized");
  },

  async onValidateConfig(config) {
    return validateConfig(config as Partial<M365Config>);
  },

  async onWebhook(input: PluginWebhookInput) {
    if (!pluginCtx) throw new Error("Plugin not initialized");

    if (input.endpointKey === WEBHOOK_KEYS.graphNotifications) {
      const config = await getConfig(pluginCtx);
      if (!plannerService || !plannerClient) {
        pluginCtx.logger.warn("Received Graph notification but Planner is not enabled");
        return;
      }
      await handleGraphNotification(pluginCtx, input, config, plannerService, plannerClient);
      return;
    }

    if (input.endpointKey === WEBHOOK_KEYS.mailNotifications) {
      const config = await getConfig(pluginCtx);
      if (!outlookClient) {
        pluginCtx.logger.warn("Received mail notification but Outlook client is not available");
        return;
      }
      await handleMailNotification(pluginCtx, input, config, outlookClient);
      return;
    }

    throw new Error(`Unsupported webhook endpoint: ${input.endpointKey}`);
  },

  async onShutdown() {
    pluginCtx?.logger.info("Microsoft 365 plugin shutting down");
  },
});

export default plugin;
runWorker(plugin, import.meta.url);

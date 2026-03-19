export const PLUGIN_ID = "paperclip.microsoft-365";
export const PLUGIN_VERSION = "0.1.0";

export const SLOT_IDS = {
  settingsPage: "m365-settings-page",
  dashboardWidget: "m365-dashboard-widget",
  issueTab: "m365-issue-tab",
  projectTab: "m365-project-tab",
} as const;

export const EXPORT_NAMES = {
  settingsPage: "M365SettingsPage",
  dashboardWidget: "M365DashboardWidget",
  issueTab: "M365IssueTab",
  projectTab: "M365ProjectTab",
} as const;

export const JOB_KEYS = {
  plannerReconcile: "planner-reconcile",
  graphSubscriptionRenew: "graph-subscription-renew",
  outlookDigest: "outlook-digest",
  tokenHealthCheck: "token-health-check",
} as const;

export const WEBHOOK_KEYS = {
  graphNotifications: "graph-notifications",
  mailNotifications: "mail-notifications",
} as const;

export const TOOL_NAMES = {
  sharepointSearch: "sharepoint-search",
  sharepointRead: "sharepoint-read",
  sharepointUpload: "sharepoint-upload",
  plannerStatus: "planner-status",
  outlookSendTaskEmail: "outlook-send-task-email",
} as const;

export const ENTITY_TYPES = {
  plannerTask: "planner-task",
  calendarEvent: "calendar-event",
  sharepointDoc: "sharepoint-doc",
  taskEmail: "task-email",
} as const;

export const STATE_KEYS = {
  lastReconcileAt: "last-reconcile-at",
  syncHealth: "sync-health",
  subscriptionId: "graph-subscription-id",
  subscriptionExpiry: "graph-subscription-expiry",
  mailSubscriptionId: "mail-subscription-id",
  mailSubscriptionExpiry: "mail-subscription-expiry",
} as const;

export type ConflictStrategy = "last_write_wins" | "paperclip_wins" | "planner_wins";

export type PaperclipIssueStatus =
  | "backlog"
  | "todo"
  | "in_progress"
  | "in_review"
  | "done"
  | "blocked"
  | "cancelled";

export interface PlannerStatusMapping {
  percentComplete: number;
  bucketName: string;
}

export const PAPERCLIP_TO_PLANNER: Record<PaperclipIssueStatus, PlannerStatusMapping> = {
  backlog: { percentComplete: 0, bucketName: "Backlog" },
  todo: { percentComplete: 0, bucketName: "To Do" },
  in_progress: { percentComplete: 50, bucketName: "In Progress" },
  in_review: { percentComplete: 50, bucketName: "In Review" },
  done: { percentComplete: 100, bucketName: "Completed" },
  blocked: { percentComplete: 50, bucketName: "Blocked" },
  cancelled: { percentComplete: 100, bucketName: "Cancelled" },
} as const;

export const PLANNER_BUCKET_TO_PAPERCLIP: Record<string, PaperclipIssueStatus> = {
  Backlog: "backlog",
  "To Do": "todo",
  "In Progress": "in_progress",
  "In Review": "in_review",
  Completed: "done",
  Blocked: "blocked",
  Cancelled: "cancelled",
} as const;

export const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
export const OAUTH_TOKEN_URL = "https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";
export const GRAPH_SCOPE = "https://graph.microsoft.com/.default";

export const DEFAULT_MAX_DOC_SIZE_BYTES = 5 * 1024 * 1024; // 5MB
export const CIRCUIT_BREAKER_THRESHOLD = 5;
export const CIRCUIT_BREAKER_COOLDOWN_MS = 5 * 60 * 1000; // 5 minutes

export type M365Config = {
  tenantId: string;
  clientId: string;
  clientSecretRef: string;
  enablePlanner: boolean;
  enableSharePoint: boolean;
  enableOutlook: boolean;
  plannerPlanId: string;
  plannerGroupId: string;
  conflictStrategy: ConflictStrategy;
  sharepointSiteId: string;
  sharepointDriveId: string;
  sharepointUploadFolderId: string;
  maxDocSizeBytes: number;
  outlookCalendarId: string;
  digestRecipients: string[];
  digestSenderUserId: string;
  outlookMailboxUserId: string;
  enableInboundEmail: boolean;
  webhookClientStateRef: string;
};

export const DEFAULT_CONFIG: M365Config = {
  tenantId: "",
  clientId: "",
  clientSecretRef: "",
  enablePlanner: false,
  enableSharePoint: false,
  enableOutlook: false,
  plannerPlanId: "",
  plannerGroupId: "",
  conflictStrategy: "last_write_wins",
  sharepointSiteId: "",
  sharepointDriveId: "",
  sharepointUploadFolderId: "",
  maxDocSizeBytes: DEFAULT_MAX_DOC_SIZE_BYTES,
  outlookCalendarId: "",
  digestRecipients: [],
  digestSenderUserId: "",
  outlookMailboxUserId: "",
  enableInboundEmail: false,
  webhookClientStateRef: "",
};

export type SyncHealthData = {
  configured: boolean;
  enablePlanner: boolean;
  enableSharePoint: boolean;
  enableOutlook: boolean;
  health: { tokenHealthy?: boolean; checkedAt?: string } | null;
  lastReconcile: string | null;
  subscriptionExpiry: string | null;
  trackedTasks: number;
};

export type PluginConfigData = {
  tenantId: string;
  clientId: string;
  hasClientSecret: boolean;
  clientSecretRef: string;
  enablePlanner: boolean;
  enableSharePoint: boolean;
  enableOutlook: boolean;
  plannerPlanId: string;
  plannerGroupId: string;
  conflictStrategy: string;
  sharepointSiteId: string;
  sharepointDriveId: string;
  sharepointUploadFolderId: string;
  maxDocSizeBytes: number;
  outlookCalendarId: string;
  digestRecipients: string[];
  digestSenderUserId: string;
  hasWebhookClientState: boolean;
};

/** Shape of the form state used by the settings page editor. */
export type ConfigFormState = {
  tenantId: string;
  clientId: string;
  clientSecretRef: string;
  enablePlanner: boolean;
  enableSharePoint: boolean;
  enableOutlook: boolean;
  plannerPlanId: string;
  plannerGroupId: string;
  conflictStrategy: string;
  sharepointSiteId: string;
  sharepointDriveId: string;
  sharepointUploadFolderId: string;
  maxDocSizeBytes: number;
  outlookCalendarId: string;
  digestRecipients: string;
  digestSenderUserId: string;
};

/** Response from the save-config action. */
export type SaveConfigResult = {
  ok: boolean;
  errors?: string[];
  warnings?: string[];
};

/** Response from the test-connection action. */
export type TestConnectionResult = {
  ok: boolean;
  error?: string | null;
};

export type IssueM365Data = {
  plannerTask: {
    id: string;
    title: string | null;
    status: string | null;
    data: { plannerTaskId?: string; lastSyncedAt?: string; bucketId?: string };
  } | null;
  calendarEvent: {
    id: string;
    title: string | null;
    data: { eventId?: string; dueDate?: string };
  } | null;
};

export type DriveItem = {
  id: string;
  name: string;
  size: number;
  webUrl: string;
  lastModifiedDateTime: string;
  file?: { mimeType: string };
  folder?: { childCount: number };
};

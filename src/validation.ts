import type { M365Config } from "./constants.js";

export interface ValidationResult {
  ok: boolean;
  warnings: string[];
  errors: string[];
}

/**
 * Validates a (potentially partial) M365Config and returns structured errors/warnings.
 * Extracted from the plugin's onValidateConfig hook so it can be reused by the
 * save-config action handler and tested in isolation.
 */
export function validateConfig(config: Partial<M365Config>): ValidationResult {
  const errors: string[] = [];
  const warnings: string[] = [];

  if (config.enablePlanner) {
    if (!config.plannerPlanId) errors.push("Planner Plan ID is required when Planner is enabled");
    if (!config.plannerGroupId) errors.push("Planner Group ID is required when Planner is enabled");
  }

  if (config.enableSharePoint) {
    if (!config.sharepointSiteId) errors.push("SharePoint Site ID is required when SharePoint is enabled");
    if (!config.sharepointDriveId) errors.push("SharePoint Drive ID is required when SharePoint is enabled");
  }

  if (config.enableOutlook) {
    if (!config.outlookCalendarId) errors.push("Outlook Calendar ID is required when Outlook is enabled");
    if (!config.digestSenderUserId) errors.push("Digest Sender User ID is required when Outlook is enabled");
  }

  if (config.enableInboundEmail) {
    if (!config.outlookMailboxUserId) {
      errors.push("Outlook Mailbox User ID is required when inbound email processing is enabled");
    }
    if (!config.webhookClientStateRef) {
      errors.push("Webhook Client State Secret is required when inbound email processing is enabled");
    }
    if (!config.enableOutlook) {
      warnings.push("Outlook integration should be enabled when inbound email processing is enabled");
    }
  }

  // Agentic Identity
  if (config.agentIdentityMap) {
    for (const [agentId, userId] of Object.entries(config.agentIdentityMap)) {
      if (!agentId || !userId) {
        errors.push("Agent identity map entries must have non-empty agent ID and M365 user ID");
        break;
      }
    }
  }
  if (
    (config.enablePlanner || config.enableSharePoint || config.enableOutlook ||
     config.enableTeams || config.enablePeople || config.enableMeetings) &&
    !config.defaultServiceUserId
  ) {
    warnings.push("No default service user ID configured — agent identity resolution will have no fallback");
  }

  // Teams — read-only (posting requires delegated auth)
  if (config.enableTeams) {
    if (!config.teamsTeamId) warnings.push("Teams Team ID is required when Teams is enabled");
  }

  // People — just needs Azure AD credentials (gated below)

  // Meetings — warn instead of error so config can be saved incrementally
  if (config.enableMeetings) {
    if (!config.meetingOrganizerUserId) {
      warnings.push("Meeting Organizer User ID is required when Meetings is enabled");
    }
    if (config.meetingDefaultDuration !== undefined && config.meetingDefaultDuration <= 0) {
      errors.push("Meeting default duration must be a positive number");
    }
  }

  if (config.enablePlanner || config.enableSharePoint || config.enableOutlook ||
      config.enableTeams || config.enablePeople || config.enableMeetings) {
    if (!config.tenantId) {
      errors.push("Azure AD Tenant ID is required");
    } else if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(config.tenantId)) {
      errors.push("Azure AD Tenant ID must be a valid UUID");
    }
    if (!config.clientId) errors.push("Azure AD Client ID is required");
    if (!config.clientSecretRef) errors.push("Client Secret Reference is required");
  }

  if (config.maxDocSizeBytes !== undefined && config.maxDocSizeBytes <= 0) {
    errors.push("Max document size must be a positive number");
  }

  return { ok: errors.length === 0, warnings, errors };
}

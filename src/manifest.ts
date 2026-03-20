import type { PaperclipPluginManifestV1 } from "@paperclipai/plugin-sdk";
import {
  DEFAULT_CONFIG,
  EXPORT_NAMES,
  JOB_KEYS,
  PLUGIN_ID,
  PLUGIN_VERSION,
  SLOT_IDS,
  TOOL_NAMES,
  WEBHOOK_KEYS,
} from "./constants.js";

const manifest: PaperclipPluginManifestV1 = {
  id: PLUGIN_ID,
  apiVersion: 1,
  version: PLUGIN_VERSION,
  displayName: "Microsoft 365",
  description:
    "Connects Paperclip AI agents and human Microsoft 365 users as teammates. Syncs Planner tasks bidirectionally, enables document collaboration through SharePoint, and provides email-based communication between agents and humans for task updates, assignments, and requests.",
  author: "Paperclip",
  categories: ["connector", "automation"],

  capabilities: [
    // Data read
    "companies.read",
    "projects.read",
    "issues.read",
    "issues.create",
    "issues.update",
    "agents.read",

    // Activity + metrics
    "activity.log.write",
    "metrics.write",

    // Plugin state
    "plugin.state.read",
    "plugin.state.write",

    // Runtime
    "events.subscribe",
    "events.emit",
    "jobs.schedule",
    "webhooks.receive",
    "http.outbound",
    "secrets.read-ref",

    // Agent tools
    "agent.tools.register",

    // UI
    "instance.settings.register",
    "ui.detailTab.register",
    "ui.dashboardWidget.register",
  ],

  entrypoints: {
    worker: "./dist/worker.js",
    ui: "./dist/ui",
  },

  instanceConfigSchema: {
    type: "object",
    properties: {
      tenantId: {
        type: "string",
        title: "Azure AD Tenant ID",
        description: "Your Azure AD tenant ID (directory ID)",
      },
      clientId: {
        type: "string",
        title: "Azure AD Client ID",
        description: "Application (client) ID from your Azure AD app registration",
      },
      clientSecretRef: {
        type: "string",
        format: "secret-ref",
        title: "Client Secret Reference",
        description: "Secret reference for the Azure AD client secret",
      },
      enablePlanner: {
        type: "boolean",
        title: "Enable Planner Sync",
        default: DEFAULT_CONFIG.enablePlanner,
      },
      enableSharePoint: {
        type: "boolean",
        title: "Enable SharePoint Integration",
        default: DEFAULT_CONFIG.enableSharePoint,
      },
      enableOutlook: {
        type: "boolean",
        title: "Enable Outlook Integration",
        default: DEFAULT_CONFIG.enableOutlook,
      },
      plannerPlanId: {
        type: "string",
        title: "Planner Plan ID",
        description: "The ID of the Planner plan to sync with",
      },
      plannerGroupId: {
        type: "string",
        title: "Planner Group ID",
        description: "The M365 Group that owns the Planner plan",
      },
      conflictStrategy: {
        type: "string",
        title: "Conflict Resolution Strategy",
        enum: ["last_write_wins", "paperclip_wins", "planner_wins"],
        default: DEFAULT_CONFIG.conflictStrategy,
      },
      sharepointSiteId: {
        type: "string",
        title: "SharePoint Site ID",
      },
      sharepointDriveId: {
        type: "string",
        title: "SharePoint Drive ID",
      },
      sharepointUploadFolderId: {
        type: "string",
        title: "SharePoint Upload Folder ID",
      },
      maxDocSizeBytes: {
        type: "number",
        title: "Max Document Size (bytes)",
        default: DEFAULT_CONFIG.maxDocSizeBytes,
      },
      outlookCalendarId: {
        type: "string",
        title: "Outlook Calendar ID",
      },
      digestRecipients: {
        type: "array",
        title: "Digest Email Recipients",
        items: { type: "string" },
        default: DEFAULT_CONFIG.digestRecipients,
      },
      digestSenderUserId: {
        type: "string",
        title: "Digest Sender User ID",
        description: "The user ID (or UPN) used to send digest emails",
      },
      webhookClientStateRef: {
        type: "string",
        format: "secret-ref",
        title: "Webhook Client State Secret",
        description: "Secret reference for Graph webhook clientState verification",
      },
      outlookMailboxUserId: {
        type: "string",
        title: "Inbound Email Mailbox User ID",
        description: "The user ID (or UPN) of the mailbox monitored for inbound task email replies",
      },
      enableInboundEmail: {
        type: "boolean",
        title: "Enable Inbound Email Processing",
        default: false,
      },
      // Agentic Identity
      agentIdentityMap: {
        type: "object",
        title: "Agent Identity Map",
        description: "Maps Paperclip agent IDs to M365 user IDs/UPNs (e.g., { 'agent-uuid': 'ceo@contoso.com' })",
        additionalProperties: { type: "string" },
        default: DEFAULT_CONFIG.agentIdentityMap,
      },
      defaultServiceUserId: {
        type: "string",
        title: "Default Service User ID",
        description: "Fallback M365 user ID for unmapped agents and background jobs",
      },
      // Teams
      enableTeams: {
        type: "boolean",
        title: "Enable Teams Integration",
        default: DEFAULT_CONFIG.enableTeams,
      },
      teamsTeamId: {
        type: "string",
        title: "Teams Team ID",
        description: "The ID of the Teams team to integrate with",
      },
      teamsDefaultChannelId: {
        type: "string",
        title: "Teams Default Channel ID",
        description: "Default channel for automated notifications",
      },
      // People & Presence
      enablePeople: {
        type: "boolean",
        title: "Enable People & Presence",
        default: DEFAULT_CONFIG.enablePeople,
      },
      // Meetings
      enableMeetings: {
        type: "boolean",
        title: "Enable Meeting Scheduling",
        default: DEFAULT_CONFIG.enableMeetings,
      },
      meetingOrganizerUserId: {
        type: "string",
        title: "Meeting Organizer User ID",
        description: "Default M365 user ID used as meeting organizer",
      },
      meetingDefaultDuration: {
        type: "number",
        title: "Default Meeting Duration (minutes)",
        default: DEFAULT_CONFIG.meetingDefaultDuration,
      },
    },
  },

  jobs: [
    {
      jobKey: JOB_KEYS.plannerReconcile,
      displayName: "Planner Reconciliation",
      description: "Full bidirectional sync between Paperclip issues and Planner tasks",
      schedule: "*/15 * * * *",
    },
    {
      jobKey: JOB_KEYS.graphSubscriptionRenew,
      displayName: "Graph Subscription Renewal",
      description: "Renews Microsoft Graph webhook subscriptions before they expire",
      schedule: "0 */12 * * *",
    },
    {
      jobKey: JOB_KEYS.outlookDigest,
      displayName: "Outlook Email Digest",
      description: "Sends a daily digest of recent Paperclip activity to configured recipients",
      schedule: "0 9 * * 1-5",
    },
    {
      jobKey: JOB_KEYS.tokenHealthCheck,
      displayName: "Token Health Check",
      description: "Verifies OAuth tokens are functional",
      schedule: "*/30 * * * *",
    },
  ],

  webhooks: [
    {
      endpointKey: WEBHOOK_KEYS.graphNotifications,
      displayName: "Graph Change Notifications",
      description: "Receives Microsoft Graph change notifications for Planner task updates",
    },
    {
      endpointKey: WEBHOOK_KEYS.mailNotifications,
      displayName: "Mail Notifications",
      description: "Receives Microsoft Graph change notifications for new emails in the configured mailbox",
    },
  ],

  tools: [
    {
      name: TOOL_NAMES.sharepointSearch,
      displayName: "SharePoint Search",
      description: "Search documents in the configured SharePoint site. Returns document titles, snippets, and IDs.",
      parametersSchema: {
        type: "object",
        properties: {
          query: { type: "string", description: "Search query string" },
          maxResults: { type: "number", description: "Maximum number of results (default: 10)" },
        },
        required: ["query"],
      },
    },
    {
      name: TOOL_NAMES.sharepointRead,
      displayName: "SharePoint Read Document",
      description: "Read the text content of a SharePoint/OneDrive document by drive and item ID.",
      parametersSchema: {
        type: "object",
        properties: {
          driveId: { type: "string", description: "The drive ID containing the document" },
          itemId: { type: "string", description: "The item ID of the document" },
        },
        required: ["driveId", "itemId"],
      },
    },
    {
      name: TOOL_NAMES.sharepointUpload,
      displayName: "SharePoint Upload",
      description: "Upload a file to the configured SharePoint upload folder.",
      parametersSchema: {
        type: "object",
        properties: {
          fileName: { type: "string", description: "Name for the uploaded file" },
          content: { type: "string", description: "File content to upload" },
          contentType: { type: "string", description: "MIME type (default: text/plain)" },
        },
        required: ["fileName", "content"],
      },
    },
    {
      name: TOOL_NAMES.plannerStatus,
      displayName: "Planner Task Status",
      description: "Check the linked Planner task status for a Paperclip issue.",
      parametersSchema: {
        type: "object",
        properties: {
          issueId: { type: "string", description: "The Paperclip issue ID to check" },
        },
        required: ["issueId"],
      },
    },
    {
      name: TOOL_NAMES.outlookSendTaskEmail,
      displayName: "Send Task Email",
      description:
        "Send an email to a person about a specific Paperclip issue. Use for assignments, status updates, blockers, or requesting input.",
      parametersSchema: {
        type: "object",
        properties: {
          issueId: { type: "string", description: "The Paperclip issue ID" },
          recipientEmail: { type: "string", description: "Email address of the recipient" },
          emailType: {
            type: "string",
            enum: ["assignment", "status_change", "blocked", "request", "custom"],
            description: "Type of email (default: custom)",
          },
          customMessage: { type: "string", description: "Custom message to include" },
        },
        required: ["issueId", "recipientEmail"],
      },
    },
    // ── Teams tools ──────────────────────────────────────────────────────────
    {
      name: TOOL_NAMES.teamsPostMessage,
      displayName: "Teams Post Message",
      description: "Post a message to a Teams channel. The message is sent as the agent's M365 identity.",
      parametersSchema: {
        type: "object",
        properties: {
          channelId: { type: "string", description: "Channel ID (defaults to configured default channel)" },
          content: { type: "string", description: "HTML message content" },
          subject: { type: "string", description: "Optional message subject" },
        },
        required: ["content"],
      },
    },
    {
      name: TOOL_NAMES.teamsReadChannel,
      displayName: "Teams Read Channel",
      description: "Read recent messages from a Teams channel.",
      parametersSchema: {
        type: "object",
        properties: {
          channelId: { type: "string", description: "The channel ID to read from" },
          maxMessages: { type: "number", description: "Maximum number of messages to return (default: 20)" },
        },
        required: ["channelId"],
      },
    },
    {
      name: TOOL_NAMES.teamsReplyThread,
      displayName: "Teams Reply to Thread",
      description: "Reply to a specific message thread in a Teams channel.",
      parametersSchema: {
        type: "object",
        properties: {
          channelId: { type: "string", description: "The channel ID" },
          messageId: { type: "string", description: "The parent message ID to reply to" },
          content: { type: "string", description: "HTML reply content" },
        },
        required: ["channelId", "messageId", "content"],
      },
    },
    {
      name: TOOL_NAMES.teamsListChannels,
      displayName: "Teams List Channels",
      description: "List all channels in the configured Teams team.",
      parametersSchema: {
        type: "object",
        properties: {},
      },
    },
    // ── People & Presence tools ──────────────────────────────────────────────
    {
      name: TOOL_NAMES.peopleLookup,
      displayName: "People Lookup",
      description: "Search for users in the M365 directory by name, email, or department.",
      parametersSchema: {
        type: "object",
        properties: {
          query: { type: "string", description: "Search query (name, email, or department)" },
        },
        required: ["query"],
      },
    },
    {
      name: TOOL_NAMES.peopleGetPresence,
      displayName: "People Get Presence",
      description: "Check a user's availability/presence status. Supports single or batch lookups.",
      parametersSchema: {
        type: "object",
        properties: {
          userId: { type: "string", description: "Single user ID to check" },
          userIds: { type: "array", items: { type: "string" }, description: "Multiple user IDs for batch lookup" },
        },
      },
    },
    {
      name: TOOL_NAMES.peopleGetManager,
      displayName: "People Get Manager",
      description: "Get a user's manager in the org hierarchy.",
      parametersSchema: {
        type: "object",
        properties: {
          userId: { type: "string", description: "The user ID to get the manager for" },
        },
        required: ["userId"],
      },
    },
    {
      name: TOOL_NAMES.peopleListTeamMembers,
      displayName: "People List Team Members",
      description: "List all members of an M365 group/team.",
      parametersSchema: {
        type: "object",
        properties: {
          groupId: { type: "string", description: "The M365 group ID" },
        },
        required: ["groupId"],
      },
    },
    // ── Meeting tools ────────────────────────────────────────────────────────
    {
      name: TOOL_NAMES.meetingSchedule,
      displayName: "Schedule Meeting",
      description: "Schedule a meeting with attendees. Optionally creates a Teams online meeting link.",
      parametersSchema: {
        type: "object",
        properties: {
          subject: { type: "string", description: "Meeting subject" },
          attendeeEmails: { type: "array", items: { type: "string" }, description: "Email addresses of attendees" },
          startDateTime: { type: "string", description: "Meeting start time (ISO 8601)" },
          endDateTime: { type: "string", description: "Meeting end time (ISO 8601, defaults to start + configured duration)" },
          body: { type: "string", description: "Meeting description/agenda" },
          createTeamsLink: { type: "boolean", description: "Create a Teams online meeting link (default: true)" },
        },
        required: ["subject", "attendeeEmails", "startDateTime"],
      },
    },
    {
      name: TOOL_NAMES.meetingFindTime,
      displayName: "Find Meeting Time",
      description: "Find available time slots for a meeting with specified attendees.",
      parametersSchema: {
        type: "object",
        properties: {
          attendeeEmails: { type: "array", items: { type: "string" }, description: "Email addresses of attendees" },
          durationMinutes: { type: "number", description: "Meeting duration in minutes (default: configured default)" },
          startRange: { type: "string", description: "Start of search range (ISO 8601)" },
          endRange: { type: "string", description: "End of search range (ISO 8601)" },
        },
        required: ["attendeeEmails"],
      },
    },
    {
      name: TOOL_NAMES.meetingCancel,
      displayName: "Cancel Meeting",
      description: "Cancel (delete) a previously scheduled meeting.",
      parametersSchema: {
        type: "object",
        properties: {
          eventId: { type: "string", description: "The calendar event ID to cancel" },
        },
        required: ["eventId"],
      },
    },
    {
      name: TOOL_NAMES.meetingList,
      displayName: "List Meetings",
      description: "List upcoming meetings in a date range.",
      parametersSchema: {
        type: "object",
        properties: {
          startDateTime: { type: "string", description: "Range start (ISO 8601, defaults to now)" },
          endDateTime: { type: "string", description: "Range end (ISO 8601, defaults to 7 days from now)" },
        },
      },
    },
  ],

  ui: {
    slots: [
      {
        type: "settingsPage",
        id: SLOT_IDS.settingsPage,
        displayName: "Microsoft 365 Settings",
        exportName: EXPORT_NAMES.settingsPage,
      },
      {
        type: "dashboardWidget",
        id: SLOT_IDS.dashboardWidget,
        displayName: "M365 Sync Health",
        exportName: EXPORT_NAMES.dashboardWidget,
      },
      {
        type: "detailTab",
        id: SLOT_IDS.issueTab,
        displayName: "Microsoft 365",
        exportName: EXPORT_NAMES.issueTab,
        entityTypes: ["issue"],
      },
      {
        type: "detailTab",
        id: SLOT_IDS.projectTab,
        displayName: "SharePoint",
        exportName: EXPORT_NAMES.projectTab,
        entityTypes: ["project"],
      },
    ],
  },
};

export default manifest;

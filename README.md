# Paperclip Microsoft 365 Plugin

A Paperclip plugin that connects AI agents and human Microsoft 365 users as teammates. Instead of treating M365 as a passive data source, this plugin makes Planner tasks, SharePoint documents, Outlook email, Teams channels, people directory, and meeting scheduling into shared workspaces where AI agents and people collaborate side by side -- assigning work, exchanging status updates, sharing files, and keeping each other in the loop.

Built on the [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview) and the Paperclip Plugin SDK.

## Features

### Agentic Identity

Each Paperclip agent can map to a dedicated M365 user account. The `AgentIdentityService` resolves which M365 user an agent acts as when sending emails, posting to Teams, or scheduling meetings. Configure `agentIdentityMap` to map agent IDs to M365 user IDs/UPNs, with `defaultServiceUserId` as the fallback for unmapped agents and background jobs.

### Planner -- Bidirectional Task Sync

- **Two-way sync** between Paperclip issues and Microsoft Planner tasks. When an agent creates or updates an issue, the linked Planner task updates automatically -- and vice versa.
- **Conflict resolution** with three strategies: last-write-wins (default), Paperclip-wins, or Planner-wins. The reconciliation engine runs every 15 minutes and resolves drift between the two systems.
- **Automatic bucket management.** Tasks are placed into Planner buckets that match Paperclip statuses (Backlog, To Do, In Progress, In Review, Blocked, Completed, Cancelled). Buckets are created on demand if they do not exist.
- **Real-time webhook updates.** Graph change notifications push Planner task changes to Paperclip within seconds, so humans moving cards on their Planner board are reflected immediately.

### SharePoint -- Document Collaboration

- **Search** across your SharePoint site from within Paperclip or through an agent tool call.
- **Read** document content directly, with configurable size limits (default 5 MB).
- **Upload** files to a designated SharePoint folder -- useful when agents generate reports, summaries, or deliverables that the human team needs.

### Outlook -- Email and Calendar

- **Task email.** Agents can send structured emails to people about specific issues: assignments, status changes, blockers, or freeform messages. Each email includes a tracking header (`X-Paperclip-Issue-Id`) for threading.
- **Inbound email parsing.** When a human replies to a task email, the plugin parses the reply to extract intent (status change keywords like "done", "blocked", "on it") and routes the action back to Paperclip as an issue update or comment.
- **Deadline calendar events.** Issues with due dates get all-day calendar events on a shared Outlook calendar. Events update or delete automatically when deadlines change.
- **Daily digest.** A weekday morning email summarizes the previous 24 hours of Paperclip issue activity for configured recipients.

### Teams -- Channel Messaging

- **Post messages** to Teams channels as the agent's M365 identity.
- **Read channel messages** to understand ongoing conversations.
- **Reply to threads** for contextual follow-up.
- **List channels** to discover available channels in the configured team.
- **Event automation.** Issue create/update events automatically post notifications to the configured default channel.

### People and Presence -- Directory Lookups

- **Search users** in the M365 directory by name, email, or department.
- **Check presence/availability** for single users or in batch.
- **Get manager** in the org hierarchy.
- **List team members** of an M365 group.

### Meetings -- Scheduling and Calendar

- **Schedule meetings** with attendees, optionally creating a Teams online meeting link.
- **Find available time slots** across multiple attendees using `findMeetingTimes`.
- **Cancel meetings** by event ID.
- **List upcoming meetings** in a date range.
- **Auto-schedule review meetings.** When an issue moves to `in_review`, the plugin automatically schedules a review meeting with relevant attendees.

## Quick Start

### 1. Install dependencies and build

```bash
npm install
npm run build
```

### 2. Register the plugin with your Paperclip instance

Use the Paperclip CLI or API to install the built plugin package. The build output in `dist/` contains the manifest, worker, and bundled UI.

### 3. Create an Azure AD app registration

The plugin authenticates using OAuth 2.0 Client Credentials (no interactive login required). You will need:

- A **Tenant ID** and **Client ID** from an Azure AD app registration
- A **Client Secret** stored as a Paperclip secret reference
- **Application permissions** granted with admin consent (see [Permissions](#permissions) below)

### 4. Configure through the Setup Wizard

Open the plugin's Settings page in Paperclip. The Setup Wizard walks you through connecting Azure AD, choosing which services to enable, and selecting your Planner plan, SharePoint site, Outlook calendar, Teams team, and meeting settings.

For detailed, step-by-step instructions, see the **[Setup Guide](docs/setup-guide.md)**.

## Permissions

Grant only the permissions for services you plan to enable. All permissions require admin consent.

| Service | Graph API Permission | Type |
|---|---|---|
| Planner | `Tasks.ReadWrite.All` | Application |
| Planner | `Group.Read.All` | Application |
| SharePoint | `Sites.Read.All` | Application |
| SharePoint | `Files.ReadWrite.All` | Application |
| Outlook | `Calendars.ReadWrite` | Application |
| Outlook | `Mail.Send` | Application |
| Teams | `Team.ReadBasic.All` | Application |
| Teams | `Channel.ReadBasic.All` | Application |
| Teams | `ChannelMessage.Read.All` | Application |
| Teams | `ChannelMessage.Send` | Application |
| People/Presence | `User.Read.All` | Application |
| People/Presence | `Presence.Read.All` | Application |
| Meetings | `Calendars.ReadWrite` | Application (shared with Outlook) |
| Meetings | `OnlineMeetings.ReadWrite.All` | Application |

## Architecture

The plugin follows the standard Paperclip plugin architecture with three entrypoints:

```
src/
  manifest.ts        Declares capabilities, jobs, webhooks, tools, UI slots
  worker.ts          Main runtime -- event handlers, jobs, data/action handlers, tool registration
  ui/                React components bundled with esbuild (settings, dashboard, issue/project tabs)
```

### Graph API Client (`src/graph/`)

`TokenManager` handles OAuth 2.0 client-credentials token acquisition with in-memory caching and request deduplication. `GraphClient` wraps `fetch` with automatic bearer token injection, 401 token refresh, 429 rate-limit backoff with retry, and a circuit breaker (5 consecutive failures triggers a 5-minute cooldown). Each M365 product gets its own `GraphClient` instance in the worker.

`validate-id.ts` exports `isValidGraphId()` to prevent path traversal -- all tool handlers validate user-supplied IDs before interpolating them into Graph API URLs.

### Agentic Identity (`src/services/identity.ts`)

`AgentIdentityService` resolves which M365 user a Paperclip agent acts as. Uses `agentIdentityMap` (agent ID to M365 user ID/UPN) with `defaultServiceUserId` as fallback. Tool handlers call `resolveActingUserId(agentId?)` to determine the M365 identity for API calls.

### Services (`src/services/`)

One service class per M365 product:

| Service | File | Responsibility |
|---|---|---|
| `PlannerService` | `planner.ts` | Task CRUD, bucket resolution and auto-creation, entity tracking |
| `SharePointService` | `sharepoint.ts` | Document search (via `/search/query`), content read with size limits, file upload |
| `OutlookService` | `outlook.ts` | Calendar event lifecycle, task email (agent-to-human), daily digest builder and sender |
| `EmailParser` | `outlook.ts` | Inbound email parsing -- extracts issue ID from headers/subject/body, detects status-change intent |
| `TeamsService` | `teams.ts` | Channel messaging: post, read, reply, list. Issue event automation to default channel |
| `PeopleService` | `people.ts` | Directory search (with OData injection escaping), presence lookups, manager chain, group members |
| `MeetingService` | `meetings.ts` | Schedule meetings (with Teams link), find available times, cancel, list upcoming |

### Sync Engine (`src/sync/`)

The bidirectional sync engine has three pieces:

- **Status mapping** -- Translates between Paperclip issue statuses and Planner's `percentComplete` + bucket name. Bucket name is the primary discriminator (since `percentComplete: 50` is ambiguous across In Progress, In Review, and Blocked).
- **Conflict resolution** -- Three strategies: `last_write_wins` (timestamp comparison, Paperclip wins ties), `paperclip_wins`, `planner_wins`.
- **Reconciliation** -- A scheduled job that paginates all tracked entities, fetches current Planner state, detects drift via status comparison, resolves conflicts, and updates the losing side.

### Webhooks (`src/webhooks/`)

Two webhook endpoints receive Microsoft Graph change notifications:

- **Graph Notifications** -- Processes Planner task changes pushed by Graph subscriptions. Validates `clientState`, fetches the updated task, maps its status back to Paperclip, and updates the linked issue.
- **Mail Notifications** -- Processes new emails arriving in the monitored mailbox. Parses inbound replies to extract issue updates or comments.

## Agent Tools

The plugin registers 17 tools that Paperclip AI agents can invoke during conversations:

### SharePoint

| Tool | Description |
|---|---|
| `sharepoint-search` | Search documents in the configured SharePoint site by keyword. Returns titles, snippets, and item IDs. |
| `sharepoint-read` | Read the text content of a specific SharePoint document by drive and item ID. |
| `sharepoint-upload` | Upload a file to the configured SharePoint upload folder. |

### Planner

| Tool | Description |
|---|---|
| `planner-status` | Check the sync status of the linked Planner task for a given Paperclip issue. |

### Outlook

| Tool | Description |
|---|---|
| `outlook-send-task-email` | Send an email to a person about a specific issue -- for assignments, status updates, blockers, requests, or custom messages. |

### Teams

| Tool | Description |
|---|---|
| `teams-post-message` | Post a message to a Teams channel. Sent as the agent's M365 identity. |
| `teams-read-channel` | Read recent messages from a Teams channel. |
| `teams-reply-thread` | Reply to a specific message thread in a Teams channel. |
| `teams-list-channels` | List all channels in the configured Teams team. |

### People and Presence

| Tool | Description |
|---|---|
| `people-lookup` | Search for users in the M365 directory by name, email, or department. |
| `people-get-presence` | Check a user's availability/presence status. Supports single or batch lookups. |
| `people-get-manager` | Get a user's manager in the org hierarchy. |
| `people-list-team-members` | List all members of an M365 group/team. |

### Meetings

| Tool | Description |
|---|---|
| `meeting-schedule` | Schedule a meeting with attendees. Optionally creates a Teams online meeting link. |
| `meeting-find-time` | Find available time slots for a meeting with specified attendees. |
| `meeting-cancel` | Cancel (delete) a previously scheduled meeting. |
| `meeting-list` | List upcoming meetings in a date range. |

## Scheduled Jobs

| Job | Schedule | What it does |
|---|---|---|
| Planner Reconciliation | Every 15 minutes | Full bidirectional sync: compares all tracked Paperclip issues against their linked Planner tasks and resolves any drift. |
| Token Health Check | Every 30 minutes | Verifies that OAuth credentials are valid and tokens can be acquired. Updates the dashboard health status. |
| Graph Subscription Renewal | Every 12 hours | Renews Microsoft Graph webhook subscriptions (Planner and mail) before they expire. |
| Outlook Email Digest | Weekdays at 9:00 AM | Sends an HTML summary of the previous 24 hours of issue activity to configured recipients. |

## UI Components

The plugin ships four React components rendered in Paperclip's UI:

| Component | Slot | Purpose |
|---|---|---|
| `M365SettingsPage` | Settings page | Configuration form and Setup Wizard for Azure AD credentials and service options |
| `M365DashboardWidget` | Dashboard widget | Sync health overview: connection status, tracked task count, last reconciliation time |
| `M365IssueTab` | Issue detail tab | Shows the linked Planner task and calendar event for a given issue |
| `M365ProjectTab` | Project detail tab | SharePoint document context for a project |

## Development

```bash
npm run build        # TypeScript compilation + esbuild UI bundle
npm run typecheck    # Type-check only (tsc --noEmit)
npm test             # Run all tests (Vitest)
npx vitest run tests/conflict.spec.ts   # Run a single test file
```

Build output goes to `dist/` with compiled worker JS, the manifest, and `dist/ui/` (the esbuild-bundled React UI).

Tests are pure unit tests against the sync logic, status mapping, conflict resolution, email parsing, and utility functions. They run in a Node environment with Vitest and do not require mocking the plugin SDK or Graph API.

## Configuration Reference

All configuration is managed through the plugin's Settings page or Setup Wizard. Key fields:

### Core (required)

| Field | Description |
|---|---|
| `tenantId` | Azure AD tenant (directory) ID |
| `clientId` | Azure AD application (client) ID |
| `clientSecretRef` | Paperclip secret reference for the client secret |

### Agentic Identity

| Field | Description |
|---|---|
| `agentIdentityMap` | Maps Paperclip agent IDs to M365 user IDs/UPNs (e.g., `{ "agent-uuid": "ceo@contoso.com" }`) |
| `defaultServiceUserId` | Fallback M365 user ID for unmapped agents and background jobs |

### Planner

| Field | Default | Description |
|---|---|---|
| `enablePlanner` | `false` | Toggle Planner bidirectional sync |
| `plannerPlanId` | | The Planner plan to sync with |
| `plannerGroupId` | | The M365 Group that owns the plan |
| `conflictStrategy` | `last_write_wins` | `last_write_wins`, `paperclip_wins`, or `planner_wins` |

### SharePoint

| Field | Default | Description |
|---|---|---|
| `enableSharePoint` | `false` | Toggle SharePoint document tools |
| `sharepointSiteId` | | SharePoint site ID |
| `sharepointDriveId` | | Document library drive ID |
| `sharepointUploadFolderId` | | Target folder for uploads |
| `maxDocSizeBytes` | `5242880` | Max document read size (5 MB) |

### Outlook

| Field | Default | Description |
|---|---|---|
| `enableOutlook` | `false` | Toggle Outlook email and calendar |
| `outlookCalendarId` | | Shared calendar for deadline events |
| `digestRecipients` | `[]` | Email addresses for the daily digest |
| `digestSenderUserId` | | M365 user ID used to send digest emails |
| `outlookMailboxUserId` | | Mailbox monitored for inbound email replies |
| `enableInboundEmail` | `false` | Toggle inbound email reply processing |
| `webhookClientStateRef` | | Secret reference for Graph webhook verification |

### Teams

| Field | Default | Description |
|---|---|---|
| `enableTeams` | `false` | Toggle Teams integration |
| `teamsTeamId` | | The Teams team to integrate with |
| `teamsDefaultChannelId` | | Default channel for automated notifications |

### People and Presence

| Field | Default | Description |
|---|---|---|
| `enablePeople` | `false` | Toggle People and Presence lookups |

### Meetings

| Field | Default | Description |
|---|---|---|
| `enableMeetings` | `false` | Toggle meeting scheduling |
| `meetingOrganizerUserId` | | Default M365 user ID used as meeting organizer |
| `meetingDefaultDuration` | `30` | Default meeting duration in minutes |

See the **[Setup Guide](docs/setup-guide.md)** for step-by-step instructions.

## License

MIT

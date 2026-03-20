# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Workflow

After completing a user's request, always commit the changes with a descriptive commit message. Stage only the files you changed — do not use `git add -A` or `git add .`.

## Build & Development Commands

```bash
npm run build        # tsc + esbuild UI bundle (scripts/build-ui.mjs)
npm run typecheck    # tsc --noEmit (type-check only, no output)
npm test             # vitest run (all tests)
npx vitest run tests/conflict.spec.ts   # run a single test file
```

Build produces `dist/` with: compiled worker JS, manifest, and `dist/ui/` (esbuild-bundled React UI).

## Architecture

This is a **Paperclip plugin** (`@paperclipai/plugin-sdk`) that integrates Microsoft 365 services (Planner, SharePoint, Outlook, Teams, People/Presence, Meetings) with the Paperclip platform. It follows the Paperclip plugin architecture with three entrypoints:

- **Manifest** (`src/manifest.ts`) — Declares capabilities, jobs, webhooks, agent tools (17 total), and UI slots. Exported via `src/index.ts`.
- **Worker** (`src/worker.ts`) — The main runtime. Uses `definePlugin()` / `runWorker()` from the SDK. Registers all event handlers, scheduled jobs, data providers, actions, tool handlers, and lifecycle hooks (health, config change, webhook dispatch, shutdown). Module-level singletons hold service instances initialized during `setup()`. Each M365 product gets its own `GraphClient` instance.
- **UI** (`src/ui/index.tsx`) — React components bundled separately via esbuild. Exports 4 named components matching manifest UI slot declarations: `M365SettingsPage`, `M365DashboardWidget`, `M365IssueTab`, `M365ProjectTab`. Uses SDK hooks (`usePluginData`, `usePluginAction`, `useHostContext`).

### Key layers

**Graph API client** (`src/graph/`) — `TokenManager` handles OAuth2 client-credentials flow with in-memory caching and deduplication. `GraphClient` wraps fetch with bearer token injection, 401 auto-refresh, 429 rate-limit backoff, and a circuit breaker (5 failures = 5-min pause). All Graph API types live in `src/graph/types.ts`. `validate-id.ts` exports `isValidGraphId()` which validates IDs before URL path interpolation to prevent path traversal attacks.

**Agentic Identity** (`src/services/identity.ts`) — `AgentIdentityService` maps each Paperclip agent to a dedicated M365 user account via `agentIdentityMap` (a `Record<string, string>` in config). `resolveActingUserId(agentId?)` returns the mapped user or falls back to `defaultServiceUserId`. All tool handlers that act on behalf of an agent use this to determine the M365 identity.

**Services** (`src/services/`) — One service class per M365 product:
- `PlannerService` — Task CRUD with bucket resolution/auto-creation, entity tracking via `ctx.entities.upsert()`
- `SharePointService` — Search (via `/search/query`), document read with size limit enforcement, file upload
- `OutlookService` — Calendar event lifecycle, email digest with HTML builder
- `TeamsService` — Channel messaging: post, read, reply, list channels. Event automation posts issue create/update notifications to the configured default channel.
- `PeopleService` — Directory search (with OData `$search` injection escaping), single/batch presence lookups, manager chain, group member listing.
- `MeetingService` — Schedule meetings (with optional Teams online link), find available time slots via `findMeetingTimes`, cancel, list upcoming. Auto-schedules a review meeting when an issue moves to `in_review`.

**Sync engine** (`src/sync/`) — Bidirectional Planner sync:
- `status-map.ts` — Bidirectional status mapping. Bucket name is the primary discriminator for Planner-to-Paperclip (percentComplete is ambiguous for 50% statuses). Mappings defined in `constants.ts` (`PAPERCLIP_TO_PLANNER` / `PLANNER_BUCKET_TO_PAPERCLIP`).
- `conflict.ts` — Three strategies: `last_write_wins` (timestamp comparison, Paperclip wins ties), `paperclip_wins`, `planner_wins`.
- `reconcile.ts` — Scheduled job that paginates all tracked entities, fetches current Planner state, detects drift via `isStatusInSync()`, resolves conflicts, and updates the losing side.

**Agent tools** (`src/tools/`) — 17 tool handlers, each a standalone function receiving typed params, a `ToolRunContext`, and the relevant service. Organized by product: `sharepoint-*` (3), `planner-*` (1), `outlook-*` (1), `teams-*` (4), `people-*` (4), `meeting-*` (4). All handlers that accept user-supplied IDs validate them via `isValidGraphId()` before URL interpolation. Registered in the worker.

**Webhooks** (`src/webhooks/`) — Two endpoints:
- `graph-notifications.ts` — Processes Planner task change notifications. Validates `clientState` secret, fetches updated task, maps status back to Paperclip, updates issue.
- `mail-notifications.ts` — Processes inbound emails. Parses replies to extract issue updates or comments.

Note: Graph subscription validation (validationToken echo) must be handled by the host server, not the plugin worker.

### Security

- **ID validation**: `isValidGraphId()` (`src/graph/validate-id.ts`) rejects path traversal characters (`/`, `\`, `..`) and whitespace. All tool handlers validate IDs before Graph API URL interpolation.
- **OData search injection**: `PeopleService.lookupUser()` escapes double-quotes in search queries before interpolating into `$search` parameters.
- **Webhook verification**: Graph notification handlers validate `clientState` secret against the configured secret reference.

### Constants & config

`src/constants.ts` is the single source of truth for: plugin ID, all key registries (jobs, webhooks, tools, entity types, state keys, slot IDs, export names), status mapping tables, Graph API URLs, circuit breaker thresholds, the `M365Config` type, and `DEFAULT_CONFIG`.

`M365Config` includes feature toggles (`enablePlanner`, `enableSharePoint`, `enableOutlook`, `enableTeams`, `enablePeople`, `enableMeetings`), agentic identity settings (`agentIdentityMap`, `defaultServiceUserId`), Teams settings (`teamsTeamId`, `teamsDefaultChannelId`), and meeting settings (`meetingOrganizerUserId`, `meetingDefaultDuration`).

## Testing

Tests are in `tests/*.spec.ts`, run with Vitest in Node environment. Tests are pure unit tests against sync logic and utility functions — they don't mock the plugin SDK or Graph API. Test files import directly from `src/` via `.js` extensions (NodeNext module resolution).

## AI Team Configuration (autogenerated by team-configurator, 2026-03-18)

**Important: YOU MUST USE subagents when available for the task.**

### Detected Stack

- **Language:** TypeScript 5.7 (ES2023 target, NodeNext modules, strict mode)
- **Runtime:** Node.js (ESM)
- **Frontend:** React 19 (JSX, SDK hooks) bundled with esbuild 0.27
- **Backend/Worker:** Paperclip Plugin SDK (`@paperclipai/plugin-sdk` ^2026.318.0)
- **External API:** Microsoft Graph API (OAuth2 client-credentials, Planner, SharePoint, Outlook, Teams, People/Presence, Meetings)
- **Build:** `tsc` + `esbuild` (UI bundle)
- **Test:** Vitest 3 (Node environment, unit tests)
- **Architecture:** Plugin model with three entrypoints (manifest, worker, UI). Bidirectional sync engine, circuit-breaker HTTP client, webhook consumer, agentic identity resolution.

### Agent Assignments

| Task | Agent | Notes |
|------|-------|-------|
| React UI components (settings, dashboard, tabs) | `react-component-architect` | React 19 with SDK hooks; use for all `src/ui/` work |
| Backend worker logic, services, sync engine | `backend-developer` | TypeScript Node.js; covers `src/worker.ts`, `src/services/`, `src/sync/` |
| Graph API contract design, tool handler schemas | `api-architect` | REST/Graph contract work; covers `src/graph/`, `src/tools/` |
| Frontend markup and styling (non-React-specific) | `frontend-developer` | Fallback for general UI concerns outside React patterns |
| Code reviews and pull requests | `code-reviewer` | Run before every merge; routes security/perf issues to specialists |
| Performance profiling and optimization | `performance-optimizer` | Circuit breaker tuning, sync pagination, rate-limit backoff, bundle size |
| Documentation updates (README, API docs, guides) | `documentation-specialist` | Keep CLAUDE.md, README, and inline docs current |
| Codebase exploration and onboarding | `code-archaeologist` | Use when navigating unfamiliar areas or planning refactors |
| Multi-step feature planning and task breakdown | `tech-lead-orchestrator` | Coordinate cross-cutting work across graph, services, sync, and UI |
| Project stack analysis | `project-analyst` | Re-run when dependencies or architecture change significantly |

### Agent Inventory Reference

| Agent | Location | Tags |
|-------|----------|------|
| `code-reviewer` | `~/.claude/agents/core/code-reviewer.md` | core, review, security |
| `performance-optimizer` | `~/.claude/agents/core/performance-optimizer.md` | core, performance, scaling |
| `documentation-specialist` | `~/.claude/agents/core/documentation-specialist.md` | core, docs, onboarding |
| `code-archaeologist` | `~/.claude/agents/core/code-archaeologist.md` | core, exploration, audit |
| `frontend-developer` | `~/.claude/agents/universal/frontend-developer.md` | universal, UI, accessibility |
| `backend-developer` | `~/.claude/agents/universal/backend-developer.md` | universal, server-side, any stack |
| `api-architect` | `~/.claude/agents/universal/api-architect.md` | universal, REST, GraphQL, contracts |
| `react-component-architect` | `~/.claude/agents/specialized/react/react-component-architect.md` | react, components, hooks |
| `tech-lead-orchestrator` | `~/.claude/agents/orchestrators/tech-lead-orchestrator.md` | orchestrator, planning, coordination |
| `project-analyst` | `~/.claude/agents/orchestrators/project-analyst.md` | orchestrator, stack detection |

*Timestamp: 2026-03-18T00:00:00Z*

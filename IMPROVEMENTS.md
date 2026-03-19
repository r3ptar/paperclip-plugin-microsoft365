# Microsoft 365 Plugin — UX Improvements Plan

## Context

The plugin works end-to-end but the setup experience is poor. Users must manually hunt down Azure AD GUIDs and there's no interactive configuration UI. This plan covers making the plugin easy to install and configure.

## Current State

- Settings page is **read-only** — displays config values but can't edit them
- Users must manually find and paste: tenant ID, client ID, plan ID, group ID, site ID, drive ID, calendar ID, folder ID, sender user ID
- No guided setup flow
- No OAuth consent flow — users must manually create an Azure AD app registration

---

## Phase 1: Interactive Settings Page

**Goal:** Users can configure everything from within Paperclip's UI.

### 1.1 — Editable Config Form

Replace the read-only settings page with a form that saves config via `usePluginAction`.

**File:** `src/ui/SettingsPage.tsx` (extract from `src/ui/index.tsx`)

- Text inputs for: Tenant ID, Client ID, Client Secret Reference
- Toggle switches for: Enable Planner, Enable SharePoint, Enable Outlook
- "Save" button that calls a new `save-config` action handler
- Validation feedback inline (red borders, error messages)
- Success toast on save via `usePluginToast()`

**Worker changes (`src/worker.ts`):**
- Register a `save-config` action handler that validates and persists config
- The action should call `onValidateConfig` logic before saving
- Return validation errors to the UI if invalid

### 1.2 — Connection Test UX

The "Test Connection" button already exists but needs better feedback:
- Show a spinner during the test
- On success: green checkmark + "Connected as [app display name]"
- On failure: red error with the specific Azure AD error message
- Disable feature toggles until connection test passes

---

## Phase 2: Setup Wizard

**Goal:** Step-by-step guided setup that walks users through configuration.

### 2.1 — Wizard Component

**File:** `src/ui/SetupWizard.tsx`

Multi-step wizard with these steps:

**Step 1: Azure AD Connection**
- Input fields: Tenant ID, Client ID, Client Secret Reference
- "Test Connection" button — must pass before proceeding
- Help text with a link to Azure AD app registration docs
- Instructions for which API permissions to grant:
  - `Tasks.ReadWrite.All`, `Group.Read.All` (Planner)
  - `Sites.Read.All`, `Files.ReadWrite.All` (SharePoint)
  - `Calendars.ReadWrite`, `Mail.Send` (Outlook)

**Step 2: Choose Services**
- Three toggle cards (Planner, SharePoint, Outlook) with descriptions
- Each card shows the required Azure AD permissions for that service
- User enables only what they need

**Step 3: Planner Configuration** (if enabled)
- Fetch available Groups via `GET /groups?$filter=groupTypes/any(c:c eq 'Unified')` and show as a dropdown
- Once a Group is selected, fetch Plans via `GET /groups/{groupId}/planner/plans` and show as a dropdown
- Auto-populate `plannerGroupId` and `plannerPlanId` from selections
- Dropdown for conflict resolution strategy with explanations:
  - "Last write wins" — most recent change takes priority (recommended)
  - "Paperclip wins" — Paperclip always overwrites Planner
  - "Planner wins" — Planner always overwrites Paperclip

**Step 4: SharePoint Configuration** (if enabled)
- Fetch available Sites via `GET /sites?search=*` and show as a dropdown
- Once a Site is selected, fetch Drives via `GET /sites/{siteId}/drives` and show as a dropdown
- Once a Drive is selected, fetch root folder children via `GET /drives/{driveId}/root/children` for upload folder selection
- Auto-populate `sharepointSiteId`, `sharepointDriveId`, `sharepointUploadFolderId`

**Step 5: Outlook Configuration** (if enabled)
- Input field for Sender User ID (email or UPN) — used for calendar events and sending digests
- Fetch calendars via `GET /users/{userId}/calendars` and show as a dropdown
- Multi-input for digest recipient email addresses (add/remove chips)
- Auto-populate `outlookCalendarId`, `digestSenderUserId`, `digestRecipients`

**Step 6: Review & Save**
- Summary of all configured values
- "Save & Activate" button
- On save, run the full validation, persist config, and show success

### 2.2 — Worker Data Handlers for Wizard Dropdowns

Register new data handlers that the wizard UI calls to fetch Graph API data:

```
ctx.data.register("m365-groups", ...)       → GET /groups (filtered to M365 groups)
ctx.data.register("m365-plans", ...)        → GET /groups/{groupId}/planner/plans
ctx.data.register("m365-sites", ...)        → GET /sites?search=*
ctx.data.register("m365-drives", ...)       → GET /sites/{siteId}/drives
ctx.data.register("m365-folders", ...)      → GET /drives/{driveId}/root/children
ctx.data.register("m365-calendars", ...)    → GET /users/{userId}/calendars
```

Each handler:
- Requires that Azure AD credentials are already configured and connection is valid
- Returns a simple `{ id, name }[]` array for the dropdown
- Handles errors gracefully (returns `{ error: string }` if Graph call fails)

---

## Phase 3: OAuth Consent Flow (Longer Term)

**Goal:** Users click "Connect to Microsoft 365" and go through a browser-based OAuth flow instead of manually creating an Azure AD app registration.

### 3.1 — OAuth Redirect Flow

This requires server-side support in Paperclip's host:

1. Plugin registers an OAuth redirect URL with the host
2. Settings page shows a "Connect to Microsoft 365" button
3. Button opens a popup/redirect to:
   ```
   https://login.microsoftonline.com/common/adminconsent
     ?client_id={clientId}
     &redirect_uri={pluginRedirectUrl}
     &state={csrf_token}
   ```
4. After admin consent, Azure AD redirects back with the tenant ID
5. Plugin stores the tenant ID from the redirect

**Prerequisites:**
- Paperclip must publish a well-known app registration (multi-tenant Azure AD app) that customers consent to
- Or: provide a "Bring Your Own App" option (current manual flow) as fallback

### 3.2 — Auto-Detection

After OAuth consent succeeds:
- Auto-detect user's default calendar via `GET /me/calendar`
- Auto-detect available Planner plans via `GET /me/planner/plans`
- Pre-select the most likely options in the wizard
- Show a "Recommended setup" with one-click accept

---

## File Structure After Changes

```
src/ui/
  index.tsx                 # Re-exports all UI components
  SettingsPage.tsx          # Interactive config form (extracted from index.tsx)
  SetupWizard.tsx           # Multi-step setup wizard
  DashboardWidget.tsx       # Sync health widget (extracted from index.tsx)
  IssueTab.tsx              # Issue detail tab (extracted from index.tsx)
  ProjectTab.tsx            # Project detail tab (extracted from index.tsx)
  components/
    ConnectionStatus.tsx    # Reusable connection test + status display
    ServiceCard.tsx         # Toggle card for enabling a service
    GraphDropdown.tsx       # Dropdown that fetches options from Graph API
    EmailChips.tsx          # Multi-email input with add/remove
    WizardStep.tsx          # Wizard step wrapper with navigation
```

## Implementation Notes

- All Graph API calls for the wizard go through the existing `GraphClient` — no direct fetch from the UI
- The wizard should be shown automatically on first install (when config is empty), and accessible from the settings page afterward via "Reconfigure" button
- Each wizard step should be independently saveable (partial config is OK)
- The wizard state should be held in React state, not persisted until "Save" on the final step
- Use `usePluginToast()` for success/error notifications throughout
- All dropdowns should show a loading spinner while fetching and handle empty states ("No plans found — create one in Microsoft Planner first")

## Acceptance Criteria

1. User installs plugin → sees setup wizard automatically
2. User enters Azure AD credentials → tests connection with one click
3. User picks Planner plan from a dropdown (no GUIDs)
4. User picks SharePoint site/drive from dropdowns (no GUIDs)
5. User picks Outlook calendar from a dropdown (no GUIDs)
6. User saves config → plugin begins syncing immediately
7. User can return to settings page to change config at any time
8. All validation errors are shown inline with clear messages

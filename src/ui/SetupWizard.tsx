import { useState, useCallback, useMemo } from "react";
import { usePluginAction } from "@paperclipai/plugin-sdk/ui";
import { WizardStep } from "./components/WizardStep.js";
import { ConnectionStatus } from "./components/ConnectionStatus.js";
import { ServiceCard } from "./components/ServiceCard.js";
import { GraphDropdown } from "./components/GraphDropdown.js";
import { EmailChips } from "./components/EmailChips.js";
import {
  card,
  label,
  fieldRow,
  fieldLabel,
  textInput,
  selectInput,
  successBanner,
  errorBanner,
} from "./styles.js";
import type { SaveConfigResult } from "./types.js";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface SetupWizardProps {
  companyId: string | null;
  onComplete: () => void;
}

interface WizardState {
  step: number;
  // Step 1: Credentials
  tenantId: string;
  clientId: string;
  clientSecret: string;
  clientSecretRef: string;
  connectionTested: boolean;
  // Step 2: Services
  enablePlanner: boolean;
  enableSharePoint: boolean;
  enableOutlook: boolean;
  // Step 3: Planner
  plannerGroupId: string;
  plannerGroupName: string;
  plannerPlanId: string;
  plannerPlanName: string;
  conflictStrategy: string;
  // Step 4: SharePoint
  sharepointSiteId: string;
  sharepointSiteName: string;
  sharepointDriveId: string;
  sharepointDriveName: string;
  sharepointUploadFolderId: string;
  sharepointUploadFolderName: string;
  // Step 5: Outlook
  digestSenderUserId: string;
  outlookCalendarId: string;
  outlookCalendarName: string;
  digestRecipients: string[];
  // Teams
  enableTeams: boolean;
  teamsTeamId: string;
  teamsTeamName: string;
  teamsDefaultChannelId: string;
  teamsDefaultChannelName: string;
  // People
  enablePeople: boolean;
  // Meetings
  enableMeetings: boolean;
  meetingOrganizerUserId: string;
  meetingDefaultDuration: number;
  // Identity
  defaultServiceUserId: string;
}

// ---------------------------------------------------------------------------
// Initial state
// ---------------------------------------------------------------------------

const initialState: WizardState = {
  step: 1,
  tenantId: "",
  clientId: "",
  clientSecret: "",
  clientSecretRef: "",
  connectionTested: false,
  enablePlanner: false,
  enableSharePoint: false,
  enableOutlook: false,
  plannerGroupId: "",
  plannerGroupName: "",
  plannerPlanId: "",
  plannerPlanName: "",
  conflictStrategy: "last_write_wins",
  sharepointSiteId: "",
  sharepointSiteName: "",
  sharepointDriveId: "",
  sharepointDriveName: "",
  sharepointUploadFolderId: "",
  sharepointUploadFolderName: "",
  digestSenderUserId: "",
  outlookCalendarId: "",
  outlookCalendarName: "",
  digestRecipients: [],
  enableTeams: false,
  teamsTeamId: "",
  teamsTeamName: "",
  teamsDefaultChannelId: "",
  teamsDefaultChannelName: "",
  enablePeople: false,
  enableMeetings: false,
  meetingOrganizerUserId: "",
  meetingDefaultDuration: 30,
  defaultServiceUserId: "",
};

// ---------------------------------------------------------------------------
// All required Graph permissions
// ---------------------------------------------------------------------------

const ALL_PERMISSIONS = [
  "Tasks.ReadWrite.All",
  "Group.Read.All",
  "Sites.Read.All",
  "Files.ReadWrite.All",
  "Calendars.ReadWrite",
  "Mail.Send",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "Teamwork.Migrate.All",
  "User.Read.All",
  "Presence.Read.All",
  "OnlineMeetings.ReadWrite.All",
];

const PLANNER_PERMISSIONS = ["Tasks.ReadWrite.All", "Group.Read.All"];
const SHAREPOINT_PERMISSIONS = ["Sites.Read.All", "Files.ReadWrite.All"];
const OUTLOOK_PERMISSIONS = ["Calendars.ReadWrite", "Mail.Send"];
const TEAMS_PERMISSIONS = ["Team.ReadBasic.All", "Channel.ReadBasic.All", "ChannelMessage.Read.All", "Teamwork.Migrate.All"];
const PEOPLE_PERMISSIONS = ["User.Read.All", "Presence.Read.All"];
const MEETINGS_PERMISSIONS = ["Calendars.ReadWrite", "OnlineMeetings.ReadWrite.All"];

// ---------------------------------------------------------------------------
// Step definitions
// ---------------------------------------------------------------------------

type StepId =
  | "credentials"
  | "services"
  | "planner"
  | "sharepoint"
  | "outlook"
  | "teams"
  | "meetings"
  | "review";

interface StepDef {
  id: StepId;
  title: string;
  description?: string;
}

const ALL_STEPS: StepDef[] = [
  {
    id: "credentials",
    title: "Azure AD Connection",
    description:
      "Create an Azure AD app registration and grant the required API permissions.",
  },
  {
    id: "services",
    title: "Choose Services",
    description:
      "Select which Microsoft 365 services you want to integrate with Paperclip.",
  },
  {
    id: "planner",
    title: "Planner Configuration",
    description:
      "Select the Microsoft 365 group and plan to sync tasks with.",
  },
  {
    id: "sharepoint",
    title: "SharePoint Configuration",
    description:
      "Select the SharePoint site, document library, and upload folder.",
  },
  {
    id: "outlook",
    title: "Outlook Configuration",
    description:
      "Configure calendar sync and email digest settings.",
  },
  {
    id: "teams",
    title: "Teams Configuration",
    description: "Select the team and default channel for notifications.",
  },
  {
    id: "meetings",
    title: "Meetings Configuration",
    description: "Configure meeting organizer and default duration.",
  },
  {
    id: "review",
    title: "Review & Save",
    description:
      "Review your configuration and activate the integration.",
  },
];

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function getActiveSteps(state: WizardState): StepDef[] {
  return ALL_STEPS.filter((s) => {
    if (s.id === "planner") return state.enablePlanner;
    if (s.id === "sharepoint") return state.enableSharePoint;
    if (s.id === "outlook") return state.enableOutlook;
    if (s.id === "teams") return state.enableTeams;
    if (s.id === "meetings") return state.enableMeetings;
    return true;
  });
}

// ---------------------------------------------------------------------------
// Inline styles for the review step
// ---------------------------------------------------------------------------

const reviewSection: React.CSSProperties = {
  marginBottom: "12px",
};

const reviewRow: React.CSSProperties = {
  display: "flex",
  justifyContent: "space-between",
  padding: "6px 0",
  borderBottom: "1px solid var(--border)",
  fontSize: "13px",
};

const reviewLabel: React.CSSProperties = {
  color: "var(--muted-foreground)",
  fontWeight: 500,
};

const reviewValue: React.CSSProperties = {
  color: "var(--foreground)",
  fontWeight: 500,
  textAlign: "right",
  maxWidth: "60%",
  wordBreak: "break-all",
};

const permBadge: React.CSSProperties = {
  display: "inline-block",
  padding: "2px 8px",
  borderRadius: "4px",
  fontSize: "11px",
  fontWeight: 500,
  backgroundColor: "var(--muted)",
  color: "var(--muted-foreground)",
  marginRight: "6px",
  marginBottom: "4px",
};

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export function SetupWizard(props: SetupWizardProps) {
  const { companyId, onComplete } = props;

  const [state, setState] = useState<WizardState>(initialState);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);
  const [saveSuccess, setSaveSuccess] = useState(false);

  const saveConfigAction = usePluginAction("save-config");

  // -- Active steps based on service selection --------------------------------
  const activeSteps = useMemo(() => getActiveSteps(state), [
    state.enablePlanner,
    state.enableSharePoint,
    state.enableOutlook,
    state.enableTeams,
    state.enableMeetings,
  ]);

  const currentStepDef = activeSteps[state.step - 1];
  const totalSteps = activeSteps.length;

  // -- Wizard credentials for data handlers (before config is saved) ----------
  const wizardCredentials = useMemo(() => ({
    tenantId: state.tenantId,
    clientId: state.clientId,
    clientSecret: state.clientSecret,
  }), [state.tenantId, state.clientId, state.clientSecret]);

  // -- Field updater ----------------------------------------------------------
  const update = useCallback(
    <K extends keyof WizardState>(key: K, value: WizardState[K]) => {
      setState((prev) => ({ ...prev, [key]: value }));
    },
    [],
  );

  // -- Navigation -------------------------------------------------------------
  const goNext = useCallback(() => {
    setState((prev) => {
      const steps = getActiveSteps(prev);
      if (prev.step < steps.length) {
        return { ...prev, step: prev.step + 1 };
      }
      return prev;
    });
  }, []);

  const goBack = useCallback(() => {
    setState((prev) => {
      if (prev.step > 1) {
        return { ...prev, step: prev.step - 1 };
      }
      return prev;
    });
  }, []);

  // -- Save handler -----------------------------------------------------------
  const handleSave = useCallback(async () => {
    setSaving(true);
    setSaveError(null);
    setSaveSuccess(false);

    try {
      const payload = {
        companyId,
        tenantId: state.tenantId,
        clientId: state.clientId,
        clientSecretRef: state.clientSecretRef,
        enablePlanner: state.enablePlanner,
        enableSharePoint: state.enableSharePoint,
        enableOutlook: state.enableOutlook,
        plannerGroupId: state.plannerGroupId,
        plannerPlanId: state.plannerPlanId,
        conflictStrategy: state.conflictStrategy,
        sharepointSiteId: state.sharepointSiteId,
        sharepointDriveId: state.sharepointDriveId,
        sharepointUploadFolderId: state.sharepointUploadFolderId,
        outlookCalendarId: state.outlookCalendarId,
        digestSenderUserId: state.digestSenderUserId,
        digestRecipients: state.digestRecipients,
        enableTeams: state.enableTeams,
        teamsTeamId: state.teamsTeamId,
        teamsDefaultChannelId: state.teamsDefaultChannelId,
        enablePeople: state.enablePeople,
        enableMeetings: state.enableMeetings,
        meetingOrganizerUserId: state.meetingOrganizerUserId,
        meetingDefaultDuration: state.meetingDefaultDuration,
        defaultServiceUserId: state.defaultServiceUserId,
      };

      const result = (await saveConfigAction(payload)) as SaveConfigResult;

      if (result.ok) {
        setSaveSuccess(true);
        // Small delay so user sees success, then switch to settings form
        setTimeout(() => onComplete(), 1200);
      } else {
        setSaveError(
          result.errors?.join("; ") ?? "Unknown error saving configuration",
        );
      }
    } catch (err) {
      setSaveError(
        err instanceof Error ? err.message : "Unexpected error saving configuration",
      );
    } finally {
      setSaving(false);
    }
  }, [state, companyId, saveConfigAction, onComplete]);

  // -- canProceed per step ----------------------------------------------------
  const canProceed = useMemo(() => {
    if (!currentStepDef) return false;

    switch (currentStepDef.id) {
      case "credentials":
        return state.connectionTested;
      case "services":
        return state.enablePlanner || state.enableSharePoint || state.enableOutlook ||
               state.enableTeams || state.enablePeople || state.enableMeetings;
      case "planner":
        return state.plannerGroupId.length > 0 && state.plannerPlanId.length > 0;
      case "sharepoint":
        return state.sharepointSiteId.length > 0 && state.sharepointDriveId.length > 0;
      case "outlook":
        return (
          state.digestSenderUserId.trim().length > 0 &&
          state.outlookCalendarId.length > 0
        );
      case "teams":
        return state.teamsTeamId.length > 0 && state.teamsDefaultChannelId.length > 0;
      case "meetings":
        return state.meetingOrganizerUserId.trim().length > 0;
      case "review":
        return true;
      default:
        return false;
    }
  }, [currentStepDef, state]);

  // -- onNext handler: save on last step, otherwise advance -------------------
  const handleNext = useCallback(() => {
    if (currentStepDef?.id === "review") {
      handleSave();
    } else {
      goNext();
    }
  }, [currentStepDef, handleSave, goNext]);

  // -- Guard: if step configuration is out of range ---------------------------
  if (!currentStepDef) {
    return <div style={{ padding: "20px" }}>Initializing wizard...</div>;
  }

  // ---------------------------------------------------------------------------
  // Render step content
  // ---------------------------------------------------------------------------

  const renderStepContent = () => {
    switch (currentStepDef.id) {
      // ── Step 1: Credentials ────────────────────────────────────────────
      case "credentials":
        return (
          <>
            <div style={fieldRow}>
              <span style={fieldLabel}>Tenant ID</span>
              <input
                type="text"
                style={textInput}
                placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
                value={state.tenantId}
                onChange={(e) => {
                  update("tenantId", e.target.value);
                  update("connectionTested", false);
                }}
              />
            </div>

            <div style={fieldRow}>
              <span style={fieldLabel}>Client ID</span>
              <input
                type="text"
                style={textInput}
                placeholder="Application (client) ID"
                value={state.clientId}
                onChange={(e) => {
                  update("clientId", e.target.value);
                  update("connectionTested", false);
                }}
              />
            </div>

            <div style={fieldRow}>
              <span style={fieldLabel}>Client Secret</span>
              <input
                type="password"
                style={textInput}
                placeholder="Paste your Azure AD client secret"
                value={state.clientSecret}
                onChange={(e) => {
                  update("clientSecret", e.target.value);
                  update("clientSecretRef", "");
                  update("connectionTested", false);
                }}
              />
            </div>

            <ConnectionStatusWrapper
              tenantId={state.tenantId}
              clientId={state.clientId}
              clientSecret={state.clientSecret}
              clientSecretRef={state.clientSecretRef}
              companyId={companyId}
              onSuccess={() => update("connectionTested", true)}
              onSecretStored={(ref) => update("clientSecretRef", ref)}
            />

            <div style={{ marginTop: "16px" }}>
              <div style={label}>Required API Permissions</div>
              <div style={{ marginTop: "6px" }}>
                {ALL_PERMISSIONS.map((perm) => (
                  <span key={perm} style={permBadge}>
                    {perm}
                  </span>
                ))}
              </div>
            </div>
          </>
        );

      // ── Step 2: Services ───────────────────────────────────────────────
      case "services":
        return (
          <div style={{ display: "flex", flexDirection: "column", gap: "12px" }}>
            <ServiceCard
              name="Planner"
              description="Sync Paperclip issues with Microsoft Planner tasks. Create, update, and reconcile tasks bidirectionally."
              permissions={PLANNER_PERMISSIONS}
              enabled={state.enablePlanner}
              onToggle={(v) => update("enablePlanner", v)}
            />
            <ServiceCard
              name="SharePoint"
              description="Search and upload documents to SharePoint document libraries. Attach files directly from your conversations."
              permissions={SHAREPOINT_PERMISSIONS}
              enabled={state.enableSharePoint}
              onToggle={(v) => update("enableSharePoint", v)}
            />
            <ServiceCard
              name="Outlook"
              description="Create calendar events for issue deadlines and send email digest summaries to your team."
              permissions={OUTLOOK_PERMISSIONS}
              enabled={state.enableOutlook}
              onToggle={(v) => update("enableOutlook", v)}
            />
            <ServiceCard
              name="Teams"
              description="Post messages to Teams channels, read conversations, and receive automated issue notifications."
              permissions={TEAMS_PERMISSIONS}
              enabled={state.enableTeams}
              onToggle={(v) => update("enableTeams", v)}
            />
            <ServiceCard
              name="People & Presence"
              description="Look up users in the directory, check availability, and explore org charts."
              permissions={PEOPLE_PERMISSIONS}
              enabled={state.enablePeople}
              onToggle={(v) => update("enablePeople", v)}
            />
            <ServiceCard
              name="Meetings"
              description="Schedule meetings with Teams links, find available times, and manage calendars."
              permissions={MEETINGS_PERMISSIONS}
              enabled={state.enableMeetings}
              onToggle={(v) => update("enableMeetings", v)}
            />

            <div style={{ ...card, marginTop: "8px" }}>
              <div style={label}>Default Service User</div>
              <div style={fieldRow}>
                <span style={fieldLabel}>Default Service User ID</span>
                <input
                  type="text"
                  style={textInput}
                  placeholder="service-account@yourtenant.com"
                  value={state.defaultServiceUserId}
                  onChange={(e) => update("defaultServiceUserId", e.target.value)}
                />
                <span style={{ fontSize: "12px", color: "var(--muted-foreground)" }}>
                  Fallback M365 identity for unmapped agents and background jobs. Recommended when any service is enabled.
                </span>
              </div>
            </div>
          </div>
        );

      // ── Step 3: Planner ────────────────────────────────────────────────
      case "planner":
        return (
          <>
            <GraphDropdown
              label="Microsoft 365 Group"
              dataHandler="m365-groups"
              value={state.plannerGroupId}
              onChange={(id, name) => {
                setState((prev) => ({
                  ...prev,
                  plannerGroupId: id,
                  plannerGroupName: name,
                  // Reset plan when group changes
                  plannerPlanId: "",
                  plannerPlanName: "",
                }));
              }}
              companyId={companyId}
              placeholder="Select a group..."
              credentials={wizardCredentials}
            />
            <GraphDropdown
              label="Planner Plan"
              dataHandler="m365-plans"
              params={{ groupId: state.plannerGroupId }}
              value={state.plannerPlanId}
              onChange={(id, name) => {
                update("plannerPlanId", id);
                update("plannerPlanName", name);
              }}
              disabled={!state.plannerGroupId}
              companyId={companyId}
              placeholder="Select a plan..."
              credentials={wizardCredentials}
            />
            <div style={fieldRow}>
              <span style={fieldLabel}>Conflict Strategy</span>
              <select
                style={selectInput}
                value={state.conflictStrategy}
                onChange={(e) => update("conflictStrategy", e.target.value)}
              >
                <option value="last_write_wins">Last Write Wins</option>
                <option value="paperclip_wins">Paperclip Wins</option>
                <option value="planner_wins">Planner Wins</option>
              </select>
              <span style={{ fontSize: "12px", color: "var(--muted-foreground)" }}>
                Determines which side wins when both Paperclip and Planner have
                changed the same task.
              </span>
            </div>
          </>
        );

      // ── Step 4: SharePoint ─────────────────────────────────────────────
      case "sharepoint":
        return (
          <>
            <GraphDropdown
              label="SharePoint Site"
              dataHandler="m365-sites"
              value={state.sharepointSiteId}
              onChange={(id, name) => {
                setState((prev) => ({
                  ...prev,
                  sharepointSiteId: id,
                  sharepointSiteName: name,
                  // Reset children when parent changes
                  sharepointDriveId: "",
                  sharepointDriveName: "",
                  sharepointUploadFolderId: "",
                  sharepointUploadFolderName: "",
                }));
              }}
              companyId={companyId}
              placeholder="Select a site..."
              credentials={wizardCredentials}
            />
            <GraphDropdown
              label="Document Library (Drive)"
              dataHandler="m365-drives"
              params={{ siteId: state.sharepointSiteId }}
              value={state.sharepointDriveId}
              onChange={(id, name) => {
                setState((prev) => ({
                  ...prev,
                  sharepointDriveId: id,
                  sharepointDriveName: name,
                  // Reset folder when drive changes
                  sharepointUploadFolderId: "",
                  sharepointUploadFolderName: "",
                }));
              }}
              disabled={!state.sharepointSiteId}
              companyId={companyId}
              placeholder="Select a drive..."
              credentials={wizardCredentials}
            />
            <GraphDropdown
              label="Upload Folder (optional)"
              dataHandler="m365-folders"
              params={{ driveId: state.sharepointDriveId }}
              value={state.sharepointUploadFolderId}
              onChange={(id, name) => {
                update("sharepointUploadFolderId", id);
                update("sharepointUploadFolderName", name);
              }}
              disabled={!state.sharepointDriveId}
              companyId={companyId}
              placeholder="Select a folder (optional)..."
              credentials={wizardCredentials}
            />
          </>
        );

      // ── Step 5: Outlook ────────────────────────────────────────────────
      case "outlook":
        return (
          <>
            <div style={fieldRow}>
              <span style={fieldLabel}>Sender User ID</span>
              <input
                type="text"
                style={textInput}
                placeholder="user@yourtenant.com (email or UPN)"
                value={state.digestSenderUserId}
                onChange={(e) => {
                  update("digestSenderUserId", e.target.value);
                  // Reset calendar when sender changes
                  update("outlookCalendarId", "");
                  update("outlookCalendarName", "");
                }}
              />
              <span style={{ fontSize: "12px", color: "var(--muted-foreground)" }}>
                The user whose calendar will be used for events and who will
                send digest emails.
              </span>
            </div>
            <GraphDropdown
              label="Calendar"
              dataHandler="m365-calendars"
              params={{ userId: state.digestSenderUserId }}
              value={state.outlookCalendarId}
              onChange={(id, name) => {
                update("outlookCalendarId", id);
                update("outlookCalendarName", name);
              }}
              disabled={!state.digestSenderUserId.trim()}
              companyId={companyId}
              placeholder="Select a calendar..."
              credentials={wizardCredentials}
            />
            <div style={{ ...fieldRow, marginTop: "4px" }}>
              <span style={fieldLabel}>Digest Recipients</span>
              <EmailChips
                emails={state.digestRecipients}
                onChange={(emails) => update("digestRecipients", emails)}
              />
              <span style={{ fontSize: "12px", color: "var(--muted-foreground)" }}>
                Email addresses that will receive periodic digest summaries.
              </span>
            </div>
          </>
        );

      // ── Teams ────────────────────────────────────────────────────────
      case "teams":
        return (
          <>
            <GraphDropdown
              label="Team"
              dataHandler="m365-teams"
              value={state.teamsTeamId}
              onChange={(id, name) => {
                setState((prev) => ({
                  ...prev,
                  teamsTeamId: id,
                  teamsTeamName: name,
                  teamsDefaultChannelId: "",
                  teamsDefaultChannelName: "",
                }));
              }}
              companyId={companyId}
              placeholder="Select a team..."
              credentials={wizardCredentials}
            />
            <GraphDropdown
              label="Default Channel"
              dataHandler="m365-teams-channels"
              params={{ teamId: state.teamsTeamId }}
              value={state.teamsDefaultChannelId}
              onChange={(id, name) => {
                setState((prev) => ({
                  ...prev,
                  teamsDefaultChannelId: id,
                  teamsDefaultChannelName: name,
                }));
              }}
              disabled={!state.teamsTeamId}
              companyId={companyId}
              placeholder="Select a channel..."
              credentials={wizardCredentials}
            />
          </>
        );

      // ── Meetings ──────────────────────────────────────────────────────
      case "meetings":
        return (
          <>
            <div style={fieldRow}>
              <span style={fieldLabel}>Meeting Organizer</span>
              <input
                type="text"
                style={textInput}
                placeholder="user@yourtenant.com"
                value={state.meetingOrganizerUserId}
                onChange={(e) => update("meetingOrganizerUserId", e.target.value)}
              />
              <span style={{ fontSize: "12px", color: "var(--muted-foreground)" }}>
                The M365 user whose calendar is used to create meetings.
              </span>
            </div>
            <div style={fieldRow}>
              <span style={fieldLabel}>Default Duration (minutes)</span>
              <input
                type="number"
                style={{ ...textInput, width: "180px" }}
                min={5}
                max={480}
                value={state.meetingDefaultDuration}
                onChange={(e) => update("meetingDefaultDuration", Math.max(5, Math.min(480, parseInt(e.target.value, 10) || 30)))}
              />
            </div>
          </>
        );

      // ── Step 6: Review & Save ──────────────────────────────────────────
      case "review":
        return (
          <>
            {saveSuccess && (
              <div style={successBanner}>
                Configuration saved and activated successfully.
              </div>
            )}
            {saveError && <div style={errorBanner}>{saveError}</div>}

            {/* Credentials */}
            <div style={reviewSection}>
              <div style={label}>Azure AD Connection</div>
              <div style={reviewRow}>
                <span style={reviewLabel}>Tenant ID</span>
                <span style={reviewValue}>{state.tenantId}</span>
              </div>
              <div style={reviewRow}>
                <span style={reviewLabel}>Client ID</span>
                <span style={reviewValue}>{state.clientId}</span>
              </div>
              <div style={reviewRow}>
                <span style={reviewLabel}>Client Secret</span>
                <span style={reviewValue}>
                  {state.clientSecretRef ? "Stored securely" : "Not configured"}
                </span>
              </div>
            </div>

            {/* Services */}
            <div style={reviewSection}>
              <div style={label}>Enabled Services</div>
              <div style={{ ...reviewRow, gap: "8px" }}>
                {state.enablePlanner && (
                  <span style={permBadge}>Planner</span>
                )}
                {state.enableSharePoint && (
                  <span style={permBadge}>SharePoint</span>
                )}
                {state.enableOutlook && (
                  <span style={permBadge}>Outlook</span>
                )}
                {state.enableTeams && <span style={permBadge}>Teams</span>}
                {state.enablePeople && <span style={permBadge}>People</span>}
                {state.enableMeetings && <span style={permBadge}>Meetings</span>}
              </div>
            </div>

            {/* Planner details */}
            {state.enablePlanner && (
              <div style={reviewSection}>
                <div style={label}>Planner</div>
                <div style={reviewRow}>
                  <span style={reviewLabel}>Group</span>
                  <span style={reviewValue}>
                    {state.plannerGroupName || state.plannerGroupId}
                  </span>
                </div>
                <div style={reviewRow}>
                  <span style={reviewLabel}>Plan</span>
                  <span style={reviewValue}>
                    {state.plannerPlanName || state.plannerPlanId}
                  </span>
                </div>
                <div style={reviewRow}>
                  <span style={reviewLabel}>Conflict Strategy</span>
                  <span style={reviewValue}>{state.conflictStrategy}</span>
                </div>
              </div>
            )}

            {/* SharePoint details */}
            {state.enableSharePoint && (
              <div style={reviewSection}>
                <div style={label}>SharePoint</div>
                <div style={reviewRow}>
                  <span style={reviewLabel}>Site</span>
                  <span style={reviewValue}>
                    {state.sharepointSiteName || state.sharepointSiteId}
                  </span>
                </div>
                <div style={reviewRow}>
                  <span style={reviewLabel}>Drive</span>
                  <span style={reviewValue}>
                    {state.sharepointDriveName || state.sharepointDriveId}
                  </span>
                </div>
                {state.sharepointUploadFolderId && (
                  <div style={reviewRow}>
                    <span style={reviewLabel}>Upload Folder</span>
                    <span style={reviewValue}>
                      {state.sharepointUploadFolderName ||
                        state.sharepointUploadFolderId}
                    </span>
                  </div>
                )}
              </div>
            )}

            {/* Outlook details */}
            {state.enableOutlook && (
              <div style={reviewSection}>
                <div style={label}>Outlook</div>
                <div style={reviewRow}>
                  <span style={reviewLabel}>Sender</span>
                  <span style={reviewValue}>{state.digestSenderUserId}</span>
                </div>
                <div style={reviewRow}>
                  <span style={reviewLabel}>Calendar</span>
                  <span style={reviewValue}>
                    {state.outlookCalendarName || state.outlookCalendarId}
                  </span>
                </div>
                {state.digestRecipients.length > 0 && (
                  <div style={reviewRow}>
                    <span style={reviewLabel}>Digest Recipients</span>
                    <span style={reviewValue}>
                      {state.digestRecipients.join(", ")}
                    </span>
                  </div>
                )}
              </div>
            )}

            {state.enableTeams && (
              <div style={reviewSection}>
                <div style={label}>Teams</div>
                <div style={reviewRow}>
                  <span style={reviewLabel}>Team</span>
                  <span style={reviewValue}>{state.teamsTeamName || state.teamsTeamId}</span>
                </div>
                <div style={reviewRow}>
                  <span style={reviewLabel}>Default Channel</span>
                  <span style={reviewValue}>{state.teamsDefaultChannelName || state.teamsDefaultChannelId}</span>
                </div>
              </div>
            )}

            {state.enablePeople && (
              <div style={reviewSection}>
                <div style={label}>People &amp; Presence</div>
                <div style={reviewRow}>
                  <span style={reviewLabel}>Status</span>
                  <span style={reviewValue}>Enabled</span>
                </div>
              </div>
            )}

            {state.enableMeetings && (
              <div style={reviewSection}>
                <div style={label}>Meetings</div>
                <div style={reviewRow}>
                  <span style={reviewLabel}>Organizer</span>
                  <span style={reviewValue}>{state.meetingOrganizerUserId}</span>
                </div>
                <div style={reviewRow}>
                  <span style={reviewLabel}>Default Duration</span>
                  <span style={reviewValue}>{state.meetingDefaultDuration} minutes</span>
                </div>
              </div>
            )}

            {state.defaultServiceUserId && (
              <div style={reviewSection}>
                <div style={label}>Agentic Identity</div>
                <div style={reviewRow}>
                  <span style={reviewLabel}>Default Service User</span>
                  <span style={reviewValue}>{state.defaultServiceUserId}</span>
                </div>
              </div>
            )}

            {saving && (
              <div style={{ fontSize: "14px", color: "var(--muted-foreground)", marginTop: "8px" }}>
                Saving configuration...
              </div>
            )}
          </>
        );

      default:
        return null;
    }
  };

  return (
    <div style={{ padding: "20px", maxWidth: "720px" }}>
      <h2 style={{ margin: "0 0 20px" }}>Microsoft 365 Setup Wizard</h2>

      <WizardStep
        title={currentStepDef.title}
        description={currentStepDef.description}
        stepNumber={state.step}
        totalSteps={totalSteps}
        canProceed={currentStepDef.id === "review" ? !saving : canProceed}
        onNext={handleNext}
        onBack={state.step > 1 ? goBack : undefined}
      >
        {renderStepContent()}
      </WizardStep>
    </div>
  );
}

// ---------------------------------------------------------------------------
// ConnectionStatus wrapper that also signals success to the wizard
// ---------------------------------------------------------------------------

interface ConnectionStatusWrapperProps {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  clientSecretRef: string;
  companyId: string | null;
  onSuccess: () => void;
  onSecretStored: (ref: string) => void;
}

/**
 * Stores a raw secret value via the Paperclip secrets API and returns the
 * generated UUID reference.  If a secret with the same name already exists
 * (409 Conflict), a timestamp-suffixed name is used as a fallback.
 */
async function storeSecret(
  companyId: string,
  name: string,
  value: string,
): Promise<string> {
  const create = async (secretName: string): Promise<string> => {
    const res = await fetch(`/api/companies/${companyId}/secrets`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name: secretName, value }),
    });

    if (res.status === 409) {
      // Secret name already exists — find it and rotate its value
      const listRes = await fetch(`/api/companies/${companyId}/secrets`);
      if (listRes.ok) {
        const secrets = (await listRes.json()) as Array<{ id: string; name: string }>;
        const existing = secrets.find((s) => s.name === secretName);
        if (existing) {
          await fetch(`/api/companies/${companyId}/secrets/${existing.id}/rotate`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ value }),
          });
          return existing.id;
        }
      }
      // Fallback: return conflict error
      throw new Error("Secret already exists and could not be updated");
    }

    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Failed to store secret: ${text}`);
    }
    const data = (await res.json()) as { id: string };
    return data.id;
  };

  return create(name);
}

function ConnectionStatusWrapper(props: ConnectionStatusWrapperProps) {
  const {
    tenantId,
    clientId,
    clientSecret,
    clientSecretRef,
    companyId,
    onSuccess,
    onSecretStored,
  } = props;

  const testConnectionAction = usePluginAction("test-connection");
  const [testing, setTesting] = useState(false);
  const [result, setResult] = useState<{
    ok: boolean;
    error?: string | null;
  } | null>(null);
  const [secretStatus, setSecretStatus] = useState<string | null>(null);

  const canTest =
    tenantId.trim().length > 0 &&
    clientId.trim().length > 0 &&
    (clientSecret.trim().length > 0 || clientSecretRef.trim().length > 0);

  const handleTest = useCallback(async () => {
    setTesting(true);
    setResult(null);
    setSecretStatus(null);

    try {
      let resolvedRef = clientSecretRef;

      // If user entered a raw secret and we have not yet stored it, store it now
      if (clientSecret.trim().length > 0 && !clientSecretRef && companyId) {
        setSecretStatus("Storing secret...");
        resolvedRef = await storeSecret(companyId, "m365-client-secret", clientSecret);
        onSecretStored(resolvedRef);
        setSecretStatus("Secret stored securely");
      }

      const res = (await testConnectionAction({
        companyId,
        tenantId,
        clientId,
        clientSecretRef: resolvedRef,
        clientSecret: clientSecret.trim() || undefined,
      })) as {
        ok: boolean;
        error?: string | null;
      };
      setResult(res);
      if (res.ok) {
        onSuccess();
      }
    } catch (err) {
      setSecretStatus(null);
      setResult({
        ok: false,
        error: err instanceof Error ? err.message : "Unknown error",
      });
    } finally {
      setTesting(false);
    }
  }, [
    testConnectionAction,
    companyId,
    tenantId,
    clientId,
    clientSecret,
    clientSecretRef,
    onSuccess,
    onSecretStored,
  ]);

  return (
    <div
      style={{
        marginTop: "4px",
        display: "flex",
        flexDirection: "column",
        gap: "8px",
      }}
    >
      <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
        <button
          disabled={testing || !canTest}
          onClick={handleTest}
          style={{
            padding: "6px 16px",
            borderRadius: "6px",
            border: "1px solid var(--border)",
            backgroundColor: "var(--secondary)",
            color: "var(--secondary-foreground)",
            fontSize: "14px",
            cursor: testing || !canTest ? "not-allowed" : "pointer",
            opacity: testing || !canTest ? 0.6 : 1,
          }}
        >
          {testing ? "Testing..." : "Test Connection"}
        </button>
        {result && (
          <span
            style={{
              color: result.ok ? "#16a34a" : "var(--destructive)",
              fontSize: "14px",
            }}
          >
            {result.ok
              ? "Connection successful"
              : result.error ?? "Connection failed"}
          </span>
        )}
      </div>
      {secretStatus && (
        <span style={{ fontSize: "12px", color: "var(--muted-foreground)" }}>
          {secretStatus}
        </span>
      )}
    </div>
  );
}

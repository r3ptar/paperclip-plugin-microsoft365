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
}

// ---------------------------------------------------------------------------
// Initial state
// ---------------------------------------------------------------------------

const initialState: WizardState = {
  step: 1,
  tenantId: "",
  clientId: "",
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
];

const PLANNER_PERMISSIONS = ["Tasks.ReadWrite.All", "Group.Read.All"];
const SHAREPOINT_PERMISSIONS = ["Sites.Read.All", "Files.ReadWrite.All"];
const OUTLOOK_PERMISSIONS = ["Calendars.ReadWrite", "Mail.Send"];

// ---------------------------------------------------------------------------
// Step definitions
// ---------------------------------------------------------------------------

type StepId =
  | "credentials"
  | "services"
  | "planner"
  | "sharepoint"
  | "outlook"
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
  borderBottom: "1px solid #f1f5f9",
  fontSize: "13px",
};

const reviewLabel: React.CSSProperties = {
  color: "#64748b",
  fontWeight: 500,
};

const reviewValue: React.CSSProperties = {
  color: "#0f172a",
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
  backgroundColor: "#f1f5f9",
  color: "#475569",
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
  ]);

  const currentStepDef = activeSteps[state.step - 1];
  const totalSteps = activeSteps.length;

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
        return state.enablePlanner || state.enableSharePoint || state.enableOutlook;
      case "planner":
        return state.plannerGroupId.length > 0 && state.plannerPlanId.length > 0;
      case "sharepoint":
        return state.sharepointSiteId.length > 0 && state.sharepointDriveId.length > 0;
      case "outlook":
        return (
          state.digestSenderUserId.trim().length > 0 &&
          state.outlookCalendarId.length > 0
        );
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
              <span style={fieldLabel}>Client Secret Reference</span>
              <input
                type="text"
                style={textInput}
                placeholder="secret-ref://..."
                value={state.clientSecretRef}
                onChange={(e) => {
                  update("clientSecretRef", e.target.value);
                  update("connectionTested", false);
                }}
              />
            </div>

            <ConnectionStatusWrapper
              tenantId={state.tenantId}
              clientId={state.clientId}
              clientSecretRef={state.clientSecretRef}
              companyId={companyId}
              onSuccess={() => update("connectionTested", true)}
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
              <span style={{ fontSize: "12px", color: "#94a3b8" }}>
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
              <span style={{ fontSize: "12px", color: "#94a3b8" }}>
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
            />
            <div style={{ ...fieldRow, marginTop: "4px" }}>
              <span style={fieldLabel}>Digest Recipients</span>
              <EmailChips
                emails={state.digestRecipients}
                onChange={(emails) => update("digestRecipients", emails)}
              />
              <span style={{ fontSize: "12px", color: "#94a3b8" }}>
                Email addresses that will receive periodic digest summaries.
              </span>
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
                <span style={reviewLabel}>Client Secret Ref</span>
                <span style={reviewValue}>{state.clientSecretRef}</span>
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

            {saving && (
              <div style={{ fontSize: "14px", color: "#64748b", marginTop: "8px" }}>
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
  clientSecretRef: string;
  companyId: string | null;
  onSuccess: () => void;
}

function ConnectionStatusWrapper(props: ConnectionStatusWrapperProps) {
  const { tenantId, clientId, clientSecretRef, companyId, onSuccess } = props;

  const testConnectionAction = usePluginAction("test-connection");
  const [testing, setTesting] = useState(false);
  const [result, setResult] = useState<{
    ok: boolean;
    error?: string | null;
  } | null>(null);

  const canTest =
    tenantId.trim().length > 0 &&
    clientId.trim().length > 0 &&
    clientSecretRef.trim().length > 0;

  const handleTest = useCallback(async () => {
    setTesting(true);
    setResult(null);
    try {
      const res = (await testConnectionAction({ companyId, tenantId, clientId, clientSecretRef })) as {
        ok: boolean;
        error?: string | null;
      };
      setResult(res);
      if (res.ok) {
        onSuccess();
      }
    } catch (err) {
      setResult({
        ok: false,
        error: err instanceof Error ? err.message : "Unknown error",
      });
    } finally {
      setTesting(false);
    }
  }, [testConnectionAction, companyId, tenantId, clientId, clientSecretRef, onSuccess]);

  return (
    <div
      style={{
        marginTop: "4px",
        display: "flex",
        alignItems: "center",
        gap: "12px",
      }}
    >
      <button
        disabled={testing || !canTest}
        onClick={handleTest}
        style={{
          padding: "6px 16px",
          borderRadius: "6px",
          border: "1px solid #e2e8f0",
          backgroundColor: "#f8fafc",
          color: "#334155",
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
            color: result.ok ? "#16a34a" : "#dc2626",
            fontSize: "14px",
          }}
        >
          {result.ok
            ? "Connection successful"
            : result.error ?? "Connection failed"}
        </span>
      )}
    </div>
  );
}

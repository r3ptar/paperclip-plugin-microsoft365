import { useState, useCallback, useEffect } from "react";
import {
  usePluginAction,
  usePluginData,
  type PluginSettingsPageProps,
} from "@paperclipai/plugin-sdk/ui";
import {
  card,
  label,
  fieldRow,
  fieldLabel,
  textInput,
  selectInput,
  numberInput,
  toggleRow,
  toggleLabel,
  successBanner,
  warningBanner,
  errorBanner,
  primaryButton,
  primaryButtonDisabled,
  secondaryButton,
  secondaryButtonDisabled,
} from "./styles.js";
import type {
  PluginConfigData,
  ConfigFormState,
  SaveConfigResult,
  TestConnectionResult,
} from "./types.js";
import { SetupWizard } from "./SetupWizard.js";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Build initial form state from the loaded plugin config. */
function configToFormState(cfg: PluginConfigData): ConfigFormState {
  return {
    tenantId: cfg.tenantId ?? "",
    clientId: cfg.clientId ?? "",
    clientSecret: "",
    clientSecretRef: cfg.clientSecretRef ?? "",
    enablePlanner: cfg.enablePlanner ?? false,
    enableSharePoint: cfg.enableSharePoint ?? false,
    enableOutlook: cfg.enableOutlook ?? false,
    plannerPlanId: cfg.plannerPlanId ?? "",
    plannerGroupId: cfg.plannerGroupId ?? "",
    conflictStrategy: cfg.conflictStrategy ?? "last_write_wins",
    sharepointSiteId: cfg.sharepointSiteId ?? "",
    sharepointDriveId: cfg.sharepointDriveId ?? "",
    sharepointUploadFolderId: cfg.sharepointUploadFolderId ?? "",
    maxDocSizeBytes: cfg.maxDocSizeBytes ?? 5242880,
    outlookCalendarId: cfg.outlookCalendarId ?? "",
    digestRecipients: (cfg.digestRecipients ?? []).join(", "),
    digestSenderUserId: cfg.digestSenderUserId ?? "",
  };
}

/** Check whether the form has been modified relative to the last-saved snapshot. */
function isDirty(current: ConfigFormState, saved: ConfigFormState): boolean {
  return (Object.keys(current) as Array<keyof ConfigFormState>).some(
    (key) => String(current[key]) !== String(saved[key]),
  );
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
      const uniqueName = `${name}-${Date.now()}`;
      const retry = await fetch(`/api/companies/${companyId}/secrets`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: uniqueName, value }),
      });
      if (!retry.ok) {
        const text = await retry.text();
        throw new Error(`Failed to store secret: ${text}`);
      }
      const data = (await retry.json()) as { id: string };
      return data.id;
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

/** Check whether the config is "empty" (first-time setup). */
function isConfigEmpty(cfg: PluginConfigData | null | undefined): boolean {
  if (!cfg) return true;
  return !cfg.tenantId && !cfg.clientId;
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export function M365SettingsPage(props: PluginSettingsPageProps) {
  const { context } = props;

  // -- Remote data -----------------------------------------------------------
  const { data, loading, error, refresh } = usePluginData<PluginConfigData>(
    "plugin-config",
    { companyId: context.companyId },
  );
  const saveConfigAction = usePluginAction("save-config");
  const testConnectionAction = usePluginAction("test-connection");

  // -- Local state -----------------------------------------------------------
  const [form, setForm] = useState<ConfigFormState | null>(null);
  const [savedSnapshot, setSavedSnapshot] = useState<ConfigFormState | null>(null);

  // Save state
  const [saving, setSaving] = useState(false);
  const [saveErrors, setSaveErrors] = useState<string[]>([]);
  const [saveWarnings, setSaveWarnings] = useState<string[]>([]);
  const [saveSuccess, setSaveSuccess] = useState(false);

  // Connection test state
  const [testing, setTesting] = useState(false);
  const [testResult, setTestResult] = useState<TestConnectionResult | null>(null);

  // Wizard vs form view
  const [showWizard, setShowWizard] = useState(false);
  const [showForm, setShowForm] = useState(false);

  // -- Initialise form from loaded config ------------------------------------
  useEffect(() => {
    if (data && !form) {
      const initial = configToFormState(data);
      setForm(initial);
      setSavedSnapshot(initial);
    }
  }, [data, form]);

  // -- Derived ---------------------------------------------------------------
  const dirty = form && savedSnapshot ? isDirty(form, savedSnapshot) : false;

  const canTestConnection =
    form != null &&
    form.tenantId.trim().length > 0 &&
    form.clientId.trim().length > 0 &&
    (form.clientSecret.trim().length > 0 || form.clientSecretRef.trim().length > 0);

  // Determine if we should show the wizard: config is empty and user has not
  // dismissed it, OR user explicitly clicked "Reconfigure".
  const configEmpty = isConfigEmpty(data);
  const shouldShowWizard = showWizard || (configEmpty && !showForm);

  // -- Callbacks -------------------------------------------------------------

  const updateField = useCallback(
    <K extends keyof ConfigFormState>(key: K, value: ConfigFormState[K]) => {
      setForm((prev) => (prev ? { ...prev, [key]: value } : prev));
      // Clear stale save feedback when user starts editing
      setSaveSuccess(false);
      setSaveErrors([]);
      setSaveWarnings([]);
    },
    [],
  );

  const handleSave = useCallback(async () => {
    if (!form) return;
    setSaving(true);
    setSaveErrors([]);
    setSaveWarnings([]);
    setSaveSuccess(false);

    try {
      // If user entered a new raw secret, store it first
      let resolvedSecretRef = form.clientSecretRef;
      if (form.clientSecret.trim().length > 0 && !form.clientSecretRef && context.companyId) {
        resolvedSecretRef = await storeSecret(context.companyId, "m365-client-secret", form.clientSecret);
        setForm((prev) => prev ? { ...prev, clientSecretRef: resolvedSecretRef, clientSecret: "" } : prev);
      }

      // Convert digestRecipients from comma-separated string to array
      const payload = {
        ...form,
        clientSecretRef: resolvedSecretRef,
        clientSecret: undefined,
        digestRecipients: form.digestRecipients
          .split(",")
          .map((s) => s.trim())
          .filter(Boolean),
      };

      const result = (await saveConfigAction({
        companyId: context.companyId,
        ...payload,
      })) as SaveConfigResult;

      if (result.ok) {
        setSaveSuccess(true);
        if (result.warnings && result.warnings.length > 0) {
          setSaveWarnings(result.warnings);
        }
        // Reset dirty tracking — snapshot becomes the current form
        setSavedSnapshot({ ...form });
        // Reload remote config data so everything is consistent
        refresh();
      } else {
        setSaveErrors(result.errors ?? ["Unknown error saving configuration"]);
      }
    } catch (err) {
      setSaveErrors([err instanceof Error ? err.message : "Unexpected error saving configuration"]);
    } finally {
      setSaving(false);
    }
  }, [form, saveConfigAction, context.companyId, refresh]);

  const handleTestConnection = useCallback(async () => {
    if (!form) return;
    setTesting(true);
    setTestResult(null);
    try {
      // If user entered a new raw secret, store it before testing
      let resolvedSecretRef = form.clientSecretRef;
      if (form.clientSecret.trim().length > 0 && !form.clientSecretRef && context.companyId) {
        resolvedSecretRef = await storeSecret(context.companyId, "m365-client-secret", form.clientSecret);
        setForm((prev) => prev ? { ...prev, clientSecretRef: resolvedSecretRef, clientSecret: "" } : prev);
      }

      const result = (await testConnectionAction({
        companyId: context.companyId,
        tenantId: form.tenantId,
        clientId: form.clientId,
        clientSecretRef: resolvedSecretRef,
      })) as TestConnectionResult;
      setTestResult(result);
    } catch (err) {
      setTestResult({
        ok: false,
        error: err instanceof Error ? err.message : "Unknown error",
      });
    } finally {
      setTesting(false);
    }
  }, [testConnectionAction, context.companyId, form]);

  const handleWizardComplete = useCallback(() => {
    // Wizard finished — switch to form view and reload data
    setShowWizard(false);
    setShowForm(true);
    // Reset form state so it reloads from remote data
    setForm(null);
    setSavedSnapshot(null);
    refresh();
  }, [refresh]);

  const handleReconfigure = useCallback(() => {
    setShowWizard(true);
  }, []);

  // -- Render: loading / error states ----------------------------------------

  if (loading) {
    return <div style={{ padding: "20px" }}>Loading configuration...</div>;
  }
  if (error) {
    return (
      <div style={{ padding: "20px", color: "#dc2626" }}>
        Error: {error.message}
      </div>
    );
  }

  // -- Render: Setup Wizard --------------------------------------------------

  if (shouldShowWizard) {
    return (
      <SetupWizard
        companyId={context.companyId}
        onComplete={handleWizardComplete}
      />
    );
  }

  // -- Render: waiting for form init -----------------------------------------

  if (!form) {
    return <div style={{ padding: "20px" }}>Initializing...</div>;
  }

  // -- Render: form ----------------------------------------------------------

  return (
    <div style={{ padding: "20px", maxWidth: "720px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "20px" }}>
        <h2 style={{ margin: 0 }}>Microsoft 365 Settings</h2>
        <button
          style={secondaryButton}
          onClick={handleReconfigure}
        >
          Reconfigure
        </button>
      </div>

      {/* Save feedback banners */}
      {saveSuccess && (
        <div style={successBanner}>Configuration saved successfully.</div>
      )}
      {saveWarnings.length > 0 && (
        <div style={warningBanner}>
          {saveWarnings.map((w, i) => (
            <div key={i}>{w}</div>
          ))}
        </div>
      )}
      {saveErrors.length > 0 && (
        <div style={errorBanner}>
          {saveErrors.map((e, i) => (
            <div key={i}>{e}</div>
          ))}
        </div>
      )}

      {/* ── Azure AD Connection ──────────────────────────────────────────── */}
      <div style={card}>
        <div style={label}>Azure AD Connection</div>

        <div style={fieldRow}>
          <span style={fieldLabel}>Tenant ID</span>
          <input
            type="text"
            style={textInput}
            placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
            value={form.tenantId}
            onChange={(e) => updateField("tenantId", e.target.value)}
          />
        </div>

        <div style={fieldRow}>
          <span style={fieldLabel}>Client ID</span>
          <input
            type="text"
            style={textInput}
            placeholder="Application (client) ID"
            value={form.clientId}
            onChange={(e) => updateField("clientId", e.target.value)}
          />
        </div>

        <div style={fieldRow}>
          <span style={fieldLabel}>Client Secret</span>
          <input
            type="password"
            style={textInput}
            placeholder={form.clientSecretRef ? "Secret already configured (enter new value to replace)" : "Paste your Azure AD client secret"}
            value={form.clientSecret}
            onChange={(e) => {
              updateField("clientSecret", e.target.value);
              // Clear the existing ref so save will store the new value
              if (e.target.value.trim().length > 0) {
                updateField("clientSecretRef", "");
              }
            }}
          />
          {form.clientSecretRef && !form.clientSecret && (
            <span style={{ fontSize: "12px", color: "#16a34a" }}>
              Secret stored securely
            </span>

          )}
        </div>

        {/* Test Connection */}
        <div style={{ marginTop: "4px" }}>
          <button
            disabled={testing || !canTestConnection}
            onClick={handleTestConnection}
            style={
              testing || !canTestConnection
                ? secondaryButtonDisabled
                : secondaryButton
            }
          >
            {testing ? "Testing..." : "Test Connection"}
          </button>
          {testResult && (
            <span
              style={{
                marginLeft: "12px",
                color: testResult.ok ? "#16a34a" : "#dc2626",
                fontSize: "14px",
              }}
            >
              {testResult.ok
                ? "Connection successful"
                : testResult.error ?? "Connection failed"}
            </span>
          )}
        </div>
      </div>

      {/* ── Feature Toggles ──────────────────────────────────────────────── */}
      <div style={card}>
        <div style={label}>Feature Toggles</div>

        <div style={{ marginTop: "8px" }}>
          <label style={toggleRow}>
            <input
              type="checkbox"
              checked={form.enablePlanner}
              onChange={(e) => updateField("enablePlanner", e.target.checked)}
            />
            <span style={toggleLabel}>Enable Planner</span>
          </label>

          <label style={toggleRow}>
            <input
              type="checkbox"
              checked={form.enableSharePoint}
              onChange={(e) => updateField("enableSharePoint", e.target.checked)}
            />
            <span style={toggleLabel}>Enable SharePoint</span>
          </label>

          <label style={toggleRow}>
            <input
              type="checkbox"
              checked={form.enableOutlook}
              onChange={(e) => updateField("enableOutlook", e.target.checked)}
            />
            <span style={toggleLabel}>Enable Outlook</span>
          </label>
        </div>
      </div>

      {/* ── Planner Configuration (conditional) ─────────────────────────── */}
      {form.enablePlanner && (
        <div style={card}>
          <div style={label}>Planner Configuration</div>

          <div style={fieldRow}>
            <span style={fieldLabel}>Plan ID</span>
            <input
              type="text"
              style={textInput}
              placeholder="Planner Plan ID"
              value={form.plannerPlanId}
              onChange={(e) => updateField("plannerPlanId", e.target.value)}
            />
          </div>

          <div style={fieldRow}>
            <span style={fieldLabel}>Group ID</span>
            <input
              type="text"
              style={textInput}
              placeholder="M365 Group ID"
              value={form.plannerGroupId}
              onChange={(e) => updateField("plannerGroupId", e.target.value)}
            />
          </div>

          <div style={fieldRow}>
            <span style={fieldLabel}>Conflict Strategy</span>
            <select
              style={selectInput}
              value={form.conflictStrategy}
              onChange={(e) => updateField("conflictStrategy", e.target.value)}
            >
              <option value="last_write_wins">Last Write Wins</option>
              <option value="paperclip_wins">Paperclip Wins</option>
              <option value="planner_wins">Planner Wins</option>
            </select>
          </div>
        </div>
      )}

      {/* ── SharePoint Configuration (conditional) ──────────────────────── */}
      {form.enableSharePoint && (
        <div style={card}>
          <div style={label}>SharePoint Configuration</div>

          <div style={fieldRow}>
            <span style={fieldLabel}>Site ID</span>
            <input
              type="text"
              style={textInput}
              placeholder="SharePoint Site ID"
              value={form.sharepointSiteId}
              onChange={(e) => updateField("sharepointSiteId", e.target.value)}
            />
          </div>

          <div style={fieldRow}>
            <span style={fieldLabel}>Drive ID</span>
            <input
              type="text"
              style={textInput}
              placeholder="SharePoint Drive ID"
              value={form.sharepointDriveId}
              onChange={(e) => updateField("sharepointDriveId", e.target.value)}
            />
          </div>

          <div style={fieldRow}>
            <span style={fieldLabel}>Upload Folder ID</span>
            <input
              type="text"
              style={textInput}
              placeholder="SharePoint Upload Folder ID"
              value={form.sharepointUploadFolderId}
              onChange={(e) =>
                updateField("sharepointUploadFolderId", e.target.value)
              }
            />
          </div>

          <div style={fieldRow}>
            <span style={fieldLabel}>Max Document Size (bytes)</span>
            <input
              type="number"
              style={numberInput}
              min={1}
              value={form.maxDocSizeBytes}
              onChange={(e) =>
                updateField(
                  "maxDocSizeBytes",
                  parseInt(e.target.value, 10) || 0,
                )
              }
            />
          </div>
        </div>
      )}

      {/* ── Outlook Configuration (conditional) ─────────────────────────── */}
      {form.enableOutlook && (
        <div style={card}>
          <div style={label}>Outlook Configuration</div>

          <div style={fieldRow}>
            <span style={fieldLabel}>Calendar ID</span>
            <input
              type="text"
              style={textInput}
              placeholder="Outlook Calendar ID"
              value={form.outlookCalendarId}
              onChange={(e) => updateField("outlookCalendarId", e.target.value)}
            />
          </div>

          <div style={fieldRow}>
            <span style={fieldLabel}>Sender User ID</span>
            <input
              type="text"
              style={textInput}
              placeholder="User ID or UPN for sending digests"
              value={form.digestSenderUserId}
              onChange={(e) =>
                updateField("digestSenderUserId", e.target.value)
              }
            />
          </div>

          <div style={fieldRow}>
            <span style={fieldLabel}>Digest Recipients</span>
            <input
              type="text"
              style={textInput}
              placeholder="user1@example.com, user2@example.com"
              value={form.digestRecipients}
              onChange={(e) => updateField("digestRecipients", e.target.value)}
            />
            <span style={{ fontSize: "12px", opacity: 0.5, color: "inherit" }}>
              Comma-separated email addresses
            </span>
          </div>
        </div>
      )}

      {/* ── Save Button ──────────────────────────────────────────────────── */}
      <div style={{ marginTop: "8px", display: "flex", alignItems: "center", gap: "12px" }}>
        <button
          disabled={!dirty || saving}
          onClick={handleSave}
          style={!dirty || saving ? primaryButtonDisabled : primaryButton}
        >
          {saving ? "Saving..." : "Save"}
        </button>
        {!dirty && !saveSuccess && (
          <span style={{ fontSize: "13px", opacity: 0.5, color: "inherit" }}>
            No unsaved changes
          </span>
        )}
      </div>
    </div>
  );
}

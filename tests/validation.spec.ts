import { describe, expect, it } from "vitest";
import { validateConfig, type ValidationResult } from "../src/validation.js";
import { DEFAULT_CONFIG, type M365Config } from "../src/constants.js";

/** Helper: returns a fully populated valid config with all services enabled. */
function validConfig(overrides: Partial<M365Config> = {}): Partial<M365Config> {
  return {
    ...DEFAULT_CONFIG,
    tenantId: "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
    clientId: "my-client-id",
    clientSecretRef: "secret-ref-123",
    enablePlanner: true,
    enableSharePoint: true,
    enableOutlook: true,
    plannerPlanId: "plan-1",
    plannerGroupId: "group-1",
    sharepointSiteId: "site-1",
    sharepointDriveId: "drive-1",
    outlookCalendarId: "calendar-1",
    digestSenderUserId: "user-1",
    ...overrides,
  };
}

describe("validateConfig", () => {
  it("accepts a fully valid config with all services enabled", () => {
    const result = validateConfig(validConfig());
    expect(result.ok).toBe(true);
    expect(result.errors).toHaveLength(0);
  });

  it("accepts a config with no services enabled", () => {
    const result = validateConfig({
      enablePlanner: false,
      enableSharePoint: false,
      enableOutlook: false,
    });
    expect(result.ok).toBe(true);
    expect(result.errors).toHaveLength(0);
  });

  // ── Planner validation ───────────────────────────────────────────────────

  it("requires plannerPlanId when Planner is enabled", () => {
    const result = validateConfig(validConfig({ plannerPlanId: "" }));
    expect(result.ok).toBe(false);
    expect(result.errors).toContain("Planner Plan ID is required when Planner is enabled");
  });

  it("requires plannerGroupId when Planner is enabled", () => {
    const result = validateConfig(validConfig({ plannerGroupId: "" }));
    expect(result.ok).toBe(false);
    expect(result.errors).toContain("Planner Group ID is required when Planner is enabled");
  });

  it("does not require Planner fields when Planner is disabled", () => {
    const result = validateConfig(validConfig({
      enablePlanner: false,
      plannerPlanId: "",
      plannerGroupId: "",
    }));
    expect(result.ok).toBe(true);
  });

  // ── SharePoint validation ────────────────────────────────────────────────

  it("requires sharepointSiteId when SharePoint is enabled", () => {
    const result = validateConfig(validConfig({ sharepointSiteId: "" }));
    expect(result.ok).toBe(false);
    expect(result.errors).toContain("SharePoint Site ID is required when SharePoint is enabled");
  });

  it("requires sharepointDriveId when SharePoint is enabled", () => {
    const result = validateConfig(validConfig({ sharepointDriveId: "" }));
    expect(result.ok).toBe(false);
    expect(result.errors).toContain("SharePoint Drive ID is required when SharePoint is enabled");
  });

  it("does not require SharePoint fields when SharePoint is disabled", () => {
    const result = validateConfig(validConfig({
      enableSharePoint: false,
      sharepointSiteId: "",
      sharepointDriveId: "",
    }));
    expect(result.ok).toBe(true);
  });

  // ── Outlook validation ───────────────────────────────────────────────────

  it("requires outlookCalendarId when Outlook is enabled", () => {
    const result = validateConfig(validConfig({ outlookCalendarId: "" }));
    expect(result.ok).toBe(false);
    expect(result.errors).toContain("Outlook Calendar ID is required when Outlook is enabled");
  });

  it("requires digestSenderUserId when Outlook is enabled", () => {
    const result = validateConfig(validConfig({ digestSenderUserId: "" }));
    expect(result.ok).toBe(false);
    expect(result.errors).toContain("Digest Sender User ID is required when Outlook is enabled");
  });

  it("does not require Outlook fields when Outlook is disabled", () => {
    const result = validateConfig(validConfig({
      enableOutlook: false,
      outlookCalendarId: "",
      digestSenderUserId: "",
    }));
    expect(result.ok).toBe(true);
  });

  // ── Azure AD credential validation ──────────────────────────────────────

  it("requires tenantId when any service is enabled", () => {
    const result = validateConfig(validConfig({ tenantId: "" }));
    expect(result.ok).toBe(false);
    expect(result.errors).toContain("Azure AD Tenant ID is required");
  });

  it("validates tenantId is a UUID", () => {
    const result = validateConfig(validConfig({ tenantId: "not-a-uuid" }));
    expect(result.ok).toBe(false);
    expect(result.errors).toContain("Azure AD Tenant ID must be a valid UUID");
  });

  it("accepts a valid uppercase UUID for tenantId", () => {
    const result = validateConfig(validConfig({
      tenantId: "A1B2C3D4-E5F6-7890-ABCD-EF1234567890",
    }));
    expect(result.ok).toBe(true);
  });

  it("requires clientId when any service is enabled", () => {
    const result = validateConfig(validConfig({ clientId: "" }));
    expect(result.ok).toBe(false);
    expect(result.errors).toContain("Azure AD Client ID is required");
  });

  it("requires clientSecretRef when any service is enabled", () => {
    const result = validateConfig(validConfig({ clientSecretRef: "" }));
    expect(result.ok).toBe(false);
    expect(result.errors).toContain("Client Secret Reference is required");
  });

  it("does not require Azure AD fields when no service is enabled", () => {
    const result = validateConfig({
      enablePlanner: false,
      enableSharePoint: false,
      enableOutlook: false,
      tenantId: "",
      clientId: "",
      clientSecretRef: "",
    });
    expect(result.ok).toBe(true);
  });

  // ── maxDocSizeBytes ──────────────────────────────────────────────────────

  it("rejects zero maxDocSizeBytes", () => {
    const result = validateConfig(validConfig({ maxDocSizeBytes: 0 }));
    expect(result.ok).toBe(false);
    expect(result.errors).toContain("Max document size must be a positive number");
  });

  it("rejects negative maxDocSizeBytes", () => {
    const result = validateConfig(validConfig({ maxDocSizeBytes: -100 }));
    expect(result.ok).toBe(false);
    expect(result.errors).toContain("Max document size must be a positive number");
  });

  it("accepts positive maxDocSizeBytes", () => {
    const result = validateConfig(validConfig({ maxDocSizeBytes: 1024 }));
    expect(result.ok).toBe(true);
  });

  it("does not error when maxDocSizeBytes is undefined", () => {
    const cfg = validConfig();
    delete (cfg as Record<string, unknown>).maxDocSizeBytes;
    const result = validateConfig(cfg);
    expect(result.ok).toBe(true);
  });

  // ── Multiple errors ──────────────────────────────────────────────────────

  it("accumulates multiple errors", () => {
    const result = validateConfig({
      enablePlanner: true,
      enableSharePoint: true,
      enableOutlook: true,
      tenantId: "",
      clientId: "",
      clientSecretRef: "",
      plannerPlanId: "",
      plannerGroupId: "",
      sharepointSiteId: "",
      sharepointDriveId: "",
      outlookCalendarId: "",
      digestSenderUserId: "",
    });
    expect(result.ok).toBe(false);
    // Planner (2) + SharePoint (2) + Outlook (2) + Azure AD (2: tenantId + clientId + secretRef minus tenantId UUID check) = at least 8
    expect(result.errors.length).toBeGreaterThanOrEqual(8);
  });

  // ── Agentic Identity validation ─────────────────────────────────────────

  it("warns when no defaultServiceUserId is set and services are enabled", () => {
    const result = validateConfig(validConfig({ defaultServiceUserId: "" }));
    expect(result.warnings).toContain(
      "No default service user ID configured — agent identity resolution will have no fallback",
    );
  });

  it("does not warn about defaultServiceUserId when no services are enabled", () => {
    const result = validateConfig({
      enablePlanner: false,
      enableSharePoint: false,
      enableOutlook: false,
      enableTeams: false,
      enablePeople: false,
      enableMeetings: false,
      defaultServiceUserId: "",
    });
    expect(result.warnings).not.toContain(
      "No default service user ID configured — agent identity resolution will have no fallback",
    );
  });

  it("rejects agent identity map with empty values", () => {
    const result = validateConfig(validConfig({
      agentIdentityMap: { "agent-1": "" },
    }));
    expect(result.ok).toBe(false);
    expect(result.errors).toContain(
      "Agent identity map entries must have non-empty agent ID and M365 user ID",
    );
  });

  it("accepts a valid agent identity map", () => {
    const result = validateConfig(validConfig({
      agentIdentityMap: { "agent-1": "ceo@contoso.com" },
      defaultServiceUserId: "service@contoso.com",
    }));
    expect(result.ok).toBe(true);
  });

  // ── Teams validation ──────────────────────────────────────────────────────

  it("warns when teamsTeamId is missing and Teams is enabled", () => {
    const result = validateConfig(validConfig({ enableTeams: true, teamsTeamId: "", teamsDefaultChannelId: "ch-1" }));
    expect(result.ok).toBe(true);
    expect(result.warnings).toContain("Teams Team ID is required when Teams is enabled");
  });

  it("warns when teamsDefaultChannelId is missing and Teams is enabled", () => {
    const result = validateConfig(validConfig({ enableTeams: true, teamsTeamId: "team-1", teamsDefaultChannelId: "" }));
    expect(result.ok).toBe(true);
    expect(result.warnings).toContain("Teams Default Channel ID is required when Teams is enabled");
  });

  it("does not require Teams fields when Teams is disabled", () => {
    const result = validateConfig(validConfig({
      enableTeams: false,
      teamsTeamId: "",
      teamsDefaultChannelId: "",
    }));
    expect(result.ok).toBe(true);
  });

  it("requires Azure AD credentials when Teams is enabled", () => {
    const result = validateConfig({
      enableTeams: true,
      teamsTeamId: "team-1",
      teamsDefaultChannelId: "ch-1",
      tenantId: "",
      clientId: "",
      clientSecretRef: "",
    });
    expect(result.ok).toBe(false);
    expect(result.errors).toContain("Azure AD Tenant ID is required");
  });

  // ── Meetings validation ───────────────────────────────────────────────────

  it("warns when meetingOrganizerUserId is missing and Meetings is enabled", () => {
    const result = validateConfig(validConfig({ enableMeetings: true, meetingOrganizerUserId: "" }));
    expect(result.ok).toBe(true);
    expect(result.warnings).toContain("Meeting Organizer User ID is required when Meetings is enabled");
  });

  it("rejects non-positive meetingDefaultDuration", () => {
    const result = validateConfig(validConfig({
      enableMeetings: true,
      meetingOrganizerUserId: "user-1",
      meetingDefaultDuration: 0,
    }));
    expect(result.ok).toBe(false);
    expect(result.errors).toContain("Meeting default duration must be a positive number");
  });

  it("accepts valid meeting configuration", () => {
    const result = validateConfig(validConfig({
      enableMeetings: true,
      meetingOrganizerUserId: "user-1",
      meetingDefaultDuration: 30,
    }));
    expect(result.ok).toBe(true);
  });

  it("does not require meeting fields when Meetings is disabled", () => {
    const result = validateConfig(validConfig({
      enableMeetings: false,
      meetingOrganizerUserId: "",
    }));
    expect(result.ok).toBe(true);
  });

  // ── Return shape ─────────────────────────────────────────────────────────

  it("always returns a warnings array even when empty", () => {
    const result = validateConfig(validConfig());
    expect(Array.isArray(result.warnings)).toBe(true);
  });
});

import { describe, expect, it } from "vitest";
import { validateConfig } from "../src/validation.js";
import {
  DEFAULT_CONFIG,
  DEFAULT_MAX_DOC_SIZE_BYTES,
  type M365Config,
} from "../src/constants.js";

/**
 * Simulates the merge logic from the save-config action handler in worker.ts:
 *
 *   const merged = { ...DEFAULT_CONFIG, ...incoming };
 *   const validation = validateConfig(merged);
 *
 * This lets us test the integration behavior (merge + validate) without
 * requiring the full plugin SDK context.
 */
function mergeAndValidate(incoming: Partial<M365Config>) {
  const merged = { ...DEFAULT_CONFIG, ...incoming };
  const validation = validateConfig(merged);
  return { merged, validation };
}

/** Helper: a complete valid config with all three services enabled. */
function fullValidConfig(overrides: Partial<M365Config> = {}): Partial<M365Config> {
  return {
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
    sharepointUploadFolderId: "folder-1",
    outlookCalendarId: "calendar-1",
    digestRecipients: ["user-a@contoso.com"],
    digestSenderUserId: "user-1",
    webhookClientStateRef: "webhook-secret-ref",
    maxDocSizeBytes: 10 * 1024 * 1024,
    conflictStrategy: "last_write_wins",
    ...overrides,
  };
}

// ── Config Merge Behavior ──────────────────────────────────────────────────

describe("save-config: config merge behavior", () => {
  it("fills all missing fields with DEFAULT_CONFIG values when only credentials are provided", () => {
    const { merged, validation } = mergeAndValidate({
      tenantId: "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
      clientId: "my-client-id",
    });

    // No services are enabled (defaults to false) so validation passes
    expect(validation.ok).toBe(true);
    expect(validation.errors).toHaveLength(0);

    // Verify defaults are applied
    expect(merged.enablePlanner).toBe(false);
    expect(merged.enableSharePoint).toBe(false);
    expect(merged.enableOutlook).toBe(false);
    expect(merged.conflictStrategy).toBe("last_write_wins");
    expect(merged.maxDocSizeBytes).toBe(DEFAULT_MAX_DOC_SIZE_BYTES);
    expect(merged.digestRecipients).toEqual([]);
    expect(merged.plannerPlanId).toBe("");
    expect(merged.plannerGroupId).toBe("");
    expect(merged.sharepointSiteId).toBe("");
    expect(merged.sharepointDriveId).toBe("");
    expect(merged.outlookCalendarId).toBe("");
    expect(merged.digestSenderUserId).toBe("");
    expect(merged.clientSecretRef).toBe("");
  });

  it("fails validation when enablePlanner is true but plan/group IDs and credentials are missing from defaults", () => {
    const { validation } = mergeAndValidate({ enablePlanner: true });

    expect(validation.ok).toBe(false);
    expect(validation.errors).toContain("Planner Plan ID is required when Planner is enabled");
    expect(validation.errors).toContain("Planner Group ID is required when Planner is enabled");
    expect(validation.errors).toContain("Azure AD Tenant ID is required");
    expect(validation.errors).toContain("Azure AD Client ID is required");
    expect(validation.errors).toContain("Client Secret Reference is required");
  });

  it("passes validation for a complete valid config after merge", () => {
    const { validation } = mergeAndValidate(fullValidConfig());

    expect(validation.ok).toBe(true);
    expect(validation.errors).toHaveLength(0);
  });
});

// ── Round-Trip Validation ──────────────────────────────────────────────────

describe("save-config: round-trip validation", () => {
  it("validates a fully populated config with all services enabled as ok", () => {
    const { merged, validation } = mergeAndValidate(fullValidConfig());

    expect(validation.ok).toBe(true);
    expect(validation.errors).toHaveLength(0);
    expect(validation.warnings).toEqual([]);

    // Verify incoming values survive the merge
    expect(merged.tenantId).toBe("a1b2c3d4-e5f6-7890-abcd-ef1234567890");
    expect(merged.clientId).toBe("my-client-id");
    expect(merged.clientSecretRef).toBe("secret-ref-123");
    expect(merged.enablePlanner).toBe(true);
    expect(merged.enableSharePoint).toBe(true);
    expect(merged.enableOutlook).toBe(true);
    expect(merged.plannerPlanId).toBe("plan-1");
    expect(merged.plannerGroupId).toBe("group-1");
    expect(merged.sharepointSiteId).toBe("site-1");
    expect(merged.sharepointDriveId).toBe("drive-1");
    expect(merged.outlookCalendarId).toBe("calendar-1");
    expect(merged.digestSenderUserId).toBe("user-1");
  });

  it("re-validates the merged config identically to the first pass", () => {
    const { merged } = mergeAndValidate(fullValidConfig());

    // Validate the merged config a second time (simulating a re-open of the form)
    const secondPass = validateConfig(merged);
    expect(secondPass.ok).toBe(true);
    expect(secondPass.errors).toHaveLength(0);
  });
});

// ── Partial Update Scenarios (typical UI form submissions) ─────────────────

describe("save-config: partial update scenarios", () => {
  it("passes when user fills in Azure AD credentials only (no services enabled)", () => {
    const { validation } = mergeAndValidate({
      tenantId: "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
      clientId: "my-client-id",
      clientSecretRef: "secret-ref-123",
    });

    expect(validation.ok).toBe(true);
    expect(validation.errors).toHaveLength(0);
  });

  it("fails when user enables Planner but forgets Plan ID", () => {
    const { validation } = mergeAndValidate({
      tenantId: "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
      clientId: "my-client-id",
      clientSecretRef: "secret-ref-123",
      enablePlanner: true,
      plannerGroupId: "group-1",
      // plannerPlanId deliberately omitted
    });

    expect(validation.ok).toBe(false);
    expect(validation.errors).toHaveLength(1);
    expect(validation.errors[0]).toBe("Planner Plan ID is required when Planner is enabled");
  });

  it("passes when user enables all 3 services with all required fields", () => {
    const { validation } = mergeAndValidate(fullValidConfig());

    expect(validation.ok).toBe(true);
    expect(validation.errors).toHaveLength(0);
  });

  it("fails with UUID error when user provides invalid tenant ID format", () => {
    const { validation } = mergeAndValidate({
      tenantId: "not-a-valid-uuid",
      clientId: "my-client-id",
      clientSecretRef: "secret-ref-123",
      enablePlanner: true,
      plannerPlanId: "plan-1",
      plannerGroupId: "group-1",
    });

    expect(validation.ok).toBe(false);
    expect(validation.errors).toContain("Azure AD Tenant ID must be a valid UUID");
  });

  it("fails when user enables SharePoint but omits Site ID and Drive ID", () => {
    const { validation } = mergeAndValidate({
      tenantId: "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
      clientId: "my-client-id",
      clientSecretRef: "secret-ref-123",
      enableSharePoint: true,
      // sharepointSiteId and sharepointDriveId default to ""
    });

    expect(validation.ok).toBe(false);
    expect(validation.errors).toContain("SharePoint Site ID is required when SharePoint is enabled");
    expect(validation.errors).toContain("SharePoint Drive ID is required when SharePoint is enabled");
  });

  it("fails when user enables Outlook but omits Calendar ID and Sender User ID", () => {
    const { validation } = mergeAndValidate({
      tenantId: "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
      clientId: "my-client-id",
      clientSecretRef: "secret-ref-123",
      enableOutlook: true,
      // outlookCalendarId and digestSenderUserId default to ""
    });

    expect(validation.ok).toBe(false);
    expect(validation.errors).toContain("Outlook Calendar ID is required when Outlook is enabled");
    expect(validation.errors).toContain("Digest Sender User ID is required when Outlook is enabled");
  });

  it("fails when user enables multiple services but provides no credentials at all", () => {
    const { validation } = mergeAndValidate({
      enablePlanner: true,
      enableSharePoint: true,
      enableOutlook: true,
    });

    expect(validation.ok).toBe(false);
    // Service-specific errors
    expect(validation.errors).toContain("Planner Plan ID is required when Planner is enabled");
    expect(validation.errors).toContain("Planner Group ID is required when Planner is enabled");
    expect(validation.errors).toContain("SharePoint Site ID is required when SharePoint is enabled");
    expect(validation.errors).toContain("SharePoint Drive ID is required when SharePoint is enabled");
    expect(validation.errors).toContain("Outlook Calendar ID is required when Outlook is enabled");
    expect(validation.errors).toContain("Digest Sender User ID is required when Outlook is enabled");
    // Credential errors
    expect(validation.errors).toContain("Azure AD Tenant ID is required");
    expect(validation.errors).toContain("Azure AD Client ID is required");
    expect(validation.errors).toContain("Client Secret Reference is required");
  });
});

// ── DEFAULT_CONFIG Merge Edge Cases ────────────────────────────────────────

describe("save-config: DEFAULT_CONFIG merge edge cases", () => {
  it("defaults maxDocSizeBytes to DEFAULT_MAX_DOC_SIZE_BYTES (5MB) when not provided", () => {
    const { merged } = mergeAndValidate({
      tenantId: "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
      clientId: "my-client-id",
      clientSecretRef: "secret-ref-123",
    });

    expect(merged.maxDocSizeBytes).toBe(DEFAULT_MAX_DOC_SIZE_BYTES);
    expect(merged.maxDocSizeBytes).toBe(5 * 1024 * 1024);
  });

  it("allows overriding maxDocSizeBytes with a custom positive value", () => {
    const { merged, validation } = mergeAndValidate({
      maxDocSizeBytes: 20 * 1024 * 1024,
    });

    expect(validation.ok).toBe(true);
    expect(merged.maxDocSizeBytes).toBe(20 * 1024 * 1024);
  });

  it("rejects maxDocSizeBytes of zero even after merge", () => {
    const { validation } = mergeAndValidate({
      maxDocSizeBytes: 0,
    });

    expect(validation.ok).toBe(false);
    expect(validation.errors).toContain("Max document size must be a positive number");
  });

  it("rejects negative maxDocSizeBytes even after merge", () => {
    const { validation } = mergeAndValidate({
      maxDocSizeBytes: -1,
    });

    expect(validation.ok).toBe(false);
    expect(validation.errors).toContain("Max document size must be a positive number");
  });

  it("defaults conflictStrategy to 'last_write_wins' when not provided", () => {
    const { merged } = mergeAndValidate({});

    expect(merged.conflictStrategy).toBe("last_write_wins");
  });

  it("preserves an explicitly provided conflictStrategy", () => {
    const { merged } = mergeAndValidate({
      conflictStrategy: "paperclip_wins",
    });

    expect(merged.conflictStrategy).toBe("paperclip_wins");
  });

  it("defaults digestRecipients to an empty array when not provided", () => {
    const { merged } = mergeAndValidate({});

    expect(merged.digestRecipients).toEqual([]);
  });

  it("preserves an explicitly provided digestRecipients list", () => {
    const { merged } = mergeAndValidate({
      digestRecipients: ["alice@contoso.com", "bob@contoso.com"],
    });

    expect(merged.digestRecipients).toEqual(["alice@contoso.com", "bob@contoso.com"]);
  });

  it("incoming values override DEFAULT_CONFIG values for every field", () => {
    const overrides: M365Config = {
      tenantId: "11111111-2222-3333-4444-555555555555",
      clientId: "custom-client",
      clientSecretRef: "custom-secret",
      enablePlanner: true,
      enableSharePoint: true,
      enableOutlook: true,
      plannerPlanId: "custom-plan",
      plannerGroupId: "custom-group",
      conflictStrategy: "planner_wins",
      sharepointSiteId: "custom-site",
      sharepointDriveId: "custom-drive",
      sharepointUploadFolderId: "custom-folder",
      maxDocSizeBytes: 1024,
      outlookCalendarId: "custom-calendar",
      digestRecipients: ["test@example.com"],
      digestSenderUserId: "custom-sender",
      webhookClientStateRef: "custom-webhook-ref",
    };

    const { merged } = mergeAndValidate(overrides);

    for (const key of Object.keys(overrides) as Array<keyof M365Config>) {
      expect(merged[key]).toEqual(overrides[key]);
    }
  });

  it("merging an empty object results in exactly DEFAULT_CONFIG", () => {
    const { merged } = mergeAndValidate({});

    for (const key of Object.keys(DEFAULT_CONFIG) as Array<keyof M365Config>) {
      expect(merged[key]).toEqual(DEFAULT_CONFIG[key]);
    }
  });

  it("defaults webhookClientStateRef to empty string when not provided", () => {
    const { merged } = mergeAndValidate({});

    expect(merged.webhookClientStateRef).toBe("");
  });

  it("defaults sharepointUploadFolderId to empty string when not provided", () => {
    const { merged } = mergeAndValidate({});

    expect(merged.sharepointUploadFolderId).toBe("");
  });
});

import { describe, expect, it } from "vitest";
import { createTestHarness } from "@paperclipai/plugin-sdk/testing";
import manifest from "../src/manifest.js";
import plugin from "../src/worker.js";

describe("sharepoint tools", () => {
  function createHarness() {
    const harness = createTestHarness({
      manifest,
      capabilities: [...manifest.capabilities],
      config: {
        tenantId: "test-tenant",
        clientId: "test-client",
        clientSecretRef: "secret:m365-client-secret",
        enableSharePoint: true,
        sharepointSiteId: "site-123",
        sharepointDriveId: "drive-456",
        sharepointUploadFolderId: "folder-789",
      },
    });
    return harness;
  }

  it("registers all sharepoint tools", async () => {
    const harness = createHarness();
    await plugin.definition.setup(harness.ctx);

    // sharepoint-search tool should be registered — calling with empty query
    // should return an error
    const searchResult = await harness.executeTool("sharepoint-search", {});
    expect(searchResult.error).toBe("query is required");

    // sharepoint-read tool validation
    const readResult = await harness.executeTool("sharepoint-read", {});
    expect(readResult.error).toBe("driveId and itemId are required");

    // sharepoint-upload tool validation
    const uploadResult = await harness.executeTool("sharepoint-upload", {});
    expect(uploadResult.error).toBe("fileName and content are required");
  });

  it("planner-status tool returns not-linked for unknown issue", async () => {
    const harness = createHarness();
    harness.setConfig({
      ...harness.ctx.manifest,
      tenantId: "test-tenant",
      clientId: "test-client",
      clientSecretRef: "secret:m365-client-secret",
      enablePlanner: true,
      plannerPlanId: "plan-123",
      plannerGroupId: "group-456",
    });
    await plugin.definition.setup(harness.ctx);

    const result = await harness.executeTool("planner-status", { issueId: "nonexistent" });
    expect(result.content).toContain("No Planner task linked");
    expect(result.data).toEqual({ linked: false });
  });
});

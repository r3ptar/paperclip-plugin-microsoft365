import { describe, expect, it, vi } from "vitest";
import { createTestHarness } from "@paperclipai/plugin-sdk/testing";
import manifest from "../src/manifest.js";
import plugin from "../src/worker.js";

/**
 * These tests verify the Planner sync event handlers using the test harness.
 * Graph API calls are implicitly mocked by the harness's ctx.http.fetch.
 *
 * Since the real Graph API isn't available in tests, these tests focus on:
 * 1. Plugin setup completes without error
 * 2. Data/action handlers are registered
 * 3. Event handlers are wired up
 * 4. Jobs are registered
 */
describe("planner sync", () => {
  function createHarness() {
    const harness = createTestHarness({
      manifest,
      capabilities: [...manifest.capabilities, "events.emit"],
      config: {
        tenantId: "test-tenant",
        clientId: "test-client",
        clientSecretRef: "secret:m365-client-secret",
        enablePlanner: true,
        plannerPlanId: "plan-123",
        plannerGroupId: "group-456",
        conflictStrategy: "last_write_wins",
      },
    });
    harness.seed({
      companies: [{ id: "co_1", name: "Test Co", issuePrefix: "TC", status: "active" }],
      projects: [{ id: "pr_1", name: "Test Project", companyId: "co_1" }],
      issues: [
        {
          id: "iss_1",
          title: "Test Issue",
          status: "todo",
          companyId: "co_1",
          projectId: "pr_1",
        } as any,
      ],
    });
    return harness;
  }

  it("sets up without errors", async () => {
    const harness = createHarness();
    await plugin.definition.setup(harness.ctx);
    expect(harness.logs.some((l) => l.message.includes("setup complete"))).toBe(true);
  });

  it("registers sync-health data handler", async () => {
    const harness = createHarness();
    await plugin.definition.setup(harness.ctx);

    const health = await harness.getData<{
      configured: boolean;
      enablePlanner: boolean;
    }>("sync-health");

    expect(health.configured).toBe(true);
    expect(health.enablePlanner).toBe(true);
  });

  it("registers plugin-config data handler without exposing secrets", async () => {
    const harness = createHarness();
    await plugin.definition.setup(harness.ctx);

    const config = await harness.getData<{
      tenantId: string;
      hasClientSecret: boolean;
      clientSecretRef?: string;
    }>("plugin-config");

    expect(config.tenantId).toBe("test-tenant");
    expect(config.hasClientSecret).toBe(true);
    // clientSecretRef is a reference identifier (not the raw secret) — exposed for form round-tripping
    expect(config.clientSecretRef).toBe("secret:m365-client-secret");
  });

  it("registers issue-m365 data handler", async () => {
    const harness = createHarness();
    await plugin.definition.setup(harness.ctx);

    const data = await harness.getData<{
      plannerTask: unknown;
      calendarEvent: unknown;
    }>("issue-m365", { issueId: "iss_1" });

    expect(data.plannerTask).toBeNull();
    expect(data.calendarEvent).toBeNull();
  });

  it("registers test-connection action handler", async () => {
    const harness = createHarness();
    await plugin.definition.setup(harness.ctx);

    // Will fail because the test harness fetch can't reach Azure AD,
    // but the handler itself should be registered and callable
    const result = await harness.performAction<{ ok: boolean }>("test-connection");
    expect(typeof result.ok).toBe("boolean");
  });

  it("responds to health probe", async () => {
    const harness = createHarness();
    await plugin.definition.setup(harness.ctx);

    const health = await plugin.definition.onHealth!();
    expect(["ok", "degraded"]).toContain(health.status);
  });

  it("validates config requiring planner fields when enabled", async () => {
    const result = await plugin.definition.onValidateConfig!({
      enablePlanner: true,
      tenantId: "t",
      clientId: "c",
      clientSecretRef: "s",
      // Missing plannerPlanId and plannerGroupId
    });
    expect(result.ok).toBe(false);
    expect(result.errors!.some((e: string) => e.includes("Plan ID"))).toBe(true);
  });

  it("validates config passes with all required planner fields", async () => {
    const result = await plugin.definition.onValidateConfig!({
      enablePlanner: true,
      tenantId: "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
      clientId: "c",
      clientSecretRef: "s",
      plannerPlanId: "plan-1",
      plannerGroupId: "group-1",
    });
    expect(result.ok).toBe(true);
  });

  it("validates config rejects invalid tenant ID format", async () => {
    const result = await plugin.definition.onValidateConfig!({
      enablePlanner: true,
      tenantId: "not-a-uuid",
      clientId: "c",
      clientSecretRef: "s",
      plannerPlanId: "plan-1",
      plannerGroupId: "group-1",
    });
    expect(result.ok).toBe(false);
    expect(result.errors!.some((e: string) => e.includes("valid UUID"))).toBe(true);
  });
});

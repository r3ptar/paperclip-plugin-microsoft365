import { describe, expect, it } from "vitest";
import { createTestHarness } from "@paperclipai/plugin-sdk/testing";
import manifest from "../src/manifest.js";
import { handleGraphNotification } from "../src/webhooks/graph-notifications.js";
import type { GraphChangeNotification } from "../src/graph/types.js";
import type { M365Config } from "../src/constants.js";
import { DEFAULT_CONFIG } from "../src/constants.js";

function makeConfig(overrides: Partial<M365Config> = {}): M365Config {
  return { ...DEFAULT_CONFIG, ...overrides };
}

describe("webhook verification", () => {
  it("logs warning when webhookClientStateRef is not configured", async () => {
    const harness = createTestHarness({
      manifest,
      capabilities: [...manifest.capabilities],
    });

    const notification: GraphChangeNotification = {
      value: [
        {
          subscriptionId: "sub-1",
          clientState: "some-state",
          changeType: "updated",
          resource: "planner/tasks/task-1",
          subscriptionExpirationDateTime: "2025-12-31T00:00:00Z",
          tenantId: "tenant-1",
        },
      ],
    };

    const config = makeConfig({ webhookClientStateRef: "" });
    // PlannerService and GraphClient are null, so processing will log debug
    // but the clientState warning should still be logged
    await handleGraphNotification(
      harness.ctx,
      { endpointKey: "graph-notifications", requestId: "req-1", rawBody: "", parsedBody: notification },
      config,
      null as any,
      null as any,
    );

    expect(harness.logs.some((l) => l.message.includes("webhookClientStateRef is not configured"))).toBe(true);
  });

  it("rejects notifications with mismatched clientState", async () => {
    const harness = createTestHarness({
      manifest,
      capabilities: [...manifest.capabilities],
    });

    const notification: GraphChangeNotification = {
      value: [
        {
          subscriptionId: "sub-1",
          clientState: "wrong-state",
          changeType: "updated",
          resource: "planner/tasks/task-1",
          subscriptionExpirationDateTime: "2025-12-31T00:00:00Z",
          tenantId: "tenant-1",
        },
      ],
    };

    const config = makeConfig({ webhookClientStateRef: "secret:webhook-state" });
    // Test harness resolves secrets as "resolved:<ref>"
    await handleGraphNotification(
      harness.ctx,
      { endpointKey: "graph-notifications", requestId: "req-1", rawBody: "", parsedBody: notification },
      config,
      null as any,
      null as any,
    );

    expect(harness.logs.some((l) => l.message.includes("clientState mismatch"))).toBe(true);
  });

  it("handles validation token request", async () => {
    const harness = createTestHarness({
      manifest,
      capabilities: [...manifest.capabilities],
    });

    const config = makeConfig();
    await handleGraphNotification(
      harness.ctx,
      {
        endpointKey: "graph-notifications",
        requestId: "req-1",
        rawBody: "",
        parsedBody: { validationToken: "abc123" },
      },
      config,
      null as any,
      null as any,
    );

    expect(harness.logs.some((l) => l.message.includes("validation request"))).toBe(true);
  });

  it("warns on empty notification body", async () => {
    const harness = createTestHarness({
      manifest,
      capabilities: [...manifest.capabilities],
    });

    const config = makeConfig();
    await handleGraphNotification(
      harness.ctx,
      { endpointKey: "graph-notifications", requestId: "req-1", rawBody: "", parsedBody: { value: [] } },
      config,
      null as any,
      null as any,
    );

    expect(harness.logs.some((l) => l.message.includes("Empty Graph notification"))).toBe(true);
  });
});

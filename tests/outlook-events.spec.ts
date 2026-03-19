import { describe, expect, it } from "vitest";
import { createTestHarness } from "@paperclipai/plugin-sdk/testing";
import manifest from "../src/manifest.js";
import plugin from "../src/worker.js";
import { OutlookService } from "../src/services/outlook.js";

describe("outlook", () => {
  it("builds digest HTML with issue data", () => {
    // OutlookService.buildDigestHtml is a pure function, testable without Graph
    const harness = createTestHarness({ manifest });
    const service = new (OutlookService as any)(
      harness.ctx,
      null, // graph client not needed for buildDigestHtml
      {},
    ) as OutlookService;

    const html = service.buildDigestHtml([
      { id: "iss_1", title: "Fix login bug", status: "in_progress", updatedAt: "2025-03-15T10:00:00Z" },
      { id: "iss_2", title: "Add tests", status: "done", updatedAt: "2025-03-15T11:00:00Z" },
    ]);

    expect(html).toContain("Paperclip Daily Digest");
    expect(html).toContain("Fix login bug");
    expect(html).toContain("Add tests");
    expect(html).toContain("in_progress");
    expect(html).toContain("done");
  });

  it("builds digest HTML with no issues", () => {
    const harness = createTestHarness({ manifest });
    const service = new (OutlookService as any)(harness.ctx, null, {}) as OutlookService;

    const html = service.buildDigestHtml([]);
    expect(html).toContain("No recent activity");
  });

  it("escapes HTML in issue titles", () => {
    const harness = createTestHarness({ manifest });
    const service = new (OutlookService as any)(harness.ctx, null, {}) as OutlookService;

    const html = service.buildDigestHtml([
      { id: "iss_1", title: '<script>alert("xss")</script>', status: "todo", updatedAt: "2025-03-15T10:00:00Z" },
    ]);

    expect(html).not.toContain("<script>");
    expect(html).toContain("&lt;script&gt;");
  });

  it("sets up with outlook enabled", async () => {
    const harness = createTestHarness({
      manifest,
      config: {
        tenantId: "test-tenant",
        clientId: "test-client",
        clientSecretRef: "secret:s",
        enableOutlook: true,
        outlookCalendarId: "cal-1",
        digestSenderUserId: "user-1",
        digestRecipients: ["user@example.com"],
      },
    });
    await plugin.definition.setup(harness.ctx);
    expect(harness.logs.some((l) => l.message.includes("setup complete"))).toBe(true);
  });

  it("validates outlook config requires calendar and sender", async () => {
    const result = await plugin.definition.onValidateConfig!({
      enableOutlook: true,
      tenantId: "t",
      clientId: "c",
      clientSecretRef: "s",
    });
    expect(result.ok).toBe(false);
    expect(result.errors!.some((e: string) => e.includes("Calendar ID"))).toBe(true);
    expect(result.errors!.some((e: string) => e.includes("Sender User ID"))).toBe(true);
  });
});

import { describe, expect, it } from "vitest";
import { resolveConflict } from "../src/sync/conflict.js";

describe("conflict resolution", () => {
  it("paperclip_wins always returns paperclip", () => {
    expect(
      resolveConflict({
        paperclipUpdatedAt: "2024-01-01T00:00:00Z",
        plannerUpdatedAt: "2025-01-01T00:00:00Z",
        strategy: "paperclip_wins",
      }),
    ).toBe("paperclip");
  });

  it("planner_wins always returns planner", () => {
    expect(
      resolveConflict({
        paperclipUpdatedAt: "2025-01-01T00:00:00Z",
        plannerUpdatedAt: "2024-01-01T00:00:00Z",
        strategy: "planner_wins",
      }),
    ).toBe("planner");
  });

  it("last_write_wins picks the more recent timestamp", () => {
    expect(
      resolveConflict({
        paperclipUpdatedAt: "2025-03-15T10:00:00Z",
        plannerUpdatedAt: "2025-03-15T09:00:00Z",
        strategy: "last_write_wins",
      }),
    ).toBe("paperclip");

    expect(
      resolveConflict({
        paperclipUpdatedAt: "2025-03-15T09:00:00Z",
        plannerUpdatedAt: "2025-03-15T10:00:00Z",
        strategy: "last_write_wins",
      }),
    ).toBe("planner");
  });

  it("last_write_wins picks paperclip on equal timestamps", () => {
    expect(
      resolveConflict({
        paperclipUpdatedAt: "2025-03-15T10:00:00Z",
        plannerUpdatedAt: "2025-03-15T10:00:00Z",
        strategy: "last_write_wins",
      }),
    ).toBe("paperclip");
  });
});

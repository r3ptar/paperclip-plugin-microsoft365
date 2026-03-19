import { describe, expect, it } from "vitest";
import { isStatusInSync, toPaperclipStatus, toPlannerStatus } from "../src/sync/status-map.js";

describe("status-map", () => {
  describe("toPlannerStatus", () => {
    it("maps backlog to 0% / Backlog bucket", () => {
      const result = toPlannerStatus("backlog");
      expect(result.percentComplete).toBe(0);
      expect(result.bucketName).toBe("Backlog");
    });

    it("maps todo to 0% / To Do bucket", () => {
      const result = toPlannerStatus("todo");
      expect(result.percentComplete).toBe(0);
      expect(result.bucketName).toBe("To Do");
    });

    it("maps in_progress to 50% / In Progress bucket", () => {
      const result = toPlannerStatus("in_progress");
      expect(result.percentComplete).toBe(50);
      expect(result.bucketName).toBe("In Progress");
    });

    it("maps in_review to 50% / In Review bucket", () => {
      const result = toPlannerStatus("in_review");
      expect(result.percentComplete).toBe(50);
      expect(result.bucketName).toBe("In Review");
    });

    it("maps done to 100% / Completed bucket", () => {
      const result = toPlannerStatus("done");
      expect(result.percentComplete).toBe(100);
      expect(result.bucketName).toBe("Completed");
    });

    it("maps blocked to 50% / Blocked bucket", () => {
      const result = toPlannerStatus("blocked");
      expect(result.percentComplete).toBe(50);
      expect(result.bucketName).toBe("Blocked");
    });

    it("maps cancelled to 100% / Cancelled bucket", () => {
      const result = toPlannerStatus("cancelled");
      expect(result.percentComplete).toBe(100);
      expect(result.bucketName).toBe("Cancelled");
    });
  });

  describe("toPaperclipStatus", () => {
    it("uses bucket name as primary discriminator", () => {
      expect(toPaperclipStatus(50, "In Progress")).toBe("in_progress");
      expect(toPaperclipStatus(50, "In Review")).toBe("in_review");
      expect(toPaperclipStatus(50, "Blocked")).toBe("blocked");
    });

    it("distinguishes 0% buckets by name", () => {
      expect(toPaperclipStatus(0, "Backlog")).toBe("backlog");
      expect(toPaperclipStatus(0, "To Do")).toBe("todo");
    });

    it("distinguishes 100% buckets by name", () => {
      expect(toPaperclipStatus(100, "Completed")).toBe("done");
      expect(toPaperclipStatus(100, "Cancelled")).toBe("cancelled");
    });

    it("falls back to percent heuristic for unknown buckets", () => {
      expect(toPaperclipStatus(0, "Custom Bucket")).toBe("todo");
      expect(toPaperclipStatus(50, "Custom Bucket")).toBe("in_progress");
      expect(toPaperclipStatus(100, "Custom Bucket")).toBe("done");
    });
  });

  describe("isStatusInSync", () => {
    it("returns true when status matches", () => {
      expect(isStatusInSync("in_progress", 50, "In Progress")).toBe(true);
      expect(isStatusInSync("done", 100, "Completed")).toBe(true);
      expect(isStatusInSync("todo", 0, "To Do")).toBe(true);
    });

    it("returns false when percent differs", () => {
      expect(isStatusInSync("in_progress", 0, "In Progress")).toBe(false);
    });

    it("returns false when bucket differs", () => {
      expect(isStatusInSync("in_progress", 50, "Blocked")).toBe(false);
    });
  });
});

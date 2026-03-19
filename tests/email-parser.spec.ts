import { describe, expect, it } from "vitest";
import { parseInboundEmail } from "../src/services/email-parser.js";
import type { GraphMailMessage } from "../src/graph/types.js";

// ---------------------------------------------------------------------------
// Test helpers
// ---------------------------------------------------------------------------

function makeMessage(overrides: Partial<GraphMailMessage> = {}): GraphMailMessage {
  return {
    id: "msg-1",
    subject: "Re: Task update",
    body: { contentType: "text", content: "Some reply text" },
    from: { emailAddress: { name: "Alice", address: "alice@example.com" } },
    toRecipients: [{ emailAddress: { name: "Bot", address: "bot@example.com" } }],
    receivedDateTime: "2026-03-19T12:00:00Z",
    ...overrides,
  };
}

function makeMessageWithHeader(issueId: string, body = "Some reply text"): GraphMailMessage {
  return makeMessage({
    internetMessageHeaders: [{ name: "X-Paperclip-Issue-Id", value: issueId }],
    body: { contentType: "text", content: body },
  });
}

function makeMessageWithSubjectTag(issueId: string, body = "Some reply text"): GraphMailMessage {
  return makeMessage({
    subject: `Re: [PC-${issueId}] Task update`,
    body: { contentType: "text", content: body },
  });
}

function makeMessageWithBodyTag(issueId: string, extraBody = ""): GraphMailMessage {
  return makeMessage({
    subject: "Re: Task update",
    body: { contentType: "text", content: `[PC-${issueId}] ${extraBody}`.trim() },
  });
}

// ---------------------------------------------------------------------------
// Issue ID extraction
// ---------------------------------------------------------------------------

describe("email-parser", () => {
  describe("issue ID extraction", () => {
    it("extracts issue ID from X-Paperclip-Issue-Id header", () => {
      const result = parseInboundEmail(makeMessageWithHeader("issue-42"));
      expect(result).toMatchObject({ issueId: "issue-42" });
    });

    it("falls back to subject line [PC-{id}] pattern when no header", () => {
      const result = parseInboundEmail(makeMessageWithSubjectTag("issue-99"));
      expect(result).toMatchObject({ issueId: "issue-99" });
    });

    it("falls back to body [PC-{id}] pattern when no header or subject match", () => {
      const result = parseInboundEmail(makeMessageWithBodyTag("issue-7", "looks good"));
      expect(result).toMatchObject({ issueId: "issue-7" });
    });

    it('returns "unrecognized" when no issue ID found anywhere', () => {
      const result = parseInboundEmail(
        makeMessage({
          subject: "Hello",
          body: { contentType: "text", content: "Just a random email" },
        }),
      );
      expect(result.kind).toBe("unrecognized");
    });

    it("handles header with extra whitespace", () => {
      const result = parseInboundEmail(
        makeMessage({
          internetMessageHeaders: [{ name: "X-Paperclip-Issue-Id", value: "  issue-55  " }],
          body: { contentType: "text", content: "done" },
        }),
      );
      // The header value is used as-is; the function does not trim it
      expect(result).toMatchObject({ issueId: "  issue-55  " });
    });
  });

  // ---------------------------------------------------------------------------
  // Intent detection - status changes
  // ---------------------------------------------------------------------------

  describe("intent detection - status changes", () => {
    it('"done" -> status_change to "done"', () => {
      const result = parseInboundEmail(makeMessageWithHeader("id-1", "done"));
      expect(result).toMatchObject({ kind: "status_change", newStatus: "done" });
    });

    it('"complete" -> status_change to "done"', () => {
      const result = parseInboundEmail(makeMessageWithHeader("id-1", "complete"));
      expect(result).toMatchObject({ kind: "status_change", newStatus: "done" });
    });

    it('"finished" -> status_change to "done"', () => {
      const result = parseInboundEmail(makeMessageWithHeader("id-1", "finished"));
      expect(result).toMatchObject({ kind: "status_change", newStatus: "done" });
    });

    it('"lgtm" -> status_change to "done"', () => {
      const result = parseInboundEmail(makeMessageWithHeader("id-1", "lgtm"));
      expect(result).toMatchObject({ kind: "status_change", newStatus: "done" });
    });

    it('"blocked" -> status_change to "blocked"', () => {
      const result = parseInboundEmail(makeMessageWithHeader("id-1", "blocked"));
      expect(result).toMatchObject({ kind: "status_change", newStatus: "blocked" });
    });

    it('"stuck" -> status_change to "blocked"', () => {
      const result = parseInboundEmail(makeMessageWithHeader("id-1", "stuck"));
      expect(result).toMatchObject({ kind: "status_change", newStatus: "blocked" });
    });

    it('"in review" -> status_change to "in_review"', () => {
      const result = parseInboundEmail(makeMessageWithHeader("id-1", "in review"));
      expect(result).toMatchObject({ kind: "status_change", newStatus: "in_review" });
    });

    it('"started" -> status_change to "in_progress"', () => {
      const result = parseInboundEmail(makeMessageWithHeader("id-1", "started"));
      expect(result).toMatchObject({ kind: "status_change", newStatus: "in_progress" });
    });

    it('"working on it" -> status_change to "in_progress"', () => {
      const result = parseInboundEmail(makeMessageWithHeader("id-1", "working on it"));
      expect(result).toMatchObject({ kind: "status_change", newStatus: "in_progress" });
    });

    it('"on it" -> status_change to "in_progress"', () => {
      const result = parseInboundEmail(makeMessageWithHeader("id-1", "on it"));
      expect(result).toMatchObject({ kind: "status_change", newStatus: "in_progress" });
    });

    it("is case insensitive for status keywords", () => {
      expect(parseInboundEmail(makeMessageWithHeader("id-1", "DONE"))).toMatchObject({
        kind: "status_change",
        newStatus: "done",
      });
      expect(parseInboundEmail(makeMessageWithHeader("id-1", "Done"))).toMatchObject({
        kind: "status_change",
        newStatus: "done",
      });
      expect(parseInboundEmail(makeMessageWithHeader("id-1", "dOnE"))).toMatchObject({
        kind: "status_change",
        newStatus: "done",
      });
    });
  });

  // ---------------------------------------------------------------------------
  // Intent detection - comments
  // ---------------------------------------------------------------------------

  describe("intent detection - comments", () => {
    it("regular text body becomes a comment", () => {
      const result = parseInboundEmail(
        makeMessageWithHeader("id-1", "I will look into this tomorrow"),
      );
      expect(result).toMatchObject({
        kind: "comment",
        issueId: "id-1",
        body: "I will look into this tomorrow",
      });
    });

    it("multi-line body is preserved as a comment", () => {
      const body = "First line\nSecond line\nThird line";
      const result = parseInboundEmail(makeMessageWithHeader("id-1", body));
      expect(result).toMatchObject({
        kind: "comment",
        issueId: "id-1",
        body,
      });
    });
  });

  // ---------------------------------------------------------------------------
  // Reply stripping
  // ---------------------------------------------------------------------------

  describe("reply stripping", () => {
    it('strips content after "On ... wrote:" line', () => {
      const body = "Looks good to me\n\nOn March 18, 2026 Alice wrote:\n> old content";
      const result = parseInboundEmail(makeMessageWithHeader("id-1", body));
      expect(result).toMatchObject({
        kind: "comment",
        body: "Looks good to me",
      });
    });

    it('strips content after "---" separator', () => {
      const body = "My reply here\n---\nOriginal message follows";
      const result = parseInboundEmail(makeMessageWithHeader("id-1", body));
      expect(result).toMatchObject({
        kind: "comment",
        body: "My reply here",
      });
    });

    it('strips consecutive quoted lines starting with ">"', () => {
      const body = "My thoughts\n> quoted line 1\n> quoted line 2\nmore text";
      const result = parseInboundEmail(makeMessageWithHeader("id-1", body));
      expect(result).toMatchObject({
        kind: "comment",
        body: "My thoughts",
      });
    });

    it("handles HTML body content type by stripping tags", () => {
      const result = parseInboundEmail(
        makeMessage({
          internetMessageHeaders: [{ name: "X-Paperclip-Issue-Id", value: "id-1" }],
          body: {
            contentType: "html",
            content: "<div><p>This is my reply</p><p>Second paragraph</p></div>",
          },
        }),
      );
      // HTML tags stripped, block-level closing tags replaced with newlines
      expect(result.kind).not.toBe("unrecognized");
      if (result.kind === "comment") {
        expect(result.body).toContain("This is my reply");
        expect(result.body).toContain("Second paragraph");
      }
    });

    it("handles body with only quoted content after reply text", () => {
      const body = "Agreed\n> previous message line 1\n> previous message line 2";
      const result = parseInboundEmail(makeMessageWithHeader("id-1", body));
      expect(result.kind).not.toBe("unrecognized");
      if (result.kind === "comment" || result.kind === "status_change") {
        expect("issueId" in result && result.issueId).toBe("id-1");
      }
    });
  });

  // ---------------------------------------------------------------------------
  // Edge cases
  // ---------------------------------------------------------------------------

  describe("edge cases", () => {
    it("empty body returns unrecognized", () => {
      const result = parseInboundEmail(makeMessageWithHeader("id-1", ""));
      expect(result.kind).toBe("unrecognized");
    });

    it("body that is just whitespace returns unrecognized", () => {
      const result = parseInboundEmail(makeMessageWithHeader("id-1", "   \n  \n  "));
      expect(result.kind).toBe("unrecognized");
    });

    it("handles very long issue ID in header", () => {
      const longId = "a".repeat(500);
      const result = parseInboundEmail(makeMessageWithHeader(longId, "some reply"));
      expect(result).toMatchObject({ issueId: longId });
    });

    it("subject with multiple [PC-...] tags picks the first one", () => {
      const result = parseInboundEmail(
        makeMessage({
          subject: "Re: [PC-first-id] and [PC-second-id] tasks",
          body: { contentType: "text", content: "my comment" },
        }),
      );
      expect(result).toMatchObject({ issueId: "first-id" });
    });
  });
});

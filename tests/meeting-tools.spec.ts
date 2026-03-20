import { describe, expect, it } from "vitest";
import { handleMeetingSchedule } from "../src/tools/meeting-schedule.js";
import { handleMeetingFindTime } from "../src/tools/meeting-find-time.js";
import { handleMeetingCancel } from "../src/tools/meeting-cancel.js";
import { handleMeetingList } from "../src/tools/meeting-list.js";
import { AgentIdentityService } from "../src/services/identity.js";
import { DEFAULT_CONFIG } from "../src/constants.js";
import type { ToolRunContext } from "@paperclipai/plugin-sdk";

const mockRunCtx = { agentId: "agent-1", companyId: "co-1" } as ToolRunContext;

const identityService = new AgentIdentityService({
  ...DEFAULT_CONFIG,
  agentIdentityMap: { "agent-1": "ceo@contoso.com" },
  defaultServiceUserId: "service@contoso.com",
});

const noIdentityService = new AgentIdentityService({
  ...DEFAULT_CONFIG,
  agentIdentityMap: {},
  defaultServiceUserId: "",
});

describe("meeting-schedule handler", () => {
  it("returns error when subject is missing", async () => {
    const result = await handleMeetingSchedule(
      { attendeeEmails: ["a@b.com"], startDateTime: "2026-03-25T10:00:00Z" },
      mockRunCtx,
      {} as never,
      identityService,
    );
    expect(result).toHaveProperty("error", "subject is required");
  });

  it("returns error when attendeeEmails is empty", async () => {
    const result = await handleMeetingSchedule(
      { subject: "Test", attendeeEmails: [], startDateTime: "2026-03-25T10:00:00Z" },
      mockRunCtx,
      {} as never,
      identityService,
    );
    expect(result).toHaveProperty("error", "attendeeEmails is required");
  });

  it("returns error when startDateTime is missing", async () => {
    const result = await handleMeetingSchedule(
      { subject: "Test", attendeeEmails: ["a@b.com"] },
      mockRunCtx,
      {} as never,
      identityService,
    );
    expect(result).toHaveProperty("error", "startDateTime is required");
  });

  it("returns error for invalid startDateTime", async () => {
    const result = await handleMeetingSchedule(
      { subject: "Test", attendeeEmails: ["a@b.com"], startDateTime: "not-a-date" },
      mockRunCtx,
      {} as never,
      identityService,
    );
    expect(result).toHaveProperty("error", "startDateTime must be a valid ISO 8601 date string");
  });

  it("returns error when no identity can be resolved", async () => {
    const result = await handleMeetingSchedule(
      { subject: "Test", attendeeEmails: ["a@b.com"], startDateTime: "2026-03-25T10:00:00Z" },
      mockRunCtx,
      {} as never,
      noIdentityService,
    );
    expect(result).toHaveProperty("error", "No M365 user identity configured for this agent");
  });
});

describe("meeting-find-time handler", () => {
  it("returns error when attendeeEmails is empty", async () => {
    const result = await handleMeetingFindTime(
      { attendeeEmails: [] },
      mockRunCtx,
      {} as never,
    );
    expect(result).toHaveProperty("error", "attendeeEmails is required");
  });

  it("returns no-results message when empty suggestions", async () => {
    const mockService = {
      findMeetingTimes: () => Promise.resolve({
        meetingTimeSuggestions: [],
        emptySuggestionsReason: "AttendeesUnavailable",
      }),
    } as never;
    const result = await handleMeetingFindTime(
      { attendeeEmails: ["a@b.com"] },
      mockRunCtx,
      mockService,
    );
    expect(result).toHaveProperty("content");
    expect((result as { content: string }).content).toContain("AttendeesUnavailable");
  });
});

describe("meeting-cancel handler", () => {
  it("returns error when eventId is missing", async () => {
    const result = await handleMeetingCancel({}, mockRunCtx, {} as never, identityService);
    expect(result).toHaveProperty("error", "eventId is required");
  });

  it("returns error for invalid eventId", async () => {
    const result = await handleMeetingCancel(
      { eventId: "../../evil" },
      mockRunCtx,
      {} as never,
      identityService,
    );
    expect(result).toHaveProperty("error", "Invalid eventId format");
  });

  it("returns error when no identity can be resolved", async () => {
    const result = await handleMeetingCancel(
      { eventId: "event-123" },
      mockRunCtx,
      {} as never,
      noIdentityService,
    );
    expect(result).toHaveProperty("error", "No M365 user identity configured for this agent");
  });
});

describe("meeting-list handler", () => {
  it("returns error when no identity can be resolved", async () => {
    const result = await handleMeetingList({}, mockRunCtx, {} as never, noIdentityService);
    expect(result).toHaveProperty("error", "No M365 user identity configured for this agent");
  });

  it("returns empty list gracefully", async () => {
    const mockService = {
      listMeetings: () => Promise.resolve([]),
    } as never;
    const result = await handleMeetingList({}, mockRunCtx, mockService, identityService);
    expect(result).toHaveProperty("content", "No upcoming meetings found");
  });
});

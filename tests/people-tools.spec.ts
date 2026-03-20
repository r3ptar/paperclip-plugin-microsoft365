import { describe, expect, it } from "vitest";
import { handlePeopleLookup } from "../src/tools/people-lookup.js";
import { handlePeopleGetPresence } from "../src/tools/people-get-presence.js";
import { handlePeopleGetManager } from "../src/tools/people-get-manager.js";
import { handlePeopleListTeamMembers } from "../src/tools/people-list-team-members.js";
import type { ToolRunContext } from "@paperclipai/plugin-sdk";

const mockRunCtx = { agentId: "agent-1", companyId: "co-1" } as ToolRunContext;

describe("people-lookup handler", () => {
  it("returns error when query is missing", async () => {
    const result = await handlePeopleLookup({}, mockRunCtx, {} as never);
    expect(result).toHaveProperty("error", "query is required");
  });

  it("propagates service errors", async () => {
    const mockService = {
      lookupUser: () => Promise.reject(new Error("Graph 403")),
    } as never;
    const result = await handlePeopleLookup({ query: "John" }, mockRunCtx, mockService);
    expect(result).toHaveProperty("error");
    expect((result as { error: string }).error).toContain("Graph 403");
  });

  it("returns formatted results", async () => {
    const mockService = {
      lookupUser: () => Promise.resolve([
        { id: "u-1", displayName: "John Doe", mail: "john@contoso.com", userPrincipalName: "john@contoso.com" },
      ]),
    } as never;
    const result = await handlePeopleLookup({ query: "John" }, mockRunCtx, mockService);
    expect(result).toHaveProperty("content");
    expect((result as { content: string }).content).toContain("John Doe");
  });
});

describe("people-get-presence handler", () => {
  it("returns error when neither userId nor userIds provided", async () => {
    const result = await handlePeopleGetPresence({}, mockRunCtx, {} as never);
    expect(result).toHaveProperty("error", "Either userId or userIds is required");
  });

  it("returns error for invalid userId", async () => {
    const result = await handlePeopleGetPresence(
      { userId: "../bad" },
      mockRunCtx,
      {} as never,
    );
    expect(result).toHaveProperty("error", "Invalid userId format");
  });

  it("returns error for invalid userId in batch", async () => {
    const result = await handlePeopleGetPresence(
      { userIds: ["valid-id", "../../bad"] },
      mockRunCtx,
      {} as never,
    );
    expect(result).toHaveProperty("error");
    expect((result as { error: string }).error).toContain("Invalid userId format");
  });
});

describe("people-get-manager handler", () => {
  it("returns error when userId is missing", async () => {
    const result = await handlePeopleGetManager({}, mockRunCtx, {} as never);
    expect(result).toHaveProperty("error", "userId is required");
  });

  it("returns error for invalid userId", async () => {
    const result = await handlePeopleGetManager(
      { userId: "../../bad" },
      mockRunCtx,
      {} as never,
    );
    expect(result).toHaveProperty("error", "Invalid userId format");
  });

  it("returns null manager gracefully", async () => {
    const mockService = { getManager: () => Promise.resolve(null) } as never;
    const result = await handlePeopleGetManager({ userId: "user-1" }, mockRunCtx, mockService);
    expect(result).toHaveProperty("content");
    expect((result as { content: string }).content).toContain("No manager found");
  });
});

describe("people-list-team-members handler", () => {
  it("returns error when groupId is missing", async () => {
    const result = await handlePeopleListTeamMembers({}, mockRunCtx, {} as never);
    expect(result).toHaveProperty("error", "groupId is required");
  });

  it("returns error for invalid groupId", async () => {
    const result = await handlePeopleListTeamMembers(
      { groupId: "../traversal" },
      mockRunCtx,
      {} as never,
    );
    expect(result).toHaveProperty("error", "Invalid groupId format");
  });

  it("returns empty member list gracefully", async () => {
    const mockService = { listGroupMembers: () => Promise.resolve([]) } as never;
    const result = await handlePeopleListTeamMembers({ groupId: "group-1" }, mockRunCtx, mockService);
    expect(result).toHaveProperty("content", "No members found");
  });
});

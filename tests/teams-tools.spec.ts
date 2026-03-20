import { describe, expect, it } from "vitest";
import { handleTeamsPostMessage } from "../src/tools/teams-post-message.js";
import { handleTeamsReadChannel } from "../src/tools/teams-read-channel.js";
import { handleTeamsReplyThread } from "../src/tools/teams-reply-thread.js";
import { handleTeamsListChannels } from "../src/tools/teams-list-channels.js";
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

describe("teams-post-message handler", () => {
  it("returns error when content is missing", async () => {
    const result = await handleTeamsPostMessage(
      { channelId: "ch-1" },
      mockRunCtx,
      {} as never,
      identityService,
    );
    expect(result).toHaveProperty("error", "content is required");
  });

  it("returns error when channelId is invalid", async () => {
    const result = await handleTeamsPostMessage(
      { channelId: "../evil", content: "hi" },
      mockRunCtx,
      {} as never,
      identityService,
    );
    expect(result).toHaveProperty("error", "Invalid channelId format");
  });

  it("returns error when no identity can be resolved", async () => {
    const mockTeams = { defaultChannelId: "ch-1" } as never;
    const result = await handleTeamsPostMessage(
      { content: "hi" },
      mockRunCtx,
      mockTeams,
      noIdentityService,
    );
    expect(result).toHaveProperty("error", "No M365 user identity configured for this agent");
  });

  it("falls back to default channel when channelId is omitted", async () => {
    const mockTeams = { defaultChannelId: "" } as never;
    const result = await handleTeamsPostMessage(
      { content: "hi" },
      mockRunCtx,
      mockTeams,
      identityService,
    );
    expect(result).toHaveProperty("error", "channelId is required (no default channel configured)");
  });
});

describe("teams-read-channel handler", () => {
  it("returns error when channelId is missing", async () => {
    const result = await handleTeamsReadChannel({}, mockRunCtx, {} as never);
    expect(result).toHaveProperty("error", "channelId is required");
  });

  it("returns error when channelId is invalid", async () => {
    const result = await handleTeamsReadChannel(
      { channelId: "../../hack" },
      mockRunCtx,
      {} as never,
    );
    expect(result).toHaveProperty("error", "Invalid channelId format");
  });
});

describe("teams-reply-thread handler", () => {
  it("returns error when channelId is missing", async () => {
    const result = await handleTeamsReplyThread(
      { messageId: "m-1", content: "reply" },
      mockRunCtx,
      {} as never,
      identityService,
    );
    expect(result).toHaveProperty("error", "channelId is required");
  });

  it("returns error when messageId is missing", async () => {
    const result = await handleTeamsReplyThread(
      { channelId: "ch-1", content: "reply" },
      mockRunCtx,
      {} as never,
      identityService,
    );
    expect(result).toHaveProperty("error", "messageId is required");
  });

  it("returns error when content is missing", async () => {
    const result = await handleTeamsReplyThread(
      { channelId: "ch-1", messageId: "m-1" },
      mockRunCtx,
      {} as never,
      identityService,
    );
    expect(result).toHaveProperty("error", "content is required");
  });

  it("returns error for invalid channelId", async () => {
    const result = await handleTeamsReplyThread(
      { channelId: "../bad", messageId: "m-1", content: "hi" },
      mockRunCtx,
      {} as never,
      identityService,
    );
    expect(result).toHaveProperty("error", "Invalid channelId format");
  });

  it("returns error for invalid messageId", async () => {
    const result = await handleTeamsReplyThread(
      { channelId: "ch-1", messageId: "../../bad", content: "hi" },
      mockRunCtx,
      {} as never,
      identityService,
    );
    expect(result).toHaveProperty("error", "Invalid messageId format");
  });
});

describe("teams-list-channels handler", () => {
  it("propagates service errors", async () => {
    const mockTeams = {
      listChannels: () => Promise.reject(new Error("Graph 403")),
    } as never;
    const result = await handleTeamsListChannels({}, mockRunCtx, mockTeams);
    expect(result).toHaveProperty("error");
    expect((result as { error: string }).error).toContain("Graph 403");
  });
});

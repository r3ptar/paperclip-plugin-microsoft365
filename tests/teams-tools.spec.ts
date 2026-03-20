import { describe, expect, it } from "vitest";
import { handleTeamsReadChannel } from "../src/tools/teams-read-channel.js";
import { handleTeamsListChannels } from "../src/tools/teams-list-channels.js";
import type { ToolRunContext } from "@paperclipai/plugin-sdk";

const mockRunCtx = { agentId: "agent-1", companyId: "co-1" } as ToolRunContext;

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

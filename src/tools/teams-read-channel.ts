import type { ToolResult, ToolRunContext } from "@paperclipai/plugin-sdk";
import type { TeamsService } from "../services/teams.js";
import { isValidGraphId } from "../graph/validate-id.js";

export interface TeamsReadChannelParams {
  channelId: string;
  maxMessages?: number;
}

export async function handleTeamsReadChannel(
  params: unknown,
  _runCtx: ToolRunContext,
  teamsService: TeamsService,
): Promise<ToolResult> {
  const { channelId, maxMessages } = params as TeamsReadChannelParams;

  if (!channelId) {
    return { error: "channelId is required" };
  }
  if (!isValidGraphId(channelId)) {
    return { error: "Invalid channelId format" };
  }

  try {
    const messages = await teamsService.readMessages(channelId, maxMessages ?? 20);

    if (messages.length === 0) {
      return { content: "No messages found in channel", data: { messages: [] } };
    }

    const summary = messages
      .map((m) => {
        const from = m.from?.user?.displayName ?? m.from?.application?.displayName ?? "Unknown";
        const text = m.body.content.replace(/<[^>]+>/g, "").slice(0, 200);
        return `[${m.createdDateTime}] ${from}: ${text}`;
      })
      .join("\n");

    return {
      content: `${messages.length} messages:\n${summary}`,
      data: { messages },
    };
  } catch (err) {
    return {
      error: `Failed to read channel: ${err instanceof Error ? err.message : String(err)}`,
    };
  }
}

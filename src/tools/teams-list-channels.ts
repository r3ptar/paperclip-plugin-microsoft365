import type { ToolResult, ToolRunContext } from "@paperclipai/plugin-sdk";
import type { TeamsService } from "../services/teams.js";

export async function handleTeamsListChannels(
  _params: unknown,
  _runCtx: ToolRunContext,
  teamsService: TeamsService,
): Promise<ToolResult> {
  try {
    const channels = await teamsService.listChannels();

    if (channels.length === 0) {
      return { content: "No channels found", data: { channels: [] } };
    }

    const summary = channels
      .map((c) => `- ${c.displayName} (${c.id})${c.description ? `: ${c.description}` : ""}`)
      .join("\n");

    return {
      content: `${channels.length} channels:\n${summary}`,
      data: { channels },
    };
  } catch (err) {
    return {
      error: `Failed to list channels: ${err instanceof Error ? err.message : String(err)}`,
    };
  }
}

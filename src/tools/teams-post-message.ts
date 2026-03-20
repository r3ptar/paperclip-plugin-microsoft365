import type { ToolResult, ToolRunContext } from "@paperclipai/plugin-sdk";
import type { TeamsService } from "../services/teams.js";
import type { AgentIdentityService } from "../services/identity.js";
import { isValidGraphId } from "../graph/validate-id.js";

export interface TeamsPostMessageParams {
  channelId?: string;
  content: string;
  subject?: string;
}

export async function handleTeamsPostMessage(
  params: unknown,
  runCtx: ToolRunContext,
  teamsService: TeamsService,
  identityService: AgentIdentityService,
): Promise<ToolResult> {
  const { channelId, content, subject } = params as TeamsPostMessageParams;

  if (!content) {
    return { error: "content is required" };
  }

  const resolvedChannelId = channelId || teamsService.defaultChannelId;
  if (!resolvedChannelId) {
    return { error: "channelId is required (no default channel configured)" };
  }
  if (!isValidGraphId(resolvedChannelId)) {
    return { error: "Invalid channelId format" };
  }

  const actAsUserId = identityService.resolveActingUserId(runCtx.agentId);
  if (!actAsUserId) {
    return { error: "No M365 user identity configured for this agent" };
  }

  try {
    const message = await teamsService.postMessage(
      resolvedChannelId,
      content,
      subject,
    );

    return {
      content: `Message posted to Teams channel${subject ? ` with subject "${subject}"` : ""}`,
      data: { messageId: message.id, channelId: resolvedChannelId },
    };
  } catch (err) {
    return {
      error: `Failed to post Teams message: ${err instanceof Error ? err.message : String(err)}`,
    };
  }
}

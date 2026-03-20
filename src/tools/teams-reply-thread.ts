import type { ToolResult, ToolRunContext } from "@paperclipai/plugin-sdk";
import type { TeamsService } from "../services/teams.js";
import type { AgentIdentityService } from "../services/identity.js";
import { isValidGraphId } from "../graph/validate-id.js";

export interface TeamsReplyThreadParams {
  channelId: string;
  messageId: string;
  content: string;
}

export async function handleTeamsReplyThread(
  params: unknown,
  runCtx: ToolRunContext,
  teamsService: TeamsService,
  identityService: AgentIdentityService,
): Promise<ToolResult> {
  const { channelId, messageId, content } = params as TeamsReplyThreadParams;

  if (!channelId) return { error: "channelId is required" };
  if (!messageId) return { error: "messageId is required" };
  if (!content) return { error: "content is required" };
  if (!isValidGraphId(channelId)) return { error: "Invalid channelId format" };
  if (!isValidGraphId(messageId)) return { error: "Invalid messageId format" };

  const actAsUserId = identityService.resolveActingUserId(runCtx.agentId);
  if (!actAsUserId) {
    return { error: "No M365 user identity configured for this agent" };
  }

  try {
    const reply = await teamsService.replyToThread(channelId, messageId, content);

    return {
      content: `Reply posted to thread ${messageId}`,
      data: { replyId: reply.id, channelId, parentMessageId: messageId },
    };
  } catch (err) {
    return {
      error: `Failed to reply to thread: ${err instanceof Error ? err.message : String(err)}`,
    };
  }
}

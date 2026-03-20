import type { ToolResult, ToolRunContext } from "@paperclipai/plugin-sdk";
import type { MeetingService } from "../services/meetings.js";
import type { AgentIdentityService } from "../services/identity.js";
import { isValidGraphId } from "../graph/validate-id.js";

export interface MeetingCancelParams {
  eventId: string;
}

export async function handleMeetingCancel(
  params: unknown,
  runCtx: ToolRunContext,
  meetingService: MeetingService,
  identityService: AgentIdentityService,
): Promise<ToolResult> {
  const { eventId } = params as MeetingCancelParams;

  if (!eventId) {
    return { error: "eventId is required" };
  }
  if (!isValidGraphId(eventId)) {
    return { error: "Invalid eventId format" };
  }

  const userId = identityService.resolveActingUserId(runCtx.agentId);
  if (!userId) {
    return { error: "No M365 user identity configured for this agent" };
  }

  try {
    await meetingService.cancelMeeting(eventId, userId);

    return {
      content: `Meeting ${eventId} cancelled`,
      data: { eventId, cancelled: true },
    };
  } catch (err) {
    return {
      error: `Failed to cancel meeting: ${err instanceof Error ? err.message : String(err)}`,
    };
  }
}

import type { ToolResult, ToolRunContext } from "@paperclipai/plugin-sdk";
import type { MeetingService } from "../services/meetings.js";
import type { AgentIdentityService } from "../services/identity.js";

export interface MeetingListParams {
  startDateTime?: string;
  endDateTime?: string;
}

export async function handleMeetingList(
  params: unknown,
  runCtx: ToolRunContext,
  meetingService: MeetingService,
  identityService: AgentIdentityService,
): Promise<ToolResult> {
  const { startDateTime, endDateTime } = params as MeetingListParams;

  // Default to next 7 days if no range specified
  const start = startDateTime || new Date().toISOString();
  const end = endDateTime || new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();

  const userId = identityService.resolveActingUserId(runCtx.agentId);
  if (!userId) {
    return { error: "No M365 user identity configured for this agent" };
  }

  try {
    const meetings = await meetingService.listMeetings({
      startDateTime: start,
      endDateTime: end,
      userId,
    });

    if (meetings.length === 0) {
      return { content: "No upcoming meetings found", data: { meetings: [] } };
    }

    const summary = meetings
      .map((m) => {
        const attendeeCount = m.attendees?.length ?? 0;
        const online = m.isOnlineMeeting ? " [Teams]" : "";
        return `- ${m.subject} (${m.start.dateTime} — ${m.end.dateTime})${online} [${attendeeCount} attendees]`;
      })
      .join("\n");

    return {
      content: `${meetings.length} meetings:\n${summary}`,
      data: { meetings },
    };
  } catch (err) {
    return {
      error: `Failed to list meetings: ${err instanceof Error ? err.message : String(err)}`,
    };
  }
}

import type { ToolResult, ToolRunContext } from "@paperclipai/plugin-sdk";
import type { MeetingService } from "../services/meetings.js";
import type { AgentIdentityService } from "../services/identity.js";

export interface MeetingScheduleParams {
  subject: string;
  attendeeEmails: string[];
  startDateTime: string;
  endDateTime?: string;
  body?: string;
  createTeamsLink?: boolean;
}

export async function handleMeetingSchedule(
  params: unknown,
  runCtx: ToolRunContext,
  meetingService: MeetingService,
  identityService: AgentIdentityService,
): Promise<ToolResult> {
  const { subject, attendeeEmails, startDateTime, endDateTime, body, createTeamsLink } =
    params as MeetingScheduleParams;

  if (!subject) return { error: "subject is required" };
  if (!attendeeEmails || attendeeEmails.length === 0) return { error: "attendeeEmails is required" };
  if (!startDateTime) return { error: "startDateTime is required" };

  // Validate startDateTime is a parseable date
  if (isNaN(new Date(startDateTime).getTime())) {
    return { error: "startDateTime must be a valid ISO 8601 date string" };
  }
  if (endDateTime && isNaN(new Date(endDateTime).getTime())) {
    return { error: "endDateTime must be a valid ISO 8601 date string" };
  }

  const organizerUserId = identityService.resolveActingUserId(runCtx.agentId);
  if (!organizerUserId) {
    return { error: "No M365 user identity configured for this agent" };
  }

  try {
    const event = await meetingService.scheduleMeeting({
      subject,
      attendeeEmails,
      startDateTime,
      endDateTime,
      body,
      createTeamsLink,
      organizerUserId,
    });

    const joinUrl = event.onlineMeeting?.joinUrl;
    const summary = `Meeting "${subject}" scheduled` +
      (joinUrl ? `\nTeams link: ${joinUrl}` : "") +
      `\nAttendees: ${attendeeEmails.join(", ")}`;

    return {
      content: summary,
      data: {
        eventId: event.id,
        subject: event.subject,
        start: event.start,
        end: event.end,
        joinUrl,
        webLink: event.webLink,
      },
    };
  } catch (err) {
    return {
      error: `Failed to schedule meeting: ${err instanceof Error ? err.message : String(err)}`,
    };
  }
}

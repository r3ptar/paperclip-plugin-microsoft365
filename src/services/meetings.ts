import type { PluginContext } from "@paperclipai/plugin-sdk";
import type { M365Config } from "../constants.js";
import type { GraphClient } from "../graph/client.js";
import type {
  CalendarEventFull,
  FindMeetingTimesResponse,
  GraphListResponse,
} from "../graph/types.js";

export interface ScheduleMeetingParams {
  subject: string;
  attendeeEmails: string[];
  startDateTime: string;
  endDateTime?: string;
  body?: string;
  createTeamsLink?: boolean;
  organizerUserId?: string;
}

export interface FindMeetingTimesParams {
  attendeeEmails: string[];
  durationMinutes?: number;
  startRange?: string;
  endRange?: string;
}

/**
 * Meeting scheduling and calendar view via Graph API.
 */
export class MeetingService {
  constructor(
    private readonly ctx: PluginContext,
    private readonly graph: GraphClient,
    private readonly config: M365Config,
  ) {}

  /** Schedule a new meeting, optionally with a Teams online link. */
  async scheduleMeeting(params: ScheduleMeetingParams): Promise<CalendarEventFull> {
    const userId = params.organizerUserId || this.config.meetingOrganizerUserId;
    const durationMs = this.config.meetingDefaultDuration * 60 * 1000;

    const startDt = params.startDateTime;
    const endDt = params.endDateTime ?? new Date(new Date(startDt).getTime() + durationMs).toISOString();

    const eventBody: Record<string, unknown> = {
      subject: params.subject,
      start: { dateTime: startDt, timeZone: "UTC" },
      end: { dateTime: endDt, timeZone: "UTC" },
      attendees: params.attendeeEmails.map((email) => ({
        emailAddress: { address: email },
        type: "required",
      })),
    };

    if (params.body) {
      eventBody.body = { contentType: "text", content: params.body };
    }

    if (params.createTeamsLink !== false) {
      eventBody.isOnlineMeeting = true;
      eventBody.onlineMeetingProvider = "teamsForBusiness";
    }

    const event = await this.graph.post<CalendarEventFull>(
      `/users/${userId}/events`,
      eventBody,
    );

    this.ctx.logger.info("Scheduled meeting", {
      eventId: event.id,
      subject: params.subject,
      organizer: userId,
      attendees: params.attendeeEmails.length,
    });

    return event;
  }

  /** Find available meeting times for a set of attendees. */
  async findMeetingTimes(params: FindMeetingTimesParams): Promise<FindMeetingTimesResponse> {
    const userId = this.config.meetingOrganizerUserId;
    const durationMinutes = params.durationMinutes ?? this.config.meetingDefaultDuration;

    const body: Record<string, unknown> = {
      attendees: params.attendeeEmails.map((email) => ({
        emailAddress: { address: email },
        type: "required",
      })),
      meetingDuration: `PT${durationMinutes}M`,
      maxCandidates: 5,
    };

    if (params.startRange && params.endRange) {
      body.timeConstraint = {
        timeslots: [
          {
            start: { dateTime: params.startRange, timeZone: "UTC" },
            end: { dateTime: params.endRange, timeZone: "UTC" },
          },
        ],
      };
    }

    return this.graph.post<FindMeetingTimesResponse>(
      `/users/${userId}/findMeetingTimes`,
      body,
    );
  }

  /** Cancel (delete) a meeting. */
  async cancelMeeting(eventId: string, userId?: string): Promise<void> {
    const actingUser = userId || this.config.meetingOrganizerUserId;
    await this.graph.delete(`/users/${actingUser}/events/${eventId}`);

    this.ctx.logger.info("Cancelled meeting", { eventId, userId: actingUser });
  }

  /** List upcoming meetings in a date range. */
  async listMeetings(params: {
    startDateTime: string;
    endDateTime: string;
    userId?: string;
  }): Promise<CalendarEventFull[]> {
    const userId = params.userId || this.config.meetingOrganizerUserId;
    const start = encodeURIComponent(params.startDateTime);
    const end = encodeURIComponent(params.endDateTime);

    const res = await this.graph.get<GraphListResponse<CalendarEventFull>>(
      `/users/${userId}/calendarView?startDateTime=${start}&endDateTime=${end}&$select=id,subject,start,end,webLink,isOnlineMeeting,onlineMeeting,attendees,organizer&$top=50`,
    );
    return res.value;
  }
}

import type { ToolResult, ToolRunContext } from "@paperclipai/plugin-sdk";
import type { MeetingService } from "../services/meetings.js";

export interface MeetingFindTimeParams {
  attendeeEmails: string[];
  durationMinutes?: number;
  startRange?: string;
  endRange?: string;
}

export async function handleMeetingFindTime(
  params: unknown,
  _runCtx: ToolRunContext,
  meetingService: MeetingService,
): Promise<ToolResult> {
  const { attendeeEmails, durationMinutes, startRange, endRange } =
    params as MeetingFindTimeParams;

  if (!attendeeEmails || attendeeEmails.length === 0) {
    return { error: "attendeeEmails is required" };
  }

  try {
    const result = await meetingService.findMeetingTimes({
      attendeeEmails,
      durationMinutes,
      startRange,
      endRange,
    });

    if (result.meetingTimeSuggestions.length === 0) {
      return {
        content: `No available times found${result.emptySuggestionsReason ? `: ${result.emptySuggestionsReason}` : ""}`,
        data: result,
      };
    }

    const summary = result.meetingTimeSuggestions
      .map((s, i) => {
        const start = s.meetingTimeSlot.start.dateTime;
        const end = s.meetingTimeSlot.end.dateTime;
        return `${i + 1}. ${start} — ${end} (confidence: ${Math.round(s.confidence)}%)`;
      })
      .join("\n");

    return {
      content: `${result.meetingTimeSuggestions.length} suggested times:\n${summary}`,
      data: result,
    };
  } catch (err) {
    return {
      error: `Failed to find meeting times: ${err instanceof Error ? err.message : String(err)}`,
    };
  }
}

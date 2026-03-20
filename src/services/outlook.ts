import type { PluginContext } from "@paperclipai/plugin-sdk";
import type { Issue } from "@paperclipai/shared";
import { ENTITY_TYPES, type M365Config } from "../constants.js";
import type { GraphClient } from "../graph/client.js";
import type { CalendarEvent, GraphMessage } from "../graph/types.js";

export type TaskEmailType = "assignment" | "status_change" | "blocked" | "request" | "custom";

/**
 * Outlook calendar and email operations via Graph API.
 */
export class OutlookService {
  constructor(
    private readonly ctx: PluginContext,
    private readonly graph: GraphClient,
    private readonly config: M365Config,
  ) {}

  /**
   * Create a calendar event for an issue deadline.
   */
  async createDeadlineEvent(issue: Issue, dueDate: string): Promise<CalendarEvent> {
    const { outlookCalendarId, digestSenderUserId } = this.config;
    const userId = digestSenderUserId;

    // Graph API all-day events use exclusive end date (end = start + 1 day)
    const endDate = new Date(dueDate);
    endDate.setUTCDate(endDate.getUTCDate() + 1);
    const endDateStr = endDate.toISOString().split("T")[0]!;
    const startDateStr = dueDate.split("T")[0] ?? dueDate;

    const event = await this.graph.post<CalendarEvent>(
      `/users/${userId}/calendars/${outlookCalendarId}/events`,
      {
        subject: `[Paperclip] ${issue.title} — Deadline`,
        start: { dateTime: startDateStr, timeZone: "UTC" },
        end: { dateTime: endDateStr, timeZone: "UTC" },
        isAllDay: true,
        body: {
          contentType: "text",
          content: `Deadline for Paperclip issue: ${issue.title}\nStatus: ${issue.status}`,
        },
      },
    );

    await this.ctx.entities.upsert({
      entityType: ENTITY_TYPES.calendarEvent,
      scopeKind: "issue",
      scopeId: issue.id,
      externalId: event.id,
      title: event.subject,
      status: "active",
      data: {
        eventId: event.id,
        dueDate,
        calendarId: outlookCalendarId,
      },
    });

    this.ctx.logger.info("Created calendar event", {
      issueId: issue.id,
      eventId: event.id,
    });

    return event;
  }

  /**
   * Update or delete a calendar event when deadline changes.
   */
  async updateDeadlineEvent(
    issueId: string,
    eventId: string,
    dueDate: string | null,
  ): Promise<void> {
    const { outlookCalendarId, digestSenderUserId } = this.config;
    const userId = digestSenderUserId;

    if (!dueDate) {
      // Deadline removed — delete the event
      await this.graph.delete(
        `/users/${userId}/calendars/${outlookCalendarId}/events/${eventId}`,
      );
      this.ctx.logger.info("Deleted calendar event", { issueId, eventId });
      return;
    }

    await this.graph.patch(
      `/users/${userId}/calendars/${outlookCalendarId}/events/${eventId}`,
      {
        start: { dateTime: dueDate, timeZone: "UTC" },
        end: { dateTime: dueDate, timeZone: "UTC" },
      },
    );

    this.ctx.logger.info("Updated calendar event", { issueId, eventId, dueDate });
  }

  /**
   * Send an HTML email digest to configured recipients.
   */
  async sendDigest(subject: string, htmlBody: string, actAsUserId?: string): Promise<void> {
    const { digestRecipients } = this.config;
    const digestSenderUserId = actAsUserId || this.config.digestSenderUserId;

    if (digestRecipients.length === 0) {
      this.ctx.logger.warn("No digest recipients configured — skipping");
      return;
    }

    const message: GraphMessage = {
      subject,
      body: { contentType: "html", content: htmlBody },
      toRecipients: digestRecipients.map((address) => ({
        emailAddress: { address },
      })),
    };

    await this.graph.post(`/users/${digestSenderUserId}/sendMail`, {
      message,
      saveToSentItems: false,
    });

    try {
      await this.ctx.activity.log({
        companyId: "",
        message: `Sent Outlook digest to ${digestRecipients.length} recipients`,
        metadata: { recipients: digestRecipients.length, subject },
      });
    } catch {
      // Activity logging is best-effort
    }

    this.ctx.logger.info("Sent email digest", {
      recipients: digestRecipients.length,
    });
  }

  /**
   * Send a task-specific email to an individual recipient.
   *
   * Uses the configured `digestSenderUserId` as the sender and includes
   * a tracking header (`X-Paperclip-Issue-Id`) for threading/filtering.
   */
  async sendTaskEmail(
    issue: { id: string; title: string; status: string },
    recipientEmail: string,
    emailType: TaskEmailType,
    customMessage?: string,
    actAsUserId?: string,
  ): Promise<void> {
    const digestSenderUserId = actAsUserId || this.config.digestSenderUserId;

    const subject = this.buildTaskEmailSubject(issue, emailType);
    const htmlContent = this.buildTaskEmailHtml(issue, emailType, customMessage);

    const message: GraphMessage = {
      subject,
      body: { contentType: "HTML", content: htmlContent },
      toRecipients: [{ emailAddress: { address: recipientEmail } }],
      internetMessageHeaders: [
        { name: "X-Paperclip-Issue-Id", value: issue.id },
      ],
    };

    await this.graph.post(`/users/${digestSenderUserId}/sendMail`, {
      message,
      saveToSentItems: false,
    });

    try {
      await this.ctx.activity.log({
        companyId: "",
        message: `Sent ${emailType} email for issue ${issue.id} to ${recipientEmail}`,
        metadata: { issueId: issue.id, emailType, recipient: recipientEmail },
      });
    } catch {
      // Activity logging is best-effort
    }

    this.ctx.logger.info("Sent task email", {
      issueId: issue.id,
      emailType,
      recipient: recipientEmail,
    });
  }

  private buildTaskEmailSubject(
    issue: { id: string; title: string; status: string },
    emailType: TaskEmailType,
  ): string {
    const tag = `[PC-${issue.id}]`;

    switch (emailType) {
      case "assignment":
        return `${tag} You've been assigned: ${issue.title}`;
      case "status_change":
        return `${tag} Status updated to ${issue.status}: ${issue.title}`;
      case "blocked":
        return `${tag} Blocked: ${issue.title}`;
      case "request":
        return `${tag} Input requested: ${issue.title}`;
      case "custom":
        return `${tag} ${issue.title}`;
    }
  }

  private buildTaskEmailHtml(
    issue: { id: string; title: string; status: string },
    emailType: TaskEmailType,
    customMessage?: string,
  ): string {
    const contextLine = this.buildTaskEmailContextLine(emailType);

    const customSection = customMessage
      ? `<p style="margin:12px 0;padding:10px;background:#f5f5f5;border-left:3px solid #0078d4;">${escapeHtml(customMessage)}</p>`
      : "";

    return `
<html>
  <body style="font-family:Segoe UI,Helvetica,Arial,sans-serif;color:#333;max-width:600px;margin:0 auto;">
    <h2 style="color:#0078d4;margin-bottom:4px;">${escapeHtml(issue.title)}</h2>
    <p style="margin:4px 0 16px;color:#666;font-size:14px;">Status: <strong>${escapeHtml(issue.status)}</strong></p>
    <p style="margin:12px 0;font-size:15px;">${escapeHtml(contextLine)}</p>
    ${customSection}
    <hr style="border:none;border-top:1px solid #ddd;margin:24px 0 12px;" />
    <p style="font-size:12px;color:#888;"><em>Reply to this email to update this task in Paperclip.</em></p>
  </body>
</html>`.trim();
  }

  private buildTaskEmailContextLine(emailType: TaskEmailType): string {
    switch (emailType) {
      case "assignment":
        return "You have been assigned to this task.";
      case "status_change":
        return "The status of this task has been updated.";
      case "blocked":
        return "This task has been marked as blocked and needs attention.";
      case "request":
        return "Your input has been requested on this task.";
      case "custom":
        return "You have a notification about this task.";
    }
  }

  /**
   * Build a digest email body from recent activity.
   */
  buildDigestHtml(
    issues: Array<{ id: string; title: string; status: string; updatedAt: string }>,
  ): string {
    const rows = issues
      .map(
        (i) =>
          `<tr><td>${escapeHtml(i.title)}</td><td>${escapeHtml(i.status)}</td><td>${escapeHtml(i.updatedAt)}</td></tr>`,
      )
      .join("");

    return `
      <html>
        <body>
          <h2>Paperclip Daily Digest</h2>
          <p>Recent issue activity:</p>
          <table border="1" cellpadding="6" cellspacing="0">
            <thead><tr><th>Issue</th><th>Status</th><th>Updated</th></tr></thead>
            <tbody>${rows || "<tr><td colspan=\"3\">No recent activity</td></tr>"}</tbody>
          </table>
          <p><em>Sent by Paperclip Microsoft 365 Plugin</em></p>
        </body>
      </html>
    `.trim();
  }
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

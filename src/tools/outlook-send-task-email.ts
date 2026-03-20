import type { PluginContext, ToolResult, ToolRunContext } from "@paperclipai/plugin-sdk";
import type { OutlookService, TaskEmailType } from "../services/outlook.js";

export interface OutlookSendTaskEmailParams {
  issueId: string;
  recipientEmail: string;
  emailType?: TaskEmailType;
  customMessage?: string;
}

const VALID_EMAIL_TYPES: ReadonlySet<string> = new Set([
  "assignment",
  "status_change",
  "blocked",
  "request",
  "custom",
]);

export async function handleOutlookSendTaskEmail(
  params: unknown,
  runCtx: ToolRunContext,
  ctx: PluginContext,
  outlookService: OutlookService,
): Promise<ToolResult> {
  const { issueId, recipientEmail, emailType, customMessage } =
    params as OutlookSendTaskEmailParams;

  if (!issueId) {
    return { error: "issueId is required" };
  }
  if (!recipientEmail) {
    return { error: "recipientEmail is required" };
  }

  const resolvedEmailType: TaskEmailType =
    emailType && VALID_EMAIL_TYPES.has(emailType) ? emailType : "custom";

  // Fetch the issue to get title and status for the email template
  const issue = await ctx.issues.get(issueId, runCtx.companyId);
  if (!issue) {
    return { error: `Issue ${issueId} not found` };
  }

  try {
    await outlookService.sendTaskEmail(
      { id: issue.id, title: issue.title, status: issue.status },
      recipientEmail,
      resolvedEmailType,
      customMessage,
    );

    return {
      content: `Email sent to ${recipientEmail} about issue ${issueId}`,
      data: {
        issueId: issue.id,
        recipientEmail,
        emailType: resolvedEmailType,
      },
    };
  } catch (err) {
    return {
      error: `Failed to send email: ${err instanceof Error ? err.message : String(err)}`,
    };
  }
}

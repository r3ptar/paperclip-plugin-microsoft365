import type { PluginContext } from "@paperclipai/plugin-sdk";
import type { M365Config } from "../constants.js";
import type { GraphClient } from "../graph/client.js";
import type {
  GraphListResponse,
  TeamsChannel,
  TeamsChannelMessage,
} from "../graph/types.js";

/**
 * Teams channel messaging via Graph API.
 */
export class TeamsService {
  private readonly teamId: string;
  private readonly _defaultChannelId: string;

  constructor(
    private readonly ctx: PluginContext,
    private readonly graph: GraphClient,
    config: M365Config,
  ) {
    this.teamId = config.teamsTeamId;
    this._defaultChannelId = config.teamsDefaultChannelId;
  }

  /** Get the default channel ID for fallback. */
  get defaultChannelId(): string {
    return this._defaultChannelId;
  }

  /** Post a message to a Teams channel. */
  async postMessage(
    channelId: string,
    content: string,
    subject?: string,
  ): Promise<TeamsChannelMessage> {
    const body: Record<string, unknown> = {
      body: { contentType: "html", content },
    };
    if (subject) body.subject = subject;

    // Application permissions post as the app. Per-user delegation requires
    // delegated auth which is not supported in the client-credentials flow.
    const message = await this.graph.post<TeamsChannelMessage>(
      `/teams/${this.teamId}/channels/${channelId}/messages`,
      body,
    );

    this.ctx.logger.info("Posted Teams message", {
      channelId,
      messageId: message.id,
    });

    return message;
  }

  /** Read recent messages from a Teams channel. */
  async readMessages(
    channelId: string,
    top = 20,
  ): Promise<TeamsChannelMessage[]> {
    const clamped = Math.min(Math.max(top, 1), 50);
    const res = await this.graph.get<GraphListResponse<TeamsChannelMessage>>(
      `/teams/${this.teamId}/channels/${channelId}/messages?$top=${clamped}`,
    );
    return res.value;
  }

  /** Reply to a specific thread in a Teams channel. */
  async replyToThread(
    channelId: string,
    messageId: string,
    content: string,
  ): Promise<TeamsChannelMessage> {
    const reply = await this.graph.post<TeamsChannelMessage>(
      `/teams/${this.teamId}/channels/${channelId}/messages/${messageId}/replies`,
      { body: { contentType: "html", content } },
    );

    this.ctx.logger.info("Posted Teams reply", {
      channelId,
      parentMessageId: messageId,
      replyId: reply.id,
    });

    return reply;
  }

  /** List all channels in the configured team. */
  async listChannels(): Promise<TeamsChannel[]> {
    const res = await this.graph.get<GraphListResponse<TeamsChannel>>(
      `/teams/${this.teamId}/channels?$select=id,displayName,description,membershipType`,
    );
    return res.value;
  }

  /** Post a formatted issue update notification to a channel. */
  async postIssueUpdate(
    channelId: string,
    issue: { id: string; title: string; status: string },
    changeType: "created" | "updated",
  ): Promise<void> {
    const verb = changeType === "created" ? "New issue created" : "Issue updated";
    const content = `<strong>${verb}</strong><br/>` +
      `<b>${escapeHtml(issue.title)}</b><br/>` +
      `Status: ${escapeHtml(issue.status)}<br/>` +
      `<em>Issue ID: ${escapeHtml(issue.id)}</em>`;

    await this.postMessage(channelId, content, verb);
  }
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

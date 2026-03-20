import type { PluginContext } from "@paperclipai/plugin-sdk";
import type { M365Config } from "../constants.js";
import type { GraphClient } from "../graph/client.js";
import type {
  GraphListResponse,
  TeamsChannel,
  TeamsChannelMessage,
} from "../graph/types.js";

/**
 * Teams channel read operations via Graph API (app-only tokens).
 *
 * Note: Posting messages requires delegated permissions which are not
 * supported in the client-credentials flow. This service is read-only.
 */
export class TeamsService {
  private readonly teamId: string;

  constructor(
    private readonly ctx: PluginContext,
    private readonly graph: GraphClient,
    config: M365Config,
  ) {
    this.teamId = config.teamsTeamId;
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

  /** List all channels in the configured team. */
  async listChannels(): Promise<TeamsChannel[]> {
    const res = await this.graph.get<GraphListResponse<TeamsChannel>>(
      `/teams/${this.teamId}/channels?$select=id,displayName,description,membershipType`,
    );
    return res.value;
  }
}

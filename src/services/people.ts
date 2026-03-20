import type { PluginContext } from "@paperclipai/plugin-sdk";
import type { GraphClient } from "../graph/client.js";
import type {
  GraphListResponse,
  GraphUser,
  GraphPresence,
  GraphGroupMember,
} from "../graph/types.js";

/**
 * People directory and presence lookups via Graph API.
 */
export class PeopleService {
  constructor(
    private readonly ctx: PluginContext,
    private readonly graph: GraphClient,
  ) {}

  /** Search for users by name, email, or department. */
  async lookupUser(query: string): Promise<GraphUser[]> {
    // Escape double-quotes to prevent OData $search injection
    const sanitized = query.replace(/"/g, '\\"');
    const encoded = encodeURIComponent(`"displayName:${sanitized}" OR "mail:${sanitized}" OR "department:${sanitized}"`);
    const res = await this.graph.get<GraphListResponse<GraphUser>>(
      `/users?$search=${encoded}&$select=id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation&$top=10`,
      {
        headers: { ConsistencyLevel: "eventual" },
      },
    );
    return res.value;
  }

  /** Get presence/availability for a single user. */
  async getPresence(userId: string): Promise<GraphPresence> {
    return this.graph.get<GraphPresence>(`/users/${userId}/presence`);
  }

  /** Get presence for multiple users in a batch. */
  async getPresences(userIds: string[]): Promise<GraphPresence[]> {
    const res = await this.graph.post<GraphListResponse<GraphPresence>>(
      "/communications/getPresencesByUserId",
      { ids: userIds },
    );
    return res.value;
  }

  /** Get a user's manager. Returns null if no manager is set. */
  async getManager(userId: string): Promise<GraphUser | null> {
    try {
      return await this.graph.get<GraphUser>(
        `/users/${userId}/manager?$select=id,displayName,mail,userPrincipalName,jobTitle,department`,
      );
    } catch (err: unknown) {
      // 404 means no manager is assigned
      if (err instanceof Error && err.message.includes("404")) {
        return null;
      }
      throw err;
    }
  }

  /** List members of a group (team). */
  async listGroupMembers(groupId: string): Promise<GraphGroupMember[]> {
    const res = await this.graph.get<GraphListResponse<GraphGroupMember>>(
      `/groups/${groupId}/members?$select=id,displayName,mail,userPrincipalName,jobTitle&$top=100`,
    );
    return res.value;
  }
}

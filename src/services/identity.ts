import type { M365Config } from "../constants.js";

/**
 * Resolves which M365 user account an agent should act as.
 *
 * Each Paperclip agent can be mapped to a dedicated M365 user (e.g., ceo-agent@contoso.com).
 * When no mapping exists, falls back to the default service user ID.
 */
export class AgentIdentityService {
  private readonly map: ReadonlyMap<string, string>;
  private readonly defaultUserId: string;

  constructor(config: M365Config) {
    this.map = new Map(Object.entries(config.agentIdentityMap));
    this.defaultUserId = config.defaultServiceUserId;
  }

  /** Resolve which M365 user a Paperclip agent should act as. */
  resolveUserId(agentId: string): string | null {
    return this.map.get(agentId) ?? null;
  }

  /** Get the fallback service user for non-agent operations. */
  getDefaultUserId(): string {
    return this.defaultUserId;
  }

  /** Check if an agent has a mapped identity. */
  hasIdentity(agentId: string): boolean {
    return this.map.has(agentId);
  }

  /** List all mapped agents. */
  listMappings(): Array<{ agentId: string; m365UserId: string }> {
    return Array.from(this.map.entries()).map(([agentId, m365UserId]) => ({
      agentId,
      m365UserId,
    }));
  }

  /**
   * Resolve the acting user for a tool call.
   * Returns the mapped user for the agent, or the default service user as fallback.
   * Returns null if no identity can be resolved (no mapping and no default).
   */
  resolveActingUserId(agentId?: string): string | null {
    if (agentId) {
      const mapped = this.resolveUserId(agentId);
      if (mapped) return mapped;
    }
    return this.defaultUserId || null;
  }
}

import type { PluginContext } from "@paperclipai/plugin-sdk";
import { GRAPH_SCOPE, OAUTH_TOKEN_URL } from "../constants.js";
import type { GraphTokenResponse } from "./types.js";

/**
 * Manages OAuth 2.0 Client Credentials tokens for Microsoft Graph.
 * Tokens are cached in-memory only and refreshed proactively before expiry.
 */
export class TokenManager {
  private accessToken: string | null = null;
  private expiresAt = 0;
  private refreshPromise: Promise<string> | null = null;

  constructor(
    private readonly ctx: PluginContext,
    private readonly tenantId: string,
    private readonly clientId: string,
    private readonly clientSecretRef: string,
    private readonly rawSecret?: string,
  ) {}

  async getToken(): Promise<string> {
    // Return cached token if still valid (with 5-min buffer)
    if (this.accessToken && Date.now() < this.expiresAt - 5 * 60 * 1000) {
      return this.accessToken;
    }
    // Deduplicate concurrent refresh requests
    if (this.refreshPromise) {
      return this.refreshPromise;
    }
    this.refreshPromise = this.acquireToken();
    try {
      return await this.refreshPromise;
    } finally {
      this.refreshPromise = null;
    }
  }

  /** Force a token refresh (e.g., after a 401). */
  async forceRefresh(): Promise<string> {
    this.accessToken = null;
    this.expiresAt = 0;
    this.refreshPromise = null;
    return this.getToken();
  }

  /** Check whether credentials can obtain a valid token. */
  async healthCheck(): Promise<{ ok: boolean; error?: string }> {
    try {
      await this.forceRefresh();
      return { ok: true };
    } catch (err) {
      return { ok: false, error: err instanceof Error ? err.message : String(err) };
    }
  }

  private async acquireToken(): Promise<string> {
    const clientSecret = this.rawSecret
      ? this.rawSecret
      : await this.ctx.secrets.resolve(this.clientSecretRef);
    const tokenUrl = OAUTH_TOKEN_URL.replace("{tenantId}", this.tenantId);

    const body = new URLSearchParams({
      grant_type: "client_credentials",
      client_id: this.clientId,
      client_secret: clientSecret,
      scope: GRAPH_SCOPE,
    });

    const response = await this.ctx.http.fetch(tokenUrl, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    });

    if (!response.ok) {
      const text = await response.text();
      this.ctx.logger.error("OAuth token acquisition failed", {
        status: response.status,
        body: text.slice(0, 500),
      });
      throw new Error(`OAuth token request failed: ${response.status}`);
    }

    const data = (await response.json()) as GraphTokenResponse;
    this.accessToken = data.access_token;
    this.expiresAt = Date.now() + data.expires_in * 1000;

    this.ctx.logger.debug("OAuth token acquired", {
      expiresIn: data.expires_in,
    });

    return this.accessToken;
  }
}

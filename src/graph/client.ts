import type { PluginContext } from "@paperclipai/plugin-sdk";
import {
  CIRCUIT_BREAKER_COOLDOWN_MS,
  CIRCUIT_BREAKER_THRESHOLD,
  GRAPH_BASE_URL,
} from "../constants.js";
import type { TokenManager } from "./auth.js";
import type { GraphError } from "./types.js";

export interface GraphRequestOptions {
  method?: string;
  headers?: Record<string, string>;
  body?: string;
  /** Skip audit logging for this call (e.g., health checks). */
  silent?: boolean;
}

interface CircuitState {
  failures: number;
  openUntil: number;
}

/**
 * Wrapper around Microsoft Graph API that handles:
 * - Bearer token injection
 * - 429 rate-limit backoff with Retry-After
 * - 401 automatic token refresh (once)
 * - Circuit breaker (5 consecutive failures -> 5-min pause)
 * - Audit logging via ctx.activity.log()
 */
export class GraphClient {
  private circuit: CircuitState = { failures: 0, openUntil: 0 };

  /** Optional company ID for activity logging. May be null during wizard setup. */
  companyId: string | null = null;

  constructor(
    private readonly ctx: PluginContext,
    private readonly tokenManager: TokenManager,
    private readonly serviceName: string,
  ) {}

  async request<T>(path: string, options: GraphRequestOptions = {}): Promise<T> {
    this.checkCircuit();

    const method = options.method ?? "GET";
    let token = await this.tokenManager.getToken();

    const doFetch = async (authToken: string): Promise<Response> => {
      const headers: Record<string, string> = {
        ...options.headers,
        Authorization: `Bearer ${authToken}`,
        "Content-Type": options.headers?.["Content-Type"] ?? "application/json",
      };
      const url = path.startsWith("http") ? path : `${GRAPH_BASE_URL}${path}`;
      return this.ctx.http.fetch(url, {
        method,
        headers,
        body: options.body,
      });
    };

    if (!options.silent) {
      try {
        await this.ctx.activity.log({
          companyId: this.companyId ?? "",
          message: `Graph API ${method} ${path}`,
          metadata: { service: this.serviceName, method, path },
        });
      } catch {
        // Activity logging is best-effort — skip if companyId is unavailable
        // (e.g., during Setup Wizard before config is saved)
      }
    }

    let response = await doFetch(token);

    // Handle 401 — refresh token once and retry
    if (response.status === 401) {
      this.ctx.logger.warn("Graph 401 — refreshing token", { path });
      token = await this.tokenManager.forceRefresh();
      response = await doFetch(token);
    }

    // Handle 429 — respect Retry-After (capped at 120s)
    if (response.status === 429) {
      const rawRetryAfter = Number(response.headers.get("Retry-After") ?? "10");
      const retryAfter = Math.min(Math.max(rawRetryAfter, 1), 120);
      this.ctx.logger.warn("Graph 429 — backing off", { path, retryAfter });
      await this.sleep(retryAfter * 1000);
      response = await doFetch(token);
    }

    if (!response.ok) {
      this.recordFailure();
      const errorBody = await response.text();
      let graphError: GraphError | undefined;
      try {
        graphError = JSON.parse(errorBody) as GraphError;
      } catch {
        // not JSON
      }
      const message = graphError?.error?.message ?? errorBody.slice(0, 500);
      throw new GraphApiError(response.status, message, path);
    }

    this.recordSuccess();

    // 204 No Content
    if (response.status === 204) {
      return undefined as T;
    }

    // Some Graph endpoints (e.g., sendMail) return 202 with no body
    const text = await response.text();
    if (!text) return undefined as T;
    return JSON.parse(text) as T;
  }

  async get<T>(path: string, options?: GraphRequestOptions): Promise<T> {
    return this.request<T>(path, { ...options, method: "GET" });
  }

  async post<T>(path: string, body: unknown, options?: GraphRequestOptions): Promise<T> {
    return this.request<T>(path, {
      ...options,
      method: "POST",
      body: JSON.stringify(body),
    });
  }

  async patch<T>(path: string, body: unknown, options?: GraphRequestOptions): Promise<T> {
    return this.request<T>(path, {
      ...options,
      method: "PATCH",
      body: JSON.stringify(body),
    });
  }

  async delete(path: string, options?: GraphRequestOptions): Promise<void> {
    await this.request<void>(path, { ...options, method: "DELETE" });
  }

  /**
   * Fetch raw response text (for binary/content downloads like /content endpoints).
   */
  async requestRaw(path: string, options: GraphRequestOptions = {}): Promise<string> {
    this.checkCircuit();

    const method = options.method ?? "GET";
    const token = await this.tokenManager.getToken();
    const url = path.startsWith("http") ? path : `${GRAPH_BASE_URL}${path}`;

    const response = await this.ctx.http.fetch(url, {
      method,
      headers: {
        ...options.headers,
        Authorization: `Bearer ${token}`,
      },
      body: options.body,
    });

    if (!response.ok) {
      this.recordFailure();
      throw new GraphApiError(response.status, `Raw request failed`, path);
    }

    this.recordSuccess();
    return response.text();
  }

  /**
   * Paginate through all pages of a Graph API list response.
   */
  async listAll<T>(path: string, options?: GraphRequestOptions, maxPages = 50): Promise<T[]> {
    type PageResponse = { value: T[]; "@odata.nextLink"?: string };
    const results: T[] = [];
    let url: string | undefined = path;
    let pageCount = 0;

    while (url && pageCount < maxPages) {
      pageCount += 1;
      const page: PageResponse = await this.get<PageResponse>(url, options);
      results.push(...page.value);
      url = page["@odata.nextLink"];
    }

    if (url) {
      this.ctx.logger.warn("listAll reached page limit — results may be incomplete", {
        path,
        maxPages,
        fetched: results.length,
      });
    }

    return results;
  }

  private checkCircuit(): void {
    if (this.circuit.failures >= CIRCUIT_BREAKER_THRESHOLD) {
      if (Date.now() < this.circuit.openUntil) {
        throw new CircuitBreakerOpenError(this.serviceName, this.circuit.openUntil);
      }
      // Cooldown expired, reset
      this.circuit.failures = 0;
      this.circuit.openUntil = 0;
    }
  }

  private recordFailure(): void {
    this.circuit.failures += 1;
    if (this.circuit.failures >= CIRCUIT_BREAKER_THRESHOLD) {
      this.circuit.openUntil = Date.now() + CIRCUIT_BREAKER_COOLDOWN_MS;
      this.ctx.logger.error("Circuit breaker opened", {
        service: this.serviceName,
        failures: this.circuit.failures,
        cooldownMs: CIRCUIT_BREAKER_COOLDOWN_MS,
      });
    }
  }

  private recordSuccess(): void {
    if (this.circuit.failures > 0) {
      this.circuit.failures = 0;
      this.circuit.openUntil = 0;
    }
  }

  private sleep(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }
}

export class GraphApiError extends Error {
  constructor(
    public readonly status: number,
    message: string,
    public readonly path: string,
  ) {
    super(`Graph API error ${status} on ${path}: ${message}`);
    this.name = "GraphApiError";
  }
}

export class CircuitBreakerOpenError extends Error {
  constructor(
    public readonly service: string,
    public readonly openUntil: number,
  ) {
    super(`Circuit breaker open for ${service} until ${new Date(openUntil).toISOString()}`);
    this.name = "CircuitBreakerOpenError";
  }
}

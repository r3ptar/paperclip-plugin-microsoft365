import { describe, expect, it, vi, beforeEach } from "vitest";
import { createTestHarness } from "@paperclipai/plugin-sdk/testing";
import manifest from "../src/manifest.js";
import plugin from "../src/worker.js";

/**
 * Tests for the 6 Setup Wizard data handlers registered in registerDataHandlers():
 *   m365-groups, m365-plans, m365-sites, m365-drives, m365-folders, m365-calendars
 *
 * Each handler:
 *  - Returns { error: "Azure AD credentials not configured" } when configClient is null
 *  - Returns { error: "...Id is required" } for parameterized handlers when the param is missing
 *  - Returns { items: Array<{ id, name }> } on success
 *  - Returns { error: string } if the Graph API call fails
 */

// ── Helpers ────────────────────────────────────────────────────────────────────

/** Harness without Azure AD credentials — configClient will be null. */
function createUnconfiguredHarness() {
  return createTestHarness({
    manifest,
    capabilities: [...manifest.capabilities, "events.emit"],
    config: {
      tenantId: "",
      clientId: "",
      clientSecretRef: "",
    },
  });
}

/** Harness with valid Azure AD credentials — configClient will be initialised. */
function createConfiguredHarness() {
  return createTestHarness({
    manifest,
    capabilities: [...manifest.capabilities, "events.emit"],
    config: {
      tenantId: "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
      clientId: "test-client-id",
      clientSecretRef: "secret:m365-client-secret",
    },
  });
}

/**
 * Build a minimal successful Graph API list response.
 * The harness delegates ctx.http.fetch to globalThis.fetch, so we stub it.
 */
function mockFetchJson(body: unknown): ReturnType<typeof vi.fn> {
  return vi.fn().mockResolvedValue(
    new Response(JSON.stringify(body), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    }),
  );
}

/**
 * Build a failing Graph API response.
 */
function mockFetchError(status: number, message: string): ReturnType<typeof vi.fn> {
  return vi.fn().mockResolvedValue(
    new Response(JSON.stringify({ error: { code: "BadRequest", message } }), {
      status,
      headers: { "Content-Type": "application/json" },
    }),
  );
}

// ── Tests ──────────────────────────────────────────────────────────────────────

describe("wizard data handlers", () => {
  const ALL_HANDLERS = [
    "m365-groups",
    "m365-plans",
    "m365-sites",
    "m365-drives",
    "m365-folders",
    "m365-calendars",
  ] as const;

  // ── 1. Error when credentials are not configured ──────────────────────────

  describe("credentials not configured (configClient is null)", () => {
    for (const handlerKey of ALL_HANDLERS) {
      it(`${handlerKey} returns credential error`, async () => {
        const harness = createUnconfiguredHarness();
        await plugin.definition.setup(harness.ctx);

        const result = await harness.getData<{ error: string }>(handlerKey, {});
        expect(result).toEqual({ error: "Azure AD credentials not configured" });
      });
    }
  });

  // ── 2. Parameter validation for parameterized handlers ────────────────────

  describe("missing required parameters", () => {
    // For these tests we need configClient to be non-null so that the handler
    // proceeds past the credentials check and reaches the parameter validation.
    // We mock globalThis.fetch to return a token for the TokenManager, since
    // initServices will create a TokenManager + GraphClient when credentials
    // are present.

    // Token endpoint response used during setup / first token acquisition
    const tokenResponse = {
      access_token: "fake-token",
      token_type: "Bearer",
      expires_in: 3600,
    };

    const parameterizedHandlers: Array<{
      key: string;
      paramName: string;
      errorMessage: string;
    }> = [
      { key: "m365-plans", paramName: "groupId", errorMessage: "groupId is required" },
      { key: "m365-drives", paramName: "siteId", errorMessage: "siteId is required" },
      { key: "m365-folders", paramName: "driveId", errorMessage: "driveId is required" },
      { key: "m365-calendars", paramName: "userId", errorMessage: "userId is required" },
    ];

    for (const { key, paramName, errorMessage } of parameterizedHandlers) {
      it(`${key} returns error when ${paramName} is missing`, async () => {
        const harness = createConfiguredHarness();

        // Stub fetch so TokenManager can acquire a token and activity.log calls work
        const originalFetch = globalThis.fetch;
        globalThis.fetch = mockFetchJson(tokenResponse);
        try {
          await plugin.definition.setup(harness.ctx);

          // Call with empty params — param is missing entirely
          const result = await harness.getData<{ error: string }>(key, {});
          expect(result).toEqual({ error: errorMessage });
        } finally {
          globalThis.fetch = originalFetch;
        }
      });

      it(`${key} returns error when ${paramName} is empty string`, async () => {
        const harness = createConfiguredHarness();

        const originalFetch = globalThis.fetch;
        globalThis.fetch = mockFetchJson(tokenResponse);
        try {
          await plugin.definition.setup(harness.ctx);

          // Call with the param set to an empty string
          const result = await harness.getData<{ error: string }>(key, {
            [paramName]: "",
          });
          expect(result).toEqual({ error: errorMessage });
        } finally {
          globalThis.fetch = originalFetch;
        }
      });

      it(`${key} returns error when ${paramName} is non-string type`, async () => {
        const harness = createConfiguredHarness();

        const originalFetch = globalThis.fetch;
        globalThis.fetch = mockFetchJson(tokenResponse);
        try {
          await plugin.definition.setup(harness.ctx);

          // Call with the param set to a non-string value
          const result = await harness.getData<{ error: string }>(key, {
            [paramName]: 12345,
          });
          expect(result).toEqual({ error: errorMessage });
        } finally {
          globalThis.fetch = originalFetch;
        }
      });
    }
  });

  // ── 3. Successful response shape validation ───────────────────────────────

  describe("successful Graph API responses", () => {
    const tokenResponse = {
      access_token: "fake-token",
      token_type: "Bearer",
      expires_in: 3600,
    };

    /**
     * Helper: set up the plugin with credentials, then replace globalThis.fetch
     * with a mock that returns a token for the initial auth call and then the
     * specified Graph API body for subsequent calls.
     *
     * Because the harness's ctx.http.fetch delegates to globalThis.fetch, and
     * GraphClient prepends the Graph base URL, we can intercept all calls.
     */
    async function setupWithGraphResponse(graphBody: unknown) {
      const harness = createConfiguredHarness();

      // During setup, TokenManager may not acquire a token yet (lazy), but
      // initServices creates TokenManager + GraphClient. We still need fetch
      // stubbed for activity.log calls that go through http.fetch — actually
      // activity.log does NOT use http.fetch. But let's keep it safe.
      const originalFetch = globalThis.fetch;

      // We use a sequential mock: first calls are token acquisition,
      // subsequent calls are Graph API calls.
      const fetchMock = vi.fn().mockImplementation(async (url: string) => {
        if (typeof url === "string" && url.includes("oauth2/v2.0/token")) {
          return new Response(JSON.stringify(tokenResponse), {
            status: 200,
            headers: { "Content-Type": "application/json" },
          });
        }
        // Graph API call
        return new Response(JSON.stringify(graphBody), {
          status: 200,
          headers: { "Content-Type": "application/json" },
        });
      });

      globalThis.fetch = fetchMock;

      await plugin.definition.setup(harness.ctx);

      return { harness, fetchMock, originalFetch };
    }

    it("m365-groups returns items with { id, name } from displayName", async () => {
      const graphResponse = {
        value: [
          { id: "grp-1", displayName: "Engineering", groupTypes: ["Unified"] },
          { id: "grp-2", displayName: "Marketing", groupTypes: ["Unified"] },
        ],
      };
      const { harness, originalFetch } = await setupWithGraphResponse(graphResponse);
      try {
        const result = await harness.getData<{ items: Array<{ id: string; name: string }> }>(
          "m365-groups",
        );

        expect(result.items).toBeDefined();
        expect(result.items).toHaveLength(2);
        expect(result.items[0]).toEqual({ id: "grp-1", name: "Engineering" });
        expect(result.items[1]).toEqual({ id: "grp-2", name: "Marketing" });
      } finally {
        globalThis.fetch = originalFetch;
      }
    });

    it("m365-plans returns items with { id, name } from title", async () => {
      const graphResponse = {
        value: [
          { id: "plan-1", title: "Sprint Board", owner: "grp-1" },
          { id: "plan-2", title: "Backlog", owner: "grp-1" },
        ],
      };
      const { harness, originalFetch } = await setupWithGraphResponse(graphResponse);
      try {
        const result = await harness.getData<{ items: Array<{ id: string; name: string }> }>(
          "m365-plans",
          { groupId: "grp-1" },
        );

        expect(result.items).toBeDefined();
        expect(result.items).toHaveLength(2);
        expect(result.items[0]).toEqual({ id: "plan-1", name: "Sprint Board" });
        expect(result.items[1]).toEqual({ id: "plan-2", name: "Backlog" });
      } finally {
        globalThis.fetch = originalFetch;
      }
    });

    it("m365-sites returns items with { id, name } from displayName", async () => {
      const graphResponse = {
        value: [
          { id: "site-1", displayName: "Team Site", webUrl: "https://example.sharepoint.com/sites/team" },
        ],
      };
      const { harness, originalFetch } = await setupWithGraphResponse(graphResponse);
      try {
        const result = await harness.getData<{ items: Array<{ id: string; name: string }> }>(
          "m365-sites",
        );

        expect(result.items).toBeDefined();
        expect(result.items).toHaveLength(1);
        expect(result.items[0]).toEqual({ id: "site-1", name: "Team Site" });
      } finally {
        globalThis.fetch = originalFetch;
      }
    });

    it("m365-drives returns items with { id, name } from name", async () => {
      const graphResponse = {
        value: [
          { id: "drive-1", name: "Documents", driveType: "documentLibrary" },
          { id: "drive-2", name: "Site Assets", driveType: "documentLibrary" },
        ],
      };
      const { harness, originalFetch } = await setupWithGraphResponse(graphResponse);
      try {
        const result = await harness.getData<{ items: Array<{ id: string; name: string }> }>(
          "m365-drives",
          { siteId: "site-1" },
        );

        expect(result.items).toBeDefined();
        expect(result.items).toHaveLength(2);
        expect(result.items[0]).toEqual({ id: "drive-1", name: "Documents" });
        expect(result.items[1]).toEqual({ id: "drive-2", name: "Site Assets" });
      } finally {
        globalThis.fetch = originalFetch;
      }
    });

    it("m365-folders returns only items with folder property", async () => {
      const graphResponse = {
        value: [
          { id: "folder-1", name: "Reports", folder: { childCount: 5 }, size: 0, webUrl: "", lastModifiedDateTime: "" },
          { id: "file-1", name: "readme.txt", file: { mimeType: "text/plain" }, size: 100, webUrl: "", lastModifiedDateTime: "" },
          { id: "folder-2", name: "Archives", folder: { childCount: 0 }, size: 0, webUrl: "", lastModifiedDateTime: "" },
        ],
      };
      const { harness, originalFetch } = await setupWithGraphResponse(graphResponse);
      try {
        const result = await harness.getData<{ items: Array<{ id: string; name: string }> }>(
          "m365-folders",
          { driveId: "drive-1" },
        );

        expect(result.items).toBeDefined();
        // The file item should be filtered out — only folders remain
        expect(result.items).toHaveLength(2);
        expect(result.items[0]).toEqual({ id: "folder-1", name: "Reports" });
        expect(result.items[1]).toEqual({ id: "folder-2", name: "Archives" });
      } finally {
        globalThis.fetch = originalFetch;
      }
    });

    it("m365-calendars returns items with { id, name } from calendar name", async () => {
      const graphResponse = {
        value: [
          { id: "cal-1", name: "Calendar", isDefaultCalendar: true },
          { id: "cal-2", name: "Team Meetings", isDefaultCalendar: false },
        ],
      };
      const { harness, originalFetch } = await setupWithGraphResponse(graphResponse);
      try {
        const result = await harness.getData<{ items: Array<{ id: string; name: string }> }>(
          "m365-calendars",
          { userId: "user-1" },
        );

        expect(result.items).toBeDefined();
        expect(result.items).toHaveLength(2);
        expect(result.items[0]).toEqual({ id: "cal-1", name: "Calendar" });
        expect(result.items[1]).toEqual({ id: "cal-2", name: "Team Meetings" });
      } finally {
        globalThis.fetch = originalFetch;
      }
    });

    it("returns empty items array when Graph API returns no results", async () => {
      const graphResponse = { value: [] };
      const { harness, originalFetch } = await setupWithGraphResponse(graphResponse);
      try {
        const result = await harness.getData<{ items: Array<{ id: string; name: string }> }>(
          "m365-groups",
        );

        expect(result.items).toBeDefined();
        expect(result.items).toHaveLength(0);
        expect(result.items).toEqual([]);
      } finally {
        globalThis.fetch = originalFetch;
      }
    });
  });

  // ── 4. Graph API error propagation ────────────────────────────────────────

  describe("Graph API errors", () => {
    const tokenResponse = {
      access_token: "fake-token",
      token_type: "Bearer",
      expires_in: 3600,
    };

    async function setupWithGraphError(status: number, errorMessage: string) {
      const harness = createConfiguredHarness();
      const originalFetch = globalThis.fetch;

      let callCount = 0;
      const fetchMock = vi.fn().mockImplementation(async (url: string) => {
        if (typeof url === "string" && url.includes("oauth2/v2.0/token")) {
          return new Response(JSON.stringify(tokenResponse), {
            status: 200,
            headers: { "Content-Type": "application/json" },
          });
        }
        callCount++;
        // First Graph call fails, which triggers 401 retry with another token call,
        // so we fail on every non-token call to ensure the error propagates.
        return new Response(
          JSON.stringify({ error: { code: "Forbidden", message: errorMessage } }),
          { status, headers: { "Content-Type": "application/json" } },
        );
      });

      globalThis.fetch = fetchMock;
      await plugin.definition.setup(harness.ctx);

      return { harness, originalFetch };
    }

    it("m365-groups returns error string when Graph API call fails", async () => {
      const { harness, originalFetch } = await setupWithGraphError(403, "Insufficient privileges");
      try {
        const result = await harness.getData<{ error: string }>("m365-groups");

        expect(result.error).toBeDefined();
        expect(typeof result.error).toBe("string");
        expect(result.error.length).toBeGreaterThan(0);
      } finally {
        globalThis.fetch = originalFetch;
      }
    });

    it("m365-sites returns error string when Graph API call fails", async () => {
      const { harness, originalFetch } = await setupWithGraphError(500, "Internal server error");
      try {
        const result = await harness.getData<{ error: string }>("m365-sites");

        expect(result.error).toBeDefined();
        expect(typeof result.error).toBe("string");
        expect(result.error.length).toBeGreaterThan(0);
      } finally {
        globalThis.fetch = originalFetch;
      }
    });

    it("m365-plans returns error string when Graph API call fails (with valid groupId)", async () => {
      const { harness, originalFetch } = await setupWithGraphError(403, "Access denied");
      try {
        const result = await harness.getData<{ error: string }>("m365-plans", {
          groupId: "grp-1",
        });

        expect(result.error).toBeDefined();
        expect(typeof result.error).toBe("string");
      } finally {
        globalThis.fetch = originalFetch;
      }
    });

    it("m365-drives returns error string when Graph API call fails (with valid siteId)", async () => {
      const { harness, originalFetch } = await setupWithGraphError(404, "Site not found");
      try {
        const result = await harness.getData<{ error: string }>("m365-drives", {
          siteId: "site-1",
        });

        expect(result.error).toBeDefined();
        expect(typeof result.error).toBe("string");
      } finally {
        globalThis.fetch = originalFetch;
      }
    });

    it("m365-folders returns error string when Graph API call fails (with valid driveId)", async () => {
      const { harness, originalFetch } = await setupWithGraphError(403, "Forbidden");
      try {
        const result = await harness.getData<{ error: string }>("m365-folders", {
          driveId: "drive-1",
        });

        expect(result.error).toBeDefined();
        expect(typeof result.error).toBe("string");
      } finally {
        globalThis.fetch = originalFetch;
      }
    });

    it("m365-calendars returns error string when Graph API call fails (with valid userId)", async () => {
      const { harness, originalFetch } = await setupWithGraphError(403, "Forbidden");
      try {
        const result = await harness.getData<{ error: string }>("m365-calendars", {
          userId: "user-1",
        });

        expect(result.error).toBeDefined();
        expect(typeof result.error).toBe("string");
      } finally {
        globalThis.fetch = originalFetch;
      }
    });
  });

  // ── 5. Response shape contracts ───────────────────────────────────────────

  describe("response shape contracts", () => {
    it("error responses always have a single 'error' string property", async () => {
      const harness = createUnconfiguredHarness();
      await plugin.definition.setup(harness.ctx);

      for (const handlerKey of ALL_HANDLERS) {
        const result = await harness.getData<Record<string, unknown>>(handlerKey, {});
        expect(Object.keys(result)).toEqual(["error"]);
        expect(typeof result.error).toBe("string");
        expect((result.error as string).length).toBeGreaterThan(0);
      }
    });

    it("success responses have an 'items' array of objects with id and name strings", async () => {
      const harness = createConfiguredHarness();
      const originalFetch = globalThis.fetch;

      const tokenResponse = {
        access_token: "fake-token",
        token_type: "Bearer",
        expires_in: 3600,
      };

      // Return different shapes based on the Graph API path to validate
      // that each handler maps the correct field to 'name'.
      globalThis.fetch = vi.fn().mockImplementation(async (url: string) => {
        if (typeof url === "string" && url.includes("oauth2/v2.0/token")) {
          return new Response(JSON.stringify(tokenResponse), {
            status: 200,
            headers: { "Content-Type": "application/json" },
          });
        }

        // Groups — displayName
        if (url.includes("/groups?")) {
          return new Response(
            JSON.stringify({
              value: [{ id: "g1", displayName: "Group 1", groupTypes: ["Unified"] }],
            }),
            { status: 200, headers: { "Content-Type": "application/json" } },
          );
        }

        // Plans — title
        if (url.includes("/planner/plans")) {
          return new Response(
            JSON.stringify({
              value: [{ id: "p1", title: "Plan 1", owner: "o1" }],
            }),
            { status: 200, headers: { "Content-Type": "application/json" } },
          );
        }

        // Sites — displayName
        if (url.includes("/sites?")) {
          return new Response(
            JSON.stringify({
              value: [{ id: "s1", displayName: "Site 1", webUrl: "https://example.com" }],
            }),
            { status: 200, headers: { "Content-Type": "application/json" } },
          );
        }

        // Drives — name
        if (url.includes("/drives?")) {
          return new Response(
            JSON.stringify({
              value: [{ id: "d1", name: "Drive 1", driveType: "documentLibrary" }],
            }),
            { status: 200, headers: { "Content-Type": "application/json" } },
          );
        }

        // Folders — name (with folder property)
        if (url.includes("/root/children")) {
          return new Response(
            JSON.stringify({
              value: [
                { id: "f1", name: "Folder 1", folder: { childCount: 0 }, size: 0, webUrl: "", lastModifiedDateTime: "" },
              ],
            }),
            { status: 200, headers: { "Content-Type": "application/json" } },
          );
        }

        // Calendars — name
        if (url.includes("/calendars")) {
          return new Response(
            JSON.stringify({
              value: [{ id: "c1", name: "Calendar 1", isDefaultCalendar: true }],
            }),
            { status: 200, headers: { "Content-Type": "application/json" } },
          );
        }

        // Fallback
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { "Content-Type": "application/json" },
        });
      });

      try {
        await plugin.definition.setup(harness.ctx);

        // Test each handler and validate shape
        const handlersWithParams: Record<string, Record<string, unknown>> = {
          "m365-groups": {},
          "m365-plans": { groupId: "grp-1" },
          "m365-sites": {},
          "m365-drives": { siteId: "site-1" },
          "m365-folders": { driveId: "drive-1" },
          "m365-calendars": { userId: "user-1" },
        };

        for (const [key, params] of Object.entries(handlersWithParams)) {
          const result = await harness.getData<{ items: Array<{ id: string; name: string }> }>(
            key,
            params,
          );

          expect(result.items).toBeDefined();
          expect(Array.isArray(result.items)).toBe(true);

          for (const item of result.items) {
            expect(typeof item.id).toBe("string");
            expect(typeof item.name).toBe("string");
            expect(item.id.length).toBeGreaterThan(0);
            expect(item.name.length).toBeGreaterThan(0);
          }
        }
      } finally {
        globalThis.fetch = originalFetch;
      }
    });
  });
});

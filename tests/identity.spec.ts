import { describe, expect, it } from "vitest";
import { AgentIdentityService } from "../src/services/identity.js";
import { DEFAULT_CONFIG, type M365Config } from "../src/constants.js";

function makeConfig(overrides: Partial<M365Config> = {}): M365Config {
  return { ...DEFAULT_CONFIG, ...overrides };
}

describe("AgentIdentityService", () => {
  it("resolves a mapped agent to the configured M365 user", () => {
    const svc = new AgentIdentityService(makeConfig({
      agentIdentityMap: { "agent-1": "ceo@contoso.com" },
    }));
    expect(svc.resolveUserId("agent-1")).toBe("ceo@contoso.com");
  });

  it("returns null for an unmapped agent", () => {
    const svc = new AgentIdentityService(makeConfig({
      agentIdentityMap: { "agent-1": "ceo@contoso.com" },
    }));
    expect(svc.resolveUserId("agent-unknown")).toBeNull();
  });

  it("returns the default service user ID", () => {
    const svc = new AgentIdentityService(makeConfig({
      defaultServiceUserId: "service@contoso.com",
    }));
    expect(svc.getDefaultUserId()).toBe("service@contoso.com");
  });

  it("hasIdentity returns true for mapped agents", () => {
    const svc = new AgentIdentityService(makeConfig({
      agentIdentityMap: { "agent-1": "ceo@contoso.com" },
    }));
    expect(svc.hasIdentity("agent-1")).toBe(true);
    expect(svc.hasIdentity("agent-unknown")).toBe(false);
  });

  it("listMappings returns all entries", () => {
    const svc = new AgentIdentityService(makeConfig({
      agentIdentityMap: {
        "agent-1": "ceo@contoso.com",
        "agent-2": "dev@contoso.com",
      },
    }));
    const mappings = svc.listMappings();
    expect(mappings).toHaveLength(2);
    expect(mappings).toContainEqual({ agentId: "agent-1", m365UserId: "ceo@contoso.com" });
    expect(mappings).toContainEqual({ agentId: "agent-2", m365UserId: "dev@contoso.com" });
  });

  it("resolveActingUserId uses mapped identity for known agent", () => {
    const svc = new AgentIdentityService(makeConfig({
      agentIdentityMap: { "agent-1": "ceo@contoso.com" },
      defaultServiceUserId: "service@contoso.com",
    }));
    expect(svc.resolveActingUserId("agent-1")).toBe("ceo@contoso.com");
  });

  it("resolveActingUserId falls back to default for unknown agent", () => {
    const svc = new AgentIdentityService(makeConfig({
      agentIdentityMap: { "agent-1": "ceo@contoso.com" },
      defaultServiceUserId: "service@contoso.com",
    }));
    expect(svc.resolveActingUserId("agent-unknown")).toBe("service@contoso.com");
  });

  it("resolveActingUserId falls back to default when no agentId provided", () => {
    const svc = new AgentIdentityService(makeConfig({
      defaultServiceUserId: "service@contoso.com",
    }));
    expect(svc.resolveActingUserId(undefined)).toBe("service@contoso.com");
    expect(svc.resolveActingUserId("")).toBe("service@contoso.com");
  });

  it("resolveActingUserId returns null when no mapping and no default", () => {
    const svc = new AgentIdentityService(makeConfig({
      agentIdentityMap: {},
      defaultServiceUserId: "",
    }));
    expect(svc.resolveActingUserId("agent-unknown")).toBeNull();
    expect(svc.resolveActingUserId(undefined)).toBeNull();
  });

  it("handles empty identity map", () => {
    const svc = new AgentIdentityService(makeConfig({
      agentIdentityMap: {},
    }));
    expect(svc.listMappings()).toHaveLength(0);
    expect(svc.resolveUserId("any")).toBeNull();
  });
});

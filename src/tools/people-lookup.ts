import type { ToolResult, ToolRunContext } from "@paperclipai/plugin-sdk";
import type { PeopleService } from "../services/people.js";

export interface PeopleLookupParams {
  query: string;
}

export async function handlePeopleLookup(
  params: unknown,
  _runCtx: ToolRunContext,
  peopleService: PeopleService,
): Promise<ToolResult> {
  const { query } = params as PeopleLookupParams;

  if (!query) {
    return { error: "query is required" };
  }

  try {
    const users = await peopleService.lookupUser(query);

    if (users.length === 0) {
      return { content: `No users found for "${query}"`, data: { users: [] } };
    }

    const summary = users
      .map((u) => {
        const parts = [u.displayName];
        if (u.mail) parts.push(`<${u.mail}>`);
        if (u.jobTitle) parts.push(`— ${u.jobTitle}`);
        if (u.department) parts.push(`(${u.department})`);
        return parts.join(" ");
      })
      .join("\n");

    return {
      content: `Found ${users.length} users:\n${summary}`,
      data: { users },
    };
  } catch (err) {
    return {
      error: `Failed to lookup users: ${err instanceof Error ? err.message : String(err)}`,
    };
  }
}

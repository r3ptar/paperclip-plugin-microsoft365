import type { ToolResult, ToolRunContext } from "@paperclipai/plugin-sdk";
import type { PeopleService } from "../services/people.js";
import { isValidGraphId } from "../graph/validate-id.js";

export interface PeopleListTeamMembersParams {
  groupId: string;
}

export async function handlePeopleListTeamMembers(
  params: unknown,
  _runCtx: ToolRunContext,
  peopleService: PeopleService,
): Promise<ToolResult> {
  const { groupId } = params as PeopleListTeamMembersParams;

  if (!groupId) {
    return { error: "groupId is required" };
  }
  if (!isValidGraphId(groupId)) {
    return { error: "Invalid groupId format" };
  }

  try {
    const members = await peopleService.listGroupMembers(groupId);

    if (members.length === 0) {
      return { content: "No members found", data: { members: [] } };
    }

    const summary = members
      .map((m) => {
        const parts = [m.displayName];
        if (m.mail) parts.push(`<${m.mail}>`);
        if (m.jobTitle) parts.push(`— ${m.jobTitle}`);
        return parts.join(" ");
      })
      .join("\n");

    return {
      content: `${members.length} members:\n${summary}`,
      data: { members },
    };
  } catch (err) {
    return {
      error: `Failed to list team members: ${err instanceof Error ? err.message : String(err)}`,
    };
  }
}

import type { ToolResult, ToolRunContext } from "@paperclipai/plugin-sdk";
import type { PeopleService } from "../services/people.js";
import { isValidGraphId } from "../graph/validate-id.js";

export interface PeopleGetManagerParams {
  userId: string;
}

export async function handlePeopleGetManager(
  params: unknown,
  _runCtx: ToolRunContext,
  peopleService: PeopleService,
): Promise<ToolResult> {
  const { userId } = params as PeopleGetManagerParams;

  if (!userId) {
    return { error: "userId is required" };
  }
  if (!isValidGraphId(userId)) {
    return { error: "Invalid userId format" };
  }

  try {
    const manager = await peopleService.getManager(userId);

    if (!manager) {
      return {
        content: `No manager found for user ${userId}`,
        data: { manager: null },
      };
    }

    const parts = [manager.displayName];
    if (manager.mail) parts.push(`<${manager.mail}>`);
    if (manager.jobTitle) parts.push(`— ${manager.jobTitle}`);

    return {
      content: `Manager: ${parts.join(" ")}`,
      data: { manager },
    };
  } catch (err) {
    return {
      error: `Failed to get manager: ${err instanceof Error ? err.message : String(err)}`,
    };
  }
}

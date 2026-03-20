import type { ToolResult, ToolRunContext } from "@paperclipai/plugin-sdk";
import type { PeopleService } from "../services/people.js";
import { isValidGraphId } from "../graph/validate-id.js";

export interface PeopleGetPresenceParams {
  userId?: string;
  userIds?: string[];
}

export async function handlePeopleGetPresence(
  params: unknown,
  _runCtx: ToolRunContext,
  peopleService: PeopleService,
): Promise<ToolResult> {
  const { userId, userIds } = params as PeopleGetPresenceParams;

  if (!userId && (!userIds || userIds.length === 0)) {
    return { error: "Either userId or userIds is required" };
  }

  try {
    if (userIds && userIds.length > 0) {
      for (const id of userIds) {
        if (!isValidGraphId(id)) return { error: `Invalid userId format: ${id}` };
      }
      const presences = await peopleService.getPresences(userIds);
      const summary = presences
        .map((p) => `${p.id}: ${p.availability} (${p.activity})`)
        .join("\n");

      return {
        content: `Presence for ${presences.length} users:\n${summary}`,
        data: { presences },
      };
    }

    if (!isValidGraphId(userId!)) return { error: "Invalid userId format" };
    const presence = await peopleService.getPresence(userId!);
    return {
      content: `${presence.availability} (${presence.activity})`,
      data: { presence },
    };
  } catch (err) {
    return {
      error: `Failed to get presence: ${err instanceof Error ? err.message : String(err)}`,
    };
  }
}

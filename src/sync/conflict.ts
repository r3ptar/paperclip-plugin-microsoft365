import type { ConflictStrategy } from "../constants.js";

export interface ConflictInput {
  paperclipUpdatedAt: string;
  plannerUpdatedAt: string;
  strategy: ConflictStrategy;
}

export type ConflictWinner = "paperclip" | "planner";

/**
 * Determine which side wins a sync conflict.
 */
export function resolveConflict(input: ConflictInput): ConflictWinner {
  switch (input.strategy) {
    case "paperclip_wins":
      return "paperclip";
    case "planner_wins":
      return "planner";
    case "last_write_wins":
    default: {
      const pcTime = new Date(input.paperclipUpdatedAt).getTime();
      const plTime = new Date(input.plannerUpdatedAt).getTime();
      return pcTime >= plTime ? "paperclip" : "planner";
    }
  }
}

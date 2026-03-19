import {
  PAPERCLIP_TO_PLANNER,
  PLANNER_BUCKET_TO_PAPERCLIP,
  type PaperclipIssueStatus,
  type PlannerStatusMapping,
} from "../constants.js";

/**
 * Map a Paperclip issue status to the corresponding Planner task state.
 */
export function toPlannerStatus(paperclipStatus: PaperclipIssueStatus): PlannerStatusMapping {
  return PAPERCLIP_TO_PLANNER[paperclipStatus] ?? PAPERCLIP_TO_PLANNER.todo;
}

/**
 * Map a Planner task back to a Paperclip issue status.
 * Bucket name is the primary discriminator because percentComplete is ambiguous
 * (e.g., 50% could be in_progress, in_review, or blocked).
 */
export function toPaperclipStatus(
  percentComplete: number,
  bucketName: string,
): PaperclipIssueStatus {
  // Bucket name takes priority — it's unambiguous
  const bucketMatch = PLANNER_BUCKET_TO_PAPERCLIP[bucketName];
  if (bucketMatch) {
    return bucketMatch;
  }

  // Fall back to percentComplete heuristic
  if (percentComplete === 100) return "done";
  if (percentComplete >= 50) return "in_progress";
  return "todo";
}

/**
 * Check whether a Paperclip status and a Planner state are already in sync.
 */
export function isStatusInSync(
  paperclipStatus: PaperclipIssueStatus,
  percentComplete: number,
  bucketName: string,
): boolean {
  const expected = toPlannerStatus(paperclipStatus);
  return expected.percentComplete === percentComplete && expected.bucketName === bucketName;
}

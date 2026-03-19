import type { ToolResult, ToolRunContext } from "@paperclipai/plugin-sdk";
import type { SharePointService } from "../services/sharepoint.js";

export interface SharePointReadParams {
  driveId: string;
  itemId: string;
}

export async function handleSharePointRead(
  params: unknown,
  runCtx: ToolRunContext,
  sharepoint: SharePointService,
): Promise<ToolResult> {
  const { driveId, itemId } = params as SharePointReadParams;
  if (!driveId || !itemId) {
    return { error: "driveId and itemId are required" };
  }

  const content = await sharepoint.readDocument(driveId, itemId);

  const truncated = content.length > 50_000;
  return {
    content: truncated ? content.slice(0, 50_000) : content,
    data: { driveId, itemId, length: content.length, truncated },
  };
}

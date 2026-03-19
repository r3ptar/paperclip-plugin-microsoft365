import type { ToolResult, ToolRunContext } from "@paperclipai/plugin-sdk";
import type { SharePointService } from "../services/sharepoint.js";

export interface SharePointUploadParams {
  fileName: string;
  content: string;
  contentType?: string;
}

const SAFE_FILENAME_RE = /^[a-zA-Z0-9._\-\s()]{1,255}$/;

export async function handleSharePointUpload(
  params: unknown,
  runCtx: ToolRunContext,
  sharepoint: SharePointService,
): Promise<ToolResult> {
  const { fileName, content, contentType } = params as SharePointUploadParams;
  if (!fileName || !content) {
    return { error: "fileName and content are required" };
  }

  if (!SAFE_FILENAME_RE.test(fileName)) {
    return {
      error: "Invalid fileName: must be 1-255 characters, alphanumeric, dots, hyphens, underscores, spaces, or parentheses only",
    };
  }

  const item = await sharepoint.uploadFile(fileName, content, contentType);

  return {
    content: `Uploaded "${fileName}" to SharePoint (${item.size} bytes)`,
    data: {
      itemId: item.id,
      name: item.name,
      size: item.size,
      webUrl: item.webUrl,
    },
  };
}

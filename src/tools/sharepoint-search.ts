import type { ToolResult, ToolRunContext } from "@paperclipai/plugin-sdk";
import type { SharePointService } from "../services/sharepoint.js";

export interface SharePointSearchParams {
  query: string;
  maxResults?: number;
}

export async function handleSharePointSearch(
  params: unknown,
  runCtx: ToolRunContext,
  sharepoint: SharePointService,
): Promise<ToolResult> {
  const { query, maxResults } = params as SharePointSearchParams;
  if (!query) {
    return { error: "query is required" };
  }

  const results = await sharepoint.search(query, maxResults ?? 10);

  if (results.length === 0) {
    return { content: `No documents found for "${query}"`, data: { results: [] } };
  }

  const summary = results
    .map((r, i) => `${i + 1}. ${r.name} — ${(r.summary ?? "").slice(0, 100)}`)
    .join("\n");

  return {
    content: `Found ${results.length} documents:\n${summary}`,
    data: { results },
  };
}

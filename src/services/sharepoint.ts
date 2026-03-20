import type { PluginContext } from "@paperclipai/plugin-sdk";
import type { M365Config } from "../constants.js";
import type { GraphClient } from "../graph/client.js";
import type {
  DriveItem,
  GraphListResponse,
  SharePointSearchResponse,
} from "../graph/types.js";

export interface SearchResult {
  id: string;
  name: string;
  summary: string;
  webUrl: string;
  size: number;
  lastModified: string;
}

/**
 * SharePoint / OneDrive document operations via Graph API.
 */
export class SharePointService {
  constructor(
    private readonly ctx: PluginContext,
    private readonly graph: GraphClient,
    private readonly config: M365Config,
  ) {}

  /**
   * Search documents in the configured SharePoint site.
   */
  async search(query: string, maxResults = 10): Promise<SearchResult[]> {
    const response = await this.graph.post<SharePointSearchResponse>(
      "/search/query",
      {
        requests: [
          {
            entityTypes: ["driveItem"],
            query: { queryString: query },
            from: 0,
            size: maxResults,
            fields: ["id", "name", "size", "webUrl", "lastModifiedDateTime"],
            region: "NAM",
          },
        ],
      },
    );

    const hits = response.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
    return hits.map((hit) => ({
      id: hit.hitId,
      name: hit.resource.name,
      summary: hit.summary,
      webUrl: hit.resource.webUrl,
      size: hit.resource.size,
      lastModified: hit.resource.lastModifiedDateTime,
    }));
  }

  /**
   * Read a document's text content by drive item ID.
   * Respects the configured max document size.
   */
  async readDocument(driveId: string, itemId: string): Promise<string> {
    // First, get metadata to check size
    const item = await this.graph.get<DriveItem>(
      `/drives/${driveId}/items/${itemId}`,
    );

    if (item.size > this.config.maxDocSizeBytes) {
      throw new Error(
        `Document "${item.name}" (${item.size} bytes) exceeds size limit of ${this.config.maxDocSizeBytes} bytes`,
      );
    }

    if (this.graph.companyId) {
      try {
        await this.ctx.activity.log({
          companyId: this.graph.companyId,
          entityType: "document",
          entityId: itemId,
          message: `Reading SharePoint document: ${item.name}`,
          metadata: { driveId, itemId, size: item.size },
        });
      } catch {
        // Activity logging is best-effort
      }
    }

    // Download content as raw text (not JSON)
    const content = await this.graph.requestRaw(
      `/drives/${driveId}/items/${itemId}/content`,
    );

    return content;
  }

  /**
   * Upload a file to the configured SharePoint upload folder.
   */
  async uploadFile(
    fileName: string,
    content: string,
    contentType = "text/plain",
  ): Promise<DriveItem> {
    const { sharepointDriveId, sharepointUploadFolderId } = this.config;

    // Build upload path — use folder if configured, otherwise upload to root
    const uploadPath = sharepointUploadFolderId
      ? `/drives/${sharepointDriveId}/items/${sharepointUploadFolderId}:/${encodeURIComponent(fileName)}:/content`
      : `/drives/${sharepointDriveId}/root:/${encodeURIComponent(fileName)}:/content`;

    const item = await this.graph.request<DriveItem>(
      uploadPath,
      {
        method: "PUT",
        headers: { "Content-Type": contentType },
        body: content,
      },
    );

    if (this.graph.companyId) {
      try {
        await this.ctx.activity.log({
          companyId: this.graph.companyId,
          entityType: "document",
          entityId: item.id,
          message: `Uploaded document to SharePoint: ${fileName}`,
          metadata: {
            driveId: sharepointDriveId,
            itemId: item.id,
            size: item.size,
          },
        });
      } catch {
        // Activity logging is best-effort
      }
    }

    return item;
  }

  /**
   * List files in a drive folder.
   */
  async listFolder(driveId: string, folderId: string): Promise<DriveItem[]> {
    const response = await this.graph.get<GraphListResponse<DriveItem>>(
      `/drives/${driveId}/items/${folderId}/children`,
    );
    return response.value;
  }
}

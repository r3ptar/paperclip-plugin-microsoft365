/**
 * Validates that an ID is safe for Graph API URL path interpolation.
 * Allows GUIDs, opaque Graph IDs (alphanumeric, dots, hyphens, colons),
 * and SharePoint site IDs (which contain commas, e.g. "contoso.sharepoint.com,guid,guid").
 * Rejects path traversal characters (/, \, ..) and whitespace.
 */
const SAFE_GRAPH_ID_RE = /^[a-zA-Z0-9._:,@-]+$/;

export function isValidGraphId(id: string): boolean {
  return SAFE_GRAPH_ID_RE.test(id) && !id.includes("..");
}

import type { PaperclipIssueStatus } from "../constants.js";
import type { GraphMailMessage } from "../graph/types.js";

export type ParsedEmailAction =
  | { kind: "status_change"; issueId: string; newStatus: PaperclipIssueStatus; replyBody: string }
  | { kind: "comment"; issueId: string; body: string }
  | { kind: "unrecognized"; reason: string };

/**
 * Parse an inbound Graph mail message into a Paperclip action.
 *
 * Three-layer issue ID extraction:
 *   1. `X-Paperclip-Issue-Id` internet message header
 *   2. `[PC-{id}]` pattern in the subject
 *   3. `[PC-{id}]` pattern in the body
 *
 * Intent detection is based on keyword matching in the cleaned reply body
 * (quoted content stripped).
 */
export function parseInboundEmail(message: GraphMailMessage): ParsedEmailAction {
  const issueId = extractIssueId(message);
  if (!issueId) {
    return { kind: "unrecognized", reason: "no issue ID found in headers, subject, or body" };
  }

  const bodyText = extractPlainText(message.body);
  const replyBody = stripQuotedContent(bodyText).trim();

  if (!replyBody) {
    return { kind: "unrecognized", reason: "empty reply body after stripping quoted content" };
  }

  const statusIntent = detectStatusIntent(replyBody);
  if (statusIntent) {
    return { kind: "status_change", issueId, newStatus: statusIntent, replyBody };
  }

  return { kind: "comment", issueId, body: replyBody };
}

// ---------------------------------------------------------------------------
// Issue ID extraction
// ---------------------------------------------------------------------------

const PC_TAG_RE = /\[PC-([^\]]+)\]/;

function extractIssueId(message: GraphMailMessage): string | undefined {
  // Layer 1: custom header
  if (message.internetMessageHeaders) {
    const header = message.internetMessageHeaders.find(
      (h) => h.name.toLowerCase() === "x-paperclip-issue-id",
    );
    if (header?.value) {
      return header.value;
    }
  }

  // Layer 2: subject line
  const subjectMatch = message.subject.match(PC_TAG_RE);
  if (subjectMatch?.[1]) {
    return subjectMatch[1];
  }

  // Layer 3: body text
  const bodyText = extractPlainText(message.body);
  const bodyMatch = bodyText.match(PC_TAG_RE);
  if (bodyMatch?.[1]) {
    return bodyMatch[1];
  }

  return undefined;
}

// ---------------------------------------------------------------------------
// Body text helpers
// ---------------------------------------------------------------------------

/**
 * Convert body content to plain text. When the contentType is "html",
 * strip tags with a simple regex (no external parsing library).
 */
function extractPlainText(body: { contentType: string; content: string }): string {
  if (body.contentType.toLowerCase() === "html") {
    return stripHtmlTags(body.content);
  }
  return body.content;
}

function stripHtmlTags(html: string): string {
  // Replace <br>, <br/>, <br /> and block-level closing tags with newlines
  let text = html.replace(/<br\s*\/?>/gi, "\n");
  text = text.replace(/<\/(?:p|div|tr|li|h[1-6])>/gi, "\n");
  // Remove all remaining tags
  text = text.replace(/<[^>]*>/g, "");
  // Decode common HTML entities
  text = text
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, " ");
  return text;
}

/**
 * Strip quoted email content. Removes everything after the first occurrence of:
 *   - A line matching "On ... wrote:" (standard reply header)
 *   - A line starting with "---" (separator)
 *   - Two or more consecutive lines starting with ">"
 */
function stripQuotedContent(text: string): string {
  const lines = text.split("\n");
  const result: string[] = [];
  let consecutiveQuotes = 0;

  for (const line of lines) {
    // Check for "On ... wrote:" pattern
    if (/^On\s.+wrote:\s*$/i.test(line.trim())) {
      break;
    }

    // Check for "---" separator
    if (/^-{3,}\s*$/.test(line.trim())) {
      break;
    }

    // Track consecutive ">" quoted lines
    if (line.trimStart().startsWith(">")) {
      consecutiveQuotes++;
      if (consecutiveQuotes >= 2) {
        // Remove the previous ">" line we already added
        result.pop();
        break;
      }
    } else {
      consecutiveQuotes = 0;
    }

    result.push(line);
  }

  return result.join("\n");
}

// ---------------------------------------------------------------------------
// Intent detection
// ---------------------------------------------------------------------------

const STATUS_KEYWORDS: Array<{ pattern: RegExp; status: PaperclipIssueStatus }> = [
  { pattern: /\b(?:done|complete|finished|completed)\b/i, status: "done" },
  { pattern: /\b(?:approved|lgtm)\b/i, status: "done" },
  { pattern: /\b(?:blocked|stuck)\b/i, status: "blocked" },
  { pattern: /\b(?:in\s+review|reviewing)\b/i, status: "in_review" },
  { pattern: /\b(?:started|working\s+on\s+it|on\s+it)\b/i, status: "in_progress" },
];

function detectStatusIntent(replyBody: string): PaperclipIssueStatus | undefined {
  const normalized = replyBody.toLowerCase().trim();

  for (const { pattern, status } of STATUS_KEYWORDS) {
    if (pattern.test(normalized)) {
      return status;
    }
  }

  return undefined;
}

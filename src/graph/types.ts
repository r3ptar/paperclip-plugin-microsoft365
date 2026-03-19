/** Microsoft Graph API response types */

export interface GraphTokenResponse {
  access_token: string;
  token_type: string;
  expires_in: number;
}

export interface GraphError {
  error: {
    code: string;
    message: string;
    innerError?: {
      "request-id"?: string;
      date?: string;
    };
  };
}

export interface GraphListResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

export interface PlannerTask {
  id: string;
  planId: string;
  bucketId: string;
  title: string;
  percentComplete: number;
  createdDateTime: string;
  lastModifiedDateTime?: string;
  dueDateTime?: string | null;
  details?: PlannerTaskDetails;
  "@odata.etag"?: string;
}

export interface PlannerTaskDetails {
  id: string;
  description: string;
  "@odata.etag"?: string;
}

export interface PlannerBucket {
  id: string;
  name: string;
  planId: string;
  orderHint: string;
}

export interface PlannerPlan {
  id: string;
  title: string;
  owner: string;
}

export interface GraphSubscription {
  id: string;
  resource: string;
  changeType: string;
  clientState: string;
  notificationUrl: string;
  expirationDateTime: string;
}

export interface GraphChangeNotification {
  value: Array<{
    subscriptionId: string;
    clientState?: string;
    changeType: string;
    resource: string;
    resourceData?: {
      id: string;
      "@odata.type": string;
      "@odata.id": string;
      "@odata.etag"?: string;
    };
    subscriptionExpirationDateTime: string;
    tenantId: string;
  }>;
}

export interface SharePointSearchHit {
  hitId: string;
  rank: number;
  summary: string;
  resource: {
    id: string;
    name: string;
    size: number;
    webUrl: string;
    lastModifiedDateTime: string;
    parentReference?: {
      driveId: string;
      id: string;
    };
  };
}

export interface SharePointSearchResponse {
  value: Array<{
    searchTerms: string[];
    hitsContainers: Array<{
      total: number;
      moreResultsAvailable: boolean;
      hits: SharePointSearchHit[];
    }>;
  }>;
}

export interface DriveItem {
  id: string;
  name: string;
  size: number;
  webUrl: string;
  file?: { mimeType: string };
  folder?: { childCount: number };
  lastModifiedDateTime: string;
  parentReference?: {
    driveId: string;
    id: string;
    path?: string;
  };
}

export interface CalendarEvent {
  id: string;
  subject: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  body?: { contentType: string; content: string };
  webLink?: string;
}

export interface GraphMessage {
  subject: string;
  body: { contentType: string; content: string };
  toRecipients: Array<{
    emailAddress: { address: string };
  }>;
  internetMessageHeaders?: Array<{
    name: string;
    value: string;
  }>;
}

export interface GraphGroup {
  id: string;
  displayName: string;
  groupTypes: string[];
}

export interface GraphSite {
  id: string;
  displayName: string;
  webUrl: string;
}

export interface GraphDrive {
  id: string;
  name: string;
  driveType: string;
}

export interface GraphCalendar {
  id: string;
  name: string;
  isDefaultCalendar?: boolean;
}

export interface GraphMailMessage {
  id: string;
  subject: string;
  body: { contentType: string; content: string };
  from: { emailAddress: { name: string; address: string } };
  toRecipients: Array<{ emailAddress: { name: string; address: string } }>;
  internetMessageHeaders?: Array<{ name: string; value: string }>;
  internetMessageId?: string;
  conversationId?: string;
  receivedDateTime: string;
}

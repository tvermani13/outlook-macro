export interface EmailAddress {
  name?: string;
  address?: string;
}

export interface Recipient {
  emailAddress?: EmailAddress;
}

export interface ItemBody {
  contentType?: "html" | "text" | string;
  content?: string;
}

export interface InternetMessageHeader {
  name?: string;
  value?: string;
}

export interface GraphMessage {
  id: string;
  subject?: string;
  sentDateTime?: string;
  receivedDateTime?: string;
  conversationId?: string;
  conversationIndex?: string;
  internetMessageId?: string;
  internetMessageHeaders?: InternetMessageHeader[];
  from?: Recipient;
  sender?: Recipient;
  toRecipients?: Recipient[];
  ccRecipients?: Recipient[];
  bccRecipients?: Recipient[];
  replyTo?: Recipient[];
  body?: ItemBody;
  uniqueBody?: ItemBody;
  hasAttachments?: boolean;
  webLink?: string;
}

export interface GraphAttachment {
  "@odata.type"?: string;
  id: string;
  name?: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  lastModifiedDateTime?: string;
  contentBytes?: string;
}

export interface GraphUser {
  id: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
}

export interface GraphCollection<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

export interface MailboxConfig {
  id: string;
  label: string;
  expectedUsername?: string;
}

export type ExtractionFailureMode = "skip" | "useSubject" | "error";

export interface SubjectRuleConfig {
  name: string;
  enabled?: boolean;
  matchPattern: string;
  matchFlags?: string;
  extractPattern?: string;
  extractFlags?: string;
  extractGroup?: string | number;
  onExtractionFailure?: ExtractionFailureMode;
}

export interface ScanConfig {
  lookbackDays: number;
  timeZone: string;
  sentFolders: string[];
  responseFolders: string[];
  onlyResponsesFromOriginalRecipients: boolean;
  excludeSenderAddresses: string[];
  maxMessagesPerFolder: number;
  includeNoResponseRows: boolean;
}

export type AttachmentMode = "metadata" | "download";

export interface OutputConfig {
  directory: string;
  csvFile: string;
  jsonFile: string;
  attachmentMode: AttachmentMode;
}

export interface AppConfig {
  version: 1;
  mailboxes: MailboxConfig[];
  subjectRules: SubjectRuleConfig[];
  scan: ScanConfig;
  output: OutputConfig;
}

export interface LoadedConfig {
  config: AppConfig;
  path: string;
  directory: string;
}

export interface SubjectMatch {
  ruleName: string;
  portion: string;
}

export interface TrackedSentMessage {
  mailboxId: string;
  message: GraphMessage;
  subjectMatch: SubjectMatch;
  originalRecipientAddresses: Set<string>;
}

export interface DetailedResponse {
  message: GraphMessage;
  attachments: GraphAttachment[];
  downloadedAttachmentPaths: string[];
}

export type ResponseMatchBasis =
  | "in-reply-to"
  | "references"
  | "conversation-and-recipient";

export interface AssignedResponse {
  sent: TrackedSentMessage;
  response: GraphMessage;
  basis: ResponseMatchBasis;
}

export interface OutputRow {
  mailbox_id: string;
  mailbox_label: string;
  mailbox_address: string;
  rule_name: string;
  status: "responded" | "no_response";
  sent_date: string;
  sent_at: string;
  sent_subject_portion: string;
  sent_subject: string;
  sent_sender: string;
  sent_recipients: string;
  response_subject: string;
  response_date: string;
  response_at: string;
  response_sender: string;
  response_recipients: string;
  response_content: string;
  response_attachments: string;
  response_attachment_files: string;
  match_basis: ResponseMatchBasis | "";
  sent_message_id: string;
  response_message_id: string;
  response_web_link: string;
}

export interface MailboxScanSummary {
  mailboxId: string;
  mailboxLabel: string;
  mailboxAddress: string;
  sentMessagesScanned: number;
  responseMessagesScanned: number;
  trackedSentMessages: number;
  responsesFound: number;
}

export interface ScanOutput {
  generatedAt: string;
  lookbackStart: string;
  timeZone: string;
  summaries: MailboxScanSummary[];
  rows: OutputRow[];
}

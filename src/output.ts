import { createHash } from "node:crypto";
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { rowsToCsv } from "./csv.js";
import type {
  GraphAttachment,
  GraphMessage,
  GraphUser,
  LoadedConfig,
  MailboxConfig,
  OutputRow,
  ScanOutput,
  TrackedSentMessage,
} from "./types.js";

export function formatAddress(
  recipient:
    | { emailAddress?: { name?: string; address?: string } }
    | undefined,
): string {
  const name = recipient?.emailAddress?.name?.trim();
  const address = recipient?.emailAddress?.address?.trim();
  if (name !== undefined && name !== "" && address !== undefined && address !== "") {
    return `${name} <${address}>`;
  }
  return address ?? name ?? "";
}

export function formatRecipients(message: GraphMessage): string {
  return [
    ...(message.toRecipients ?? []),
    ...(message.ccRecipients ?? []),
    ...(message.bccRecipients ?? []),
  ]
    .map(formatAddress)
    .filter((value) => value !== "")
    .join("; ");
}

export function dateInTimeZone(iso: string, timeZone: string): string {
  if (iso === "") {
    return "";
  }
  const date = new Date(iso);
  if (!Number.isFinite(date.getTime())) {
    return "";
  }
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(date);
  const values = Object.fromEntries(
    parts
      .filter((part) => part.type !== "literal")
      .map((part) => [part.type, part.value]),
  );
  return `${values.year}-${values.month}-${values.day}`;
}

export function messageContent(message: GraphMessage): string {
  return (
    message.uniqueBody?.content?.trim() ??
    message.body?.content?.trim() ??
    ""
  );
}

function stableId(value: string): string {
  return createHash("sha256").update(value).digest("hex").slice(0, 12);
}

function safeFilename(filename: string): string {
  const base = path.basename(filename).replaceAll(/[\u0000-\u001f]/g, "");
  const safe = base.replaceAll(/[<>:"/\\|?*]/g, "_").trim();
  return safe === "" ? "attachment" : safe;
}

export function attachmentDirectory(
  outputDirectory: string,
  mailboxId: string,
  messageId: string,
): string {
  return path.join(
    outputDirectory,
    "attachments",
    mailboxId,
    stableId(messageId),
  );
}

export function attachmentPath(
  directory: string,
  attachment: GraphAttachment,
  index: number,
): string {
  return path.join(
    directory,
    `${String(index + 1).padStart(2, "0")}-${safeFilename(attachment.name ?? "attachment")}`,
  );
}

export async function saveAttachment(
  filePath: string,
  attachment: GraphAttachment,
): Promise<void> {
  if (attachment.contentBytes === undefined) {
    throw new Error(
      `Attachment "${attachment.name ?? attachment.id}" is not a downloadable file attachment.`,
    );
  }
  await mkdir(path.dirname(filePath), { recursive: true, mode: 0o700 });
  await writeFile(filePath, Buffer.from(attachment.contentBytes, "base64"), {
    mode: 0o600,
  });
}

interface BuildRowOptions {
  mailbox: MailboxConfig;
  user: GraphUser;
  sent: TrackedSentMessage;
  response?: GraphMessage;
  attachments?: GraphAttachment[];
  downloadedAttachmentPaths?: string[];
  matchBasis?: OutputRow["match_basis"];
  timeZone: string;
}

export function buildOutputRow(options: BuildRowOptions): OutputRow {
  const sentMessage = options.sent.message;
  const response = options.response;
  const sentAt = sentMessage.sentDateTime ?? "";
  const responseAt = response?.receivedDateTime ?? response?.sentDateTime ?? "";
  const mailboxAddress =
    options.user.mail ?? options.user.userPrincipalName ?? "";
  const responseSender =
    formatAddress(response?.sender) || formatAddress(response?.from);

  return {
    mailbox_id: options.mailbox.id,
    mailbox_label: options.mailbox.label,
    mailbox_address: mailboxAddress,
    rule_name: options.sent.subjectMatch.ruleName,
    status: response === undefined ? "no_response" : "responded",
    sent_date: dateInTimeZone(sentAt, options.timeZone),
    sent_at: sentAt,
    sent_subject_portion: options.sent.subjectMatch.portion,
    sent_subject: sentMessage.subject ?? "",
    sent_sender:
      formatAddress(sentMessage.sender) || formatAddress(sentMessage.from),
    sent_recipients: formatRecipients(sentMessage),
    response_subject: response?.subject ?? "",
    response_date: dateInTimeZone(responseAt, options.timeZone),
    response_at: responseAt,
    response_sender: responseSender,
    response_recipients: response === undefined ? "" : formatRecipients(response),
    response_content: response === undefined ? "" : messageContent(response),
    response_attachments: JSON.stringify(options.attachments ?? []),
    response_attachment_files: (options.downloadedAttachmentPaths ?? []).join(
      "; ",
    ),
    match_basis: options.matchBasis ?? "",
    sent_message_id: sentMessage.id,
    response_message_id: response?.id ?? "",
    response_web_link: response?.webLink ?? "",
  };
}

export async function writeScanOutput(
  loadedConfig: LoadedConfig,
  output: ScanOutput,
): Promise<{ csvPath: string; jsonPath: string }> {
  const outputDirectory = path.resolve(
    loadedConfig.directory,
    loadedConfig.config.output.directory,
  );
  await mkdir(outputDirectory, { recursive: true, mode: 0o700 });
  const csvPath = path.join(
    outputDirectory,
    loadedConfig.config.output.csvFile,
  );
  const jsonPath = path.join(
    outputDirectory,
    loadedConfig.config.output.jsonFile,
  );
  await Promise.all([
    writeFile(csvPath, rowsToCsv(output.rows), { encoding: "utf8", mode: 0o600 }),
    writeFile(jsonPath, `${JSON.stringify(output, null, 2)}\n`, {
      encoding: "utf8",
      mode: 0o600,
    }),
  ]);
  return { csvPath, jsonPath };
}

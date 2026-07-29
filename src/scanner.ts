import path from "node:path";
import { getAccessToken } from "./auth.js";
import { GraphClient } from "./graph.js";
import {
  assignFirstResponses,
  normalizeAddress,
  recipientAddresses,
} from "./matcher.js";
import {
  attachmentDirectory,
  attachmentPath,
  buildOutputRow,
  saveAttachment,
  writeScanOutput,
} from "./output.js";
import { matchSubject } from "./subject-rules.js";
import type {
  GraphAttachment,
  GraphMessage,
  LoadedConfig,
  MailboxConfig,
  MailboxScanSummary,
  OutputRow,
  ScanOutput,
  TrackedSentMessage,
} from "./types.js";

function uniqueMessages(messages: GraphMessage[]): GraphMessage[] {
  return [...new Map(messages.map((message) => [message.id, message])).values()];
}

async function listConfiguredFolders(
  graph: GraphClient,
  folders: string[],
  timestampField: "sentDateTime" | "receivedDateTime",
  startIso: string,
  maxMessagesPerFolder: number,
): Promise<GraphMessage[]> {
  const pages = await Promise.all(
    folders.map((folder) =>
      graph.listFolderMessages(
        folder,
        timestampField,
        startIso,
        maxMessagesPerFolder,
      ),
    ),
  );
  return uniqueMessages(pages.flat());
}

function trackedSentMessages(
  mailboxId: string,
  messages: GraphMessage[],
  loadedConfig: LoadedConfig,
): TrackedSentMessage[] {
  return messages.flatMap((message) => {
    const subjectMatch = matchSubject(
      message.subject ?? "",
      loadedConfig.config.subjectRules,
    );
    if (
      subjectMatch === undefined ||
      message.sentDateTime === undefined ||
      message.conversationId === undefined
    ) {
      return [];
    }
    return [
      {
        mailboxId,
        message,
        subjectMatch,
        originalRecipientAddresses: recipientAddresses(message),
      },
    ];
  });
}

async function detailRelevantResponses(
  graph: GraphClient,
  lightweightResponses: GraphMessage[],
  sentMessages: TrackedSentMessage[],
): Promise<GraphMessage[]> {
  const conversationIds = new Set(
    sentMessages
      .map((sent) => sent.message.conversationId)
      .filter((value): value is string => value !== undefined),
  );
  const relevant = lightweightResponses.filter(
    (message) =>
      message.conversationId !== undefined &&
      conversationIds.has(message.conversationId),
  );

  const detailed: GraphMessage[] = [];
  const concurrency = 6;
  for (let index = 0; index < relevant.length; index += concurrency) {
    const batch = relevant.slice(index, index + concurrency);
    detailed.push(
      ...(await Promise.all(
        batch.map((message) => graph.getMessageDetails(message.id)),
      )),
    );
  }
  return detailed;
}

async function extractAttachments(
  graph: GraphClient,
  mailbox: MailboxConfig,
  response: GraphMessage,
  loadedConfig: LoadedConfig,
): Promise<{
  attachments: GraphAttachment[];
  downloadedAttachmentPaths: string[];
}> {
  const attachments = await graph.listAttachments(response.id);
  if (loadedConfig.config.output.attachmentMode !== "download") {
    return { attachments, downloadedAttachmentPaths: [] };
  }

  const outputDirectory = path.resolve(
    loadedConfig.directory,
    loadedConfig.config.output.directory,
  );
  const directory = attachmentDirectory(
    outputDirectory,
    mailbox.id,
    response.id,
  );
  const downloadedAttachmentPaths: string[] = [];
  for (const [index, metadata] of attachments.entries()) {
    if (
      metadata.isInline === true ||
      metadata["@odata.type"] !== "#microsoft.graph.fileAttachment"
    ) {
      continue;
    }
    const attachment = await graph.getAttachment(response.id, metadata.id);
    const filePath = attachmentPath(directory, attachment, index);
    await saveAttachment(filePath, attachment);
    downloadedAttachmentPaths.push(filePath);
  }
  return { attachments, downloadedAttachmentPaths };
}

async function scanMailbox(
  mailbox: MailboxConfig,
  loadedConfig: LoadedConfig,
  startIso: string,
  configuredMailboxAddresses: Set<string>,
): Promise<{ rows: OutputRow[]; summary: MailboxScanSummary }> {
  const cacheDirectory = path.resolve(loadedConfig.directory, ".cache");
  const accessToken = await getAccessToken(mailbox, cacheDirectory, false);
  const graph = new GraphClient(accessToken);
  const user = await graph.getMe();
  const mailboxAddress = normalizeAddress(
    user.mail ?? user.userPrincipalName,
  );
  if (mailboxAddress !== "") {
    configuredMailboxAddresses.add(mailboxAddress);
  }

  const [sentMessages, responseMessages] = await Promise.all([
    listConfiguredFolders(
      graph,
      loadedConfig.config.scan.sentFolders,
      "sentDateTime",
      startIso,
      loadedConfig.config.scan.maxMessagesPerFolder,
    ),
    listConfiguredFolders(
      graph,
      loadedConfig.config.scan.responseFolders,
      "receivedDateTime",
      startIso,
      loadedConfig.config.scan.maxMessagesPerFolder,
    ),
  ]);
  const tracked = trackedSentMessages(mailbox.id, sentMessages, loadedConfig);
  const detailedResponses = await detailRelevantResponses(
    graph,
    responseMessages,
    tracked,
  );

  const exclusions = new Set([
    ...configuredMailboxAddresses,
    ...loadedConfig.config.scan.excludeSenderAddresses.map(normalizeAddress),
  ]);
  const assignments = assignFirstResponses(
    tracked,
    detailedResponses,
    exclusions,
    loadedConfig.config.scan.onlyResponsesFromOriginalRecipients,
  );

  const rows: OutputRow[] = [];
  for (const sent of tracked) {
    const assignment = assignments.get(sent.message.id);
    if (assignment === undefined) {
      if (loadedConfig.config.scan.includeNoResponseRows) {
        rows.push(
          buildOutputRow({
            mailbox,
            user,
            sent,
            timeZone: loadedConfig.config.scan.timeZone,
          }),
        );
      }
      continue;
    }

    const attachmentResult = await extractAttachments(
      graph,
      mailbox,
      assignment.response,
      loadedConfig,
    );
    rows.push(
      buildOutputRow({
        mailbox,
        user,
        sent,
        response: assignment.response,
        attachments: attachmentResult.attachments,
        downloadedAttachmentPaths:
          attachmentResult.downloadedAttachmentPaths,
        matchBasis: assignment.basis,
        timeZone: loadedConfig.config.scan.timeZone,
      }),
    );
  }

  return {
    rows,
    summary: {
      mailboxId: mailbox.id,
      mailboxLabel: mailbox.label,
      mailboxAddress,
      sentMessagesScanned: sentMessages.length,
      responseMessagesScanned: responseMessages.length,
      trackedSentMessages: tracked.length,
      responsesFound: assignments.size,
    },
  };
}

export async function runScan(
  loadedConfig: LoadedConfig,
): Promise<{
  output: ScanOutput;
  csvPath: string;
  jsonPath: string;
}> {
  const start = new Date(
    Date.now() - loadedConfig.config.scan.lookbackDays * 24 * 60 * 60 * 1000,
  );
  const startIso = start.toISOString();
  const configuredMailboxAddresses = new Set(
    loadedConfig.config.mailboxes
      .map((mailbox) => normalizeAddress(mailbox.expectedUsername))
      .filter((address) => address !== ""),
  );

  const rows: OutputRow[] = [];
  const summaries: MailboxScanSummary[] = [];
  for (const mailbox of loadedConfig.config.mailboxes) {
    console.log(`Scanning ${mailbox.label}...`);
    const result = await scanMailbox(
      mailbox,
      loadedConfig,
      startIso,
      configuredMailboxAddresses,
    );
    rows.push(...result.rows);
    summaries.push(result.summary);
  }
  rows.sort((left, right) => left.sent_at.localeCompare(right.sent_at));

  const output: ScanOutput = {
    generatedAt: new Date().toISOString(),
    lookbackStart: startIso,
    timeZone: loadedConfig.config.scan.timeZone,
    summaries,
    rows,
  };
  const paths = await writeScanOutput(loadedConfig, output);
  return { output, ...paths };
}

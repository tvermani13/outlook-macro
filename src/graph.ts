import { setTimeout as delay } from "node:timers/promises";
import type {
  GraphAttachment,
  GraphCollection,
  GraphMessage,
  GraphUser,
} from "./types.js";

const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
const DEFAULT_PREFER =
  'IdType="ImmutableId", outlook.body-content-type="text"';

function graphUrl(pathOrUrl: string): string {
  return pathOrUrl.startsWith("https://")
    ? pathOrUrl
    : `${GRAPH_ROOT}${pathOrUrl}`;
}

function retryDelayMilliseconds(
  response: Response,
  attempt: number,
): number {
  const retryAfter = response.headers.get("retry-after");
  if (retryAfter !== null) {
    const seconds = Number(retryAfter);
    if (Number.isFinite(seconds) && seconds >= 0) {
      return Math.min(seconds * 1000, 30_000);
    }
  }
  return Math.min(2 ** attempt * 1000, 30_000);
}

export class GraphClient {
  public constructor(private readonly accessToken: string) {}

  private async request<T>(
    pathOrUrl: string,
    init: RequestInit = {},
  ): Promise<T> {
    const headers = new Headers(init.headers);
    headers.set("authorization", `Bearer ${this.accessToken}`);
    headers.set("accept", "application/json");
    headers.set("prefer", DEFAULT_PREFER);

    for (let attempt = 0; attempt < 5; attempt += 1) {
      const response = await fetch(graphUrl(pathOrUrl), { ...init, headers });
      if (response.ok) {
        if (response.status === 204) {
          return undefined as T;
        }
        return (await response.json()) as T;
      }

      if (
        (response.status === 429 ||
          response.status === 503 ||
          response.status === 504) &&
        attempt < 4
      ) {
        await delay(retryDelayMilliseconds(response, attempt));
        continue;
      }

      const body = await response.text();
      throw new Error(
        `Microsoft Graph ${response.status} ${response.statusText} for ${pathOrUrl}: ${body}`,
      );
    }

    throw new Error(`Microsoft Graph retry limit reached for ${pathOrUrl}.`);
  }

  public getMe(): Promise<GraphUser> {
    return this.request<GraphUser>(
      "/me?$select=id,displayName,mail,userPrincipalName",
    );
  }

  public async listFolderMessages(
    folder: string,
    timestampField: "sentDateTime" | "receivedDateTime",
    startIso: string,
    maxMessages: number,
  ): Promise<GraphMessage[]> {
    const select = [
      "id",
      "subject",
      "sentDateTime",
      "receivedDateTime",
      "conversationId",
      "conversationIndex",
      "internetMessageId",
      "from",
      "sender",
      "toRecipients",
      "ccRecipients",
      "bccRecipients",
      "hasAttachments",
      "webLink",
    ].join(",");
    const params = new URLSearchParams({
      $select: select,
      $filter: `${timestampField} ge ${startIso}`,
      $orderby: `${timestampField} asc`,
      $top: "100",
    });
    let nextUrl: string | undefined =
      `/me/mailFolders/${encodeURIComponent(folder)}/messages?${params.toString()}`;
    const messages: GraphMessage[] = [];

    while (nextUrl !== undefined && messages.length < maxMessages) {
      const page: GraphCollection<GraphMessage> = await this.request(nextUrl);
      const remaining = maxMessages - messages.length;
      messages.push(...page.value.slice(0, remaining));
      nextUrl = page["@odata.nextLink"];
    }

    if (nextUrl !== undefined) {
      throw new Error(
        `Folder "${folder}" contains more than scan.maxMessagesPerFolder (${maxMessages}) messages in the configured lookback window. Increase the limit or shorten scan.lookbackDays so the table is not silently incomplete.`,
      );
    }

    return messages;
  }

  public getMessageDetails(messageId: string): Promise<GraphMessage> {
    const select = [
      "id",
      "subject",
      "sentDateTime",
      "receivedDateTime",
      "conversationId",
      "conversationIndex",
      "internetMessageId",
      "internetMessageHeaders",
      "from",
      "sender",
      "toRecipients",
      "ccRecipients",
      "bccRecipients",
      "replyTo",
      "body",
      "uniqueBody",
      "hasAttachments",
      "webLink",
    ].join(",");
    const params = new URLSearchParams({ $select: select });
    return this.request<GraphMessage>(
      `/me/messages/${encodeURIComponent(messageId)}?${params.toString()}`,
    );
  }

  public async listAttachments(
    messageId: string,
  ): Promise<GraphAttachment[]> {
    const select = [
      "id",
      "name",
      "contentType",
      "size",
      "isInline",
      "lastModifiedDateTime",
    ].join(",");
    const params = new URLSearchParams({ $select: select });
    const result = await this.request<GraphCollection<GraphAttachment>>(
      `/me/messages/${encodeURIComponent(messageId)}/attachments?${params.toString()}`,
    );
    return result.value.map((attachment) => {
      const { contentBytes: _contentBytes, ...metadata } = attachment;
      return metadata;
    });
  }

  public getAttachment(
    messageId: string,
    attachmentId: string,
  ): Promise<GraphAttachment> {
    return this.request<GraphAttachment>(
      `/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`,
    );
  }
}

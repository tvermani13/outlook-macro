import type {
  AssignedResponse,
  GraphMessage,
  ResponseMatchBasis,
  TrackedSentMessage,
} from "./types.js";

export function normalizeAddress(address: string | undefined): string {
  return (address ?? "").trim().toLocaleLowerCase("en-US");
}

export function recipientAddresses(message: GraphMessage): Set<string> {
  const recipients = [
    ...(message.toRecipients ?? []),
    ...(message.ccRecipients ?? []),
    ...(message.bccRecipients ?? []),
  ];
  const addresses = recipients
    .map((recipient) => normalizeAddress(recipient.emailAddress?.address))
    .filter((address) => address !== "");
  return new Set(addresses);
}

export function senderAddress(message: GraphMessage): string {
  return normalizeAddress(
    message.sender?.emailAddress?.address ??
      message.from?.emailAddress?.address,
  );
}

function sentTimestamp(sent: TrackedSentMessage): number {
  return Date.parse(sent.message.sentDateTime ?? "");
}

function responseTimestamp(response: GraphMessage): number {
  return Date.parse(
    response.receivedDateTime ?? response.sentDateTime ?? "",
  );
}

function normalizedMessageId(value: string | undefined): string {
  return (value ?? "").trim().toLocaleLowerCase("en-US");
}

function headerValue(message: GraphMessage, headerName: string): string {
  const header = message.internetMessageHeaders?.find(
    (item) =>
      item.name?.toLocaleLowerCase("en-US") ===
      headerName.toLocaleLowerCase("en-US"),
  );
  return header?.value ?? "";
}

function extractHeaderMessageIds(value: string): Set<string> {
  const ids = value.match(/<[^<>]+>/g) ?? [];
  return new Set(ids.map(normalizedMessageId));
}

interface CandidateMatch {
  sent: TrackedSentMessage;
  basis: ResponseMatchBasis;
  basisScore: number;
}

function matchCandidateToSent(
  response: GraphMessage,
  sentMessages: TrackedSentMessage[],
  excludedSenderAddresses: Set<string>,
  onlyOriginalRecipients: boolean,
): CandidateMatch | undefined {
  const responseSender = senderAddress(response);
  if (
    responseSender === "" ||
    excludedSenderAddresses.has(responseSender) ||
    !Number.isFinite(responseTimestamp(response))
  ) {
    return undefined;
  }

  const inReplyToIds = extractHeaderMessageIds(
    headerValue(response, "in-reply-to"),
  );
  const referenceIds = extractHeaderMessageIds(
    headerValue(response, "references"),
  );
  const candidates: CandidateMatch[] = [];

  for (const sent of sentMessages) {
    const sentAt = sentTimestamp(sent);
    if (
      !Number.isFinite(sentAt) ||
      responseTimestamp(response) <= sentAt ||
      response.conversationId === undefined ||
      sent.message.conversationId !== response.conversationId
    ) {
      continue;
    }

    if (
      onlyOriginalRecipients &&
      !sent.originalRecipientAddresses.has(responseSender)
    ) {
      continue;
    }

    const sentInternetId = normalizedMessageId(
      sent.message.internetMessageId,
    );
    if (sentInternetId !== "" && inReplyToIds.has(sentInternetId)) {
      candidates.push({ sent, basis: "in-reply-to", basisScore: 3 });
      continue;
    }
    if (sentInternetId !== "" && referenceIds.has(sentInternetId)) {
      candidates.push({ sent, basis: "references", basisScore: 2 });
      continue;
    }
    candidates.push({
      sent,
      basis: "conversation-and-recipient",
      basisScore: 1,
    });
  }

  candidates.sort((left, right) => {
    if (left.basisScore !== right.basisScore) {
      return right.basisScore - left.basisScore;
    }
    return sentTimestamp(right.sent) - sentTimestamp(left.sent);
  });
  return candidates[0];
}

export function assignFirstResponses(
  sentMessages: TrackedSentMessage[],
  responses: GraphMessage[],
  excludedSenderAddresses: Set<string>,
  onlyOriginalRecipients: boolean,
): Map<string, AssignedResponse> {
  const sentByConversation = new Map<string, TrackedSentMessage[]>();
  for (const sent of sentMessages) {
    const conversationId = sent.message.conversationId;
    if (conversationId === undefined) {
      continue;
    }
    const bucket = sentByConversation.get(conversationId) ?? [];
    bucket.push(sent);
    sentByConversation.set(conversationId, bucket);
  }
  for (const bucket of sentByConversation.values()) {
    bucket.sort((left, right) => sentTimestamp(left) - sentTimestamp(right));
  }

  const sortedResponses = [...responses].sort(
    (left, right) => responseTimestamp(left) - responseTimestamp(right),
  );
  const assignments = new Map<string, AssignedResponse>();

  for (const response of sortedResponses) {
    const conversationId = response.conversationId;
    if (conversationId === undefined) {
      continue;
    }
    const sentInConversation = sentByConversation.get(conversationId);
    if (sentInConversation === undefined) {
      continue;
    }

    const candidate = matchCandidateToSent(
      response,
      sentInConversation,
      excludedSenderAddresses,
      onlyOriginalRecipients,
    );
    if (
      candidate !== undefined &&
      !assignments.has(candidate.sent.message.id)
    ) {
      assignments.set(candidate.sent.message.id, {
        sent: candidate.sent,
        response,
        basis: candidate.basis,
      });
    }
  }

  return assignments;
}

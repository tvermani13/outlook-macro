import assert from "node:assert/strict";
import { test } from "node:test";
import {
  assignFirstResponses,
  recipientAddresses,
} from "../src/matcher.js";
import type {
  GraphMessage,
  TrackedSentMessage,
} from "../src/types.js";

function address(value: string) {
  return { emailAddress: { address: value } };
}

function sent(
  id: string,
  sentAt: string,
  internetMessageId: string,
  recipient: string,
  conversationId = "conversation-1",
): TrackedSentMessage {
  const message: GraphMessage = {
    id,
    subject: `Tracked ${id}`,
    sentDateTime: sentAt,
    internetMessageId,
    conversationId,
    sender: address("friend@example.com"),
    toRecipients: [address(recipient)],
  };
  return {
    mailboxId: "friend",
    message,
    subjectMatch: { ruleName: "test", portion: id },
    originalRecipientAddresses: recipientAddresses(message),
  };
}

function response(options: {
  id: string;
  receivedAt: string;
  sender?: string;
  inReplyTo?: string;
  references?: string;
  conversationId?: string;
}): GraphMessage {
  const headers = [];
  if (options.inReplyTo !== undefined) {
    headers.push({ name: "In-Reply-To", value: options.inReplyTo });
  }
  if (options.references !== undefined) {
    headers.push({ name: "References", value: options.references });
  }
  return {
    id: options.id,
    subject: `Re: ${options.id}`,
    receivedDateTime: options.receivedAt,
    conversationId: options.conversationId ?? "conversation-1",
    sender: address(options.sender ?? "recipient@example.com"),
    internetMessageHeaders: headers,
  };
}

test("uses the earliest direct response and ignores later followups", () => {
  const tracked = [
    sent(
      "sent-1",
      "2026-07-01T10:00:00Z",
      "<sent-1@example.com>",
      "recipient@example.com",
    ),
  ];
  const responses = [
    response({
      id: "response-2",
      receivedAt: "2026-07-01T12:00:00Z",
      inReplyTo: "<sent-1@example.com>",
    }),
    response({
      id: "response-1",
      receivedAt: "2026-07-01T11:00:00Z",
      inReplyTo: "<sent-1@example.com>",
    }),
  ];

  const assignments = assignFirstResponses(
    tracked,
    responses,
    new Set(["friend@example.com"]),
    true,
  );

  assert.equal(assignments.size, 1);
  assert.equal(assignments.get("sent-1")?.response.id, "response-1");
  assert.equal(assignments.get("sent-1")?.basis, "in-reply-to");
});

test("excludes messages sent by either tracked mailbox owner", () => {
  const tracked = [
    sent(
      "sent-1",
      "2026-07-01T10:00:00Z",
      "<sent-1@example.com>",
      "recipient@example.com",
    ),
  ];
  const responses = [
    response({
      id: "internal-followup",
      receivedAt: "2026-07-01T10:30:00Z",
      sender: "teammate@example.com",
      inReplyTo: "<sent-1@example.com>",
    }),
    response({
      id: "external-response",
      receivedAt: "2026-07-01T11:00:00Z",
      inReplyTo: "<sent-1@example.com>",
    }),
  ];

  const assignments = assignFirstResponses(
    tracked,
    responses,
    new Set(["friend@example.com", "teammate@example.com"]),
    true,
  );

  assert.equal(
    assignments.get("sent-1")?.response.id,
    "external-response",
  );
});

test("references header is used when in-reply-to is unavailable", () => {
  const tracked = [
    sent(
      "sent-1",
      "2026-07-01T10:00:00Z",
      "<sent-1@example.com>",
      "recipient@example.com",
    ),
  ];
  const responses = [
    response({
      id: "response-1",
      receivedAt: "2026-07-01T11:00:00Z",
      references: "<older@example.com> <sent-1@example.com>",
    }),
  ];

  const assignments = assignFirstResponses(
    tracked,
    responses,
    new Set(["friend@example.com"]),
    true,
  );

  assert.equal(assignments.get("sent-1")?.basis, "references");
});

test("conversation fallback assigns a reply to the latest prior sent message", () => {
  const tracked = [
    sent(
      "sent-1",
      "2026-07-01T10:00:00Z",
      "<sent-1@example.com>",
      "recipient@example.com",
    ),
    sent(
      "sent-2",
      "2026-07-02T10:00:00Z",
      "<sent-2@example.com>",
      "recipient@example.com",
    ),
  ];
  const responses = [
    response({
      id: "response-1",
      receivedAt: "2026-07-02T11:00:00Z",
    }),
  ];

  const assignments = assignFirstResponses(
    tracked,
    responses,
    new Set(["friend@example.com"]),
    true,
  );

  assert.equal(assignments.size, 1);
  assert.equal(assignments.get("sent-2")?.response.id, "response-1");
  assert.equal(
    assignments.get("sent-2")?.basis,
    "conversation-and-recipient",
  );
});

test("original-recipient enforcement rejects a different sender", () => {
  const tracked = [
    sent(
      "sent-1",
      "2026-07-01T10:00:00Z",
      "<sent-1@example.com>",
      "recipient@example.com",
    ),
  ];
  const responses = [
    response({
      id: "response-1",
      receivedAt: "2026-07-01T11:00:00Z",
      sender: "other-person@example.com",
      inReplyTo: "<sent-1@example.com>",
    }),
  ];

  const assignments = assignFirstResponses(
    tracked,
    responses,
    new Set(["friend@example.com"]),
    true,
  );

  assert.equal(assignments.size, 0);
});

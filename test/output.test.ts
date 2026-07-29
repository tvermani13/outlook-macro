import assert from "node:assert/strict";
import { test } from "node:test";
import { rowsToCsv } from "../src/csv.js";
import {
  buildOutputRow,
  dateInTimeZone,
  messageContent,
} from "../src/output.js";
import type {
  GraphMessage,
  GraphUser,
  MailboxConfig,
  TrackedSentMessage,
} from "../src/types.js";

test("formats dates using the configured time zone", () => {
  assert.equal(
    dateInTimeZone("2026-07-02T02:30:00Z", "America/New_York"),
    "2026-07-01",
  );
});

test("uses uniqueBody so quoted thread history is not included when available", () => {
  const message: GraphMessage = {
    id: "response-1",
    body: { contentType: "text", content: "New response\nQuoted history" },
    uniqueBody: { contentType: "text", content: "New response" },
  };
  assert.equal(messageContent(message), "New response");
});

test("builds the required table fields and safely quotes CSV", () => {
  const mailbox: MailboxConfig = { id: "friend", label: "Friend" };
  const user: GraphUser = {
    id: "user-1",
    mail: "friend@example.com",
  };
  const sentMessage: GraphMessage = {
    id: "sent-1",
    subject: "Drew Outreach | Acme, Inc. | July",
    sentDateTime: "2026-07-01T14:00:00Z",
    conversationId: "conversation-1",
    sender: { emailAddress: { address: "friend@example.com" } },
    toRecipients: [
      { emailAddress: { name: "Alex", address: "alex@acme.example" } },
    ],
  };
  const sent: TrackedSentMessage = {
    mailboxId: "friend",
    message: sentMessage,
    subjectMatch: { ruleName: "outreach", portion: "Acme, Inc." },
    originalRecipientAddresses: new Set(["alex@acme.example"]),
  };
  const response: GraphMessage = {
    id: "response-1",
    subject: "Re: Drew Outreach | Acme, Inc. | July",
    receivedDateTime: "2026-07-01T15:00:00Z",
    sender: {
      emailAddress: { name: "Alex", address: "alex@acme.example" },
    },
    toRecipients: [{ emailAddress: { address: "friend@example.com" } }],
    uniqueBody: { contentType: "text", content: 'Hello,\n"Interested."' },
  };

  const row = buildOutputRow({
    mailbox,
    user,
    sent,
    response,
    attachments: [],
    matchBasis: "in-reply-to",
    timeZone: "America/New_York",
  });
  const csv = rowsToCsv([row]);

  assert.equal(row.sent_date, "2026-07-01");
  assert.equal(row.sent_subject_portion, "Acme, Inc.");
  assert.equal(row.response_date, "2026-07-01");
  assert.match(csv, /"Acme, Inc\."/);
  assert.match(csv, /"Hello,\n""Interested\."""/);
});

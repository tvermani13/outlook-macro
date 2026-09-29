# Outlook First-Response Tracker

This local TypeScript tool scans configured Outlook mailboxes for sent messages
whose subjects match editable rules, finds the first external response to each
message, and writes a CSV table plus a detailed JSON file.

Feature catalog: [`FEATURES.md`](FEATURES.md). Runbook: [`RUNBOOK.md`](RUNBOOK.md). Workspace architecture: [`../ECOSYSTEM.md`](../ECOSYSTEM.md).

It does not use AI or fuzzy extraction. Subject selection, subject-portion
extraction, sender/recipient extraction, body extraction, attachment metadata,
and response matching are deterministic.

## What the output contains

The first four CSV columns are:

1. `sent_date`
2. `sent_subject_portion`
3. `response_subject`
4. `response_date`

The remaining columns include mailbox, full sent subject, sender, recipients,
first-response content, attachment metadata, exact match basis, message IDs, and
the Outlook link. Relevant sent messages with no response can be retained with
the `no_response` status.

The JSON output contains the same rows plus per-mailbox scan summaries.

## How a response is matched

For each incoming message, the matcher uses this order:

1. The RFC `In-Reply-To` header exactly references a tracked sent message.
2. The RFC `References` header contains the tracked sent message.
3. The Outlook `conversationId` matches and the sender was an original
   recipient.

The matcher excludes both configured mailbox owners and any addresses in
`scan.excludeSenderAddresses`. It assigns an incoming message to only one sent
message. Once a sent message has an external response, later thread messages are
ignored.

Set `scan.onlyResponsesFromOriginalRecipients` to `false` if replies can
legitimately come from another address or a colleague of the original recipient.

## One-time Microsoft setup

Create a Microsoft Entra app registration:

1. In the Microsoft Entra admin center, open **App registrations** and select
   **New registration**.
2. Use the account type appropriate for the organization. A single-tenant
   registration is simplest when both mailboxes are in one Microsoft 365 tenant.
3. On **Authentication**, enable **Allow public client flows**. Device-code
   sign-in does not need a redirect URI or client secret.
4. Under **API permissions**, add Microsoft Graph delegated permissions:
   `User.Read` and `Mail.Read`.
5. Copy the **Application (client) ID** and **Directory (tenant) ID**.

An organization may require an administrator to approve `Mail.Read`.

## Install and configure

Node.js 22 or newer is required.

```sh
npm install
cp .env.example .env
cp config.example.json config.json
```

Edit `.env`:

```dotenv
MS_CLIENT_ID=the-application-client-id
MS_TENANT_ID=the-directory-tenant-id
```

Edit `config.json` with the two mailbox addresses and the real subject pattern.
Then validate it:

```sh
npm run validate-config
```

Both `.env` and `config.json` are ignored by Git.

## Dynamic subject rules

Rules are evaluated from top to bottom. Each rule uses JavaScript regular
expression syntax. A single expression can both filter the sent messages and
capture the portion that should appear in the table:

```json
{
  "name": "outreach",
  "enabled": true,
  "matchPattern": "^Drew Outreach\\s*\\|\\s*(?<portion>[^|]+?)\\s*\\|",
  "matchFlags": "i",
  "extractGroup": "portion",
  "onExtractionFailure": "skip"
}
```

This example turns:

```text
Drew Outreach | Acme Corporation | July
```

into `Acme Corporation` in `sent_subject_portion`.

If filtering and extraction need different expressions, add
`extractPattern` and optionally `extractFlags`. `extractGroup` can be a named
group such as `"portion"` or a numbered group such as `1`.

`onExtractionFailure` supports:

- `skip`: do not track that sent message.
- `useSubject`: use the entire sent subject as the portion.
- `error`: stop the scan so the rule can be corrected.

## Authenticate each mailbox

Each mailbox owner signs in once:

```sh
npm run auth -- --mailbox friend
npm run auth -- --mailbox teammate
```

Or run both prompts in sequence:

```sh
npm run auth -- --mailbox all
```

Microsoft displays a device-login URL and short code. The expected username in
`config.json` prevents accidentally associating the wrong account with a
mailbox. Tokens are stored using Microsoft Authentication Extensions: macOS
Keychain on Mac, DPAPI on Windows, and LibSecret on Linux. No client secret is
used.

## Scan

```sh
npm run scan
```

By default, the tool scans 90 days of Sent Items and Inbox mail and writes:

```text
output/first-responses.csv
output/first-responses.json
```

The scan is idempotent: rerunning it rebuilds the table from the configured
lookback window and still keeps only the first response.

## Attachments

The default `attachmentMode` is `metadata`. It records attachment name, MIME
type, size, inline status, and Graph ID without writing untrusted files to disk.
The attachment collection is checked for every matched response, including
messages that contain only inline attachments.

To download non-inline file attachments, change:

```json
"attachmentMode": "download"
```

Downloaded files are written below `output/attachments/`. Item attachments and
reference/cloud attachments remain metadata-only.

## Automation

After both accounts have been authenticated, `npm run scan` uses the secure token
cache silently and is suitable for a scheduled job. On macOS, use a LaunchAgent
running under the same signed-in user so it can access that user's Keychain.

Choose the desired interval only after a manual scan succeeds. A 30- or 60-minute
interval is usually enough for a response-tracking table.

## Development checks

```sh
npm run check
npm test
npm run build
```

The tests cover subject extraction, exact reply-header matching, reference
matching, internal-sender exclusion, original-recipient enforcement,
conversation fallback, first-response selection, body extraction, time zones,
and CSV escaping.

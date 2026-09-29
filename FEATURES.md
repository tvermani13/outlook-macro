# Outlook First-Response Tracker (`drew-email-macro`) — Feature Catalog

Last reviewed: 2026-09-28 (code unchanged since 2026-07-29)

README: [`README.md`](README.md). Runbook: [`RUNBOOK.md`](RUNBOOK.md). Ecosystem: [`../ECOSYSTEM.md`](../ECOSYSTEM.md).

## Purpose

Deterministic local TypeScript tool: scan configured Outlook mailboxes for sent messages matching editable subject regex rules, find the first external reply (In-Reply-To → References → conversationId), and emit CSV + JSON. **No AI.** Not in the Life Orchestrator registry.

## Feature catalog

| Feature | Behavior |
| --- | --- |
| Device-code auth | Per mailbox (`npm run auth -- --mailbox friend\|teammate\|all`) with expected-username check. Tokens in OS keystore (Keychain / DPAPI / LibSecret). |
| `validate-config` | Config lint. |
| `scan` | Idempotent rebuild of the lookback window. |
| Subject rules | Match/extract patterns, named/numbered groups, `onExtractionFailure`: skip \| useSubject \| error. |
| Reply matching | In-Reply-To → References → conversationId + original recipient. |
| Exclusions | Mailbox owners + `excludeSenderAddresses`. Optional `onlyResponsesFromOriginalRecipients`. |
| Attachments | `metadata` (default) or `download` under `output/attachments/`. |
| Output | `output/first-responses.csv|.json`; Outlook deep links; timezone-aware dates; `includeNoResponseRows`. |
| Defaults | 90-day lookback, Sent Items + Inbox, max 10k msgs/folder. |
| Tests | Matcher, subject rules, CSV escaping, TZ, body extraction. |

## Integrations

Microsoft Graph `https://graph.microsoft.com/v1.0`. Authority `https://login.microsoftonline.com/${tenantId}`. Scopes: `User.Read` + `Mail.Read`. Env names: `MS_CLIENT_ID`, `MS_TENANT_ID`. Config: `config.json` (gitignored).

## Gaps

No fuzzy matching by design. Org admin consent may be required for Mail.Read. Attachment download is still metadata-only for item/cloud attachments. Suitable for a LaunchAgent once silent token cache works. No coupling to other `~/Developer` apps.

The GitHub repo is **public** (`tvermani13/outlook-macro`). History was checked on 2026-09-28: `.env`, `config.json`, `.cache/`, and `dist/` have never been committed, and the example config uses `example.com` addresses only.

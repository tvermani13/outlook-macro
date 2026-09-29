# Outlook First-Response Tracker runbook

Last reviewed: 2026-09-28

Local CLI with no services. Configuration and matching rules are documented in
[`README.md`](README.md); features are in [`FEATURES.md`](FEATURES.md).

## Routine run

```bash
npm install                                   # Node 22+
npm run validate-config                       # lints config.json (gitignored)
npm run scan                                  # rebuilds the lookback window; idempotent
open output/first-responses.csv               # plus output/first-responses.json
```

## Sign-in

```bash
npm run auth -- --mailbox all                 # device-code flow per mailbox; checks the expected username
```

Tokens are cached through MSAL persistence. The 0-byte `.cache/<mailbox>.cache`
files are lock/markers; the secrets live in the macOS Keychain. If `scan`
reports "No cached sign-in", rerun `auth` for that mailbox. An organization may
require admin consent for `Mail.Read`.

## Scheduling

Not scheduled. Once silent token refresh has held up for a few weeks, a
LaunchAgent running `npm run scan` daily is enough. Keep the output folder
private: it contains mail content.

## Tests

```bash
npm test          # matcher, subject rules, CSV escaping, timezone, body extraction
npm run check     # TypeScript, no emit
```

## Public repository hygiene

The repo is public. Never commit `.env`, `config.json`, `.cache/`, `output/`, or
`dist/`; all are gitignored. Keep example addresses on `example.com`.

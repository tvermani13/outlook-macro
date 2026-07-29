import { existsSync } from "node:fs";
import { loadEnvFile } from "node:process";
import path from "node:path";
import { getAccessToken } from "./auth.js";
import { loadConfig } from "./config.js";
import { GraphClient } from "./graph.js";
import { runScan } from "./scanner.js";

interface ParsedArguments {
  command: string;
  configPath: string;
  mailboxId?: string;
}

function parseArguments(argv: string[]): ParsedArguments {
  const command = argv[0] ?? "";
  let configPath = "config.json";
  let mailboxId: string | undefined;

  for (let index = 1; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === "--config") {
      const value = argv[index + 1];
      if (value === undefined) {
        throw new Error("--config requires a file path.");
      }
      configPath = value;
      index += 1;
      continue;
    }
    if (argument === "--mailbox") {
      const value = argv[index + 1];
      if (value === undefined) {
        throw new Error("--mailbox requires a mailbox id or 'all'.");
      }
      mailboxId = value;
      index += 1;
      continue;
    }
    throw new Error(`Unknown argument: ${argument}`);
  }

  return {
    command,
    configPath,
    ...(mailboxId === undefined ? {} : { mailboxId }),
  };
}

function loadLocalEnvironment(): void {
  const environmentPath = path.resolve(".env");
  if (existsSync(environmentPath)) {
    loadEnvFile(environmentPath);
  }
}

function usage(): string {
  return [
    "Usage:",
    "  npm run validate-config",
    "  npm run auth -- --mailbox <mailbox-id|all>",
    "  npm run scan",
    "",
    "Optional: --config /path/to/config.json",
  ].join("\n");
}

async function authenticate(
  configPath: string,
  mailboxId: string | undefined,
): Promise<void> {
  const loaded = await loadConfig(configPath);
  const requested = mailboxId ?? "all";
  const mailboxes =
    requested === "all"
      ? loaded.config.mailboxes
      : loaded.config.mailboxes.filter((mailbox) => mailbox.id === requested);
  if (mailboxes.length === 0) {
    throw new Error(`No mailbox with id "${requested}" exists in the config.`);
  }

  const cacheDirectory = path.resolve(loaded.directory, ".cache");
  for (const mailbox of mailboxes) {
    const accessToken = await getAccessToken(mailbox, cacheDirectory, true);
    const user = await new GraphClient(accessToken).getMe();
    console.log(
      `Authenticated ${mailbox.label} as ${user.mail ?? user.userPrincipalName ?? user.displayName ?? user.id}.`,
    );
  }
}

async function main(): Promise<void> {
  loadLocalEnvironment();
  const args = parseArguments(process.argv.slice(2));

  if (args.command === "validate-config") {
    const loaded = await loadConfig(args.configPath);
    console.log(`Config is valid: ${loaded.path}`);
    return;
  }
  if (args.command === "auth") {
    await authenticate(args.configPath, args.mailboxId);
    return;
  }
  if (args.command === "scan") {
    const loaded = await loadConfig(args.configPath);
    const result = await runScan(loaded);
    const sentCount = result.output.summaries.reduce(
      (total, summary) => total + summary.trackedSentMessages,
      0,
    );
    const responseCount = result.output.summaries.reduce(
      (total, summary) => total + summary.responsesFound,
      0,
    );
    console.log(
      `Done. Tracked ${sentCount} sent messages and found ${responseCount} first responses.`,
    );
    console.log(`CSV: ${result.csvPath}`);
    console.log(`JSON: ${result.jsonPath}`);
    return;
  }

  throw new Error(usage());
}

main().catch((error: unknown) => {
  const message = error instanceof Error ? error.message : String(error);
  console.error(`\nError: ${message}`);
  process.exitCode = 1;
});

import { readFile } from "node:fs/promises";
import path from "node:path";
import type {
  AppConfig,
  AttachmentMode,
  ExtractionFailureMode,
  LoadedConfig,
  MailboxConfig,
  OutputConfig,
  ScanConfig,
  SubjectRuleConfig,
} from "./types.js";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function requireRecord(
  value: unknown,
  location: string,
): Record<string, unknown> {
  if (!isRecord(value)) {
    throw new Error(`${location} must be an object.`);
  }
  return value;
}

function requireString(value: unknown, location: string): string {
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`${location} must be a non-empty string.`);
  }
  return value;
}

function optionalString(value: unknown, location: string): string | undefined {
  if (value === undefined) {
    return undefined;
  }
  return requireString(value, location);
}

function optionalBoolean(
  value: unknown,
  location: string,
): boolean | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (typeof value !== "boolean") {
    throw new Error(`${location} must be true or false.`);
  }
  return value;
}

function requireBoolean(value: unknown, location: string): boolean {
  const result = optionalBoolean(value, location);
  if (result === undefined) {
    throw new Error(`${location} is required.`);
  }
  return result;
}

function requirePositiveInteger(value: unknown, location: string): number {
  if (
    typeof value !== "number" ||
    !Number.isInteger(value) ||
    value <= 0
  ) {
    throw new Error(`${location} must be a positive integer.`);
  }
  return value;
}

function requireStringArray(value: unknown, location: string): string[] {
  if (!Array.isArray(value) || value.length === 0) {
    throw new Error(`${location} must be a non-empty string array.`);
  }

  return value.map((item, index) =>
    requireString(item, `${location}[${index}]`),
  );
}

function parseMailbox(value: unknown, index: number): MailboxConfig {
  const item = requireRecord(value, `mailboxes[${index}]`);
  const id = requireString(item.id, `mailboxes[${index}].id`);
  if (!/^[a-z0-9_-]+$/i.test(id)) {
    throw new Error(
      `mailboxes[${index}].id may contain only letters, numbers, "_" and "-".`,
    );
  }

  const expectedUsername = optionalString(
    item.expectedUsername,
    `mailboxes[${index}].expectedUsername`,
  );

  return {
    id,
    label: requireString(item.label, `mailboxes[${index}].label`),
    ...(expectedUsername === undefined ? {} : { expectedUsername }),
  };
}

function validateRegex(pattern: string, flags: string, location: string): void {
  try {
    new RegExp(pattern, flags);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`${location} is not a valid regular expression: ${message}`);
  }
}

function parseSubjectRule(value: unknown, index: number): SubjectRuleConfig {
  const location = `subjectRules[${index}]`;
  const item = requireRecord(value, location);
  const matchPattern = requireString(
    item.matchPattern,
    `${location}.matchPattern`,
  );
  const matchFlags =
    optionalString(item.matchFlags, `${location}.matchFlags`) ?? "";
  validateRegex(matchPattern, matchFlags, `${location}.matchPattern`);

  const extractPattern = optionalString(
    item.extractPattern,
    `${location}.extractPattern`,
  );
  const extractFlags =
    optionalString(item.extractFlags, `${location}.extractFlags`) ?? matchFlags;
  if (extractPattern !== undefined) {
    validateRegex(extractPattern, extractFlags, `${location}.extractPattern`);
  }

  const rawExtractGroup = item.extractGroup;
  if (
    rawExtractGroup !== undefined &&
    typeof rawExtractGroup !== "string" &&
    typeof rawExtractGroup !== "number"
  ) {
    throw new Error(`${location}.extractGroup must be a string or number.`);
  }

  const rawFailureMode = item.onExtractionFailure ?? "useSubject";
  if (
    rawFailureMode !== "skip" &&
    rawFailureMode !== "useSubject" &&
    rawFailureMode !== "error"
  ) {
    throw new Error(
      `${location}.onExtractionFailure must be "skip", "useSubject", or "error".`,
    );
  }

  const enabled = optionalBoolean(item.enabled, `${location}.enabled`);

  return {
    name: requireString(item.name, `${location}.name`),
    matchPattern,
    ...(matchFlags === "" ? {} : { matchFlags }),
    ...(extractPattern === undefined ? {} : { extractPattern }),
    ...(extractPattern === undefined || extractFlags === ""
      ? {}
      : { extractFlags }),
    ...(rawExtractGroup === undefined
      ? {}
      : { extractGroup: rawExtractGroup }),
    onExtractionFailure: rawFailureMode as ExtractionFailureMode,
    ...(enabled === undefined ? {} : { enabled }),
  };
}

function parseScan(value: unknown): ScanConfig {
  const item = requireRecord(value, "scan");
  const timeZone = requireString(item.timeZone, "scan.timeZone");
  try {
    new Intl.DateTimeFormat("en-US", { timeZone }).format();
  } catch {
    throw new Error(`scan.timeZone "${timeZone}" is not a valid IANA time zone.`);
  }

  const excludeSenderAddresses = Array.isArray(item.excludeSenderAddresses)
    ? item.excludeSenderAddresses.map((address, index) =>
        requireString(address, `scan.excludeSenderAddresses[${index}]`),
      )
    : [];

  return {
    lookbackDays: requirePositiveInteger(
      item.lookbackDays,
      "scan.lookbackDays",
    ),
    timeZone,
    sentFolders: requireStringArray(item.sentFolders, "scan.sentFolders"),
    responseFolders: requireStringArray(
      item.responseFolders,
      "scan.responseFolders",
    ),
    onlyResponsesFromOriginalRecipients: requireBoolean(
      item.onlyResponsesFromOriginalRecipients,
      "scan.onlyResponsesFromOriginalRecipients",
    ),
    excludeSenderAddresses,
    maxMessagesPerFolder: requirePositiveInteger(
      item.maxMessagesPerFolder,
      "scan.maxMessagesPerFolder",
    ),
    includeNoResponseRows: requireBoolean(
      item.includeNoResponseRows,
      "scan.includeNoResponseRows",
    ),
  };
}

function parseOutput(value: unknown): OutputConfig {
  const item = requireRecord(value, "output");
  const attachmentMode = item.attachmentMode;
  if (attachmentMode !== "metadata" && attachmentMode !== "download") {
    throw new Error(
      'output.attachmentMode must be either "metadata" or "download".',
    );
  }

  return {
    directory: requireString(item.directory, "output.directory"),
    csvFile: requireString(item.csvFile, "output.csvFile"),
    jsonFile: requireString(item.jsonFile, "output.jsonFile"),
    attachmentMode: attachmentMode as AttachmentMode,
  };
}

function parseConfig(value: unknown): AppConfig {
  const root = requireRecord(value, "config");
  if (root.version !== 1) {
    throw new Error("config.version must be 1.");
  }
  if (!Array.isArray(root.mailboxes) || root.mailboxes.length === 0) {
    throw new Error("config.mailboxes must contain at least one mailbox.");
  }
  if (!Array.isArray(root.subjectRules) || root.subjectRules.length === 0) {
    throw new Error("config.subjectRules must contain at least one rule.");
  }

  const mailboxes = root.mailboxes.map(parseMailbox);
  const mailboxIds = new Set<string>();
  for (const mailbox of mailboxes) {
    if (mailboxIds.has(mailbox.id)) {
      throw new Error(`Duplicate mailbox id "${mailbox.id}".`);
    }
    mailboxIds.add(mailbox.id);
  }

  const subjectRules = root.subjectRules.map(parseSubjectRule);
  const enabledRules = subjectRules.filter((rule) => rule.enabled !== false);
  if (enabledRules.length === 0) {
    throw new Error("At least one subject rule must be enabled.");
  }

  const ruleNames = new Set<string>();
  for (const rule of subjectRules) {
    if (ruleNames.has(rule.name)) {
      throw new Error(`Duplicate subject rule name "${rule.name}".`);
    }
    ruleNames.add(rule.name);
  }

  return {
    version: 1,
    mailboxes,
    subjectRules,
    scan: parseScan(root.scan),
    output: parseOutput(root.output),
  };
}

export async function loadConfig(configPath: string): Promise<LoadedConfig> {
  const absolutePath = path.resolve(configPath);
  let source: string;
  try {
    source = await readFile(absolutePath, "utf8");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`Unable to read config at ${absolutePath}: ${message}`);
  }

  let json: unknown;
  try {
    json = JSON.parse(source);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`Config is not valid JSON: ${message}`);
  }

  return {
    config: parseConfig(json),
    path: absolutePath,
    directory: path.dirname(absolutePath),
  };
}

import type { OutputRow } from "./types.js";

export const OUTPUT_COLUMNS: readonly (keyof OutputRow)[] = [
  "sent_date",
  "sent_subject_portion",
  "response_subject",
  "response_date",
  "mailbox_id",
  "mailbox_label",
  "mailbox_address",
  "rule_name",
  "status",
  "sent_at",
  "sent_subject",
  "sent_sender",
  "sent_recipients",
  "response_at",
  "response_sender",
  "response_recipients",
  "response_content",
  "response_attachments",
  "response_attachment_files",
  "match_basis",
  "sent_message_id",
  "response_message_id",
  "response_web_link",
];

function escapeCsvValue(value: string): string {
  if (/[",\r\n]/.test(value)) {
    return `"${value.replaceAll('"', '""')}"`;
  }
  return value;
}

export function rowsToCsv(rows: OutputRow[]): string {
  const header = OUTPUT_COLUMNS.join(",");
  const body = rows.map((row) =>
    OUTPUT_COLUMNS.map((column) => escapeCsvValue(row[column])).join(","),
  );
  return `${[header, ...body].join("\n")}\n`;
}

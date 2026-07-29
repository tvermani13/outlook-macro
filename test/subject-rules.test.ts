import assert from "node:assert/strict";
import { test } from "node:test";
import { matchSubject } from "../src/subject-rules.js";
import type { SubjectRuleConfig } from "../src/types.js";

test("matches a configurable subject and extracts a named group", () => {
  const rules: SubjectRuleConfig[] = [
    {
      name: "outreach",
      matchPattern: "^Drew Outreach\\s*\\|\\s*(?<portion>[^|]+?)\\s*\\|",
      matchFlags: "i",
      extractGroup: "portion",
      onExtractionFailure: "skip",
    },
  ];

  assert.deepEqual(
    matchSubject("DREW OUTREACH | Acme Corporation | July", rules),
    {
      ruleName: "outreach",
      portion: "Acme Corporation",
    },
  );
});

test("uses a separate extraction pattern when configured", () => {
  const rules: SubjectRuleConfig[] = [
    {
      name: "campaign",
      matchPattern: "^Campaign:",
      matchFlags: "i",
      extractPattern: "\\[(?<key>[^\\]]+)\\]",
      extractGroup: "key",
      onExtractionFailure: "error",
    },
  ];

  assert.deepEqual(matchSubject("Campaign: Intro [ABC-123]", rules), {
    ruleName: "campaign",
    portion: "ABC-123",
  });
});

test("skips a matched rule when extraction is missing and mode is skip", () => {
  const rules: SubjectRuleConfig[] = [
    {
      name: "too-specific",
      matchPattern: "^Campaign:",
      extractPattern: "\\[(?<key>[^\\]]+)\\]",
      extractGroup: "key",
      onExtractionFailure: "skip",
    },
    {
      name: "fallback",
      matchPattern: "^Campaign:\\s*(?<portion>.+)$",
      extractGroup: "portion",
    },
  ];

  assert.deepEqual(matchSubject("Campaign: Acme", rules), {
    ruleName: "fallback",
    portion: "Acme",
  });
});

test("does not match unrelated subjects", () => {
  const rules: SubjectRuleConfig[] = [
    {
      name: "outreach",
      matchPattern: "^Drew Outreach",
    },
  ];

  assert.equal(matchSubject("Weekly team status", rules), undefined);
});

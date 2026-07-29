import type { SubjectMatch, SubjectRuleConfig } from "./types.js";

function extractGroup(
  match: RegExpExecArray,
  group: string | number | undefined,
): string | undefined {
  if (group === undefined) {
    return match[0];
  }
  if (typeof group === "number") {
    return match[group];
  }
  return match.groups?.[group];
}

export function matchSubject(
  subject: string,
  rules: SubjectRuleConfig[],
): SubjectMatch | undefined {
  for (const rule of rules) {
    if (rule.enabled === false) {
      continue;
    }

    const match = new RegExp(rule.matchPattern, rule.matchFlags ?? "").exec(
      subject,
    );
    if (match === null) {
      continue;
    }

    const extractionMatch =
      rule.extractPattern === undefined
        ? match
        : new RegExp(
            rule.extractPattern,
            rule.extractFlags ?? rule.matchFlags ?? "",
          ).exec(subject);
    const extracted =
      extractionMatch === null
        ? undefined
        : extractGroup(extractionMatch, rule.extractGroup)?.trim();

    if (extracted !== undefined && extracted !== "") {
      return { ruleName: rule.name, portion: extracted };
    }

    const failureMode = rule.onExtractionFailure ?? "useSubject";
    if (failureMode === "useSubject") {
      return { ruleName: rule.name, portion: subject.trim() };
    }
    if (failureMode === "error") {
      throw new Error(
        `Subject matched rule "${rule.name}" but its extraction did not produce a value: ${subject}`,
      );
    }
  }

  return undefined;
}

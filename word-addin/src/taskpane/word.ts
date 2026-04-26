/* global Office, Word */

import { ApplicationMode, IssueApplicationSummary, ProofreadIssue, ProofreadScope } from "./types";

export async function getSelectedText(): Promise<string> {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const selectedText = (selection.text || "").trim();
    if (selectedText.length === 0) {
      throw new Error("请先在 Word 中选中一段文字。");
    }

    return selectedText;
  });
}

export async function getDocumentBodyText(): Promise<string> {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const bodyText = (body.text || "").trim();
    if (bodyText.length === 0) {
      throw new Error("当前 Word 正文为空，无法进行全书审校。");
    }

    return bodyText;
  });
}

export function ensureWordCommentSupport(): boolean {
  return Office.context.requirements.isSetSupported("WordApi", "1.4");
}

export async function applyIssuesToScope(
  sourceText: string,
  issues: ProofreadIssue[],
  scope: ProofreadScope,
  applicationMode: ApplicationMode
): Promise<IssueApplicationSummary> {
  if (scope === "document") {
    return applyIssuesToDocument(sourceText, issues, applicationMode);
  }

  return applyIssuesToSelection(sourceText, issues, applicationMode);
}

export async function insertUnlocatedSummaryComment(
  issues: ProofreadIssue[],
  scope: ProofreadScope
): Promise<IssueApplicationSummary> {
  const unlocatedIssues = issues.filter((issue) => !hasUsableLocation(issue));
  if (unlocatedIssues.length === 0) {
    return { commentCount: 0, revisionCount: 0, fallbackCount: 0 };
  }

  return Word.run(async (context) => {
    const target =
      scope === "document" ? context.document.body.getRange() : context.document.getSelection();
    target.insertComment(formatFallbackComment(unlocatedIssues));
    await context.sync();
    return { commentCount: 0, revisionCount: 0, fallbackCount: unlocatedIssues.length };
  });
}

function applyIssuesToSelection(
  selectedText: string,
  issues: ProofreadIssue[],
  applicationMode: ApplicationMode
): Promise<IssueApplicationSummary> {
  if (applicationMode === "revision") {
    return applyRevisionsForIssues(selectedText, issues, "selection");
  }

  return insertCommentsForIssues(selectedText, issues, "selection");
}

function applyIssuesToDocument(
  documentText: string,
  issues: ProofreadIssue[],
  applicationMode: ApplicationMode
): Promise<IssueApplicationSummary> {
  if (applicationMode === "revision") {
    return applyRevisionsForIssues(documentText, issues, "document");
  }

  return insertCommentsForIssues(documentText, issues, "document");
}

function insertCommentsForIssues(
  sourceText: string,
  issues: ProofreadIssue[],
  scope: ProofreadScope
): Promise<IssueApplicationSummary> {
  return Word.run(async (context) => {
    const searchRoot =
      scope === "document" ? context.document.body : context.document.getSelection();
    const pendingSearches: Array<{
      issue: ProofreadIssue;
      occurrenceIndex: number;
      searchResults: Word.RangeCollection;
    }> = [];
    let commentCount = 0;
    let fallbackCount = 0;

    for (const issue of issues) {
      if (!isIssueLocatable(sourceText, issue)) {
        fallbackCount += 1;
        continue;
      }

      const occurrenceIndex = getOccurrenceIndexBeforeOffset(
        sourceText,
        issue.original,
        issue.start as number
      );
      const searchResults = searchRoot.search(issue.original, {
        matchCase: true,
        matchWholeWord: false,
      });
      searchResults.load("items");
      pendingSearches.push({ issue, occurrenceIndex, searchResults });
    }

    await context.sync();

    pendingSearches.forEach(({ issue, occurrenceIndex, searchResults }) => {
      const targetRange = searchResults.items[occurrenceIndex];

      if (!targetRange) {
        fallbackCount += 1;
        return;
      }

      targetRange.insertComment(formatIssueComment(issue));
      commentCount += 1;
    });

    await context.sync();
    return { commentCount, revisionCount: 0, fallbackCount };
  });
}

function applyRevisionsForIssues(
  sourceText: string,
  issues: ProofreadIssue[],
  scope: ProofreadScope
): Promise<IssueApplicationSummary> {
  return Word.run(async (context) => {
    const document = context.document;
    const searchRoot = scope === "document" ? document.body : document.getSelection();
    const pendingSearches: Array<{
      issue: ProofreadIssue;
      occurrenceIndex: number;
      searchResults: Word.RangeCollection;
    }> = [];
    let revisionCount = 0;
    let fallbackCount = 0;

    document.load("changeTrackingMode");

    for (const issue of issues) {
      if (!isIssueLocatable(sourceText, issue) || !hasReplacement(issue)) {
        fallbackCount += 1;
        continue;
      }

      const occurrenceIndex = getOccurrenceIndexBeforeOffset(
        sourceText,
        issue.original,
        issue.start as number
      );
      const searchResults = searchRoot.search(issue.original, {
        matchCase: true,
        matchWholeWord: false,
      });
      searchResults.load("items");
      pendingSearches.push({ issue, occurrenceIndex, searchResults });
    }

    await context.sync();

    const originalTrackingMode = document.changeTrackingMode;

    try {
      if (pendingSearches.length > 0) {
        document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;

        pendingSearches
          .sort((left, right) => (right.issue.start || 0) - (left.issue.start || 0))
          .forEach(({ issue, occurrenceIndex, searchResults }) => {
            const targetRange = searchResults.items[occurrenceIndex];

            if (!targetRange || !issue.replacement) {
              fallbackCount += 1;
              return;
            }

            targetRange.insertText(issue.replacement, Word.InsertLocation.replace);
            revisionCount += 1;
          });

        await context.sync();
      }
    } finally {
      document.changeTrackingMode = originalTrackingMode;
      await context.sync();
    }

    return { commentCount: 0, revisionCount, fallbackCount };
  });
}

function isIssueLocatable(sourceText: string, issue: ProofreadIssue): boolean {
  if (!hasUsableLocation(issue)) {
    return false;
  }

  return sourceText.slice(issue.start as number, issue.end as number) === issue.original;
}

function hasUsableLocation(issue: ProofreadIssue): boolean {
  if (!issue.original || typeof issue.start !== "number" || typeof issue.end !== "number") {
    return false;
  }

  return issue.start >= 0 && issue.end >= issue.start;
}

function hasReplacement(issue: ProofreadIssue): boolean {
  return typeof issue.replacement === "string" && issue.replacement.trim().length > 0;
}

function getOccurrenceIndexBeforeOffset(
  text: string,
  original: string,
  targetStart: number
): number {
  let count = 0;
  let searchFrom = 0;

  while (searchFrom < targetStart) {
    const foundAt = text.indexOf(original, searchFrom);

    if (foundAt === -1 || foundAt >= targetStart) {
      break;
    }

    count += 1;
    searchFrom = foundAt + original.length;
  }

  return count;
}

function formatIssueComment(issue: ProofreadIssue): string {
  return [
    issue.replacement ? `替换为：${issue.replacement}` : "",
    `建议：${issue.suggestion || "未提供"}`,
    `类别：${issue.category} / ${issue.severity}`,
  ]
    .filter(Boolean)
    .join("\n");
}

function formatFallbackComment(issues: ProofreadIssue[]): string {
  return `AI 审校：以下 ${issues.length} 条建议未能定位到具体原文片段，已汇总为单条批注。\n\n${formatComment(issues)}`;
}

function formatComment(issues: ProofreadIssue[]): string {
  if (issues.length === 0) {
    return "AI 审校：未发现明显问题。";
  }

  const lines = ["AI 审校建议："];

  issues.forEach((issue, index) => {
    lines.push("");
    lines.push(`${index + 1}. [${issue.severity}] ${issue.category}`);
    lines.push(`原文：${issue.original || "未提供"}`);
    if (issue.replacement) {
      lines.push(`替换为：${issue.replacement}`);
    }
    lines.push(`建议：${issue.suggestion || "未提供"}`);
  });

  return lines.join("\n");
}

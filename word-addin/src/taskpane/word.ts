/* global Office, Word */

import { ApplicationMode, IssueApplicationSummary, ProofreadIssue, ProofreadScope } from "./types";

type SearchRoot = Word.Body | Word.Range;
const FALLBACK_COMMENT_PREVIEW_LIMIT = 10;

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
    const searchRoot = getSearchRoot(context, scope);
    const pendingSearches: Array<{
      issue: ProofreadIssue;
      occurrenceIndex: number;
      searchResults: Word.RangeCollection;
    }> = [];
    const summaryIssues: ProofreadIssue[] = [];
    const fallbackAnchorSearch = createFallbackAnchorSearch(searchRoot, sourceText);
    let commentCount = 0;

    for (const issue of issues) {
      if (!isIssueLocatable(sourceText, issue)) {
        summaryIssues.push(issue);
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
        summaryIssues.push(issue);
        return;
      }

      targetRange.insertComment(formatIssueComment(issue));
      commentCount += 1;
    });

    if (summaryIssues.length > 0) {
      getFallbackAnchorRange(searchRoot, scope, fallbackAnchorSearch).insertComment(
        formatFallbackComment(summaryIssues)
      );
    }

    await context.sync();
    return { commentCount, revisionCount: 0, fallbackCount: summaryIssues.length };
  });
}

function applyRevisionsForIssues(
  sourceText: string,
  issues: ProofreadIssue[],
  scope: ProofreadScope
): Promise<IssueApplicationSummary> {
  return Word.run(async (context) => {
    const document = context.document;
    const searchRoot = getSearchRoot(context, scope);
    const pendingRevisionSearches: Array<{
      issue: ProofreadIssue;
      occurrenceIndex: number;
      searchResults: Word.RangeCollection;
    }> = [];
    const pendingCommentSearches: Array<{
      issue: ProofreadIssue;
      occurrenceIndex: number;
      searchResults: Word.RangeCollection;
    }> = [];
    const summaryIssues: ProofreadIssue[] = [];
    const fallbackAnchorSearch = createFallbackAnchorSearch(searchRoot, sourceText);
    let revisionCount = 0;
    let commentCount = 0;

    document.load("changeTrackingMode");

    for (const issue of issues) {
      if (!isIssueLocatable(sourceText, issue)) {
        summaryIssues.push(issue);
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

      if (hasReplacement(issue)) {
        pendingRevisionSearches.push({ issue, occurrenceIndex, searchResults });
      } else {
        pendingCommentSearches.push({ issue, occurrenceIndex, searchResults });
      }
    }

    await context.sync();

    const revisionApplications: Array<{ issue: ProofreadIssue; targetRange: Word.Range }> = [];

    pendingRevisionSearches.forEach(({ issue, occurrenceIndex, searchResults }) => {
      const targetRange = searchResults.items[occurrenceIndex];

      if (!targetRange) {
        summaryIssues.push(issue);
        return;
      }

      revisionApplications.push({ issue, targetRange });
    });

    pendingCommentSearches.forEach(({ issue, occurrenceIndex, searchResults }) => {
      const targetRange = searchResults.items[occurrenceIndex];

      if (!targetRange) {
        summaryIssues.push(issue);
        return;
      }

      targetRange.insertComment(formatIssueComment(issue));
      commentCount += 1;
    });

    if (summaryIssues.length > 0) {
      getFallbackAnchorRange(searchRoot, scope, fallbackAnchorSearch).insertComment(
        formatFallbackComment(summaryIssues)
      );
    }

    await context.sync();

    const originalTrackingMode = document.changeTrackingMode;

    try {
      if (revisionApplications.length > 0) {
        document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;

        revisionApplications
          .sort((left, right) => (right.issue.start || 0) - (left.issue.start || 0))
          .forEach(({ issue, targetRange }) => {
            targetRange.insertText(issue.replacement as string, Word.InsertLocation.replace);
            revisionCount += 1;
          });

        await context.sync();
      }
    } finally {
      document.changeTrackingMode = originalTrackingMode;
      await context.sync();
    }

    return { commentCount, revisionCount, fallbackCount: summaryIssues.length };
  });
}

function getSearchRoot(context: Word.RequestContext, scope: ProofreadScope): SearchRoot {
  return scope === "document" ? context.document.body : context.document.getSelection();
}

function createFallbackAnchorSearch(
  searchRoot: SearchRoot,
  sourceText: string
): { occurrenceIndex: number; searchResults: Word.RangeCollection | null } {
  const firstVisibleOffset = findFirstNonWhitespaceOffset(sourceText);

  if (firstVisibleOffset === -1) {
    return { occurrenceIndex: 0, searchResults: null };
  }

  const anchorText = sourceText[firstVisibleOffset];
  const searchResults = searchRoot.search(anchorText, {
    matchCase: true,
    matchWholeWord: false,
  });
  searchResults.load("items");

  return {
    occurrenceIndex: getOccurrenceIndexBeforeOffset(sourceText, anchorText, firstVisibleOffset),
    searchResults,
  };
}

function getFallbackAnchorRange(
  searchRoot: SearchRoot,
  scope: ProofreadScope,
  anchorSearch: { occurrenceIndex: number; searchResults: Word.RangeCollection | null }
): Word.Range {
  const anchorRange = anchorSearch.searchResults?.items[anchorSearch.occurrenceIndex];

  if (anchorRange) {
    return anchorRange;
  }

  return scope === "document" ? (searchRoot as Word.Body).getRange() : (searchRoot as Word.Range);
}

function findFirstNonWhitespaceOffset(text: string): number {
  const match = /\S/.exec(text);
  return match ? match.index : -1;
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
  const visibleIssues = issues.slice(0, FALLBACK_COMMENT_PREVIEW_LIMIT);
  const remainingCount = issues.length - visibleIssues.length;
  const remainingMessage =
    remainingCount > 0
      ? `\n\n另有 ${remainingCount} 条未定位建议，请在任务窗格中查看完整结果。`
      : "";

  return `AI 审校：以下 ${issues.length} 条建议未能定位到具体原文片段，已汇总为单条批注。\n\n${formatComment(visibleIssues)}${remainingMessage}`;
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

/* global Office, Word */

import { appendDebugLog } from "./debug";
import {
  ApplicationMode,
  IssueApplicationProgress,
  IssueApplicationSummary,
  ProofreadIssue,
  ProofreadLocator,
  ProofreadScope,
} from "./types";

type SearchRoot = Word.Body | Word.Range;

interface SearchContext {
  root: SearchRoot;
  rootKind: "body" | "selection";
  occurrenceText: string;
  canSearch: boolean;
}

interface ApplyIssuesOptions {
  onProgress?: (progress: IssueApplicationProgress) => void;
}

interface PreparedIssue {
  issue: ProofreadIssue;
  locator: ProofreadLocator;
}

interface SearchGroup {
  locator: ProofreadLocator;
  occurrenceIndex: number;
  issues: PreparedIssue[];
}

interface PendingOriginalSearch {
  prepared: PreparedIssue;
  searchResults: Word.RangeCollection;
}

interface ResolvedIssueTarget {
  issue: ProofreadIssue;
  targetRange: Word.Range;
}

interface ResolvedIssueTargets {
  targets: ResolvedIssueTarget[];
  summaryIssues: ProofreadIssue[];
  fallbackAnchorRange: Word.Range;
}

const FALLBACK_COMMENT_PREVIEW_LIMIT = 10;
const SEARCH_BATCH_SIZE = 8;
const MIN_ORIGINAL_LOCATOR_LENGTH = 6;
const MAX_LOCATOR_OCCURRENCES = 3;
const CONTEXT_LOCATOR_WINDOWS = [16, 32, 64];

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
  applicationMode: ApplicationMode,
  options: ApplyIssuesOptions = {}
): Promise<IssueApplicationSummary> {
  if (scope === "document") {
    return applyIssuesToDocument(sourceText, issues, applicationMode, options);
  }

  return applyIssuesToSelection(sourceText, issues, applicationMode, options);
}

export async function selectIssueInScope(
  sourceText: string,
  issue: ProofreadIssue,
  scope: ProofreadScope
): Promise<void> {
  return Word.run(async (context) => {
    appendDebugLog("info", "Word.run 定位开始", {
      issueId: issue.id,
      scope,
      start: issue.start,
      end: issue.end,
      original: issue.original,
      sourceTextLength: sourceText.length,
      locatorStrategy: issue.locator?.strategy,
    });

    const searchContext = await createSearchContext(context, scope, sourceText);
    const resolved = await resolveIssueTargets(context, searchContext, sourceText, [issue]);
    const target = resolved.targets[0];

    if (!target) {
      appendDebugLog("warn", "定位失败：未能解析 Word range", {
        issueId: issue.id,
        summaryCount: resolved.summaryIssues.length,
      });
      throw new Error("未能在当前 Word 范围中找到对应原文。");
    }

    target.targetRange.select();
    await context.sync();
    appendDebugLog("info", "Word range 已 select", {
      issueId: issue.id,
    });
  });
}

function applyIssuesToSelection(
  selectedText: string,
  issues: ProofreadIssue[],
  applicationMode: ApplicationMode,
  options: ApplyIssuesOptions
): Promise<IssueApplicationSummary> {
  if (applicationMode === "revision") {
    return applyRevisionsForIssues(selectedText, issues, "selection", options);
  }

  return insertCommentsForIssues(selectedText, issues, "selection", options);
}

function applyIssuesToDocument(
  documentText: string,
  issues: ProofreadIssue[],
  applicationMode: ApplicationMode,
  options: ApplyIssuesOptions
): Promise<IssueApplicationSummary> {
  if (applicationMode === "revision") {
    return applyRevisionsForIssues(documentText, issues, "document", options);
  }

  return insertCommentsForIssues(documentText, issues, "document", options);
}

function insertCommentsForIssues(
  sourceText: string,
  issues: ProofreadIssue[],
  scope: ProofreadScope,
  options: ApplyIssuesOptions
): Promise<IssueApplicationSummary> {
  return Word.run(async (context) => {
    const searchContext = await createSearchContext(context, scope, sourceText);
    const resolved = await resolveIssueTargets(context, searchContext, sourceText, issues, options);
    let commentCount = 0;

    resolved.targets.forEach(({ issue, targetRange }) => {
      targetRange.insertComment(formatIssueComment(issue));
      commentCount += 1;
    });

    if (resolved.summaryIssues.length > 0) {
      resolved.fallbackAnchorRange.insertComment(formatFallbackComment(resolved.summaryIssues));
    }

    await context.sync();
    return { commentCount, revisionCount: 0, fallbackCount: resolved.summaryIssues.length };
  });
}

function applyRevisionsForIssues(
  sourceText: string,
  issues: ProofreadIssue[],
  scope: ProofreadScope,
  options: ApplyIssuesOptions
): Promise<IssueApplicationSummary> {
  return Word.run(async (context) => {
    const document = context.document;
    document.load("changeTrackingMode");

    const searchContext = await createSearchContext(context, scope, sourceText);
    const resolved = await resolveIssueTargets(context, searchContext, sourceText, issues, options);
    const revisionApplications: ResolvedIssueTarget[] = [];
    let revisionCount = 0;
    let commentCount = 0;

    resolved.targets.forEach((target) => {
      if (hasReplacement(target.issue)) {
        revisionApplications.push(target);
        return;
      }

      target.targetRange.insertComment(formatIssueComment(target.issue));
      commentCount += 1;
    });

    if (resolved.summaryIssues.length > 0) {
      resolved.fallbackAnchorRange.insertComment(formatFallbackComment(resolved.summaryIssues));
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

    return { commentCount, revisionCount, fallbackCount: resolved.summaryIssues.length };
  });
}

async function createSearchContext(
  context: Word.RequestContext,
  scope: ProofreadScope,
  sourceText: string
): Promise<SearchContext> {
  if (scope === "document") {
    appendDebugLog("info", "使用正文范围定位，不在应用阶段读取全文", {
      sourceTextLength: sourceText.length,
    });
    return {
      root: context.document.body,
      rootKind: "body",
      occurrenceText: sourceText,
      canSearch: true,
    };
  }

  const selection = context.document.getSelection();
  selection.load("text");
  await context.sync();

  const selectionText = (selection.text || "").trim();
  const canSearch = selectionText === sourceText;

  appendDebugLog(canSearch ? "info" : "warn", "使用当前选区范围定位", {
    sourceTextLength: sourceText.length,
    selectionTextLength: selectionText.length,
    canSearch,
  });

  return {
    root: selection,
    rootKind: "selection",
    occurrenceText: sourceText,
    canSearch,
  };
}

async function resolveIssueTargets(
  context: Word.RequestContext,
  searchContext: SearchContext,
  sourceText: string,
  issues: ProofreadIssue[],
  options: ApplyIssuesOptions = {}
): Promise<ResolvedIssueTargets> {
  const summaryIssues: ProofreadIssue[] = [];
  const targets: ResolvedIssueTarget[] = [];
  let firstTargetRange: Word.Range | null = null;

  if (!searchContext.canSearch) {
    appendDebugLog("warn", "当前 Word 范围与审校文本不一致，全部降级汇总", {
      issueCount: issues.length,
      rootKind: searchContext.rootKind,
    });
    return {
      targets,
      summaryIssues: issues,
      fallbackAnchorRange: getRootRange(searchContext),
    };
  }

  const groups = buildSearchGroups(sourceText, issues, summaryIssues);
  const totalBatches = Math.ceil(groups.length / SEARCH_BATCH_SIZE);
  let completedIssues = summaryIssues.length;

  for (let batchStart = 0; batchStart < groups.length; batchStart += SEARCH_BATCH_SIZE) {
    const batch = groups.slice(batchStart, batchStart + SEARCH_BATCH_SIZE);
    const pendingKeySearches = batch.map((group) => {
      const searchResults = searchContext.root.search(group.locator.key, {
        matchCase: true,
        matchWholeWord: false,
      });
      searchResults.load("items");
      return { group, searchResults };
    });

    // eslint-disable-next-line office-addins/no-context-sync-in-loop -- Intentional batch boundary to keep Word responsive.
    await context.sync();

    const pendingOriginalSearches: PendingOriginalSearch[] = [];

    pendingKeySearches.forEach(({ group, searchResults }) => {
      const keyRange = searchResults.items[group.occurrenceIndex];

      if (!keyRange) {
        appendDebugLog("warn", "Word search 找不到 locator key", {
          keyStart: group.locator.key_start,
          keyLength: group.locator.key.length,
          occurrenceIndex: group.occurrenceIndex,
          searchResultCount: searchResults.items.length,
          issueIds: group.issues.map(({ issue }) => issue.id),
        });
        summaryIssues.push(...group.issues.map(({ issue }) => issue));
        return;
      }

      group.issues.forEach((prepared) => {
        if (prepared.locator.strategy === "original") {
          targets.push({ issue: prepared.issue, targetRange: keyRange });
          firstTargetRange = firstTargetRange || keyRange;
          return;
        }

        const searchResultsInKey = keyRange.search(prepared.issue.original, {
          matchCase: true,
          matchWholeWord: false,
        });
        searchResultsInKey.load("items");
        pendingOriginalSearches.push({ prepared, searchResults: searchResultsInKey });
      });
    });

    if (pendingOriginalSearches.length > 0) {
      // eslint-disable-next-line office-addins/no-context-sync-in-loop -- Intentional nested batch boundary for context locators.
      await context.sync();
    }

    pendingOriginalSearches.forEach(({ prepared, searchResults }) => {
      const occurrenceIndex = getOccurrenceIndexBeforeOffset(
        prepared.locator.key,
        prepared.issue.original,
        prepared.locator.original_start_in_key
      );
      const targetRange = searchResults.items[occurrenceIndex];

      if (!targetRange) {
        appendDebugLog("warn", "locator key 内找不到 original", {
          issueId: prepared.issue.id,
          occurrenceIndex,
          searchResultCount: searchResults.items.length,
        });
        summaryIssues.push(prepared.issue);
        return;
      }

      targets.push({ issue: prepared.issue, targetRange });
      firstTargetRange = firstTargetRange || targetRange;
    });

    completedIssues += batch.reduce((total, group) => total + group.issues.length, 0);
    options.onProgress?.({
      completedBatches: Math.min(Math.floor(batchStart / SEARCH_BATCH_SIZE) + 1, totalBatches),
      totalBatches,
      completedIssues: Math.min(completedIssues, issues.length),
      totalIssues: issues.length,
    });
  }

  return {
    targets,
    summaryIssues,
    fallbackAnchorRange: firstTargetRange || getRootRange(searchContext),
  };
}

function buildSearchGroups(
  sourceText: string,
  issues: ProofreadIssue[],
  summaryIssues: ProofreadIssue[]
): SearchGroup[] {
  const groupsById = new Map<string, SearchGroup>();

  issues.forEach((issue) => {
    const locator = getUsableLocator(sourceText, issue);

    if (!locator) {
      summaryIssues.push(issue);
      return;
    }

    const groupId = `${locator.key_start}:${locator.key_end}:${locator.key}`;
    const existing = groupsById.get(groupId);
    const prepared = { issue, locator };

    if (existing) {
      existing.issues.push(prepared);
      return;
    }

    groupsById.set(groupId, {
      locator,
      occurrenceIndex: getOccurrenceIndexBeforeOffset(sourceText, locator.key, locator.key_start),
      issues: [prepared],
    });
  });

  return Array.from(groupsById.values());
}

function getUsableLocator(sourceText: string, issue: ProofreadIssue): ProofreadLocator | null {
  if (!isIssueLocatable(sourceText, issue)) {
    return null;
  }

  if (isValidLocator(sourceText, issue, issue.locator || null)) {
    return issue.locator as ProofreadLocator;
  }

  return buildClientLocator(sourceText, issue);
}

function isValidLocator(
  sourceText: string,
  issue: ProofreadIssue,
  locator: ProofreadLocator | null
): boolean {
  if (
    !locator ||
    !locator.key ||
    typeof locator.key_start !== "number" ||
    typeof locator.key_end !== "number" ||
    typeof locator.original_start_in_key !== "number" ||
    typeof locator.original_end_in_key !== "number"
  ) {
    return false;
  }

  if (
    locator.key_start < 0 ||
    locator.key_end < locator.key_start ||
    locator.key_end > sourceText.length ||
    locator.original_start_in_key < 0 ||
    locator.original_end_in_key < locator.original_start_in_key ||
    locator.original_end_in_key > locator.key.length
  ) {
    return false;
  }

  const key = sourceText.slice(locator.key_start, locator.key_end);
  const originalInKey = locator.key.slice(
    locator.original_start_in_key,
    locator.original_end_in_key
  );

  return (
    key === locator.key &&
    originalInKey === issue.original &&
    locator.key_start + locator.original_start_in_key === issue.start &&
    locator.key_start + locator.original_end_in_key === issue.end
  );
}

function buildClientLocator(sourceText: string, issue: ProofreadIssue): ProofreadLocator | null {
  if (!hasUsableLocation(issue)) {
    return null;
  }

  const start = issue.start as number;
  const end = issue.end as number;

  if (
    issue.original.length >= MIN_ORIGINAL_LOCATOR_LENGTH &&
    getOccurrenceCount(sourceText, issue.original) <= MAX_LOCATOR_OCCURRENCES
  ) {
    return {
      key: issue.original,
      key_start: start,
      key_end: end,
      original_start_in_key: 0,
      original_end_in_key: issue.original.length,
      strategy: "original",
    };
  }

  for (const windowSize of CONTEXT_LOCATOR_WINDOWS) {
    const keyStart = Math.max(0, start - windowSize);
    const keyEnd = Math.min(sourceText.length, end + windowSize);
    const key = sourceText.slice(keyStart, keyEnd);

    if (key && getOccurrenceCount(sourceText, key) <= MAX_LOCATOR_OCCURRENCES) {
      return {
        key,
        key_start: keyStart,
        key_end: keyEnd,
        original_start_in_key: start - keyStart,
        original_end_in_key: end - keyStart,
        strategy: "context",
      };
    }
  }

  return null;
}

function getRootRange(searchContext: SearchContext): Word.Range {
  return searchContext.rootKind === "body"
    ? (searchContext.root as Word.Body).getRange()
    : (searchContext.root as Word.Range);
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

function getOccurrenceIndexBeforeOffset(text: string, target: string, targetStart: number): number {
  let count = 0;
  let searchFrom = 0;

  while (searchFrom < targetStart) {
    const foundAt = text.indexOf(target, searchFrom);

    if (foundAt === -1 || foundAt >= targetStart) {
      break;
    }

    count += 1;
    searchFrom = foundAt + 1;
  }

  return count;
}

function getOccurrenceCount(text: string, target: string): number {
  if (!target) {
    return 0;
  }

  let count = 0;
  let searchFrom = 0;

  while (searchFrom < text.length) {
    const foundAt = text.indexOf(target, searchFrom);

    if (foundAt === -1) {
      break;
    }

    count += 1;
    searchFrom = foundAt + target.length;
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

  return `AI 审校：以下 ${issues.length} 条建议未能精准写回，已汇总为单条批注。\n\n${formatComment(visibleIssues)}${remainingMessage}`;
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

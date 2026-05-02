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
  fallbackSummaryTruncateEnabled?: boolean;
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
  contextFailed?: boolean;
}

interface WriteBatchResult {
  successCount: number;
  failedIssues: ProofreadIssue[];
  contextFailed?: boolean;
}

interface RevisionWriteBatchResult extends WriteBatchResult {
  commentCount: number;
}

interface CommentBatchResult {
  commentCount: number;
  fallbackCount: number;
  fallbackCommentCount: number;
  failedCount: number;
  truncatedFallbackCount: number;
  retryIssues: ProofreadIssue[];
  fallbackIssues: ProofreadIssue[];
}

interface FallbackWriteResult {
  fallbackCount: number;
  fallbackCommentCount: number;
  failedCount: number;
  truncatedFallbackCount: number;
}

interface FallbackCommentChunk {
  issues: ProofreadIssue[];
  comment: string;
}

interface RevisionBatchResult {
  commentCount: number;
  revisionCount: number;
  retryIssues: ProofreadIssue[];
  fallbackIssues: ProofreadIssue[];
}

interface BestEffortCommentResult {
  commentCount: number;
}

type WordOperationStage =
  | "anchor"
  | "comment"
  | "fallback"
  | "locate"
  | "revision"
  | "search"
  | "select";

interface WordErrorLogDetails {
  batchIndex?: number;
  category?: string;
  commentLength?: number;
  issueCount?: number;
  issueId?: string;
  keyLength?: number;
  replacementLength?: number;
  severity?: string;
  totalBatches?: number;
}

const SEARCH_BATCH_SIZE = 16;
const COMMENT_RUN_BATCH_SIZE = 64;
const COMMENT_INSERT_BATCH_SIZE = 16;
const REVISION_RUN_BATCH_SIZE = 64;
const REVISION_INSERT_BATCH_SIZE = 16;
const MAX_COMMENT_LENGTH = 1500;
const FALLBACK_COMMENT_MAX_LENGTH = 1200;
const FALLBACK_COMMENT_MAX_CHUNKS = 10;
const FALLBACK_COMMENT_HEADER_RESERVE = 120;
const FALLBACK_ISSUE_MAX_LENGTH = 900;
const MIN_ORIGINAL_LOCATOR_LENGTH = 6;
const MAX_LOCATOR_OCCURRENCES = 3;
const CONTEXT_LOCATOR_WINDOWS = [16, 32, 64];

let trackedProofreadSelectionRange: Word.Range | null = null;

export async function getSelectedText(): Promise<string> {
  await clearTrackedSelectionRange();

  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    selection.track();
    await context.sync();

    const selectedText = (selection.text || "").trim();
    if (selectedText.length === 0) {
      selection.untrack();
      await context.sync();
      throw new Error("请先在 Word 中选中一段文字。");
    }

    trackedProofreadSelectionRange = selection;
    appendDebugLog("info", "已缓存送审选区范围", {
      selectedTextLength: selectedText.length,
    });
    return selectedText;
  });
}

export async function getDocumentBodyText(): Promise<string> {
  await clearTrackedSelectionRange();

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

export async function clearTrackedSelectionRange(): Promise<void> {
  const range = trackedProofreadSelectionRange;
  trackedProofreadSelectionRange = null;

  if (!range) {
    return;
  }

  try {
    await Word.run(range, async (context) => {
      range.untrack();
      await context.sync();
    });
    appendDebugLog("info", "已释放送审选区范围");
  } catch (error) {
    appendDebugLog("warn", "释放送审选区范围失败，已丢弃本地引用", {
      error: getWordErrorDetails(error),
    });
  }
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
  return runWordBatchForScope(sourceText, scope, async (context) => {
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

    try {
      target.targetRange.select();
      await context.sync();
      appendDebugLog("info", "Word range 已 select", {
        issueId: issue.id,
      });
    } catch (error) {
      logWordOperationFailure("select", error, {
        category: issue.category,
        issueId: issue.id,
        severity: issue.severity,
      });
      throw new Error("未能在当前 Word 范围中定位该问题。");
    }
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

async function insertCommentsForIssues(
  sourceText: string,
  issues: ProofreadIssue[],
  scope: ProofreadScope,
  options: ApplyIssuesOptions
): Promise<IssueApplicationSummary> {
  const summary = createEmptyApplicationSummary();
  const totalBatches = Math.ceil(issues.length / COMMENT_RUN_BATCH_SIZE);
  const deferredFallbackIssues: ProofreadIssue[] = [];
  let completedIssues = 0;

  for (let batchStart = 0; batchStart < issues.length; batchStart += COMMENT_RUN_BATCH_SIZE) {
    const batch = issues.slice(batchStart, batchStart + COMMENT_RUN_BATCH_SIZE);
    const batchIndex = Math.floor(batchStart / COMMENT_RUN_BATCH_SIZE) + 1;
    let batchResult: CommentBatchResult;

    try {
      batchResult = await insertCommentIssueBatch(sourceText, batch, scope, options);
    } catch (error) {
      logWordOperationFailure("comment", error, {
        batchIndex,
        issueCount: batch.length,
      });
      appendDebugLog("error", "批注应用批次上下文失败，改用单条重试", {
        batchIndex,
        batchSize: batch.length,
        error: getWordErrorDetails(error),
      });
      batchResult = {
        commentCount: 0,
        fallbackCount: 0,
        fallbackCommentCount: 0,
        failedCount: 0,
        truncatedFallbackCount: 0,
        retryIssues: batch,
        fallbackIssues: [],
      };
    }

    addCommentBatchResult(summary, batchResult);

    if (batchResult.retryIssues.length > 0) {
      const retryResult = await retryCommentIssuesInFreshContexts(
        sourceText,
        batchResult.retryIssues,
        scope,
        options
      );
      addCommentBatchResult(summary, retryResult);
      batchResult.fallbackIssues.push(...retryResult.fallbackIssues);
    }

    if (batchResult.fallbackIssues.length > 0) {
      deferredFallbackIssues.push(...batchResult.fallbackIssues);
    }

    completedIssues += batch.length;
    options.onProgress?.({
      stage: "commenting",
      completedBatches: Math.min(batchIndex, totalBatches),
      totalBatches,
      completedIssues: Math.min(completedIssues, issues.length),
      totalIssues: issues.length,
    });
  }

  if (deferredFallbackIssues.length > 0) {
    const fallbackResult = await insertFallbackCommentInFreshContext(
      sourceText,
      scope,
      deferredFallbackIssues,
      options
    );
    summary.fallbackCount += fallbackResult.fallbackCount;
    summary.fallbackCommentCount += fallbackResult.fallbackCommentCount;
    summary.failedCount += fallbackResult.failedCount;
    summary.truncatedFallbackCount += fallbackResult.truncatedFallbackCount;
  }

  return summary;
}

async function insertCommentIssueBatch(
  sourceText: string,
  issues: ProofreadIssue[],
  scope: ProofreadScope,
  options: ApplyIssuesOptions
): Promise<CommentBatchResult> {
  return runWordBatchForScope(sourceText, scope, async (context) => {
    const searchContext = await createSearchContext(context, scope, sourceText);
    const resolved = await resolveIssueTargets(context, searchContext, sourceText, issues, options);

    if (resolved.contextFailed) {
      return {
        commentCount: 0,
        fallbackCount: 0,
        fallbackCommentCount: 0,
        failedCount: 0,
        truncatedFallbackCount: 0,
        retryIssues: issues,
        fallbackIssues: [],
      };
    }

    const commentResult = await insertCommentsInBatches(context, resolved.targets, options);

    if (commentResult.contextFailed) {
      return {
        commentCount: commentResult.successCount,
        fallbackCount: 0,
        fallbackCommentCount: 0,
        failedCount: 0,
        truncatedFallbackCount: 0,
        retryIssues: commentResult.failedIssues,
        fallbackIssues: resolved.summaryIssues,
      };
    }

    return {
      commentCount: commentResult.successCount,
      fallbackCount: 0,
      fallbackCommentCount: 0,
      failedCount: 0,
      truncatedFallbackCount: 0,
      retryIssues: [],
      fallbackIssues: resolved.summaryIssues,
    };
  });
}

async function retryCommentIssuesInFreshContexts(
  sourceText: string,
  issues: ProofreadIssue[],
  scope: ProofreadScope,
  options: ApplyIssuesOptions
): Promise<CommentBatchResult> {
  const result: CommentBatchResult = {
    commentCount: 0,
    fallbackCount: 0,
    fallbackCommentCount: 0,
    failedCount: 0,
    truncatedFallbackCount: 0,
    retryIssues: [],
    fallbackIssues: [],
  };

  for (const issue of issues) {
    try {
      const retryResult = await insertCommentIssueBatch(sourceText, [issue], scope, options);
      result.commentCount += retryResult.commentCount;
      result.fallbackCount += retryResult.fallbackCount;
      result.fallbackCommentCount += retryResult.fallbackCommentCount;
      result.failedCount += retryResult.failedCount;
      result.truncatedFallbackCount += retryResult.truncatedFallbackCount;
      result.fallbackIssues.push(...retryResult.fallbackIssues, ...retryResult.retryIssues);
    } catch (error) {
      logWordOperationFailure("comment", error, {
        category: issue.category,
        commentLength: formatIssueComment(issue).length,
        issueId: issue.id,
        severity: issue.severity,
      });
      appendDebugLog("error", "批注单条新上下文重试失败", {
        issueId: issue.id,
        category: issue.category,
        severity: issue.severity,
        originalLength: issue.original.length,
        commentLength: formatIssueComment(issue).length,
        error: getWordErrorDetails(error),
      });
      result.fallbackIssues.push(issue);
    }
  }

  return result;
}

async function insertFallbackCommentInFreshContext(
  sourceText: string,
  scope: ProofreadScope,
  issues: ProofreadIssue[],
  options: ApplyIssuesOptions
): Promise<FallbackWriteResult> {
  if (issues.length === 0) {
    return {
      fallbackCount: 0,
      fallbackCommentCount: 0,
      failedCount: 0,
      truncatedFallbackCount: 0,
    };
  }

  try {
    return await insertFallbackCommentChunksInFreshContexts(sourceText, scope, issues, options);
  } catch (error) {
    logWordOperationFailure("fallback", error, {
      issueCount: issues.length,
    });
    appendDebugLog("error", "新上下文汇总批注写入失败", {
      issueCount: issues.length,
      error: getWordErrorDetails(error),
    });
    return {
      fallbackCount: 0,
      fallbackCommentCount: 0,
      failedCount: issues.length,
      truncatedFallbackCount: 0,
    };
  }
}

async function insertFallbackCommentChunksInFreshContexts(
  sourceText: string,
  scope: ProofreadScope,
  issues: ProofreadIssue[],
  options: ApplyIssuesOptions
): Promise<FallbackWriteResult> {
  const { chunks, truncatedFallbackCount } = buildFallbackCommentChunks(issues, options);
  let fallbackCount = 0;
  let fallbackCommentCount = 0;
  let failedCount = 0;
  let completedIssues = 0;

  if (chunks.length === 0) {
    return {
      fallbackCount: 0,
      fallbackCommentCount: 0,
      failedCount: 0,
      truncatedFallbackCount,
    };
  }

  try {
    await runWordBatchForScope(sourceText, scope, async (context) => {
      const searchContext = await createSearchContext(context, scope, sourceText);
      const anchorRange = await getFallbackAnchorRange(context, searchContext);

      for (let chunkIndex = 0; chunkIndex < chunks.length; chunkIndex += 1) {
        const chunk = chunks[chunkIndex];
        try {
          anchorRange.insertComment(chunk.comment);
        } catch (error) {
          logWordOperationFailure("fallback", error, {
            batchIndex: chunkIndex + 1,
            commentLength: chunk.comment.length,
            issueCount: chunk.issues.length,
            totalBatches: chunks.length,
          });
          throw error;
        }

        // eslint-disable-next-line office-addins/no-context-sync-in-loop -- Intentional per-chunk commit so long summaries cannot roll back prior chunks.
        await context.sync();

        fallbackCount += chunk.issues.length;
        fallbackCommentCount += 1;
        completedIssues += chunk.issues.length;
        appendDebugLog("info", "汇总批注分片写入成功", {
          chunkIndex: chunkIndex + 1,
          totalChunks: chunks.length,
          issueCount: chunk.issues.length,
          commentLength: chunk.comment.length,
          fallbackCount,
          failedCount,
          truncatedFallbackCount,
        });
        options.onProgress?.({
          stage: "fallback",
          completedBatches: chunkIndex + 1,
          totalBatches: chunks.length,
          completedIssues: Math.min(completedIssues, issues.length),
          totalIssues: issues.length,
        });
      }
    });
  } catch (error) {
    const remainingIssueCount = Math.max(issues.length - truncatedFallbackCount - fallbackCount, 0);
    failedCount += remainingIssueCount;
    logWordOperationFailure("fallback", error, {
      issueCount: remainingIssueCount,
      totalBatches: chunks.length,
    });
    appendDebugLog("error", "汇总批注写入上下文失败", {
      fallbackCount,
      failedCount,
      truncatedFallbackCount,
      totalChunks: chunks.length,
      error: getWordErrorDetails(error),
    });
  }

  return { fallbackCount, fallbackCommentCount, failedCount, truncatedFallbackCount };
}

function addCommentBatchResult(summary: IssueApplicationSummary, result: CommentBatchResult) {
  summary.commentCount += result.commentCount;
  summary.fallbackCount += result.fallbackCount;
  summary.fallbackCommentCount += result.fallbackCommentCount;
  summary.failedCount += result.failedCount;
  summary.truncatedFallbackCount += result.truncatedFallbackCount;
}

function createEmptyApplicationSummary(): IssueApplicationSummary {
  return {
    commentCount: 0,
    revisionCount: 0,
    fallbackCount: 0,
    fallbackCommentCount: 0,
    failedCount: 0,
    truncatedFallbackCount: 0,
  };
}

async function applyRevisionsForIssues(
  sourceText: string,
  issues: ProofreadIssue[],
  scope: ProofreadScope,
  options: ApplyIssuesOptions
): Promise<IssueApplicationSummary> {
  const summary = createEmptyApplicationSummary();
  const totalBatches = Math.ceil(issues.length / REVISION_RUN_BATCH_SIZE);
  const deferredFallbackIssues: ProofreadIssue[] = [];
  let completedIssues = 0;

  for (let batchStart = 0; batchStart < issues.length; batchStart += REVISION_RUN_BATCH_SIZE) {
    const batch = issues.slice(batchStart, batchStart + REVISION_RUN_BATCH_SIZE);
    const batchIndex = Math.floor(batchStart / REVISION_RUN_BATCH_SIZE) + 1;
    let batchResult: RevisionBatchResult;

    try {
      batchResult = await applyRevisionIssueBatch(sourceText, batch, scope, options);
    } catch (error) {
      logWordOperationFailure("revision", error, {
        batchIndex,
        issueCount: batch.length,
      });
      appendDebugLog("error", "修订应用批次上下文失败，整批改为汇总批注", {
        batchIndex,
        batchSize: batch.length,
        error: getWordErrorDetails(error),
      });
      batchResult = {
        commentCount: 0,
        revisionCount: 0,
        retryIssues: [],
        fallbackIssues: batch,
      };
    }

    summary.commentCount += batchResult.commentCount;
    summary.revisionCount += batchResult.revisionCount;

    if (batchResult.retryIssues.length > 0) {
      const retryResult = await retryRevisionIssuesInFreshContexts(
        sourceText,
        batchResult.retryIssues,
        scope,
        options
      );
      summary.revisionCount += retryResult.revisionCount;
      batchResult.fallbackIssues.push(...retryResult.fallbackIssues, ...retryResult.retryIssues);
    }

    if (batchResult.fallbackIssues.length > 0) {
      deferredFallbackIssues.push(...batchResult.fallbackIssues);
    }

    completedIssues += batch.length;
    options.onProgress?.({
      stage: "revising",
      completedBatches: Math.min(batchIndex, totalBatches),
      totalBatches,
      completedIssues: Math.min(completedIssues, issues.length),
      totalIssues: issues.length,
    });
  }

  if (deferredFallbackIssues.length > 0) {
    const fallbackResult = await insertFallbackCommentInFreshContext(
      sourceText,
      scope,
      deferredFallbackIssues,
      options
    );
    summary.fallbackCount += fallbackResult.fallbackCount;
    summary.fallbackCommentCount += fallbackResult.fallbackCommentCount;
    summary.failedCount += fallbackResult.failedCount;
    summary.truncatedFallbackCount += fallbackResult.truncatedFallbackCount;
  }

  return summary;
}

function applyRevisionIssueBatch(
  sourceText: string,
  issues: ProofreadIssue[],
  scope: ProofreadScope,
  options: ApplyIssuesOptions
): Promise<RevisionBatchResult> {
  return runWordBatchForScope(sourceText, scope, async (context) => {
    const document = context.document;

    const searchContext = await createSearchContext(context, scope, sourceText);
    const resolved = await resolveIssueTargets(context, searchContext, sourceText, issues, options);

    if (resolved.contextFailed) {
      return {
        commentCount: 0,
        revisionCount: 0,
        retryIssues: [],
        fallbackIssues: issues,
      };
    }

    const revisionApplications: ResolvedIssueTarget[] = [];
    const commentApplications: ResolvedIssueTarget[] = [];

    resolved.targets.forEach((target) => {
      if (hasReplacement(target.issue)) {
        revisionApplications.push(target);
        return;
      }

      commentApplications.push(target);
    });

    const commentResult = await insertCommentsInBatches(context, commentApplications, options);

    let revisionResult: RevisionWriteBatchResult = {
      successCount: 0,
      failedIssues: [],
      commentCount: 0,
    };
    let originalTrackingMode:
      | Word.ChangeTrackingMode
      | "Off"
      | "TrackAll"
      | "TrackMineOnly"
      | null = null;

    if (revisionApplications.length > 0) {
      try {
        document.load("changeTrackingMode");
        await context.sync();
        originalTrackingMode = document.changeTrackingMode;
        document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        await context.sync();
        revisionResult = await insertRevisionsInBatches(context, revisionApplications, options);
      } catch (error) {
        logWordOperationFailure("revision", error, {
          issueCount: revisionApplications.length,
        });
        appendDebugLog("error", "启用或写入 Word 修订失败，改为汇总批注", {
          revisionCount: revisionApplications.length,
          error: getWordErrorDetails(error),
        });
        revisionResult = {
          successCount: 0,
          failedIssues: revisionApplications.map(({ issue }) => issue),
          commentCount: 0,
          contextFailed: true,
        };
      } finally {
        if (originalTrackingMode !== null) {
          document.changeTrackingMode = originalTrackingMode;
          try {
            await context.sync();
          } catch (error) {
            logWordOperationFailure("revision", error);
            appendDebugLog("error", "恢复 Word 修订设置失败", {
              error: getWordErrorDetails(error),
            });
          }
        }
      }
    }

    const fallbackIssues = [
      ...resolved.summaryIssues,
      ...commentResult.failedIssues,
      ...(revisionResult.contextFailed ? [] : revisionResult.failedIssues),
    ];

    return {
      commentCount: commentResult.successCount + revisionResult.commentCount,
      revisionCount: revisionResult.successCount,
      retryIssues: revisionResult.contextFailed ? revisionResult.failedIssues : [],
      fallbackIssues,
    };
  });
}

async function insertRevisionReasonCommentsBestEffort(
  context: Word.RequestContext,
  targets: ResolvedIssueTarget[],
  options: ApplyIssuesOptions
): Promise<BestEffortCommentResult> {
  if (targets.length === 0) {
    return { commentCount: 0 };
  }

  const result = await insertCommentsInBatches(context, targets, options);
  if (result.contextFailed || result.failedIssues.length > 0) {
    appendDebugLog("warn", "修订说明批注部分写入失败，仍继续写入修订", {
      attemptedCount: targets.length,
      commentCount: result.successCount,
      failedCount: result.failedIssues.length,
    });
  }

  return { commentCount: result.successCount };
}

async function retryRevisionIssuesInFreshContexts(
  sourceText: string,
  issues: ProofreadIssue[],
  scope: ProofreadScope,
  options: ApplyIssuesOptions
): Promise<RevisionBatchResult> {
  const result: RevisionBatchResult = {
    commentCount: 0,
    revisionCount: 0,
    retryIssues: [],
    fallbackIssues: [],
  };

  for (const issue of issues) {
    try {
      const retryResult = await applyRevisionIssueBatch(sourceText, [issue], scope, options);
      result.commentCount += retryResult.commentCount;
      result.revisionCount += retryResult.revisionCount;
      result.fallbackIssues.push(...retryResult.fallbackIssues, ...retryResult.retryIssues);
    } catch (error) {
      logWordOperationFailure("revision", error, {
        category: issue.category,
        issueId: issue.id,
        replacementLength: issue.replacement?.length || 0,
        severity: issue.severity,
      });
      result.fallbackIssues.push(issue);
    }
  }

  return result;
}

async function insertCommentsInBatches(
  context: Word.RequestContext,
  targets: ResolvedIssueTarget[],
  options: ApplyIssuesOptions
): Promise<WriteBatchResult> {
  const failedIssues: ProofreadIssue[] = [];
  let successCount = 0;
  const totalBatches = Math.ceil(targets.length / COMMENT_INSERT_BATCH_SIZE);

  for (let batchStart = 0; batchStart < targets.length; batchStart += COMMENT_INSERT_BATCH_SIZE) {
    const batch = targets.slice(batchStart, batchStart + COMMENT_INSERT_BATCH_SIZE);
    const batchIndex = Math.floor(batchStart / COMMENT_INSERT_BATCH_SIZE) + 1;

    try {
      batch.forEach(({ issue, targetRange }) => {
        try {
          targetRange.insertComment(formatIssueComment(issue));
        } catch (error) {
          logWordOperationFailure("comment", error, {
            batchIndex,
            category: issue.category,
            commentLength: formatIssueComment(issue).length,
            issueId: issue.id,
            severity: issue.severity,
          });
          throw error;
        }
      });
      // eslint-disable-next-line office-addins/no-context-sync-in-loop -- Intentional write batch boundary so earlier comments stay committed.
      await context.sync();
      successCount += batch.length;
      appendDebugLog("info", "批量批注写入成功", {
        batchIndex,
        batchSize: batch.length,
        successCount,
        failedCount: failedIssues.length,
      });
    } catch (error) {
      logWordOperationFailure("comment", error, {
        batchIndex,
        issueCount: batch.length,
      });
      appendDebugLog("warn", "批量批注写入失败，将切换新上下文单条重试", {
        batchIndex,
        batchSize: batch.length,
        error: getWordErrorDetails(error),
      });
      failedIssues.push(...batch.map(({ issue }) => issue));
      return { successCount, failedIssues, contextFailed: true };
    }

    options.onProgress?.({
      stage: "commenting",
      completedBatches: Math.min(batchIndex, totalBatches),
      totalBatches,
      completedIssues: Math.min(successCount + failedIssues.length, targets.length),
      totalIssues: targets.length,
    });
  }

  return { successCount, failedIssues };
}

async function insertRevisionsInBatches(
  context: Word.RequestContext,
  targets: ResolvedIssueTarget[],
  options: ApplyIssuesOptions
): Promise<RevisionWriteBatchResult> {
  const orderedTargets = [...targets].sort((left, right) => {
    const rightStart = typeof right.issue.start === "number" ? right.issue.start : 0;
    const leftStart = typeof left.issue.start === "number" ? left.issue.start : 0;
    return rightStart - leftStart;
  });
  const failedIssues: ProofreadIssue[] = [];
  const insertedTargets: ResolvedIssueTarget[] = [];
  let successCount = 0;
  const totalBatches = Math.ceil(orderedTargets.length / REVISION_INSERT_BATCH_SIZE);

  for (
    let batchStart = 0;
    batchStart < orderedTargets.length;
    batchStart += REVISION_INSERT_BATCH_SIZE
  ) {
    const batch = orderedTargets.slice(batchStart, batchStart + REVISION_INSERT_BATCH_SIZE);
    const batchIndex = Math.floor(batchStart / REVISION_INSERT_BATCH_SIZE) + 1;
    const insertedBatchTargets: ResolvedIssueTarget[] = [];

    try {
      batch.forEach(({ issue, targetRange }) => {
        try {
          const replacementRange = targetRange.insertText(
            issue.replacement as string,
            Word.InsertLocation.replace
          );
          insertedBatchTargets.push({ issue, targetRange: replacementRange });
        } catch (error) {
          logWordOperationFailure("revision", error, {
            batchIndex,
            category: issue.category,
            issueId: issue.id,
            replacementLength: issue.replacement?.length || 0,
            severity: issue.severity,
          });
          throw error;
        }
      });
      // eslint-disable-next-line office-addins/no-context-sync-in-loop -- Intentional write batch boundary so earlier revisions stay committed.
      await context.sync();
      successCount += batch.length;
      insertedTargets.push(...insertedBatchTargets);
      appendDebugLog("info", "批量修订写入成功", {
        batchIndex,
        batchSize: batch.length,
        successCount,
        failedCount: failedIssues.length,
      });
    } catch (error) {
      logWordOperationFailure("revision", error, {
        batchIndex,
        issueCount: batch.length,
      });
      appendDebugLog("warn", "批量修订写入失败，将切换新上下文单条重试", {
        batchIndex,
        batchSize: batch.length,
        error: getWordErrorDetails(error),
      });
      failedIssues.push(...batch.map(({ issue }) => issue));
      const commentResult = await insertRevisionReasonCommentsBestEffort(
        context,
        insertedTargets,
        options
      );
      return {
        successCount,
        failedIssues,
        commentCount: commentResult.commentCount,
        contextFailed: true,
      };
    }

    options.onProgress?.({
      stage: "revising",
      completedBatches: Math.min(batchIndex, totalBatches),
      totalBatches,
      completedIssues: Math.min(successCount + failedIssues.length, orderedTargets.length),
      totalIssues: orderedTargets.length,
    });
  }

  const commentResult = await insertRevisionReasonCommentsBestEffort(
    context,
    insertedTargets,
    options
  );
  return { successCount, failedIssues, commentCount: commentResult.commentCount };
}

function getWordErrorDetails(error: unknown): object {
  if (error && typeof error === "object") {
    const candidate = error as {
      code?: unknown;
      debugInfo?: unknown;
      message?: unknown;
      name?: unknown;
      stack?: unknown;
    };

    return {
      name: candidate.name,
      code: candidate.code,
      message: candidate.message,
      debugInfo: candidate.debugInfo,
    };
  }

  return { message: String(error) };
}

function logWordOperationFailure(
  stage: WordOperationStage,
  error: unknown,
  details: WordErrorLogDetails = {}
) {
  appendDebugLog("error", "Word 操作失败", {
    stage,
    ...details,
    error: getWordErrorDetails(error),
  });
}

async function createSearchContext(
  context: Word.RequestContext,
  scope: ProofreadScope,
  sourceText: string
): Promise<SearchContext> {
  if (sourceText.length === 0) {
    appendDebugLog("info", "使用历史 locator 定位，将在正文范围内搜索", {
      scope,
      locatorOnly: true,
    });
    return {
      root: context.document.body,
      rootKind: "body",
      occurrenceText: "",
      canSearch: true,
    };
  }

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

  const trackedSelectionContext = await createTrackedSelectionSearchContext(context, sourceText);
  if (trackedSelectionContext) {
    return trackedSelectionContext;
  }

  let selection: Word.Range;

  try {
    selection = context.document.getSelection();
    selection.load("text");
    await context.sync();
  } catch (error) {
    logWordOperationFailure("locate", error, {
      issueCount: sourceText.length,
    });
    appendDebugLog("warn", "读取当前选区失败，将整批降级汇总", {
      sourceTextLength: sourceText.length,
      error: getWordErrorDetails(error),
    });
    return {
      root: context.document.body,
      rootKind: "body",
      occurrenceText: sourceText,
      canSearch: false,
    };
  }

  const selectionText = (selection.text || "").trim();
  const canSearch = sourceText.length === 0 || selectionText === sourceText;

  appendDebugLog(canSearch ? "info" : "warn", "使用当前选区范围定位", {
    sourceTextLength: sourceText.length,
    selectionTextLength: selectionText.length,
    canSearch,
    locatorOnly: sourceText.length === 0,
  });

  return {
    root: selection,
    rootKind: "selection",
    occurrenceText: sourceText,
    canSearch,
  };
}

function runWordBatchForScope<T>(
  sourceText: string,
  scope: ProofreadScope,
  batch: (context: Word.RequestContext) => Promise<T>
): Promise<T> {
  if (scope === "selection" && sourceText.length > 0 && trackedProofreadSelectionRange) {
    return Word.run(trackedProofreadSelectionRange, batch);
  }

  return Word.run(batch);
}

async function createTrackedSelectionSearchContext(
  context: Word.RequestContext,
  sourceText: string
): Promise<SearchContext | null> {
  if (!trackedProofreadSelectionRange) {
    return null;
  }

  try {
    trackedProofreadSelectionRange.load("text");
    await context.sync();
  } catch (error) {
    logWordOperationFailure("locate", error);
    appendDebugLog("warn", "读取送审选区范围失败，将尝试当前选区", {
      sourceTextLength: sourceText.length,
      error: getWordErrorDetails(error),
    });
    trackedProofreadSelectionRange = null;
    return null;
  }

  const selectionText = (trackedProofreadSelectionRange.text || "").trim();
  const canSearch = selectionText === sourceText;

  appendDebugLog(canSearch ? "info" : "warn", "使用送审选区范围定位", {
    sourceTextLength: sourceText.length,
    selectionTextLength: selectionText.length,
    canSearch,
  });

  return {
    root: trackedProofreadSelectionRange,
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
  const fallbackAnchorRange = await getFallbackAnchorRange(context, searchContext);

  if (!searchContext.canSearch) {
    appendDebugLog("warn", "当前 Word 范围与审校文本不一致，全部降级汇总", {
      issueCount: issues.length,
      rootKind: searchContext.rootKind,
    });
    return {
      targets,
      summaryIssues: issues,
      fallbackAnchorRange,
    };
  }

  const groups = buildSearchGroups(sourceText, issues, summaryIssues);
  const totalBatches = Math.ceil(groups.length / SEARCH_BATCH_SIZE);
  let completedIssues = summaryIssues.length;

  for (let batchStart = 0; batchStart < groups.length; batchStart += SEARCH_BATCH_SIZE) {
    const batch = groups.slice(batchStart, batchStart + SEARCH_BATCH_SIZE);
    let pendingKeySearches: Array<{ group: SearchGroup; searchResults: Word.RangeCollection }> = [];

    try {
      pendingKeySearches = batch.map((group) => {
        const searchResults = searchContext.root.search(group.locator.key, {
          matchCase: true,
          matchWholeWord: false,
        });
        searchResults.load("items");
        return { group, searchResults };
      });
    } catch (error) {
      logWordOperationFailure("search", error, {
        batchIndex: Math.floor(batchStart / SEARCH_BATCH_SIZE) + 1,
        issueCount: batch.reduce((total, group) => total + group.issues.length, 0),
        keyLength: Math.max(...batch.map((group) => group.locator.key.length)),
      });
      return {
        targets: [],
        summaryIssues: issues,
        fallbackAnchorRange,
        contextFailed: true,
      };
    }

    try {
      // eslint-disable-next-line office-addins/no-context-sync-in-loop -- Intentional batch boundary to keep Word responsive.
      await context.sync();
    } catch (error) {
      const unresolvedIssues = groups
        .slice(batchStart)
        .flatMap((group) => group.issues.map(({ issue }) => issue));
      logWordOperationFailure("search", error, {
        batchIndex: Math.floor(batchStart / SEARCH_BATCH_SIZE) + 1,
        issueCount: unresolvedIssues.length,
      });
      appendDebugLog("error", "Word locator key 搜索批次失败，剩余条目降级汇总", {
        batchIndex: Math.floor(batchStart / SEARCH_BATCH_SIZE) + 1,
        batchSize: batch.length,
        unresolvedIssueCount: unresolvedIssues.length,
        error: getWordErrorDetails(error),
      });
      return {
        targets: [],
        summaryIssues: issues,
        fallbackAnchorRange,
        contextFailed: true,
      };
    }

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

        try {
          const searchResultsInKey = keyRange.search(prepared.issue.original, {
            matchCase: true,
            matchWholeWord: false,
          });
          searchResultsInKey.load("items");
          pendingOriginalSearches.push({ prepared, searchResults: searchResultsInKey });
        } catch (error) {
          logWordOperationFailure("search", error, {
            category: prepared.issue.category,
            issueId: prepared.issue.id,
            keyLength: prepared.locator.key.length,
            severity: prepared.issue.severity,
          });
          summaryIssues.push(prepared.issue);
        }
      });
    });

    if (pendingOriginalSearches.length > 0) {
      try {
        // eslint-disable-next-line office-addins/no-context-sync-in-loop -- Intentional nested batch boundary for context locators.
        await context.sync();
      } catch (error) {
        logWordOperationFailure("search", error, {
          batchIndex: Math.floor(batchStart / SEARCH_BATCH_SIZE) + 1,
          issueCount: pendingOriginalSearches.length,
        });
        appendDebugLog("error", "Word locator key 内 original 搜索失败，当前批次降级汇总", {
          batchIndex: Math.floor(batchStart / SEARCH_BATCH_SIZE) + 1,
          pendingOriginalSearchCount: pendingOriginalSearches.length,
          error: getWordErrorDetails(error),
        });
        return {
          targets: [],
          summaryIssues: issues,
          fallbackAnchorRange,
          contextFailed: true,
        };
      }
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
      stage: "locating",
      completedBatches: Math.min(Math.floor(batchStart / SEARCH_BATCH_SIZE) + 1, totalBatches),
      totalBatches,
      completedIssues: Math.min(completedIssues, issues.length),
      totalIssues: issues.length,
    });
  }

  return {
    targets,
    summaryIssues,
    fallbackAnchorRange: firstTargetRange || fallbackAnchorRange,
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
      occurrenceIndex:
        sourceText.length > 0
          ? getOccurrenceIndexBeforeOffset(sourceText, locator.key, locator.key_start)
          : (locator.key_occurrence_index as number),
      issues: [prepared],
    });
  });

  return Array.from(groupsById.values());
}

function getUsableLocator(sourceText: string, issue: ProofreadIssue): ProofreadLocator | null {
  if (sourceText.length === 0) {
    return isReplayableLocator(issue.locator || null) ? (issue.locator as ProofreadLocator) : null;
  }

  if (!isIssueLocatable(sourceText, issue)) {
    return null;
  }

  if (isValidLocator(sourceText, issue, issue.locator || null)) {
    return issue.locator as ProofreadLocator;
  }

  return buildClientLocator(sourceText, issue);
}

function isReplayableLocator(locator: ProofreadLocator | null): boolean {
  return (
    Boolean(locator?.key) &&
    typeof locator?.key_occurrence_index === "number" &&
    typeof locator.original_start_in_key === "number" &&
    typeof locator.original_end_in_key === "number" &&
    locator.original_end_in_key <= locator.key.length
  );
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
      key_occurrence_index: getOccurrenceIndexBeforeOffset(sourceText, issue.original, start),
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
        key_occurrence_index: getOccurrenceIndexBeforeOffset(sourceText, key, keyStart),
      };
    }
  }

  return null;
}

async function getFallbackAnchorRange(
  context: Word.RequestContext,
  searchContext: SearchContext
): Promise<Word.Range> {
  try {
    const paragraphs =
      searchContext.rootKind === "body"
        ? (searchContext.root as Word.Body).paragraphs
        : (searchContext.root as Word.Range).paragraphs;
    const firstParagraph = paragraphs.getFirstOrNullObject();
    // eslint-disable-next-line office-addins/no-navigational-load -- Required to inspect OrNullObject before selecting a small fallback anchor.
    firstParagraph.load("isNullObject");
    await context.sync();

    if (!firstParagraph.isNullObject) {
      return firstParagraph.getRange(Word.RangeLocation.start);
    }
  } catch (error) {
    logWordOperationFailure("anchor", error);
    appendDebugLog("warn", "获取汇总批注段落锚点失败，退回范围起点", {
      error: getWordErrorDetails(error),
    });
  }

  return searchContext.rootKind === "body"
    ? (searchContext.root as Word.Body).getRange(Word.RangeLocation.start)
    : (searchContext.root as Word.Range).getRange(Word.RangeLocation.start);
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
    searchFrom = foundAt + 1;
  }

  return count;
}

function formatIssueComment(issue: ProofreadIssue): string {
  return limitCommentText(
    sanitizeCommentText(
      [
        issue.replacement ? `替换为：${issue.replacement}` : "",
        `建议：${issue.suggestion || "未提供"}`,
        `类别：${issue.category} / ${issue.severity}`,
      ]
        .filter(Boolean)
        .join("\n")
    )
  );
}

function buildFallbackCommentChunks(
  issues: ProofreadIssue[],
  options: ApplyIssuesOptions
): { chunks: FallbackCommentChunk[]; truncatedFallbackCount: number } {
  const entries = issues.map((issue, index) => ({
    issue,
    entry: formatFallbackIssueEntry(issue, index + 1),
  }));
  const budget = FALLBACK_COMMENT_MAX_LENGTH - FALLBACK_COMMENT_HEADER_RESERVE;
  const groups: Array<Array<{ issue: ProofreadIssue; entry: string }>> = [];
  let currentGroup: Array<{ issue: ProofreadIssue; entry: string }> = [];
  let currentLength = 0;

  entries.forEach((entry) => {
    const separatorLength = currentGroup.length > 0 ? 2 : 0;
    const nextLength = currentLength + separatorLength + entry.entry.length;

    if (currentGroup.length > 0 && nextLength > budget) {
      groups.push(currentGroup);
      currentGroup = [];
      currentLength = 0;
    }

    currentGroup.push(entry);
    currentLength += (currentGroup.length > 1 ? 2 : 0) + entry.entry.length;
  });

  if (currentGroup.length > 0) {
    groups.push(currentGroup);
  }

  const truncateEnabled = options.fallbackSummaryTruncateEnabled !== false;
  const visibleGroups = truncateEnabled ? groups.slice(0, FALLBACK_COMMENT_MAX_CHUNKS) : groups;
  const truncatedFallbackCount = truncateEnabled
    ? groups.slice(FALLBACK_COMMENT_MAX_CHUNKS).reduce((count, group) => count + group.length, 0)
    : 0;

  const chunks = visibleGroups.map((group, index) => {
    const title = `AI 审校汇总批注 ${index + 1}/${visibleGroups.length}`;
    const issueText = group.map(({ entry }) => entry).join("\n\n");
    const truncationNotice =
      truncatedFallbackCount > 0 && index === visibleGroups.length - 1
        ? `\n\n另有 ${truncatedFallbackCount} 条未写入汇总批注，请在任务窗格中查看。`
        : "";
    const comment = limitText(
      sanitizeCommentText(
        `${title}\n以下 ${group.length} 条建议未能精准写回：\n\n${issueText}${truncationNotice}`
      ),
      FALLBACK_COMMENT_MAX_LENGTH,
      "\n\n（本条汇总批注过长，已截断；完整建议请在任务窗格中查看。）"
    );

    return {
      issues: group.map(({ issue }) => issue),
      comment,
    };
  });

  return { chunks, truncatedFallbackCount };
}

function formatFallbackIssueEntry(issue: ProofreadIssue, displayIndex: number): string {
  const lines = [
    `${displayIndex}. [${issue.severity}] ${issue.category}`,
    `原文：${issue.original || "未提供"}`,
    issue.replacement ? `替换为：${issue.replacement}` : "",
    `建议：${issue.suggestion || "未提供"}`,
  ];

  return limitText(
    sanitizeCommentText(lines.filter(Boolean).join("\n")),
    FALLBACK_ISSUE_MAX_LENGTH,
    "\n（单条建议过长，已截断；完整建议请在任务窗格中查看。）"
  );
}

function sanitizeCommentText(text: string): string {
  return Array.from(text)
    .filter((character) => {
      const codePoint = character.codePointAt(0) || 0;
      return codePoint === 9 || codePoint === 10 || (codePoint > 31 && codePoint !== 127);
    })
    .join("");
}

function limitCommentText(text: string): string {
  return limitText(
    text,
    MAX_COMMENT_LENGTH,
    "\n\n（批注内容过长，已截断；完整建议请在任务窗格中查看。）"
  );
}

function limitText(text: string, maxLength: number, suffix: string): string {
  if (text.length <= maxLength) {
    return text;
  }

  return `${text.slice(0, Math.max(0, maxLength - suffix.length))}${suffix}`;
}

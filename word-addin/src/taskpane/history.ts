/* global Blob, File, URL, document, localStorage */

import {
  ApplicationMode,
  BookInfo,
  DEFAULT_TEMPERATURE,
  IssueApplicationSummary,
  PendingProofreadResult,
  ProofreadHistoryEntry,
  ProofreadIssue,
  ProofreadLocator,
  ProofreadMode,
  ProofreadScope,
  ProviderAPI,
  TaskState,
} from "./types";

const HISTORY_STORAGE_KEY = "word-ai-proofreader-history-v2";
const HISTORY_SCHEMA_VERSION = 7;
const MAX_HISTORY_ENTRIES = 20;
const MIN_ORIGINAL_LOCATOR_LENGTH = 6;
const MAX_LOCATOR_OCCURRENCES = 3;
const CONTEXT_LOCATOR_WINDOWS = [16, 32, 64];

export function saveHistoryEntry(input: {
  status: TaskState;
  text: string;
  book: BookInfo;
  issues: ProofreadIssue[];
  insertedComment: boolean;
  appliedToWord: boolean;
  locatedIssueCount: number;
  revisionCount: number;
  fallbackCount: number;
  failedCount?: number;
  scope: ProofreadScope;
  taskId?: string | null;
  runId?: string | null;
  totalChunks: number;
  completedChunks: number;
  failedChunks: number;
  sessionId: string;
  aiProfileId?: string | null;
  providerApi: ProviderAPI;
  proofreadMode: ProofreadMode;
  reasoningEnabled: boolean;
  temperature: number;
  applicationMode: ApplicationMode;
  selectedIssueIds: string[];
  skippedIssueCount: number;
  issueCount?: number;
  sourceFilename?: string | null;
  outputFilename?: string | null;
  downloadUrl?: string | null;
  expiresAt?: string | null;
  retentionDays?: number | null;
  errorMessage?: string;
}) {
  if (input.text.trim().length === 0) {
    return;
  }

  const entries = getHistoryEntries();
  const issues = normalizeIssuesForHistory(input.text, input.issues);
  const entry: ProofreadHistoryEntry = {
    id: createLocalId(),
    historySchemaVersion: HISTORY_SCHEMA_VERSION,
    sessionId: input.sessionId,
    createdAt: new Date().toISOString(),
    textPreview: input.text.trim().slice(0, 40),
    bookTitle: input.book.title,
    bookIntroductionPreview: (input.book.introduction || "").trim().slice(0, 40),
    status: input.status,
    issueCount: input.issueCount ?? issues.length,
    locatedIssueCount: input.locatedIssueCount,
    revisionCount: input.revisionCount,
    fallbackCount: input.fallbackCount,
    failedCount: input.failedCount || 0,
    aiProfileId: input.aiProfileId || null,
    providerApi: input.providerApi,
    proofreadMode: input.proofreadMode,
    reasoningEnabled: input.reasoningEnabled,
    temperature: input.temperature,
    applicationMode: input.applicationMode,
    scope: input.scope,
    taskId: input.taskId,
    runId: input.runId || null,
    totalChunks: input.totalChunks,
    completedChunks: input.completedChunks,
    failedChunks: input.failedChunks,
    globalLocatedIssueCount: issues.filter(
      (issue) => typeof issue.start === "number" && typeof issue.end === "number"
    ).length,
    issues,
    selectedIssueIds: input.selectedIssueIds,
    skippedIssueCount: input.skippedIssueCount,
    insertedComment: input.insertedComment,
    appliedToWord: input.appliedToWord,
    replayable: issues.some(hasReplayableLocator),
    errorMessage: input.errorMessage,
    sourceFilename: input.sourceFilename || null,
    outputFilename: input.outputFilename || null,
    downloadUrl: input.downloadUrl || null,
    expiresAt: input.expiresAt || null,
    retentionDays: input.retentionDays || null,
  };

  localStorage.setItem(
    HISTORY_STORAGE_KEY,
    JSON.stringify([entry, ...entries].slice(0, MAX_HISTORY_ENTRIES))
  );
}

export function savePendingResultHistory(input: {
  result: PendingProofreadResult;
  applicationMode: ApplicationMode;
  status: TaskState;
  errorMessage?: string;
}) {
  const selectedIssueIds = input.result.issues.map((issue) => issue.id);

  saveHistoryEntry({
    status: input.status,
    text: getHistoryText(input.result),
    book: input.result.book,
    issues: input.result.issues,
    insertedComment: false,
    appliedToWord: false,
    locatedIssueCount: input.result.issues.filter(
      (issue) => typeof issue.start === "number" && typeof issue.end === "number"
    ).length,
    revisionCount: 0,
    fallbackCount: input.result.issues.filter(
      (issue) => typeof issue.start !== "number" || typeof issue.end !== "number"
    ).length,
    failedCount: 0,
    scope: input.result.scope,
    taskId: input.result.taskId,
    runId: input.result.runId,
    totalChunks: input.result.totalChunks,
    completedChunks: input.result.completedChunks,
    failedChunks: input.result.failedChunks,
    sessionId: input.result.sessionId,
    aiProfileId: input.result.aiProfileId,
    providerApi: input.result.providerApi,
    proofreadMode: input.result.proofreadMode,
    reasoningEnabled: input.result.reasoningEnabled,
    temperature: input.result.temperature,
    applicationMode: input.applicationMode,
    selectedIssueIds,
    skippedIssueCount: input.result.issues.length - selectedIssueIds.length,
    issueCount: input.result.issueCount,
    sourceFilename: input.result.sourceFilename,
    outputFilename: input.result.outputFilename,
    downloadUrl: input.result.downloadUrl,
    expiresAt: input.result.expiresAt,
    retentionDays: input.result.retentionDays,
    errorMessage: input.errorMessage,
  });
}

export function saveAppliedResultHistory(input: {
  result: PendingProofreadResult;
  summary: IssueApplicationSummary;
  applicationMode: ApplicationMode;
  selectedIssueIds: string[];
}) {
  const selectedIssueIds = input.selectedIssueIds;
  const selectedIssueIdSet = new Set(selectedIssueIds);
  const selectedIssues = input.result.issues.filter((issue) => selectedIssueIdSet.has(issue.id));

  saveHistoryEntry({
    status: "succeeded",
    text: getHistoryText(input.result),
    book: input.result.book,
    issues: input.result.issues,
    insertedComment: input.summary.commentCount > 0 || input.summary.fallbackCount > 0,
    appliedToWord: true,
    locatedIssueCount: Math.min(
      selectedIssues.length,
      input.summary.commentCount + input.summary.revisionCount
    ),
    revisionCount: input.summary.revisionCount,
    fallbackCount: input.summary.fallbackCount,
    failedCount: input.summary.failedCount,
    scope: input.result.scope,
    taskId: input.result.taskId,
    runId: input.result.runId,
    totalChunks: input.result.totalChunks,
    completedChunks: input.result.completedChunks,
    failedChunks: input.result.failedChunks,
    sessionId: input.result.sessionId,
    aiProfileId: input.result.aiProfileId,
    providerApi: input.result.providerApi,
    proofreadMode: input.result.proofreadMode,
    reasoningEnabled: input.result.reasoningEnabled,
    temperature: input.result.temperature,
    applicationMode: input.applicationMode,
    selectedIssueIds,
    skippedIssueCount: input.result.issues.length - selectedIssues.length,
    issueCount: input.result.issueCount,
    sourceFilename: input.result.sourceFilename,
    outputFilename: input.result.outputFilename,
    downloadUrl: input.result.downloadUrl,
    expiresAt: input.result.expiresAt,
    retentionDays: input.result.retentionDays,
  });
}

export function getHistoryEntries(): ProofreadHistoryEntry[] {
  try {
    const parsed = JSON.parse(localStorage.getItem(HISTORY_STORAGE_KEY) || "[]");
    return Array.isArray(parsed) && parsed.every(isHistoryEntry)
      ? (parsed as ProofreadHistoryEntry[]).map(normalizeHistoryEntry)
      : [];
  } catch {
    return [];
  }
}

export function clearHistoryEntries() {
  localStorage.removeItem(HISTORY_STORAGE_KEY);
}

export function exportHistoryEntries() {
  const entries = getHistoryEntries();
  const blob = new Blob([JSON.stringify(entries, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = `word-ai-proofreader-history-${new Date().toISOString().slice(0, 10)}.json`;
  link.click();
  URL.revokeObjectURL(url);
  return entries.length;
}

export async function importHistoryEntries(file: File): Promise<number> {
  const parsed = JSON.parse(await file.text());

  if (!Array.isArray(parsed)) {
    throw new Error("历史文件必须是数组格式。");
  }

  if (!parsed.every(isHistoryEntry)) {
    throw new Error("历史文件格式不兼容，请导入当前开发版本导出的历史 JSON。");
  }

  const imported = parsed as ProofreadHistoryEntry[];
  localStorage.setItem(HISTORY_STORAGE_KEY, JSON.stringify(imported.slice(0, MAX_HISTORY_ENTRIES)));
  return imported.length;
}

function createLocalId(): string {
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function getHistoryText(result: PendingProofreadResult): string {
  return (
    result.sourceText || result.sourceFilename || result.historyTextPreview || result.book.title
  );
}

function normalizeIssuesForHistory(text: string, issues: ProofreadIssue[]): ProofreadIssue[] {
  return issues.map((issue) => {
    if (!issue.locator) {
      const locator = buildHistoryLocator(text, issue);
      return locator ? { ...issue, locator } : issue;
    }

    const locator = normalizeLocatorForHistory(text, issue, issue.locator);
    return locator === issue.locator ? issue : { ...issue, locator };
  });
}

function buildHistoryLocator(text: string, issue: ProofreadIssue): ProofreadLocator | null {
  if (
    !text ||
    typeof issue.start !== "number" ||
    typeof issue.end !== "number" ||
    issue.end > text.length ||
    text.slice(issue.start, issue.end) !== issue.original
  ) {
    return null;
  }

  if (
    issue.original.length >= MIN_ORIGINAL_LOCATOR_LENGTH &&
    getOccurrenceCount(text, issue.original) <= MAX_LOCATOR_OCCURRENCES
  ) {
    return {
      key: issue.original,
      key_start: issue.start,
      key_end: issue.end,
      original_start_in_key: 0,
      original_end_in_key: issue.original.length,
      strategy: "original",
      key_occurrence_index: getOccurrenceIndexBeforeOffset(text, issue.original, issue.start),
    };
  }

  for (const windowSize of CONTEXT_LOCATOR_WINDOWS) {
    const keyStart = Math.max(0, issue.start - windowSize);
    const keyEnd = Math.min(text.length, issue.end + windowSize);
    const key = text.slice(keyStart, keyEnd);

    if (key && getOccurrenceCount(text, key) <= MAX_LOCATOR_OCCURRENCES) {
      return {
        key,
        key_start: keyStart,
        key_end: keyEnd,
        original_start_in_key: issue.start - keyStart,
        original_end_in_key: issue.end - keyStart,
        strategy: "context",
        key_occurrence_index: getOccurrenceIndexBeforeOffset(text, key, keyStart),
      };
    }
  }

  return null;
}

function normalizeLocatorForHistory(
  text: string,
  issue: ProofreadIssue,
  locator: ProofreadLocator
): ProofreadLocator {
  if (typeof locator.key_occurrence_index === "number") {
    return locator;
  }

  if (!canUseTextToNormalizeLocator(text, issue, locator)) {
    return locator;
  }

  return {
    ...locator,
    key_occurrence_index: getOccurrenceIndexBeforeOffset(text, locator.key, locator.key_start),
  };
}

function canUseTextToNormalizeLocator(
  text: string,
  issue: ProofreadIssue,
  locator: ProofreadLocator
): boolean {
  return (
    text.length > 0 &&
    typeof issue.start === "number" &&
    typeof issue.end === "number" &&
    issue.end <= text.length &&
    locator.key_end <= text.length &&
    text.slice(issue.start, issue.end) === issue.original &&
    text.slice(locator.key_start, locator.key_end) === locator.key
  );
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

function hasReplayableLocator(issue: ProofreadIssue): boolean {
  return isReplayableLocator(issue.locator || null);
}

function isReplayableLocator(locator: ProofreadLocator | null): boolean {
  return (
    Boolean(locator?.key) &&
    typeof locator?.key_occurrence_index === "number" &&
    typeof locator.key_start === "number" &&
    typeof locator.key_end === "number" &&
    typeof locator.original_start_in_key === "number" &&
    typeof locator.original_end_in_key === "number" &&
    (locator.strategy === "original" || locator.strategy === "context")
  );
}

function isHistoryEntry(value: unknown): value is ProofreadHistoryEntry {
  if (!isRecord(value)) {
    return false;
  }

  return (
    typeof value.id === "string" &&
    (typeof value.historySchemaVersion === "undefined" ||
      typeof value.historySchemaVersion === "number") &&
    typeof value.sessionId === "string" &&
    typeof value.createdAt === "string" &&
    typeof value.textPreview === "string" &&
    typeof value.bookTitle === "string" &&
    typeof value.bookIntroductionPreview === "string" &&
    isTaskState(value.status) &&
    typeof value.issueCount === "number" &&
    typeof value.locatedIssueCount === "number" &&
    typeof value.revisionCount === "number" &&
    typeof value.fallbackCount === "number" &&
    (typeof value.failedCount === "undefined" || typeof value.failedCount === "number") &&
    (typeof value.aiProfileId === "undefined" ||
      value.aiProfileId === null ||
      typeof value.aiProfileId === "string") &&
    isProviderApi(value.providerApi) &&
    isProofreadMode(value.proofreadMode) &&
    typeof value.reasoningEnabled === "boolean" &&
    (typeof value.temperature === "undefined" || typeof value.temperature === "number") &&
    isApplicationMode(value.applicationMode) &&
    isProofreadScope(value.scope) &&
    typeof value.totalChunks === "number" &&
    typeof value.completedChunks === "number" &&
    typeof value.failedChunks === "number" &&
    typeof value.globalLocatedIssueCount === "number" &&
    Array.isArray(value.issues) &&
    value.issues.every(isProofreadIssue) &&
    Array.isArray(value.selectedIssueIds) &&
    value.selectedIssueIds.every((id) => typeof id === "string") &&
    typeof value.skippedIssueCount === "number" &&
    typeof value.insertedComment === "boolean" &&
    typeof value.appliedToWord === "boolean" &&
    (typeof value.replayable === "undefined" || typeof value.replayable === "boolean") &&
    (typeof value.errorMessage === "undefined" || typeof value.errorMessage === "string") &&
    (typeof value.sourceFilename === "undefined" ||
      value.sourceFilename === null ||
      typeof value.sourceFilename === "string") &&
    (typeof value.runId === "undefined" ||
      value.runId === null ||
      typeof value.runId === "string") &&
    (typeof value.outputFilename === "undefined" ||
      value.outputFilename === null ||
      typeof value.outputFilename === "string") &&
    (typeof value.downloadUrl === "undefined" ||
      value.downloadUrl === null ||
      typeof value.downloadUrl === "string") &&
    (typeof value.expiresAt === "undefined" ||
      value.expiresAt === null ||
      typeof value.expiresAt === "string") &&
    (typeof value.retentionDays === "undefined" ||
      value.retentionDays === null ||
      typeof value.retentionDays === "number")
  );
}

function normalizeHistoryEntry(entry: ProofreadHistoryEntry): ProofreadHistoryEntry {
  return {
    ...entry,
    temperature: normalizeTemperature(entry.temperature),
  };
}

function normalizeTemperature(value: unknown): number {
  if (typeof value !== "number" || Number.isNaN(value) || value < 0 || value > 1.5) {
    return DEFAULT_TEMPERATURE;
  }

  return value;
}

function isProofreadIssue(value: unknown): value is ProofreadIssue {
  if (!isRecord(value)) {
    return false;
  }

  return (
    typeof value.id === "string" &&
    typeof value.category === "string" &&
    (value.severity === "low" || value.severity === "medium" || value.severity === "high") &&
    typeof value.original === "string" &&
    (typeof value.replacement === "undefined" ||
      value.replacement === null ||
      typeof value.replacement === "string") &&
    typeof value.suggestion === "string" &&
    (typeof value.start === "undefined" ||
      value.start === null ||
      typeof value.start === "number") &&
    (typeof value.end === "undefined" || value.end === null || typeof value.end === "number") &&
    (typeof value.locator === "undefined" ||
      value.locator === null ||
      isProofreadLocator(value.locator))
  );
}

function isProofreadLocator(value: unknown): value is ProofreadLocator {
  if (!isRecord(value)) {
    return false;
  }

  return (
    typeof value.key === "string" &&
    typeof value.key_start === "number" &&
    typeof value.key_end === "number" &&
    typeof value.original_start_in_key === "number" &&
    typeof value.original_end_in_key === "number" &&
    (value.strategy === "original" || value.strategy === "context") &&
    (typeof value.key_occurrence_index === "undefined" ||
      value.key_occurrence_index === null ||
      typeof value.key_occurrence_index === "number")
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isTaskState(value: unknown): value is TaskState {
  return (
    value === "idle" ||
    value === "queued" ||
    value === "running" ||
    value === "succeeded" ||
    value === "partial_succeeded" ||
    value === "failed" ||
    value === "cancelled"
  );
}

function isProviderApi(value: unknown): value is ProviderAPI {
  return value === "responses" || value === "chat";
}

function isProofreadMode(value: unknown): value is ProofreadMode {
  return value === "fast" || value === "thinking";
}

function isApplicationMode(value: unknown): value is ApplicationMode {
  return value === "comment" || value === "revision";
}

function isProofreadScope(value: unknown): value is ProofreadScope {
  return value === "selection" || value === "document";
}

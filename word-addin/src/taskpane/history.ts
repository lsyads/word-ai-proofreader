/* global Blob, File, URL, document, localStorage */

import {
  ApplicationMode,
  BookInfo,
  IssueApplicationSummary,
  PendingProofreadResult,
  ProofreadHistoryEntry,
  ProofreadIssue,
  ProofreadMode,
  ProofreadScope,
  ProviderAPI,
  TaskState,
} from "./types";

const HISTORY_STORAGE_KEY = "word-ai-proofreader-history-v2";
const MAX_HISTORY_ENTRIES = 20;

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
  scope: ProofreadScope;
  taskId?: string | null;
  totalChunks: number;
  completedChunks: number;
  failedChunks: number;
  sessionId: string;
  providerApi: ProviderAPI;
  proofreadMode: ProofreadMode;
  reasoningEnabled: boolean;
  applicationMode: ApplicationMode;
  selectedIssueIds: string[];
  skippedIssueCount: number;
  errorMessage?: string;
}) {
  if (input.text.trim().length === 0) {
    return;
  }

  const entries = getHistoryEntries();
  const entry: ProofreadHistoryEntry = {
    id: createLocalId(),
    sessionId: input.sessionId,
    createdAt: new Date().toISOString(),
    textPreview: input.text.trim().slice(0, 40),
    bookTitle: input.book.title,
    bookIntroductionPreview: (input.book.introduction || "").trim().slice(0, 40),
    status: input.status,
    issueCount: input.issues.length,
    locatedIssueCount: input.locatedIssueCount,
    revisionCount: input.revisionCount,
    fallbackCount: input.fallbackCount,
    providerApi: input.providerApi,
    proofreadMode: input.proofreadMode,
    reasoningEnabled: input.reasoningEnabled,
    applicationMode: input.applicationMode,
    scope: input.scope,
    taskId: input.taskId,
    totalChunks: input.totalChunks,
    completedChunks: input.completedChunks,
    failedChunks: input.failedChunks,
    globalLocatedIssueCount: input.issues.filter(
      (issue) => typeof issue.start === "number" && typeof issue.end === "number"
    ).length,
    issues: input.issues,
    selectedIssueIds: input.selectedIssueIds,
    skippedIssueCount: input.skippedIssueCount,
    insertedComment: input.insertedComment,
    appliedToWord: input.appliedToWord,
    errorMessage: input.errorMessage,
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
    text: input.result.sourceText,
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
    scope: input.result.scope,
    taskId: input.result.taskId,
    totalChunks: input.result.totalChunks,
    completedChunks: input.result.completedChunks,
    failedChunks: input.result.failedChunks,
    sessionId: input.result.sessionId,
    providerApi: input.result.providerApi,
    proofreadMode: input.result.proofreadMode,
    reasoningEnabled: input.result.reasoningEnabled,
    applicationMode: input.applicationMode,
    selectedIssueIds,
    skippedIssueCount: input.result.issues.length - selectedIssueIds.length,
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
    text: input.result.sourceText,
    book: input.result.book,
    issues: input.result.issues,
    insertedComment: input.summary.commentCount > 0 || input.summary.fallbackCount > 0,
    appliedToWord: true,
    locatedIssueCount: input.summary.commentCount + input.summary.revisionCount,
    revisionCount: input.summary.revisionCount,
    fallbackCount: input.summary.fallbackCount,
    scope: input.result.scope,
    taskId: input.result.taskId,
    totalChunks: input.result.totalChunks,
    completedChunks: input.result.completedChunks,
    failedChunks: input.result.failedChunks,
    sessionId: input.result.sessionId,
    providerApi: input.result.providerApi,
    proofreadMode: input.result.proofreadMode,
    reasoningEnabled: input.result.reasoningEnabled,
    applicationMode: input.applicationMode,
    selectedIssueIds,
    skippedIssueCount: input.result.issues.length - selectedIssues.length,
  });
}

export function getHistoryEntries(): ProofreadHistoryEntry[] {
  try {
    const parsed = JSON.parse(localStorage.getItem(HISTORY_STORAGE_KEY) || "[]");
    return Array.isArray(parsed) && parsed.every(isHistoryEntry)
      ? (parsed as ProofreadHistoryEntry[])
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

function isHistoryEntry(value: unknown): value is ProofreadHistoryEntry {
  if (!isRecord(value)) {
    return false;
  }

  return (
    typeof value.id === "string" &&
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
    isProviderApi(value.providerApi) &&
    isProofreadMode(value.proofreadMode) &&
    typeof value.reasoningEnabled === "boolean" &&
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
    (typeof value.errorMessage === "undefined" || typeof value.errorMessage === "string")
  );
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
    (typeof value.end === "undefined" || value.end === null || typeof value.end === "number")
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

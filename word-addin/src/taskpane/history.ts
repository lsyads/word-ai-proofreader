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
  applicationMode: ApplicationMode;
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
    applicationMode: input.applicationMode,
    errorMessage: input.errorMessage,
  });
}

export function saveAppliedResultHistory(input: {
  result: PendingProofreadResult;
  summary: IssueApplicationSummary;
  applicationMode: ApplicationMode;
}) {
  saveHistoryEntry({
    status: "succeeded",
    text: input.result.sourceText,
    book: input.result.book,
    issues: input.result.issues,
    insertedComment: input.summary.commentCount > 0,
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
    applicationMode: input.applicationMode,
  });
}

export function getHistoryEntries(): ProofreadHistoryEntry[] {
  try {
    const parsed = JSON.parse(localStorage.getItem(HISTORY_STORAGE_KEY) || "[]");
    return Array.isArray(parsed) ? (parsed as ProofreadHistoryEntry[]) : [];
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

  const imported = parsed as ProofreadHistoryEntry[];
  localStorage.setItem(HISTORY_STORAGE_KEY, JSON.stringify(imported.slice(0, MAX_HISTORY_ENTRIES)));
  return imported.length;
}

function createLocalId(): string {
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

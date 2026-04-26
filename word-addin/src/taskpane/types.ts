export interface ProofreadIssue {
  id: string;
  category: string;
  severity: "low" | "medium" | "high";
  original: string;
  replacement?: string | null;
  suggestion: string;
  start?: number | null;
  end?: number | null;
}

export interface ProofreadResponse {
  issues: ProofreadIssue[];
}

export interface ChunkedProofreadIssue extends ProofreadIssue {
  chunk_index: number;
  global_start?: number | null;
  global_end?: number | null;
}

export interface ChunkedProofreadResponse {
  task_id?: string | null;
  scope: ProofreadScope;
  status: TaskState;
  total_chunks: number;
  completed_chunks: number;
  failed_chunks: number;
  issues: ChunkedProofreadIssue[];
  error_message?: string | null;
}

export interface ProofreadStatusEvent {
  stage: string;
  message: string;
  task_id?: string;
  status?: TaskState;
  scope?: ProofreadScope;
  total_chunks?: number;
  completed_chunks?: number;
  failed_chunks?: number;
  issue_count?: number;
  chunk_index?: number;
  chunk_start?: number;
  chunk_end?: number;
  error_message?: string;
}

export interface SessionResponse {
  session_id: string;
  created_at: string;
}

export interface BookInfo {
  title: string;
  introduction?: string | null;
}

export type TaskState =
  | "idle"
  | "queued"
  | "running"
  | "succeeded"
  | "partial_succeeded"
  | "failed"
  | "cancelled";
export type ProviderAPI = "responses" | "chat";
export type ProofreadMode = "fast" | "thinking";
export type ApplicationMode = "comment" | "revision";
export type ProofreadScope = "selection" | "document";

export interface ProofreadHistoryEntry {
  id: string;
  sessionId: string;
  createdAt: string;
  textPreview: string;
  bookTitle: string;
  bookIntroductionPreview: string;
  status: TaskState;
  issueCount: number;
  locatedIssueCount: number;
  revisionCount: number;
  fallbackCount: number;
  providerApi: ProviderAPI;
  proofreadMode: ProofreadMode;
  applicationMode: ApplicationMode;
  scope: ProofreadScope;
  taskId?: string | null;
  totalChunks: number;
  completedChunks: number;
  failedChunks: number;
  globalLocatedIssueCount: number;
  issues: ProofreadIssue[];
  insertedComment: boolean;
  appliedToWord: boolean;
  errorMessage?: string;
}

export interface IssueApplicationSummary {
  commentCount: number;
  revisionCount: number;
  fallbackCount: number;
}

export interface PendingProofreadResult {
  sessionId: string;
  sourceText: string;
  book: BookInfo;
  scope: ProofreadScope;
  taskId: string | null;
  totalChunks: number;
  completedChunks: number;
  failedChunks: number;
  providerApi: ProviderAPI;
  proofreadMode: ProofreadMode;
  issues: ProofreadIssue[];
}

export interface ControlsState {
  providerApi: ProviderAPI;
  proofreadMode: ProofreadMode;
  applicationMode: ApplicationMode;
  scope: ProofreadScope;
}

export const SELECTION_CHUNK_THRESHOLD = 5000;
export const DEFAULT_CHUNK_SIZE = 3000;

export interface ProofreadLocator {
  key: string;
  key_start: number;
  key_end: number;
  original_start_in_key: number;
  original_end_in_key: number;
  strategy: "original" | "context";
  key_occurrence_index?: number | null;
}

export interface ProofreadIssue {
  id: string;
  category: string;
  severity: "low" | "medium" | "high";
  original: string;
  replacement?: string | null;
  suggestion: string;
  start?: number | null;
  end?: number | null;
  locator?: ProofreadLocator | null;
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
  chunk_len?: number;
  elapsed_seconds?: number;
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
  historySchemaVersion?: number;
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
  reasoningEnabled: boolean;
  applicationMode: ApplicationMode;
  scope: ProofreadScope;
  taskId?: string | null;
  totalChunks: number;
  completedChunks: number;
  failedChunks: number;
  globalLocatedIssueCount: number;
  issues: ProofreadIssue[];
  selectedIssueIds: string[];
  skippedIssueCount: number;
  insertedComment: boolean;
  appliedToWord: boolean;
  replayable?: boolean;
  errorMessage?: string;
}

export interface IssueApplicationSummary {
  commentCount: number;
  revisionCount: number;
  fallbackCount: number;
}

export interface IssueApplicationProgress {
  completedBatches: number;
  totalBatches: number;
  completedIssues: number;
  totalIssues: number;
}

export interface PendingProofreadResult {
  sessionId: string;
  sourceText: string;
  sourceTextAvailable?: boolean;
  historyTextPreview?: string;
  book: BookInfo;
  scope: ProofreadScope;
  taskId: string | null;
  totalChunks: number;
  completedChunks: number;
  failedChunks: number;
  providerApi: ProviderAPI;
  proofreadMode: ProofreadMode;
  reasoningEnabled: boolean;
  issues: ProofreadIssue[];
}

export interface ControlsState {
  providerApi: ProviderAPI;
  proofreadMode: ProofreadMode;
  reasoningEnabled: boolean;
  applicationMode: ApplicationMode;
  scope: ProofreadScope;
}

export type IssueSeverityFilter = "all" | "high-medium" | "high" | "medium" | "low";
export type IssueLocationFilter = "all" | "located" | "unlocated";
export type IssueReplacementFilter = "all" | "with-replacement" | "needs-review";

export interface IssueFilterState {
  severity: IssueSeverityFilter;
  category: string;
  location: IssueLocationFilter;
  replacement: IssueReplacementFilter;
}

export interface IssueReviewState {
  selectedIssueIds: string[];
  filter: IssueFilterState;
}

export const SELECTION_CHUNK_THRESHOLD = 7000;
export const DEFAULT_CHUNK_SIZE = 5000;

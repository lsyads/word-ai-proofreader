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
  run_id?: string | null;
}

export interface ChunkedProofreadIssue extends ProofreadIssue {
  chunk_index: number;
  global_start?: number | null;
  global_end?: number | null;
}

export interface ChunkedProofreadResponse {
  task_id?: string | null;
  run_id?: string | null;
  scope: ProofreadScope;
  status: TaskState;
  total_chunks: number;
  completed_chunks: number;
  failed_chunks: number;
  issues: ChunkedProofreadIssue[];
  error_message?: string | null;
}

export interface DocxProofreadResponse {
  task_id: string;
  run_id?: string | null;
  status: TaskState;
  total_chunks: number;
  completed_chunks: number;
  failed_chunks: number;
  issue_count: number;
  source_filename: string;
  application_mode: ApplicationMode;
  output_filename?: string | null;
  download_url?: string | null;
  expires_at?: string | null;
  retention_days?: number | null;
  error_message?: string | null;
}

export interface ProofreadStatusEvent {
  stage: string;
  message: string;
  task_id?: string;
  run_id?: string | null;
  status?: TaskState;
  scope?: ProofreadScope;
  total_chunks?: number;
  completed_chunks?: number;
  failed_chunks?: number;
  issue_count?: number;
  source_filename?: string;
  output_filename?: string | null;
  download_url?: string | null;
  expires_at?: string | null;
  retention_days?: number | null;
  chunk_index?: number;
  chunk_start?: number;
  chunk_end?: number;
  chunk_len?: number;
  elapsed_seconds?: number;
  error_message?: string;
}

export interface AgentNodeTrace {
  node_name: string;
  status: "running" | "succeeded" | "failed";
  started_at: string;
  ended_at?: string | null;
  elapsed_seconds?: number | null;
  error_message?: string | null;
}

export interface AgentChunkTrace {
  chunk_index: number;
  chunk_start: number;
  chunk_end: number;
  chunk_len: number;
  status: "running" | "succeeded" | "failed";
  issue_count: number;
  retry_count: number;
  error_message?: string | null;
  started_at: string;
  ended_at?: string | null;
  elapsed_seconds?: number | null;
}

export interface AgentRunTrace {
  run_id: string;
  flow: string;
  task_id?: string | null;
  status: string;
  created_at: string;
  updated_at: string;
  total_chunks: number;
  completed_chunks: number;
  failed_chunks: number;
  issue_count: number;
  error_message?: string | null;
  metadata: Record<string, unknown>;
  nodes: AgentNodeTrace[];
  chunks: AgentChunkTrace[];
}

export type V2ProjectStatus =
  | "created"
  | "running"
  | "waiting_for_approval"
  | "succeeded"
  | "written"
  | "failed"
  | "cancelled";
export type V2RunStatus =
  | "queued"
  | "running"
  | "waiting_for_approval"
  | "succeeded"
  | "partial_succeeded"
  | "failed"
  | "cancelled";
export type V2CandidateStatus = "pending" | "approved" | "rejected" | "deferred" | "written";

export interface V2Project {
  project_id: string;
  source_type: "selection" | "docx";
  status: V2ProjectStatus;
  source_filename: string;
  text_preview?: string | null;
  book: BookInfo;
  review_goal: string;
  created_at: string;
  updated_at: string;
  run_count: number;
  latest_run_id?: string | null;
  latest_run_status?: V2RunStatus | null;
  latest_run_stage?: string | null;
  candidate_count: number;
  pending_count: number;
  approved_count: number;
  output_filename?: string | null;
  download_url?: string | null;
}

export interface V2ProjectList {
  projects: V2Project[];
}

export interface V2ProjectDeleteResponse {
  project_id: string;
  deleted: boolean;
}

export interface V2DocumentBlock {
  index: number;
  start: number;
  end: number;
  text_preview: string;
  style?: string | null;
  in_textbox: boolean;
}

export interface V2DocumentChunk {
  index: number;
  start: number;
  end: number;
  chunk_len: number;
  text_preview: string;
}

export interface V2DocumentMap {
  project_id: string;
  text_len: number;
  block_count: number;
  chunk_count: number;
  blocks: V2DocumentBlock[];
  chunks: V2DocumentChunk[];
}

export interface V2ReviewPlanStep {
  step_id: string;
  title: string;
  tool_name: string;
  status: "pending" | "running" | "succeeded" | "failed";
  description: string;
}

export interface V2ReviewPlan {
  project_id: string;
  run_id?: string | null;
  steps: V2ReviewPlanStep[];
}

export interface V2Run {
  project_id: string;
  run_id: string;
  status: V2RunStatus;
  stage: string;
  total_chunks: number;
  completed_chunks: number;
  failed_chunks: number;
  candidate_count: number;
  error_message?: string | null;
  created_at: string;
  updated_at: string;
}

export interface V2CandidateIssue {
  candidate_id: string;
  project_id: string;
  run_id: string;
  status: V2CandidateStatus;
  category: string;
  severity: "low" | "medium" | "high";
  original: string;
  replacement?: string | null;
  suggestion: string;
  evidence: string;
  chunk_index: number;
  global_start?: number | null;
  global_end?: number | null;
  locator?: ProofreadLocator | null;
  self_check?: string | null;
  pass_name: string;
  confidence: number;
  evidence_kind: "locator" | "context" | "rule" | "memory" | "document_map";
  rule_id?: string | null;
  needs_human_review: boolean;
  evaluation_note?: string | null;
  created_at: string;
  updated_at: string;
}

export interface V2CandidateList {
  project_id: string;
  candidates: V2CandidateIssue[];
  page: number;
  page_size: number;
  total: number;
  total_pages: number;
  has_previous: boolean;
  has_next: boolean;
}

export interface V2ApprovalDecision {
  candidate_id: string;
  status: "approved" | "rejected" | "deferred";
}

export interface V2ApprovalDecisionResponse {
  project_id: string;
  updated_count: number;
  candidates: V2CandidateIssue[];
}

export interface V2MarkWrittenResponse {
  project_id: string;
  updated_count: number;
  candidates: V2CandidateIssue[];
}

export interface V2WritebackResponse {
  project_id: string;
  output_filename: string;
  download_url: string;
  comment_count: number;
  revision_count: number;
  fallback_count: number;
  failed_count: number;
  written_count: number;
}

export interface V2RunEvent {
  event: string;
  data: Record<string, unknown>;
  created_at: string;
}

export interface V2RunTrace {
  project_id: string;
  run_id: string;
  status: V2RunStatus;
  events: V2RunEvent[];
}

export interface V2MemoryItem {
  memory_id: string;
  project_id: string;
  kind: "terminology" | "style_rule" | "preference" | "book_convention" | "observation";
  key: string;
  value: string;
  source: string;
  confidence: number;
  created_at: string;
  updated_at: string;
}

export interface V2MemoryList {
  project_id: string;
  memory: V2MemoryItem[];
}

export interface V2ReviewReport {
  project_id: string;
  status: V2ProjectStatus;
  source_filename: string;
  book: BookInfo;
  review_goal: string;
  issue_count: number;
  pending_count: number;
  approved_count: number;
  rejected_count: number;
  deferred_count: number;
  written_count: number;
  severity_counts: Record<string, number>;
  category_counts: Record<string, number>;
  pass_counts: Record<string, number>;
  unresolved_items: string[];
  generated_at: string;
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

export interface AIProfile {
  id: string;
  label: string;
  model: string;
  default_api: ProviderAPI;
  supported_apis: ProviderAPI[];
  configured: boolean;
}

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
  failedCount?: number;
  aiProfileId?: string | null;
  providerApi: ProviderAPI;
  proofreadMode: ProofreadMode;
  reasoningEnabled: boolean;
  temperature: number;
  applicationMode: ApplicationMode;
  scope: ProofreadScope;
  taskId?: string | null;
  runId?: string | null;
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
  sourceFilename?: string | null;
  outputFilename?: string | null;
  downloadUrl?: string | null;
  expiresAt?: string | null;
  retentionDays?: number | null;
}

export interface IssueApplicationSummary {
  commentCount: number;
  revisionCount: number;
  fallbackCount: number;
  fallbackCommentCount: number;
  failedCount: number;
  truncatedFallbackCount: number;
}

export interface IssueApplicationProgress {
  stage?: "locating" | "commenting" | "revising" | "fallback";
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
  runId?: string | null;
  totalChunks: number;
  completedChunks: number;
  failedChunks: number;
  aiProfileId?: string | null;
  providerApi: ProviderAPI;
  proofreadMode: ProofreadMode;
  reasoningEnabled: boolean;
  temperature: number;
  issues: ProofreadIssue[];
  issueCount?: number;
  sourceFilename?: string | null;
  outputFilename?: string | null;
  downloadUrl?: string | null;
  expiresAt?: string | null;
  retentionDays?: number | null;
}

export interface ControlsState {
  aiProfileId: string;
  providerApi: ProviderAPI;
  proofreadMode: ProofreadMode;
  reasoningEnabled: boolean;
  temperature: number;
  applicationMode: ApplicationMode;
  fallbackSummaryTruncateEnabled: boolean;
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
export const DEFAULT_TEMPERATURE = 0.2;

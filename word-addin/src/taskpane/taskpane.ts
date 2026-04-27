/* global AbortController, Office, clearInterval, clearTimeout, localStorage, setInterval, setTimeout */

import {
  cancelProofreadTask,
  createSession,
  formatProgressResult,
  isAbortError,
  normalizeChunkedIssuesForScope,
  requestChunkedProofreadTask,
  requestProofread,
} from "./api";
import { appendDebugLog, clearDebugLog } from "./debug";
import {
  clearHistoryEntries,
  exportHistoryEntries,
  getHistoryEntries,
  importHistoryEntries,
  saveAppliedResultHistory,
  saveHistoryEntry,
  savePendingResultHistory,
} from "./history";
import {
  appendProgressStatus,
  BulkSelectionAction,
  createDefaultFilterState,
  formatApplyButtonLabel,
  formatComment,
  formatCompletionMessage,
  getButton,
  getErrorMessage,
  getInput,
  getSelect,
  getTextArea,
  renderEmptyResult,
  renderHistory,
  renderHistoryEntry,
  renderResult,
  resetProgress,
  showMessage,
} from "./render";
import {
  ApplicationMode,
  BookInfo,
  ControlsState,
  IssueFilterState,
  IssueReviewState,
  PendingProofreadResult,
  ProofreadHistoryEntry,
  ProofreadIssue,
  ProofreadStatusEvent,
  ProofreadMode,
  ProofreadScope,
  ProviderAPI,
  SELECTION_CHUNK_THRESHOLD,
  TaskState,
} from "./types";
import {
  applyIssuesToScope,
  ensureWordCommentSupport,
  getDocumentBodyText,
  getSelectedText,
  selectIssueInScope,
} from "./word";

const PROVIDER_API_STORAGE_KEY = "word-ai-proofreader-provider-api-v2";
const PROOFREAD_MODE_STORAGE_KEY = "word-ai-proofreader-mode-v2";
const REASONING_ENABLED_STORAGE_KEY = "word-ai-proofreader-reasoning-enabled-v2";
const APPLICATION_MODE_STORAGE_KEY = "word-ai-proofreader-application-mode-v2";
const PROOFREAD_SCOPE_STORAGE_KEY = "word-ai-proofreader-scope-v2";
const BOOK_TITLE_STORAGE_KEY = "word-ai-proofreader-book-title-v2";
const BOOK_INTRODUCTION_STORAGE_KEY = "word-ai-proofreader-book-introduction-v2";
const CLEAR_HISTORY_CONFIRM_MS = 4000;

let currentSessionId: string | null = null;
let currentAbortController: AbortController | null = null;
let currentTaskId: string | null = null;
let pendingResult: PendingProofreadResult | null = null;
let issueReviewState: IssueReviewState | null = null;
let clearHistoryConfirmTimer: number | null = null;
let isClearHistoryArmed = false;
let taskState: TaskState = "idle";
let isApplyingToWord = false;
let activeChunkTimer: number | null = null;
let activeChunkStartedAtMs = 0;
let activeChunkProgress: ProofreadStatusEvent | null = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    getButton("proofread").onclick = proofreadSelection;
    getButton("apply-to-word").onclick = applyPendingResultToWord;
    getButton("new-conversation").onclick = clearCurrentResult;
    getButton("clear-debug-log").onclick = clearDebugLog;
    getButton("clear-history").onclick = clearHistory;
    getButton("export-history").onclick = exportHistory;
    getButton("import-history").onclick = () => getInput("history-file").click();
    getInput("history-file").onchange = importHistory;
    getInput("book-title").oninput = persistBookInfo;
    getTextArea("book-introduction").oninput = persistBookInfo;
    getSelect("provider-api").onchange = persistControls;
    getSelect("proofread-mode").onchange = persistControls;
    getInput("reasoning-enabled").onchange = persistControls;
    getSelect("application-mode").onchange = persistControls;
    getSelect("proofread-scope").onchange = persistControls;
    initializeControls();
    refreshHistory();
    updateActionButtons();
    initializeSession();
    return;
  }

  getButton("proofread").disabled = true;
  getButton("new-conversation").disabled = true;
  showMessage("请在 Microsoft Word 任务窗格中使用此插件。", "error");
});

async function initializeSession() {
  try {
    const session = await createSession();
    currentSessionId = session.session_id;
    showMessage("已创建审校会话，请选择 Word 文本开始审校。", "default");
  } catch (error) {
    showMessage(`创建审校会话失败：${getErrorMessage(error)}`, "error");
  }
}

async function clearCurrentResult() {
  if (taskState === "running") {
    cancelCurrentProofread();
  }

  pendingResult = null;
  issueReviewState = null;
  stopChunkElapsedTimer();
  setBusy(false);
  resetProgress();
  renderEmptyResult("尚未开始审校");
  updateActionButtons();

  try {
    const session = await createSession();
    currentSessionId = session.session_id;
    taskState = "idle";
    showMessage("已清空当前结果，请选择 Word 文本开始审校。", "success");
  } catch (error) {
    taskState = "failed";
    showMessage(`清空当前结果失败：${getErrorMessage(error)}`, "error");
  }
}

export async function proofreadSelection() {
  if (taskState === "running") {
    showMessage("正在停止审校，当前分块完成后结束。", "default");
    cancelCurrentProofread();
    return;
  }

  const book = getValidatedBookInfo();
  if (!book) {
    return;
  }

  if (!ensureWordCommentSupport()) {
    showMessage("当前 Word 环境不支持批注 API，无法完成审校。", "error");
    return;
  }

  if (!currentSessionId) {
    await initializeSession();
  }

  if (!currentSessionId) {
    return;
  }

  const controls = getControlsState();
  const abortController = new AbortController();
  currentAbortController = abortController;
  currentTaskId = null;
  pendingResult = null;
  issueReviewState = null;
  taskState = "running";
  stopChunkElapsedTimer();
  setBusy(true);
  updateActionButtons();
  showMessage(
    controls.scope === "document" ? "正在审校全书正文..." : "正在审校当前选区...",
    "default"
  );
  resetProgress();
  renderEmptyResult("审校中...");

  let sourceText = "";

  try {
    sourceText =
      controls.scope === "document" ? await getDocumentBodyText() : await getSelectedText();
    const useChunkedFlow =
      controls.scope === "document" || sourceText.length > SELECTION_CHUNK_THRESHOLD;
    let issues = [];
    let taskId: string | null = null;
    let totalChunks = 1;
    let completedChunks = 1;
    let failedChunks = 0;
    let status: TaskState = "succeeded";

    if (useChunkedFlow) {
      const chunkedResult = await requestChunkedProofreadTask(
        sourceText,
        book,
        controls.scope,
        currentSessionId,
        controls.providerApi,
        controls.proofreadMode,
        controls.reasoningEnabled,
        renderChunkedProgress,
        (createdTaskId) => {
          currentTaskId = createdTaskId;
        },
        abortController.signal
      );
      taskId = chunkedResult.task_id || null;
      totalChunks = chunkedResult.total_chunks;
      completedChunks = chunkedResult.completed_chunks;
      failedChunks = chunkedResult.failed_chunks;
      status = chunkedResult.status;
      issues = normalizeChunkedIssuesForScope(chunkedResult.issues);
    } else {
      const proofreadResult = await requestProofread(
        sourceText,
        book,
        currentSessionId,
        controls.providerApi,
        controls.proofreadMode,
        controls.reasoningEnabled,
        (progress) => {
          showMessage(progress.message, "default");
          appendProgressStatus(progress);
          renderEmptyResult(progress.message);
        },
        abortController.signal
      );
      issues = proofreadResult.issues;
    }

    pendingResult = {
      sessionId: currentSessionId,
      sourceText,
      book,
      scope: controls.scope,
      taskId,
      totalChunks,
      completedChunks,
      failedChunks,
      providerApi: controls.providerApi,
      proofreadMode: controls.proofreadMode,
      reasoningEnabled: controls.reasoningEnabled,
      issues,
    };
    issueReviewState = {
      selectedIssueIds: issues.map((issue) => issue.id),
      filter: createDefaultFilterState(),
    };
    renderCurrentPendingResult();
    taskState = status;
    savePendingResultHistory({
      result: pendingResult,
      applicationMode: controls.applicationMode,
      status,
    });
    refreshHistory();
    updateActionButtons();

    const hasUnlocatedIssues = issues.some(
      (issue) => typeof issue.start !== "number" || typeof issue.end !== "number"
    );
    showMessage(
      issues.length > 0
        ? hasUnlocatedIssues
          ? "审校完成，请确认结果后点击“应用到 Word”；未定位问题会合并为汇总批注。"
          : "审校完成，请确认结果后点击“应用到 Word”。"
        : "审校完成，未发现明显问题。",
      status === "partial_succeeded" ? "default" : "success"
    );
  } catch (error) {
    if (isAbortError(error)) {
      taskState = "cancelled";
      appendProgressStatus({ stage: "cancelled", message: "已停止当前审校。" });
      renderEmptyResult("已停止审校");
      saveStoppedOrFailedHistory({
        status: taskState,
        sourceText,
        book,
        controls,
        errorMessage: "用户停止了当前审校。",
      });
      showMessage("已停止当前审校。", "default");
      return;
    }

    taskState = "failed";
    appendProgressStatus({ stage: "failed", message: getErrorMessage(error) });
    renderEmptyResult("审校失败");
    saveStoppedOrFailedHistory({
      status: taskState,
      sourceText,
      book,
      controls,
      errorMessage: getErrorMessage(error),
    });
    showMessage(`审校失败：${getErrorMessage(error)}`, "error");
  } finally {
    stopChunkElapsedTimer();
    currentAbortController = null;
    currentTaskId = null;
    setBusy(false);
    refreshHistory();
    updateActionButtons();
  }
}

async function applyPendingResultToWord() {
  const selectedIssues = getSelectedIssues();
  if (!pendingResult || selectedIssues.length === 0 || isApplyingToWord) {
    return;
  }

  isApplyingToWord = true;
  setBusy(true, { applying: true });
  updateActionButtons();
  showMessage(`正在应用到 Word，${formatIssueApplicationPreview(selectedIssues)}。`, "default");

  try {
    const summary = await applyIssuesToScope(
      pendingResult.sourceText,
      selectedIssues,
      pendingResult.scope,
      getApplicationMode()
    );
    saveAppliedResultHistory({
      result: pendingResult,
      summary,
      applicationMode: getApplicationMode(),
      selectedIssueIds: selectedIssues.map((issue) => issue.id),
    });
    refreshHistory();
    showMessage(formatCompletionMessage(summary), "success");
  } catch (error) {
    showMessage(`应用失败：${getErrorMessage(error)}`, "error");
  } finally {
    isApplyingToWord = false;
    setBusy(false);
    updateActionButtons();
  }
}

function renderCurrentPendingResult() {
  if (!pendingResult || !issueReviewState) {
    return;
  }

  renderResult(pendingResult.issues, formatComment(pendingResult.issues), {
    sourceText: pendingResult.sourceText,
    reviewState: issueReviewState,
    applicationMode: getApplicationMode(),
    onToggleIssue: updateIssueSelection,
    onBulkSelect: updateBulkSelection,
    onFilterChange: updateIssueFilter,
    onLocateIssue: locateIssue,
  });
  updateActionButtons();
}

function updateIssueSelection(issueId: string, selected: boolean) {
  if (!issueReviewState || !pendingResult) {
    return;
  }

  const selectedIssueIds = issueReviewState.selectedIssueIds.filter((id) => id !== issueId);
  if (selected && pendingResult.issues.some((issue) => issue.id === issueId)) {
    selectedIssueIds.push(issueId);
  }

  issueReviewState = {
    ...issueReviewState,
    selectedIssueIds,
  };
  renderCurrentPendingResult();
}

function updateBulkSelection(action: BulkSelectionAction) {
  if (!issueReviewState || !pendingResult) {
    return;
  }

  let selectedIssueIds: string[] = [];
  if (action === "all") {
    selectedIssueIds = pendingResult.issues.map((issue) => issue.id);
  }
  if (action === "high-medium") {
    selectedIssueIds = pendingResult.issues
      .filter((issue) => issue.severity === "high" || issue.severity === "medium")
      .map((issue) => issue.id);
  }
  if (action === "replaceable") {
    selectedIssueIds = pendingResult.issues
      .filter(
        (issue) => typeof issue.replacement === "string" && issue.replacement.trim().length > 0
      )
      .map((issue) => issue.id);
  }

  issueReviewState = {
    ...issueReviewState,
    selectedIssueIds,
  };
  renderCurrentPendingResult();
}

function updateIssueFilter(filter: IssueFilterState) {
  if (!issueReviewState) {
    return;
  }

  issueReviewState = {
    ...issueReviewState,
    filter,
  };
  renderCurrentPendingResult();
}

async function locateIssue(issue: ProofreadIssue) {
  if (!pendingResult) {
    appendDebugLog("warn", "定位已跳过：当前没有待应用结果", {
      issueId: issue.id,
    });
    return;
  }

  appendDebugLog("info", "开始定位 Word 原文", {
    issueId: issue.id,
    category: issue.category,
    severity: issue.severity,
    scope: pendingResult.scope,
    start: issue.start,
    end: issue.end,
    originalLength: issue.original.length,
    sourceTextLength: pendingResult.sourceText.length,
  });

  try {
    await selectIssueInScope(pendingResult.sourceText, issue, pendingResult.scope);
    appendDebugLog("info", "定位完成", {
      issueId: issue.id,
    });
    showMessage("已定位到 Word 原文。", "success");
  } catch (error) {
    appendDebugLog("error", "定位失败", {
      issueId: issue.id,
      error,
    });
    showMessage(`定位失败：${getErrorMessage(error)}`, "error");
  }
}

function getSelectedIssues(): ProofreadIssue[] {
  if (!pendingResult || !issueReviewState) {
    return [];
  }

  const selectedIssueIdSet = new Set(issueReviewState.selectedIssueIds);
  return pendingResult.issues.filter((issue) => selectedIssueIdSet.has(issue.id));
}

function formatIssueApplicationPreview(issues: ProofreadIssue[]): string {
  const revisionCount =
    getApplicationMode() === "revision"
      ? issues.filter((issue) => isLocatedIssue(issue) && hasReplacement(issue)).length
      : 0;
  const commentCount = issues.filter((issue) => {
    if (!isLocatedIssue(issue)) {
      return false;
    }

    return getApplicationMode() === "comment" || !hasReplacement(issue);
  }).length;
  const fallbackCount = issues.filter((issue) => !isLocatedIssue(issue)).length;
  const skippedCount = pendingResult ? pendingResult.issues.length - issues.length : 0;

  return `预计精准批注 ${commentCount} 条，生成修订 ${revisionCount} 条，未定位汇总 ${fallbackCount} 条，跳过 ${skippedCount} 条`;
}

function isLocatedIssue(issue: ProofreadIssue): boolean {
  return typeof issue.start === "number" && typeof issue.end === "number";
}

function hasReplacement(issue: ProofreadIssue): boolean {
  return typeof issue.replacement === "string" && issue.replacement.trim().length > 0;
}

function cancelCurrentProofread() {
  stopChunkElapsedTimer();

  if (currentAbortController) {
    currentAbortController.abort();
  }

  if (currentTaskId) {
    cancelProofreadTask(currentTaskId).catch(() => {
      // The local abort is enough for UI state; task cancellation is best effort.
    });
  }
}

function renderChunkedProgress(progress: ProofreadStatusEvent) {
  showMessage(progress.message, "default");
  appendProgressStatus(progress);
  renderEmptyResult(formatProgressResult(updateChunkElapsedTimer(progress)));
}

function updateChunkElapsedTimer(progress: ProofreadStatusEvent): ProofreadStatusEvent {
  if (progress.stage === "chunk_started") {
    startChunkElapsedTimer(progress);
    return getActiveChunkProgress() || progress;
  }

  if (progress.stage === "heartbeat" && typeof progress.chunk_index === "number") {
    syncChunkElapsedTimer(progress);
    return getActiveChunkProgress() || progress;
  }

  if (isChunkElapsedTimerStopStage(progress.stage)) {
    stopChunkElapsedTimer();
  }

  return progress;
}

function startChunkElapsedTimer(progress: ProofreadStatusEvent) {
  stopChunkElapsedTimer();
  activeChunkProgress = progress;
  activeChunkStartedAtMs = Date.now() - (progress.elapsed_seconds || 0) * 1000;
  activeChunkTimer = setInterval(renderActiveChunkElapsed, 1000);
}

function syncChunkElapsedTimer(progress: ProofreadStatusEvent) {
  if (!activeChunkProgress || activeChunkProgress.chunk_index !== progress.chunk_index) {
    startChunkElapsedTimer(progress);
    return;
  }

  activeChunkProgress = progress;
  if (typeof progress.elapsed_seconds === "number") {
    activeChunkStartedAtMs = Date.now() - progress.elapsed_seconds * 1000;
  }
}

function renderActiveChunkElapsed() {
  const progress = getActiveChunkProgress();
  if (progress) {
    renderEmptyResult(formatProgressResult(progress));
  }
}

function getActiveChunkProgress(): ProofreadStatusEvent | null {
  if (!activeChunkProgress) {
    return null;
  }

  return {
    ...activeChunkProgress,
    elapsed_seconds: Math.max(0, (Date.now() - activeChunkStartedAtMs) / 1000),
  };
}

function stopChunkElapsedTimer() {
  if (activeChunkTimer !== null) {
    clearInterval(activeChunkTimer);
  }

  activeChunkTimer = null;
  activeChunkProgress = null;
  activeChunkStartedAtMs = 0;
}

function isChunkElapsedTimerStopStage(stage: string): boolean {
  return ["chunk_completed", "chunk_failed", "completed", "cancelled", "error", "failed"].includes(
    stage
  );
}

function saveStoppedOrFailedHistory(input: {
  status: TaskState;
  sourceText: string;
  book: BookInfo;
  controls: ControlsState;
  errorMessage: string;
}) {
  saveHistoryEntry({
    status: input.status,
    text: input.sourceText,
    book: input.book,
    issues: [],
    insertedComment: false,
    appliedToWord: false,
    locatedIssueCount: 0,
    revisionCount: 0,
    fallbackCount: 0,
    scope: input.controls.scope,
    taskId: currentTaskId,
    totalChunks: 0,
    completedChunks: 0,
    failedChunks: 0,
    providerApi: input.controls.providerApi,
    proofreadMode: input.controls.proofreadMode,
    reasoningEnabled: input.controls.reasoningEnabled,
    applicationMode: input.controls.applicationMode,
    sessionId: currentSessionId || "",
    selectedIssueIds: [],
    skippedIssueCount: 0,
    errorMessage: input.errorMessage,
  });
}

function refreshHistory() {
  renderHistory(getHistoryEntries(), openHistoryEntry);
}

function openHistoryEntry(entry: ProofreadHistoryEntry) {
  pendingResult = null;
  issueReviewState = {
    selectedIssueIds: entry.selectedIssueIds,
    filter: createDefaultFilterState(),
  };
  renderHistoryEntry(entry);
  updateActionButtons();
}

function initializeControls() {
  getInput("book-title").value = localStorage.getItem(BOOK_TITLE_STORAGE_KEY) || "";
  getTextArea("book-introduction").value =
    localStorage.getItem(BOOK_INTRODUCTION_STORAGE_KEY) || "";
  getSelect("provider-api").value = readStoredProviderApi();
  getSelect("proofread-mode").value = readStoredProofreadMode();
  getInput("reasoning-enabled").checked = readStoredReasoningEnabled();
  getSelect("application-mode").value = readStoredApplicationMode();
  getSelect("proofread-scope").value = readStoredProofreadScope();
}

function persistControls() {
  localStorage.setItem(PROVIDER_API_STORAGE_KEY, getProviderApi());
  localStorage.setItem(PROOFREAD_MODE_STORAGE_KEY, getProofreadMode());
  localStorage.setItem(REASONING_ENABLED_STORAGE_KEY, String(getReasoningEnabled()));
  localStorage.setItem(APPLICATION_MODE_STORAGE_KEY, getApplicationMode());
  localStorage.setItem(PROOFREAD_SCOPE_STORAGE_KEY, getProofreadScope());

  if (pendingResult && issueReviewState) {
    renderCurrentPendingResult();
  }
}

function persistBookInfo() {
  localStorage.setItem(BOOK_TITLE_STORAGE_KEY, getInput("book-title").value);
  localStorage.setItem(BOOK_INTRODUCTION_STORAGE_KEY, getTextArea("book-introduction").value);
}

function getValidatedBookInfo(): BookInfo | null {
  const titleInput = getInput("book-title");
  const title = titleInput.value.trim();
  const introduction = getTextArea("book-introduction").value.trim();

  if (!title) {
    showMessage("请先填写书名。", "error");
    titleInput.focus();
    return null;
  }

  titleInput.value = title;
  getTextArea("book-introduction").value = introduction;
  persistBookInfo();
  return {
    title,
    introduction: introduction || null,
  };
}

function getControlsState(): ControlsState {
  return {
    providerApi: getProviderApi(),
    proofreadMode: getProofreadMode(),
    reasoningEnabled: getReasoningEnabled(),
    applicationMode: getApplicationMode(),
    scope: getProofreadScope(),
  };
}

function getProviderApi(): ProviderAPI {
  const value = getSelect("provider-api").value;
  return isProviderApi(value) ? value : "responses";
}

function getProofreadMode(): ProofreadMode {
  const value = getSelect("proofread-mode").value;
  return isProofreadMode(value) ? value : "fast";
}

function getReasoningEnabled(): boolean {
  return getInput("reasoning-enabled").checked;
}

function getApplicationMode(): ApplicationMode {
  const value = getSelect("application-mode").value;
  return isApplicationMode(value) ? value : "comment";
}

function getProofreadScope(): ProofreadScope {
  const value = getSelect("proofread-scope").value;
  return isProofreadScope(value) ? value : "selection";
}

function readStoredProviderApi(): ProviderAPI {
  const value = localStorage.getItem(PROVIDER_API_STORAGE_KEY);
  return isProviderApi(value) ? value : "responses";
}

function readStoredProofreadMode(): ProofreadMode {
  const value = localStorage.getItem(PROOFREAD_MODE_STORAGE_KEY);
  return isProofreadMode(value) ? value : "fast";
}

function readStoredReasoningEnabled(): boolean {
  return localStorage.getItem(REASONING_ENABLED_STORAGE_KEY) === "true";
}

function readStoredApplicationMode(): ApplicationMode {
  const value = localStorage.getItem(APPLICATION_MODE_STORAGE_KEY);
  return isApplicationMode(value) ? value : "comment";
}

function readStoredProofreadScope(): ProofreadScope {
  const value = localStorage.getItem(PROOFREAD_SCOPE_STORAGE_KEY);
  return isProofreadScope(value) ? value : "selection";
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

function clearHistory() {
  if (!isClearHistoryArmed) {
    armClearHistoryConfirmation();
    return;
  }

  resetClearHistoryConfirmation();
  clearHistoryEntries();
  refreshHistory();
  showMessage("已清空本地历史记录。", "success");
}

function armClearHistoryConfirmation() {
  isClearHistoryArmed = true;
  getButton("clear-history").querySelector(".ms-Button-label").textContent = "确认清空";
  showMessage("再次点击“确认清空”将删除本地历史记录。", "default");

  if (clearHistoryConfirmTimer !== null) {
    clearTimeout(clearHistoryConfirmTimer);
  }

  clearHistoryConfirmTimer = setTimeout(() => {
    resetClearHistoryConfirmation();
  }, CLEAR_HISTORY_CONFIRM_MS);
}

function resetClearHistoryConfirmation() {
  isClearHistoryArmed = false;
  getButton("clear-history").querySelector(".ms-Button-label").textContent = "清空";

  if (clearHistoryConfirmTimer !== null) {
    clearTimeout(clearHistoryConfirmTimer);
    clearHistoryConfirmTimer = null;
  }
}

function exportHistory() {
  const count = exportHistoryEntries();
  showMessage(`已导出 ${count} 条历史记录。`, "success");
}

async function importHistory() {
  const input = getInput("history-file");
  const file = input.files && input.files[0];

  if (!file) {
    return;
  }

  try {
    const count = await importHistoryEntries(file);
    refreshHistory();
    showMessage(`已导入 ${count} 条历史记录。`, "success");
  } catch (error) {
    showMessage(`导入历史失败：${getErrorMessage(error)}`, "error");
  } finally {
    input.value = "";
  }
}

function setBusy(isBusy: boolean, options: { applying?: boolean } = {}) {
  const proofreadButton = getButton("proofread");
  if (isBusy) {
    resetClearHistoryConfirmation();
  }

  proofreadButton.disabled = Boolean(options.applying);
  proofreadButton.querySelector(".ms-Button-label").textContent = isBusy
    ? options.applying
      ? "应用中"
      : "停止审校"
    : "AI 审校";
  getButton("new-conversation").disabled = isBusy;
  getButton("clear-history").disabled = isBusy;
  getButton("export-history").disabled = isBusy;
  getButton("import-history").disabled = isBusy;
  getInput("book-title").disabled = isBusy;
  getTextArea("book-introduction").disabled = isBusy;
  getSelect("provider-api").disabled = isBusy;
  getSelect("proofread-mode").disabled = isBusy;
  getInput("reasoning-enabled").disabled = isBusy;
  getSelect("application-mode").disabled = isBusy;
  getSelect("proofread-scope").disabled = isBusy;
}

function updateActionButtons() {
  const selectedCount = getSelectedIssues().length;
  const canApply = Boolean(pendingResult && selectedCount > 0 && !isApplyingToWord);
  const applyButton = getButton("apply-to-word");

  applyButton.disabled = !canApply || taskState === "running";
  applyButton.querySelector(".ms-Button-label").textContent = formatApplyButtonLabel(selectedCount);
}

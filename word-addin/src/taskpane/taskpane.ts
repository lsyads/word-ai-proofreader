/* global AbortController, Blob, File, Office, Response, URL, clearInterval, clearTimeout, document, fetch, localStorage, setInterval, setTimeout */

import {
  cancelDocxProofreadTask,
  cancelProofreadTask,
  createSession,
  formatProgressResult,
  getAIProfiles,
  getAgentRunTrace,
  getDocxDownloadUrl,
  getProofreadTask,
  isAbortError,
  normalizeChunkedIssuesForScope,
  requestChunkedProofreadTask,
  requestDocxProofreadTask,
  requestProofread,
  retryCurrentDocxProofreadChunk,
  retryCurrentProofreadChunk,
  retryFailedDocxProofreadChunks,
  retryFailedProofreadChunks,
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
  getElement,
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
  AgentRunTrace,
  AIProfile,
  BookInfo,
  ControlsState,
  DEFAULT_TEMPERATURE,
  IssueApplicationSummary,
  IssueFilterState,
  IssueApplicationProgress,
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
  clearTrackedSelectionRange,
  ensureWordCommentSupport,
  getSelectedText,
  selectIssueInScope,
} from "./word";

const PROVIDER_API_STORAGE_KEY = "word-ai-proofreader-provider-api-v2";
const AI_PROFILE_STORAGE_KEY = "word-ai-proofreader-ai-profile-v1";
const PROOFREAD_MODE_STORAGE_KEY = "word-ai-proofreader-mode-v2";
const REASONING_ENABLED_STORAGE_KEY = "word-ai-proofreader-reasoning-enabled-v2";
const TEMPERATURE_STORAGE_KEY = "word-ai-proofreader-temperature-v1";
const APPLICATION_MODE_STORAGE_KEY = "word-ai-proofreader-application-mode-v2";
const FALLBACK_SUMMARY_TRUNCATE_STORAGE_KEY =
  "word-ai-proofreader-fallback-summary-truncate-enabled-v2";
const PROOFREAD_SCOPE_STORAGE_KEY = "word-ai-proofreader-scope-v2";
const BOOK_TITLE_STORAGE_KEY = "word-ai-proofreader-book-title-v2";
const BOOK_INTRODUCTION_STORAGE_KEY = "word-ai-proofreader-book-introduction-v2";
const APPLICATION_PREVIEW_BATCH_SIZE = 16;
const CLEAR_HISTORY_CONFIRM_MS = 4000;
const CURRENT_CHUNK_RETRY_THRESHOLD_SECONDS = 120;

let currentSessionId: string | null = null;
let currentAbortController: AbortController | null = null;
let currentTaskId: string | null = null;
let pendingResult: PendingProofreadResult | null = null;
let issueReviewState: IssueReviewState | null = null;
let clearHistoryConfirmTimer: number | null = null;
let isClearHistoryArmed = false;
let taskState: TaskState = "idle";
let isApplyingToWord = false;
let isRetryingCurrentChunk = false;
let isRetryingFailedChunks = false;
let isDownloadingDocx = false;
let activeChunkTimer: number | null = null;
let activeChunkStartedAtMs = 0;
let activeChunkProgress: ProofreadStatusEvent | null = null;
let currentTaskKind: "text" | "docx" = "text";
let aiProfiles: AIProfile[] = [];
let currentRunId: string | null = null;
let currentTrace: AgentRunTrace | null = null;
let currentTraceError: string | null = null;
let isRefreshingTrace = false;
let pendingTraceRefresh = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    getButton("proofread").onclick = proofreadSelection;
    getButton("apply-to-word").onclick = applyPendingResultToWord;
    getButton("download-docx").onclick = downloadCurrentDocxResult;
    getButton("retry-current-chunk").onclick = retryCurrentChunk;
    getButton("retry-failed-chunks").onclick = retryFailedChunks;
    getButton("new-conversation").onclick = clearCurrentResult;
    getButton("clear-debug-log").onclick = clearDebugLog;
    getButton("clear-history").onclick = clearHistory;
    getButton("export-history").onclick = exportHistory;
    getButton("import-history").onclick = () => getInput("history-file").click();
    getButton("refresh-trace").onclick = refreshCurrentTrace;
    getButton("copy-run-id").onclick = copyCurrentRunId;
    getInput("history-file").onchange = importHistory;
    getInput("book-title").oninput = persistBookInfo;
    getTextArea("book-introduction").oninput = persistBookInfo;
    getSelect("ai-profile").onchange = handleAIProfileChange;
    getSelect("provider-api").onchange = persistControls;
    getSelect("proofread-mode").onchange = persistControls;
    getInput("reasoning-enabled").onchange = persistControls;
    getInput("temperature").onchange = persistControls;
    getSelect("application-mode").onchange = persistControls;
    getInput("fallback-summary-truncate-enabled").onchange = persistControls;
    getSelect("proofread-scope").onchange = persistControls;
    getInput("docx-file").onchange = updateDocxFileOutput;
    initializeControls();
    initializeAIProfiles();
    syncScopeControls();
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

  await clearTrackedSelectionRange();
  pendingResult = null;
  issueReviewState = null;
  setCurrentRunId(null);
  renderAgentTraceSummary();
  stopChunkElapsedTimer();
  setBusy(false);
  resetProgress();
  renderEmptyResult("尚未开始审校");
  updateDocxFileOutput();
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

  const controls = getControlsState();

  if (controls.scope === "selection" && !ensureWordCommentSupport()) {
    showMessage("当前 Word 环境不支持批注 API，无法完成审校。", "error");
    return;
  }

  if (!currentSessionId) {
    await initializeSession();
  }

  if (!currentSessionId) {
    return;
  }

  const abortController = new AbortController();
  currentAbortController = abortController;
  currentTaskId = null;
  setCurrentRunId(null);
  renderAgentTraceSummary();
  currentTaskKind = controls.scope === "document" ? "docx" : "text";
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
    if (controls.scope === "document") {
      await proofreadDocxFile(book, controls, abortController);
      return;
    }

    sourceText = await getSelectedText();
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
        controls.aiProfileId,
        controls.providerApi,
        controls.proofreadMode,
        controls.reasoningEnabled,
        controls.temperature,
        renderChunkedProgress,
        (createdTaskId, createdRunId) => {
          currentTaskId = createdTaskId;
          setCurrentRunId(createdRunId);
          void refreshCurrentTrace({ silent: true });
        },
        abortController.signal
      );
      setCurrentRunId(chunkedResult.run_id || currentRunId);
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
        controls.aiProfileId,
        controls.providerApi,
        controls.proofreadMode,
        controls.reasoningEnabled,
        controls.temperature,
        (progress) => {
          showMessage(progress.message, "default");
          appendProgressStatus(progress);
          renderEmptyResult(progress.message);
        },
        abortController.signal
      );
      setCurrentRunId(proofreadResult.run_id || null);
      issues = proofreadResult.issues;
    }

    pendingResult = {
      sessionId: currentSessionId,
      sourceText,
      book,
      scope: controls.scope,
      taskId,
      runId: currentRunId,
      totalChunks,
      completedChunks,
      failedChunks,
      providerApi: controls.providerApi,
      aiProfileId: controls.aiProfileId,
      proofreadMode: controls.proofreadMode,
      reasoningEnabled: controls.reasoningEnabled,
      temperature: controls.temperature,
      issues,
    };
    issueReviewState = {
      selectedIssueIds: issues.map((issue) => issue.id),
      filter: createDefaultFilterState(),
    };
    renderCurrentPendingResult();
    if (status === "failed" && issues.length === 0) {
      renderEmptyResult("分块审校失败，尚未收到可用问题。可重试失败分块。");
    }
    taskState = status;
    savePendingResultHistory({
      result: pendingResult,
      applicationMode: controls.applicationMode,
      status,
    });
    await refreshCurrentTrace({ silent: true });
    refreshHistory();
    updateActionButtons();

    const hasUnlocatedIssues = issues.some(
      (issue) => typeof issue.start !== "number" || typeof issue.end !== "number"
    );
    if (status === "failed") {
      showMessage("分块审校失败，可点击“重试失败分块”。", "error");
    } else {
      showMessage(
        issues.length > 0
          ? hasUnlocatedIssues
            ? "审校完成，请确认结果后点击“应用到 Word”；未定位问题会合并为汇总批注。"
            : "审校完成，请确认结果后点击“应用到 Word”。"
          : "审校完成，未发现明显问题。",
        status === "partial_succeeded" ? "default" : "success"
      );
    }
  } catch (error) {
    if (isAbortError(error)) {
      taskState = "cancelled";
      appendProgressStatus({ stage: "cancelled", message: "已停止当前审校。" });
      const preserved = await preserveCurrentTaskSnapshotAfterStop({
        status: taskState,
        sourceText,
        book,
        controls,
        errorMessage: "用户停止了当前审校。",
      });
      if (!preserved) {
        renderEmptyResult("已停止审校");
        saveStoppedOrFailedHistory({
          status: taskState,
          sourceText,
          book,
          controls,
          errorMessage: "用户停止了当前审校。",
        });
      }
      showMessage(
        preserved ? "已停止当前审校，已保留已完成分块的问题。" : "已停止当前审校。",
        "default"
      );
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

async function proofreadDocxFile(
  book: BookInfo,
  controls: ControlsState,
  abortController: AbortController
) {
  const file = getSelectedDocxFile();
  if (!file) {
    taskState = "idle";
    renderEmptyResult("尚未开始审校");
    return;
  }

  const docxResult = await requestDocxProofreadTask(
    file,
    book,
    currentSessionId || "",
    controls.aiProfileId,
    controls.providerApi,
    controls.proofreadMode,
    controls.reasoningEnabled,
    controls.temperature,
    controls.applicationMode,
    controls.fallbackSummaryTruncateEnabled,
    renderChunkedProgress,
    (createdTaskId, createdRunId) => {
      currentTaskId = createdTaskId;
      setCurrentRunId(createdRunId);
      void refreshCurrentTrace({ silent: true });
      currentTaskKind = "docx";
    },
    abortController.signal
  );
  setCurrentRunId(docxResult.run_id || currentRunId);

  pendingResult = {
    sessionId: currentSessionId || "",
    sourceText: "",
    sourceTextAvailable: false,
    historyTextPreview: docxResult.source_filename,
    book,
    scope: "document",
    taskId: docxResult.task_id,
    runId: currentRunId,
    totalChunks: docxResult.total_chunks,
    completedChunks: docxResult.completed_chunks,
    failedChunks: docxResult.failed_chunks,
    providerApi: controls.providerApi,
    aiProfileId: controls.aiProfileId,
    proofreadMode: controls.proofreadMode,
    reasoningEnabled: controls.reasoningEnabled,
    temperature: controls.temperature,
    issues: [],
    issueCount: docxResult.issue_count,
    sourceFilename: docxResult.source_filename,
    outputFilename: docxResult.output_filename || null,
    downloadUrl: docxResult.download_url || null,
    expiresAt: docxResult.expires_at || null,
    retentionDays: docxResult.retention_days || null,
  };
  issueReviewState = null;
  taskState = docxResult.status;
  renderDocxPendingResult();
  savePendingResultHistory({
    result: pendingResult,
    applicationMode: controls.applicationMode,
    status: docxResult.status,
  });
  await refreshCurrentTrace({ silent: true });
  refreshHistory();
  updateActionButtons();

  if (docxResult.status === "failed") {
    showMessage("DOCX 全书审校失败，可点击“重试失败分块”。", "error");
    return;
  }

  const outputText = docxResult.output_filename
    ? `已生成：${docxResult.output_filename}${formatRetentionSuffix(docxResult.expires_at)}`
    : "审校完成，但尚未生成可下载文件。";
  showMessage(outputText, docxResult.failed_chunks > 0 ? "default" : "success");
}

async function retryCurrentChunk() {
  if (!currentTaskId || isRetryingCurrentChunk) {
    return;
  }

  const activeProgress = getActiveChunkProgress();
  if (
    !activeProgress ||
    typeof activeProgress.elapsed_seconds !== "number" ||
    activeProgress.elapsed_seconds < CURRENT_CHUNK_RETRY_THRESHOLD_SECONDS
  ) {
    showMessage(
      `当前分块审校超过 ${CURRENT_CHUNK_RETRY_THRESHOLD_SECONDS} 秒后可手动重试。`,
      "default"
    );
    return;
  }

  isRetryingCurrentChunk = true;
  updateActionButtons();

  try {
    if (currentTaskKind === "docx") {
      await retryCurrentDocxProofreadChunk(currentTaskId);
    } else {
      await retryCurrentProofreadChunk(currentTaskId);
    }
    appendProgressStatus({
      stage: "chunk_retry_requested",
      message: "已请求重试当前分块。",
    });
    showMessage("已请求重试当前分块，任务会继续向后处理。", "default");
  } catch (error) {
    showMessage(`重试当前分块失败：${getErrorMessage(error)}`, "error");
  } finally {
    isRetryingCurrentChunk = false;
    updateActionButtons();
  }
}

async function retryFailedChunks() {
  if (!pendingResult?.taskId || pendingResult.failedChunks <= 0 || isRetryingFailedChunks) {
    return;
  }

  const abortController = new AbortController();
  currentAbortController = abortController;
  currentTaskId = pendingResult.taskId;
  taskState = "running";
  isRetryingFailedChunks = true;
  stopChunkElapsedTimer();
  setBusy(true);
  updateActionButtons();
  showMessage("正在重试失败分块...", "default");

  try {
    let remainingFailedChunks = 0;
    if (pendingResult.sourceFilename) {
      currentTaskKind = "docx";
      const retryResult = await retryFailedDocxProofreadChunks(
        pendingResult.taskId,
        renderChunkedProgress,
        abortController.signal
      );
      setCurrentRunId(retryResult.run_id || pendingResult.runId || currentRunId);
      pendingResult = {
        ...pendingResult,
        runId: retryResult.run_id || pendingResult.runId,
        totalChunks: retryResult.total_chunks,
        completedChunks: retryResult.completed_chunks,
        failedChunks: retryResult.failed_chunks,
        issueCount: retryResult.issue_count,
        outputFilename: retryResult.output_filename || null,
        downloadUrl: retryResult.download_url || null,
        expiresAt: retryResult.expires_at || null,
        retentionDays: retryResult.retention_days || null,
      };
      issueReviewState = null;
      taskState = retryResult.status;
      remainingFailedChunks = retryResult.failed_chunks;
      renderDocxPendingResult();
    } else {
      currentTaskKind = "text";
      const retryResult = await retryFailedProofreadChunks(
        pendingResult.taskId,
        renderChunkedProgress,
        abortController.signal
      );
      setCurrentRunId(retryResult.run_id || pendingResult.runId || currentRunId);
      const issues = normalizeChunkedIssuesForScope(retryResult.issues);
      pendingResult = {
        ...pendingResult,
        runId: retryResult.run_id || pendingResult.runId,
        totalChunks: retryResult.total_chunks,
        completedChunks: retryResult.completed_chunks,
        failedChunks: retryResult.failed_chunks,
        issues,
      };
      issueReviewState = {
        selectedIssueIds: issues.map((issue) => issue.id),
        filter: issueReviewState?.filter || createDefaultFilterState(),
      };
      taskState = retryResult.status;
      remainingFailedChunks = retryResult.failed_chunks;
      renderCurrentPendingResult();
    }
    savePendingResultHistory({
      result: pendingResult,
      applicationMode: getApplicationMode(),
      status: taskState,
    });
    await refreshCurrentTrace({ silent: true });
    refreshHistory();
    showMessage(
      remainingFailedChunks > 0
        ? "失败分块已重试，仍有分块失败，可稍后再次重试。"
        : "失败分块已重试完成。",
      remainingFailedChunks > 0 ? "default" : "success"
    );
  } catch (error) {
    if (isAbortError(error)) {
      taskState = "cancelled";
      appendProgressStatus({ stage: "cancelled", message: "已停止失败分块重试。" });
      showMessage("已停止失败分块重试。", "default");
      return;
    }

    taskState = "failed";
    showMessage(`重试失败分块失败：${getErrorMessage(error)}`, "error");
  } finally {
    stopChunkElapsedTimer();
    currentAbortController = null;
    currentTaskId = null;
    isRetryingFailedChunks = false;
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
      getApplicationMode(),
      {
        fallbackSummaryTruncateEnabled: getFallbackSummaryTruncateEnabled(),
        onProgress: (progress) => {
          showMessage(formatApplicationProgress(progress), "default");
        },
      }
    );
    saveAppliedResultHistory({
      result: pendingResult,
      summary,
      applicationMode: getApplicationMode(),
      selectedIssueIds: selectedIssues.map((issue) => issue.id),
    });
    refreshHistory();
    showMessage(formatCompletionMessage(summary), getApplicationMessageType(summary));
  } catch (error) {
    showMessage(`应用失败：${getErrorMessage(error)}`, "error");
  } finally {
    isApplyingToWord = false;
    setBusy(false);
    updateActionButtons();
  }
}

function formatApplicationProgress(progress: IssueApplicationProgress): string {
  const stageLabels = {
    commenting: "正在写入批注",
    fallback: "正在写入汇总批注",
    locating: "正在定位",
    revising: "正在写入修订",
  };
  const stage = progress.stage ? stageLabels[progress.stage] : "正在应用到 Word";

  return `${stage}，第 ${progress.completedBatches}/${progress.totalBatches} 批，已处理 ${progress.completedIssues}/${progress.totalIssues} 条。`;
}

function getApplicationMessageType(
  summary: IssueApplicationSummary
): "default" | "error" | "success" {
  const successCount = summary.commentCount + summary.revisionCount + summary.fallbackCount;

  if (summary.failedCount > 0 && successCount === 0 && summary.truncatedFallbackCount === 0) {
    return "error";
  }

  if (summary.failedCount > 0 || summary.truncatedFallbackCount > 0) {
    return "default";
  }

  return "success";
}

function renderCurrentPendingResult() {
  if (!pendingResult || !issueReviewState) {
    if (pendingResult?.sourceFilename) {
      renderDocxPendingResult();
    }
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

function renderDocxPendingResult() {
  if (!pendingResult?.sourceFilename) {
    return;
  }

  const issueCount = pendingResult.issueCount || 0;
  const outputText = pendingResult.outputFilename
    ? `已生成审校后 Word：${pendingResult.outputFilename}${formatRetentionSuffix(pendingResult.expiresAt)}`
    : "尚未生成审校后 Word。";
  getElement("docx-output").textContent = outputText;
  renderEmptyResult(
    `${pendingResult.sourceFilename}：已审校 ${pendingResult.completedChunks}/${pendingResult.totalChunks} 块，失败 ${pendingResult.failedChunks} 块，累计问题 ${issueCount} 条。${outputText}`
  );
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
  const batchCount = Math.ceil(
    countPreciselyWritableIssues(issues) / APPLICATION_PREVIEW_BATCH_SIZE
  );
  const revisionCount =
    getApplicationMode() === "revision"
      ? issues.filter((issue) => isPreciselyWritableIssue(issue) && hasReplacement(issue)).length
      : 0;
  const commentCount = issues.filter((issue) => {
    if (!isPreciselyWritableIssue(issue)) {
      return false;
    }

    return getApplicationMode() === "comment" || !hasReplacement(issue);
  }).length;
  const fallbackCount = issues.filter((issue) => !isPreciselyWritableIssue(issue)).length;
  const skippedCount = pendingResult ? pendingResult.issues.length - issues.length : 0;
  const batchText = batchCount > 0 ? `，预计分 ${batchCount} 批应用` : "";

  return `预计精准批注 ${commentCount} 条，生成修订并附批注 ${revisionCount} 条，将汇总批注 ${fallbackCount} 条，跳过 ${skippedCount} 条${batchText}`;
}

function isLocatedIssue(issue: ProofreadIssue): boolean {
  return typeof issue.start === "number" && typeof issue.end === "number";
}

function isPreciselyWritableIssue(issue: ProofreadIssue): boolean {
  if (!isLocatedIssue(issue)) {
    return false;
  }

  if (pendingResult?.sourceTextAvailable === false) {
    return hasReplayableLocator(issue);
  }

  if (issue.locator) {
    return true;
  }

  return Boolean(
    pendingResult &&
    pendingResult.sourceText.slice(issue.start as number, issue.end as number) === issue.original
  );
}

function countPreciselyWritableIssues(issues: ProofreadIssue[]): number {
  return issues.filter(isPreciselyWritableIssue).length;
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
    const cancelTask = currentTaskKind === "docx" ? cancelDocxProofreadTask : cancelProofreadTask;
    cancelTask(currentTaskId).catch(() => {
      // The local abort is enough for UI state; task cancellation is best effort.
    });
  }
}

function renderChunkedProgress(progress: ProofreadStatusEvent) {
  if (progress.run_id) {
    setCurrentRunId(progress.run_id);
  }
  showMessage(progress.message, "default");
  appendProgressStatus(progress);
  renderEmptyResult(formatProgressResult(updateChunkElapsedTimer(progress)));
}

function setCurrentRunId(runId: string | null | undefined) {
  const normalizedRunId = runId || null;
  if (currentRunId === normalizedRunId) {
    updateTraceButtons();
    return;
  }

  currentRunId = normalizedRunId;
  currentTrace = null;
  currentTraceError = null;
  renderAgentTraceSummary();
}

async function refreshCurrentTrace(options: { silent?: boolean } = {}) {
  if (!currentRunId) {
    renderAgentTraceSummary();
    return;
  }

  if (isRefreshingTrace) {
    pendingTraceRefresh = true;
    return;
  }

  isRefreshingTrace = true;
  const requestedRunId = currentRunId;
  updateTraceButtons();
  if (!options.silent) {
    showMessage("正在刷新 Agent trace。", "default");
  }

  try {
    const trace = await getAgentRunTrace(requestedRunId, new AbortController().signal);
    if (currentRunId !== requestedRunId) {
      pendingTraceRefresh = true;
      return;
    }
    currentTrace = trace;
    currentTraceError = null;
    renderAgentTraceSummary();
    if (!options.silent) {
      showMessage("Agent trace 已刷新。", "success");
    }
  } catch (error) {
    if (isAbortError(error)) {
      return;
    }
    if (currentRunId !== requestedRunId) {
      pendingTraceRefresh = true;
      return;
    }
    currentTrace = null;
    currentTraceError = getErrorMessage(error);
    renderAgentTraceSummary();
    if (!options.silent) {
      showMessage(`刷新 Agent trace 失败：${currentTraceError}`, "error");
    }
  } finally {
    isRefreshingTrace = false;
    updateTraceButtons();
    if (pendingTraceRefresh) {
      pendingTraceRefresh = false;
      void refreshCurrentTrace({ silent: true });
    }
  }
}

function renderAgentTraceSummary() {
  const container = getElement("agent-trace");
  container.innerHTML = "";

  if (!currentRunId) {
    container.className = "trace-empty";
    container.textContent = "暂无 Agent trace";
    updateTraceButtons();
    return;
  }

  if (currentTraceError) {
    container.className = "trace-empty";
    container.textContent = `${formatShortRunId(currentRunId)}：${currentTraceError}`;
    updateTraceButtons();
    return;
  }

  if (!currentTrace) {
    container.className = "trace-empty";
    container.textContent = `${formatShortRunId(currentRunId)}：尚未加载 Agent trace`;
    updateTraceButtons();
    return;
  }

  container.className = "trace-summary";
  appendTraceLine(container, "trace-overview", formatTraceOverview(currentTrace));
  appendTraceLine(container, "trace-detail", formatTraceChunkSummary(currentTrace));

  const recentNodes = currentTrace.nodes.slice(-4);
  if (recentNodes.length > 0) {
    const list = document.createElement("div");
    list.className = "trace-node-list";
    recentNodes.forEach((node) => {
      appendTraceLine(
        list,
        `trace-node ${node.status === "failed" ? "is-failed" : ""}`,
        `${node.node_name} / ${formatTraceStatus(node.status)} / ${formatNullableElapsed(node.elapsed_seconds)}`
      );
    });
    container.appendChild(list);
  }

  const failedChunks = currentTrace.chunks.filter((chunk) => chunk.status === "failed").slice(0, 3);
  if (failedChunks.length > 0) {
    const errors = document.createElement("div");
    errors.className = "trace-error-list";
    failedChunks.forEach((chunk) => {
      appendTraceLine(
        errors,
        "trace-chunk-error",
        `chunk ${chunk.chunk_index + 1} 失败：${chunk.error_message || "未返回失败原因"}`
      );
    });
    container.appendChild(errors);
  }

  updateTraceButtons();
}

function appendTraceLine(parent: HTMLElement, className: string, text: string) {
  const line = document.createElement("div");
  line.className = className.trim();
  line.textContent = text;
  parent.appendChild(line);
}

function formatTraceOverview(trace: AgentRunTrace): string {
  return `${formatShortRunId(trace.run_id)} / ${trace.flow} / ${formatTraceStatus(trace.status)} / 节点 ${trace.nodes.length} 个`;
}

function formatTraceChunkSummary(trace: AgentRunTrace): string {
  const retryCount = trace.chunks.reduce((total, chunk) => total + chunk.retry_count, 0);
  return `chunk ${trace.completed_chunks}/${trace.total_chunks}，失败 ${trace.failed_chunks}，重试 ${retryCount}，问题 ${trace.issue_count}`;
}

function formatShortRunId(runId: string): string {
  return `run_id ...${runId.slice(-8)}`;
}

function formatTraceStatus(status: string): string {
  const labels: Record<string, string> = {
    queued: "排队中",
    running: "运行中",
    succeeded: "成功",
    partial_succeeded: "部分成功",
    failed: "失败",
    cancelled: "已取消",
  };
  return labels[status] || status;
}

function formatNullableElapsed(seconds?: number | null): string {
  return typeof seconds === "number" ? `${seconds.toFixed(2)} 秒` : "耗时待定";
}

function updateTraceButtons() {
  const hasRunId = Boolean(currentRunId);
  getButton("refresh-trace").disabled = !hasRunId || isRefreshingTrace;
  getButton("refresh-trace").querySelector(".ms-Button-label").textContent = isRefreshingTrace
    ? "刷新中"
    : "刷新 Trace";
  getButton("copy-run-id").disabled = !hasRunId;
}

function copyCurrentRunId() {
  if (!currentRunId) {
    return;
  }

  const textarea = document.createElement("textarea");
  textarea.value = currentRunId;
  textarea.setAttribute("readonly", "true");
  textarea.style.position = "fixed";
  textarea.style.left = "-9999px";
  document.body.appendChild(textarea);
  textarea.select();
  const copied = document.execCommand("copy");
  textarea.remove();
  showMessage(copied ? "run_id 已复制。" : currentRunId, copied ? "success" : "default");
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
    updateActionButtons();
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
    runId: currentRunId,
    totalChunks: 0,
    completedChunks: 0,
    failedChunks: 0,
    providerApi: input.controls.providerApi,
    aiProfileId: input.controls.aiProfileId,
    proofreadMode: input.controls.proofreadMode,
    reasoningEnabled: input.controls.reasoningEnabled,
    temperature: input.controls.temperature,
    applicationMode: input.controls.applicationMode,
    sessionId: currentSessionId || "",
    selectedIssueIds: [],
    skippedIssueCount: 0,
    errorMessage: input.errorMessage,
  });
}

async function preserveCurrentTaskSnapshotAfterStop(input: {
  status: TaskState;
  sourceText: string;
  book: BookInfo;
  controls: ControlsState;
  errorMessage: string;
}): Promise<boolean> {
  if (!currentTaskId || !input.sourceText.trim()) {
    return false;
  }

  try {
    const task = await getProofreadTask(currentTaskId, new AbortController().signal);
    setCurrentRunId(task.run_id || currentRunId);
    const issues = normalizeChunkedIssuesForScope(task.issues);
    if (issues.length === 0) {
      return false;
    }

    pendingResult = {
      sessionId: currentSessionId || "",
      sourceText: input.sourceText,
      book: input.book,
      scope: task.scope,
      taskId: task.task_id || currentTaskId,
      runId: task.run_id || currentRunId,
      totalChunks: task.total_chunks,
      completedChunks: task.completed_chunks,
      failedChunks: task.failed_chunks,
      providerApi: input.controls.providerApi,
      aiProfileId: input.controls.aiProfileId,
      proofreadMode: input.controls.proofreadMode,
      reasoningEnabled: input.controls.reasoningEnabled,
      temperature: input.controls.temperature,
      issues,
    };
    issueReviewState = {
      selectedIssueIds: issues.map((issue) => issue.id),
      filter: createDefaultFilterState(),
    };
    renderCurrentPendingResult();
    savePendingResultHistory({
      result: pendingResult,
      applicationMode: input.controls.applicationMode,
      status: input.status,
      errorMessage: input.errorMessage,
    });
    refreshHistory();
    return true;
  } catch {
    return false;
  }
}

function refreshHistory() {
  renderHistory(getHistoryEntries(), openHistoryEntry);
}

function openHistoryEntry(entry: ProofreadHistoryEntry) {
  setCurrentRunId(entry.runId || null);
  issueReviewState = {
    selectedIssueIds: entry.selectedIssueIds,
    filter: createDefaultFilterState(),
  };

  if (entry.sourceFilename) {
    pendingResult = {
      sessionId: entry.sessionId,
      sourceText: "",
      sourceTextAvailable: false,
      historyTextPreview: entry.textPreview,
      book: {
        title: entry.bookTitle || "历史记录",
        introduction: entry.bookIntroductionPreview || null,
      },
      scope: "document",
      taskId: entry.taskId || null,
      runId: entry.runId || null,
      totalChunks: entry.totalChunks,
      completedChunks: entry.completedChunks,
      failedChunks: entry.failedChunks,
      providerApi: entry.providerApi,
      aiProfileId: entry.aiProfileId || "default",
      proofreadMode: entry.proofreadMode,
      reasoningEnabled: entry.reasoningEnabled,
      temperature: entry.temperature,
      issues: [],
      issueCount: entry.issueCount,
      sourceFilename: entry.sourceFilename,
      outputFilename: entry.outputFilename || null,
      downloadUrl: entry.downloadUrl || null,
      expiresAt: entry.expiresAt || null,
      retentionDays: entry.retentionDays || null,
    };
    issueReviewState = null;
    taskState = entry.status;
    getSelect("application-mode").value = entry.applicationMode;
    getSelect("proofread-scope").value = "document";
    syncScopeControls();
    renderDocxPendingResult();
    showMessage("已打开 DOCX 全书审校历史记录。", "default");
    return;
  }

  if (isReplayableHistoryEntry(entry)) {
    pendingResult = {
      sessionId: entry.sessionId,
      sourceText: "",
      sourceTextAvailable: false,
      historyTextPreview: entry.textPreview,
      book: {
        title: entry.bookTitle || "历史记录",
        introduction: entry.bookIntroductionPreview || null,
      },
      scope: entry.scope,
      taskId: entry.taskId || null,
      runId: entry.runId || null,
      totalChunks: entry.totalChunks,
      completedChunks: entry.completedChunks,
      failedChunks: entry.failedChunks,
      providerApi: entry.providerApi,
      aiProfileId: entry.aiProfileId || "default",
      proofreadMode: entry.proofreadMode,
      reasoningEnabled: entry.reasoningEnabled,
      temperature: entry.temperature,
      issues: entry.issues,
    };
    taskState = entry.status;
    getSelect("application-mode").value = entry.applicationMode;
    renderCurrentPendingResult();
    showMessage("已从历史恢复，可再次应用到当前 Word 文档。", "default");
  } else {
    pendingResult = null;
    renderHistoryEntry(entry);
    showMessage("该历史记录缺少可回写定位包，已按只读方式打开。", "default");
  }

  updateActionButtons();
}

function isReplayableHistoryEntry(entry: ProofreadHistoryEntry): boolean {
  return entry.replayable === true && entry.issues.some(hasReplayableLocator);
}

function hasReplayableLocator(issue: ProofreadIssue): boolean {
  return (
    Boolean(issue.locator?.key) &&
    typeof issue.locator?.key_occurrence_index === "number" &&
    typeof issue.locator.original_start_in_key === "number" &&
    typeof issue.locator.original_end_in_key === "number"
  );
}

function initializeControls() {
  getInput("book-title").value = localStorage.getItem(BOOK_TITLE_STORAGE_KEY) || "";
  getTextArea("book-introduction").value =
    localStorage.getItem(BOOK_INTRODUCTION_STORAGE_KEY) || "";
  getSelect("ai-profile").value = readStoredAIProfileId();
  getSelect("provider-api").value = readStoredProviderApi();
  getSelect("proofread-mode").value = readStoredProofreadMode();
  getInput("reasoning-enabled").checked = readStoredReasoningEnabled();
  getInput("temperature").value = formatTemperature(readStoredTemperature());
  getSelect("application-mode").value = readStoredApplicationMode();
  getInput("fallback-summary-truncate-enabled").checked =
    readStoredFallbackSummaryTruncateEnabled();
  getSelect("proofread-scope").value = readStoredProofreadScope();
  syncScopeControls();
}

async function initializeAIProfiles() {
  try {
    aiProfiles = await getAIProfiles();
    renderAIProfileOptions();
    syncProviderApiOptionsForProfile();
    persistControls();
  } catch (error) {
    aiProfiles = [];
    syncProviderApiOptionsForProfile();
    appendDebugLog(`AI 配置列表加载失败：${getErrorMessage(error)}`);
  }
}

function renderAIProfileOptions() {
  const select = getSelect("ai-profile");
  const storedProfileId = readStoredAIProfileId();
  const selectedProfileId = aiProfiles.some((profile) => profile.id === storedProfileId)
    ? storedProfileId
    : aiProfiles[0]?.id || "default";

  select.innerHTML = "";
  const profiles = aiProfiles.length > 0 ? aiProfiles : [defaultAIProfile()];
  profiles.forEach((profile) => {
    const option = document.createElement("option");
    option.value = profile.id;
    option.textContent = `${profile.label} (${profile.model})${profile.configured ? "" : " - 未配置 Key"}`;
    select.appendChild(option);
  });
  select.value = selectedProfileId;
}

function handleAIProfileChange() {
  syncProviderApiOptionsForProfile();
  persistControls();
}

function persistControls() {
  localStorage.setItem(AI_PROFILE_STORAGE_KEY, getAIProfileId());
  localStorage.setItem(PROVIDER_API_STORAGE_KEY, getProviderApi());
  localStorage.setItem(PROOFREAD_MODE_STORAGE_KEY, getProofreadMode());
  localStorage.setItem(REASONING_ENABLED_STORAGE_KEY, String(getReasoningEnabled()));
  localStorage.setItem(TEMPERATURE_STORAGE_KEY, String(getTemperature()));
  localStorage.setItem(APPLICATION_MODE_STORAGE_KEY, getApplicationMode());
  localStorage.setItem(
    FALLBACK_SUMMARY_TRUNCATE_STORAGE_KEY,
    String(getFallbackSummaryTruncateEnabled())
  );
  localStorage.setItem(PROOFREAD_SCOPE_STORAGE_KEY, getProofreadScope());
  syncScopeControls();

  if (pendingResult && issueReviewState) {
    renderCurrentPendingResult();
  }
}

function getSelectedDocxFile(): File | null {
  const input = getInput("docx-file");
  const file = input.files && input.files[0];

  if (!file) {
    showMessage("请选择需要全书审校的 .docx 文件。", "error");
    input.focus();
    return null;
  }

  const lowerName = file.name.toLowerCase();
  if (lowerName.endsWith(".doc")) {
    showMessage(".doc 是旧二进制格式，请先另存为 .docx 后再上传。", "error");
    return null;
  }

  if (!lowerName.endsWith(".docx")) {
    showMessage("全书文件审校当前仅支持 .docx。", "error");
    return null;
  }

  return file;
}

function syncScopeControls() {
  const isDocumentScope = getProofreadScope() === "document";
  getElement("docx-file-panel").hidden = !isDocumentScope;
  getButton("apply-to-word").hidden = isDocumentScope;
  getButton("download-docx").hidden = !isDocumentScope;
  updateDocxFileOutput();
  updateActionButtons();
}

function updateDocxFileOutput() {
  const output = getElement("docx-output");
  const file = getInput("docx-file").files?.[0];
  if (pendingResult?.sourceFilename && pendingResult.outputFilename) {
    output.textContent = `已生成：${pendingResult.outputFilename}${formatRetentionSuffix(pendingResult.expiresAt)}`;
    return;
  }
  output.textContent = file ? `已选择：${file.name}` : "尚未生成审校后 Word。";
}

async function downloadCurrentDocxResult() {
  if (!pendingResult?.taskId || !pendingResult.outputFilename) {
    showMessage("尚未生成可下载的审校后 Word。", "error");
    return;
  }

  if (isDownloadingDocx) {
    return;
  }

  isDownloadingDocx = true;
  updateActionButtons();
  showMessage(`正在准备下载：${pendingResult.outputFilename}`, "default");

  try {
    const response = await fetch(
      pendingResult.downloadUrl || getDocxDownloadUrl(pendingResult.taskId)
    );
    if (!response.ok) {
      throw new Error(await getDocxDownloadErrorMessage(response));
    }

    const blob = await response.blob();
    triggerBlobDownload(blob, pendingResult.outputFilename);
    showMessage(`已开始下载：${pendingResult.outputFilename}`, "success");
  } catch (error) {
    showMessage(`下载失败：${getErrorMessage(error)}`, "error");
  } finally {
    isDownloadingDocx = false;
    updateActionButtons();
  }
}

function triggerBlobDownload(blob: Blob, filename: string) {
  const objectUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = objectUrl;
  link.download = filename;
  link.style.display = "none";
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
}

async function getDocxDownloadErrorMessage(response: Response): Promise<string> {
  try {
    const payload = (await response.json()) as { detail?: string };
    if (payload.detail) {
      if (response.status === 404 || response.status === 409) {
        return `${payload.detail}。历史记录只保存下载入口；如果结果文件已过期或已被清理，需要重新审校生成。`;
      }
      return payload.detail;
    }
  } catch {
    // Fall through to the generic HTTP status message.
  }

  return `后端返回 HTTP ${response.status}`;
}

function formatRetentionSuffix(expiresAt?: string | null): string {
  if (!expiresAt) {
    return "";
  }

  const expiresDate = new Date(expiresAt);
  if (Number.isNaN(expiresDate.getTime())) {
    return "";
  }

  return `，保留至 ${expiresDate.toLocaleString()}`;
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
    aiProfileId: getAIProfileId(),
    providerApi: getProviderApi(),
    proofreadMode: getProofreadMode(),
    reasoningEnabled: getReasoningEnabled(),
    temperature: getTemperature(),
    applicationMode: getApplicationMode(),
    fallbackSummaryTruncateEnabled: getFallbackSummaryTruncateEnabled(),
    scope: getProofreadScope(),
  };
}

function getAIProfileId(): string {
  return getSelect("ai-profile").value || "default";
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

function getTemperature(): number {
  const input = getInput("temperature");
  const rawValue = input.value.trim();
  const temperature = normalizeTemperature(rawValue);
  const wasInvalid =
    !rawValue || Number.isNaN(Number(rawValue)) || Number(rawValue) < 0 || Number(rawValue) > 1.5;
  input.value = formatTemperature(temperature);
  if (wasInvalid) {
    showMessage("Temperature 已恢复默认值 0.2。", "default");
  }
  return temperature;
}

function getApplicationMode(): ApplicationMode {
  const value = getSelect("application-mode").value;
  return isApplicationMode(value) ? value : "comment";
}

function getFallbackSummaryTruncateEnabled(): boolean {
  return getInput("fallback-summary-truncate-enabled").checked;
}

function getProofreadScope(): ProofreadScope {
  const value = getSelect("proofread-scope").value;
  return isProofreadScope(value) ? value : "selection";
}

function readStoredProviderApi(): ProviderAPI {
  const value = localStorage.getItem(PROVIDER_API_STORAGE_KEY);
  return isProviderApi(value) ? value : "responses";
}

function readStoredAIProfileId(): string {
  return localStorage.getItem(AI_PROFILE_STORAGE_KEY) || "default";
}

function readStoredProofreadMode(): ProofreadMode {
  const value = localStorage.getItem(PROOFREAD_MODE_STORAGE_KEY);
  return isProofreadMode(value) ? value : "fast";
}

function readStoredReasoningEnabled(): boolean {
  return localStorage.getItem(REASONING_ENABLED_STORAGE_KEY) === "true";
}

function readStoredTemperature(): number {
  return normalizeTemperature(localStorage.getItem(TEMPERATURE_STORAGE_KEY));
}

function readStoredApplicationMode(): ApplicationMode {
  const value = localStorage.getItem(APPLICATION_MODE_STORAGE_KEY);
  return isApplicationMode(value) ? value : "comment";
}

function readStoredFallbackSummaryTruncateEnabled(): boolean {
  const value = localStorage.getItem(FALLBACK_SUMMARY_TRUNCATE_STORAGE_KEY);
  return value === null ? true : value === "true";
}

function readStoredProofreadScope(): ProofreadScope {
  const value = localStorage.getItem(PROOFREAD_SCOPE_STORAGE_KEY);
  return isProofreadScope(value) ? value : "selection";
}

function isProviderApi(value: unknown): value is ProviderAPI {
  return value === "responses" || value === "chat";
}

function normalizeTemperature(value: string | null): number {
  const rawValue = value === null ? "" : value.trim();
  const parsed = Number(rawValue);

  if (!rawValue || Number.isNaN(parsed) || parsed < 0 || parsed > 1.5) {
    return DEFAULT_TEMPERATURE;
  }

  return Math.round(parsed * 100) / 100;
}

function formatTemperature(value: number): string {
  return Number.isInteger(value) ? String(value) : String(value);
}

function syncProviderApiOptionsForProfile() {
  const profile = aiProfiles.find((item) => item.id === getAIProfileId());
  const supportedApis = profile?.supported_apis || ["responses", "chat"];
  const select = getSelect("provider-api");

  Array.from(select.options).forEach((option) => {
    option.disabled = !supportedApis.includes(option.value as ProviderAPI);
  });

  if (!supportedApis.includes(getProviderApi())) {
    select.value = profile?.default_api || supportedApis[0] || "responses";
  }
}

function defaultAIProfile(): AIProfile {
  return {
    id: "default",
    label: "Default AI (.env)",
    model: "default",
    default_api: "responses",
    supported_apis: ["responses", "chat"],
    configured: false,
  };
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
  getSelect("ai-profile").disabled = isBusy;
  getSelect("provider-api").disabled = isBusy;
  getSelect("proofread-mode").disabled = isBusy;
  getInput("reasoning-enabled").disabled = isBusy;
  getInput("temperature").disabled = isBusy;
  getSelect("application-mode").disabled = isBusy;
  getSelect("proofread-scope").disabled = isBusy;
  getInput("docx-file").disabled = isBusy;
}

function updateActionButtons() {
  const selectedCount = getSelectedIssues().length;
  const isDocxResult = Boolean(pendingResult?.sourceFilename);
  const canApply = Boolean(
    pendingResult && selectedCount > 0 && !isApplyingToWord && !isDocxResult
  );
  const applyButton = getButton("apply-to-word");
  const downloadButton = getButton("download-docx");
  const retryCurrentButton = getButton("retry-current-chunk");
  const retryFailedButton = getButton("retry-failed-chunks");
  const activeProgress = getActiveChunkProgress();
  const canRetryCurrent = Boolean(
    currentTaskId &&
    taskState === "running" &&
    activeProgress &&
    typeof activeProgress.elapsed_seconds === "number" &&
    activeProgress.elapsed_seconds >= CURRENT_CHUNK_RETRY_THRESHOLD_SECONDS &&
    !isRetryingCurrentChunk &&
    !isRetryingFailedChunks &&
    !isApplyingToWord
  );
  const canRetryFailed = Boolean(
    pendingResult?.taskId &&
    pendingResult.failedChunks > 0 &&
    taskState !== "running" &&
    !isRetryingFailedChunks &&
    !isApplyingToWord
  );

  applyButton.disabled = !canApply || taskState === "running";
  applyButton.querySelector(".ms-Button-label").textContent = formatApplyButtonLabel(selectedCount);
  downloadButton.disabled =
    !pendingResult?.outputFilename || taskState === "running" || isDownloadingDocx;
  downloadButton.querySelector(".ms-Button-label").textContent = isDownloadingDocx
    ? "下载中"
    : "下载审校后 Word";
  retryCurrentButton.disabled = !canRetryCurrent;
  retryFailedButton.disabled = !canRetryFailed;
  updateTraceButtons();
}

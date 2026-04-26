/* global AbortController, Office, clearTimeout, localStorage, setTimeout */

import {
  cancelProofreadTask,
  createSession,
  formatProgressResult,
  isAbortError,
  normalizeChunkedIssuesForScope,
  requestChunkedProofreadTask,
  requestProofread,
} from "./api";
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
  PendingProofreadResult,
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
let clearHistoryConfirmTimer: number | null = null;
let isClearHistoryArmed = false;
let taskState: TaskState = "idle";
let isApplyingToWord = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    getButton("proofread").onclick = proofreadSelection;
    getButton("apply-to-word").onclick = applyPendingResultToWord;
    getButton("new-conversation").onclick = newConversation;
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
    showMessage("已创建 AI 对话，请选择 Word 文本开始审校。", "default");
  } catch (error) {
    showMessage(`创建 AI 对话失败：${getErrorMessage(error)}`, "error");
  }
}

async function newConversation() {
  if (taskState === "running") {
    cancelCurrentProofread();
  }

  pendingResult = null;
  setBusy(false);
  resetProgress();
  renderEmptyResult("尚未开始审校");
  updateActionButtons();

  try {
    const session = await createSession();
    currentSessionId = session.session_id;
    taskState = "idle";
    showMessage("已新建 AI 对话，请选择 Word 文本开始审校。", "success");
  } catch (error) {
    taskState = "failed";
    showMessage(`新建 AI 对话失败：${getErrorMessage(error)}`, "error");
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
  taskState = "running";
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
        (progress) => {
          showMessage(progress.message, "default");
          appendProgressStatus(progress);
          renderEmptyResult(formatProgressResult(progress));
        },
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
    renderResult(issues, formatComment(issues));
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
    currentAbortController = null;
    currentTaskId = null;
    setBusy(false);
    refreshHistory();
    updateActionButtons();
  }
}

async function applyPendingResultToWord() {
  if (!pendingResult || pendingResult.issues.length === 0 || isApplyingToWord) {
    return;
  }

  isApplyingToWord = true;
  setBusy(true, { applying: true });
  updateActionButtons();
  showMessage("正在应用到 Word...", "default");

  try {
    const summary = await applyIssuesToScope(
      pendingResult.sourceText,
      pendingResult.issues,
      pendingResult.scope,
      getApplicationMode()
    );
    saveAppliedResultHistory({
      result: pendingResult,
      summary,
      applicationMode: getApplicationMode(),
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

function cancelCurrentProofread() {
  if (currentAbortController) {
    currentAbortController.abort();
  }

  if (currentTaskId) {
    cancelProofreadTask(currentTaskId).catch(() => {
      // The local abort is enough for UI state; task cancellation is best effort.
    });
  }
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
    errorMessage: input.errorMessage,
  });
}

function refreshHistory() {
  renderHistory(getHistoryEntries(), renderHistoryEntry);
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
  const canApply = Boolean(pendingResult && pendingResult.issues.length > 0 && !isApplyingToWord);

  getButton("apply-to-word").disabled = !canApply || taskState === "running";
}

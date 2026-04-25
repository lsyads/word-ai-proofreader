/* global AbortController, AbortSignal, document, Office, Word, HTMLElement, HTMLButtonElement, Response, TextDecoder, fetch, localStorage */

interface ProofreadIssue {
  id: string;
  category: string;
  severity: "low" | "medium" | "high";
  original: string;
  suggestion: string;
  comment: string;
  start?: number | null;
  end?: number | null;
}

interface ProofreadResponse {
  issues: ProofreadIssue[];
}

interface ProofreadStatusEvent {
  stage: string;
  message: string;
}

interface SessionResponse {
  session_id: string;
  created_at: string;
}

type TaskState = "idle" | "running" | "succeeded" | "failed" | "cancelled";

interface ProofreadHistoryEntry {
  id: string;
  sessionId: string;
  createdAt: string;
  textPreview: string;
  status: TaskState;
  issueCount: number;
  issues: ProofreadIssue[];
  insertedComment: boolean;
  errorMessage?: string;
}

const API_BASE_URL = "";
const HISTORY_STORAGE_KEY = "word-ai-proofreader-history-v1";
const MAX_HISTORY_ENTRIES = 20;

let currentSessionId: string | null = null;
let currentAbortController: AbortController | null = null;
let taskState: TaskState = "idle";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    getButton("proofread").onclick = proofreadSelection;
    getButton("new-conversation").onclick = newConversation;
    renderHistory();
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

  setBusy(false);
  resetProgress();
  renderEmptyResult("尚未开始审校");

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

async function createSession(): Promise<SessionResponse> {
  const response = await fetch(`${API_BASE_URL}/api/sessions`, {
    method: "POST",
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as SessionResponse;
}

export async function proofreadSelection() {
  if (taskState === "running") {
    cancelCurrentProofread();
    return;
  }

  if (!Office.context.requirements.isSetSupported("WordApi", "1.4")) {
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
  taskState = "running";
  setBusy(true);
  showMessage("正在审校当前选区...", "default");
  resetProgress();
  renderEmptyResult("审校中...");

  let selectedText = "";

  try {
    selectedText = await getSelectedText();
    const proofreadResult = await requestProofread(
      selectedText,
      (status) => {
        showMessage(status.message, "default");
        appendProgressStatus(status);
        renderEmptyResult(status.message);
      },
      abortController.signal
    );
    const commentText = formatComment(proofreadResult.issues);
    const insertedComment = proofreadResult.issues.length > 0;

    if (insertedComment) {
      await insertCommentToSelection(commentText);
    }

    renderResult(proofreadResult.issues, commentText);
    taskState = "succeeded";
    saveHistoryEntry({
      status: taskState,
      text: selectedText,
      issues: proofreadResult.issues,
      insertedComment,
    });

    showMessage(
      proofreadResult.issues.length > 0
        ? "审校完成，已在当前选区插入批注。"
        : "审校完成，未发现明显问题，本次未插入批注。",
      "success"
    );
  } catch (error) {
    if (isAbortError(error)) {
      taskState = "cancelled";
      appendProgressStatus({ stage: "cancelled", message: "已停止当前审校。" });
      renderEmptyResult("已停止审校");
      saveHistoryEntry({
        status: taskState,
        text: selectedText,
        issues: [],
        insertedComment: false,
        errorMessage: "用户停止了当前审校。",
      });
      showMessage("已停止当前审校。", "default");
      return;
    }

    taskState = "failed";
    appendProgressStatus({ stage: "failed", message: getErrorMessage(error) });
    renderEmptyResult("审校失败");
    saveHistoryEntry({
      status: taskState,
      text: selectedText,
      issues: [],
      insertedComment: false,
      errorMessage: getErrorMessage(error),
    });
    showMessage(`审校失败：${getErrorMessage(error)}`, "error");
  } finally {
    currentAbortController = null;
    setBusy(false);
  }
}

function cancelCurrentProofread() {
  if (currentAbortController) {
    currentAbortController.abort();
  }
}

async function getSelectedText(): Promise<string> {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const selectedText = selection.text || "";
    if (selectedText.trim().length === 0) {
      throw new Error("请先在 Word 中选中一段文字。");
    }

    return selectedText;
  });
}

async function insertCommentToSelection(commentText: string): Promise<void> {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertComment(commentText);
    await context.sync();
  });
}

async function requestProofread(
  text: string,
  onStatus: (status: ProofreadStatusEvent) => void,
  signal: AbortSignal
): Promise<ProofreadResponse> {
  try {
    onStatus({ stage: "api", message: "正在请求流式接口" });
    return await requestProofreadStream(text, onStatus, signal);
  } catch (error) {
    if (isAbortError(error)) {
      throw error;
    }

    onStatus({ stage: "fallback", message: "流式接口不可用，正在回退到普通接口" });
    return requestProofreadJson(text, signal);
  }
}

async function requestProofreadJson(text: string, signal: AbortSignal): Promise<ProofreadResponse> {
  const response = await fetch(`${API_BASE_URL}/api/proofread`, {
    method: "POST",
    signal,
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      text,
      session_id: currentSessionId,
      context: {
        source: "word-addin",
      },
    }),
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as ProofreadResponse;
}

async function requestProofreadStream(
  text: string,
  onStatus: (status: ProofreadStatusEvent) => void,
  signal: AbortSignal
): Promise<ProofreadResponse> {
  const response = await fetch(`${API_BASE_URL}/api/proofread/stream`, {
    method: "POST",
    signal,
    headers: {
      "Content-Type": "application/json",
      Accept: "text/event-stream",
    },
    body: JSON.stringify({
      text,
      session_id: currentSessionId,
      context: {
        source: "word-addin",
      },
    }),
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  if (!response.body) {
    throw new Error("当前 Word WebView 不支持流式读取。");
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";
  let result: ProofreadResponse | null = null;

  while (true) {
    const { done, value } = await reader.read();
    buffer += decoder.decode(value || new Uint8Array(), { stream: !done });

    const events = buffer.split("\n\n");
    buffer = events.pop() || "";

    events.forEach((eventText) => {
      const event = parseSseEvent(eventText);

      if (!event) {
        return;
      }

      if (event.name === "status") {
        onStatus(event.data as ProofreadStatusEvent);
      }

      if (event.name === "result") {
        result = event.data as ProofreadResponse;
        onStatus({ stage: "result", message: "已收到结构化审校结果。" });
      }

      if (event.name === "error") {
        throw new Error((event.data as { message?: string }).message || "AI 审校失败");
      }
    });

    if (done) {
      break;
    }
  }

  if (!result) {
    throw new Error("流式审校没有返回最终结果。");
  }

  return result;
}

function parseSseEvent(eventText: string): { name: string; data: unknown } | null {
  const lines = eventText.split("\n");
  const eventLine = lines.find((line) => line.startsWith("event: "));
  const dataLine = lines.find((line) => line.startsWith("data: "));

  if (!eventLine || !dataLine) {
    return null;
  }

  return {
    name: eventLine.replace("event: ", ""),
    data: JSON.parse(dataLine.replace("data: ", "")),
  };
}

function formatComment(issues: ProofreadIssue[]): string {
  if (issues.length === 0) {
    return "AI 审校：未发现明显问题，本次未插入批注。";
  }

  const lines = ["AI 审校建议："];

  issues.forEach((issue, index) => {
    lines.push("");
    lines.push(`${index + 1}. [${issue.severity}] ${issue.category}`);
    lines.push(`原文：${issue.original || "未提供"}`);
    lines.push(`建议：${issue.suggestion || "未提供"}`);
    lines.push(`说明：${issue.comment || "未提供"}`);
  });

  return lines.join("\n");
}

function renderResult(issues: ProofreadIssue[], commentText: string) {
  const result = getElement("result");

  if (issues.length === 0) {
    result.className = "result-empty";
    result.textContent = commentText;
    return;
  }

  result.className = "result-list";
  result.innerHTML = issues
    .map(
      (issue, index) => `
        <article class="result-item">
          <p class="result-item-title">${index + 1}. ${escapeHtml(issue.category)} / ${escapeHtml(issue.severity)}</p>
          <p><b>原文：</b>${escapeHtml(issue.original || "未提供")}</p>
          <p><b>建议：</b>${escapeHtml(issue.suggestion || "未提供")}</p>
          <p><b>说明：</b>${escapeHtml(issue.comment || "未提供")}</p>
        </article>
      `
    )
    .join("");
}

function renderEmptyResult(message: string) {
  const result = getElement("result");
  result.className = "result-empty";
  result.textContent = message;
}

function resetProgress() {
  const progress = getElement("progress");
  progress.className = "progress-log";
  progress.innerHTML = "";
}

function appendProgressStatus(status: ProofreadStatusEvent) {
  const progress = getElement("progress");
  progress.className = "progress-log";

  const existingItem = progress.querySelector(`[data-stage="${status.stage}"]`);

  if (existingItem) {
    const existingMessage = existingItem.querySelector(".progress-message");

    if (existingMessage) {
      existingMessage.textContent = status.message;
      return;
    }
  }

  const item = document.createElement("div");
  item.className = "progress-item";
  item.setAttribute("data-stage", status.stage);

  const stage = document.createElement("span");
  stage.className = "progress-stage";
  stage.textContent = formatStage(status.stage);

  const message = document.createElement("span");
  message.className = "progress-message";
  message.textContent = status.message;

  item.appendChild(stage);
  item.appendChild(message);
  progress.appendChild(item);
}

function formatStage(stage: string): string {
  const stageLabels: Record<string, string> = {
    api: "接口",
    calling_ai: "AI",
    cancelled: "停止",
    completed: "完成",
    failed: "失败",
    fallback: "回退",
    normalizing: "整理",
    received: "接收",
    result: "结果",
  };

  return `[${stageLabels[stage] || stage}]`;
}

function saveHistoryEntry(input: {
  status: TaskState;
  text: string;
  issues: ProofreadIssue[];
  insertedComment: boolean;
  errorMessage?: string;
}) {
  if (!currentSessionId || input.text.trim().length === 0) {
    return;
  }

  const entries = getHistoryEntries();
  const entry: ProofreadHistoryEntry = {
    id: createLocalId(),
    sessionId: currentSessionId,
    createdAt: new Date().toISOString(),
    textPreview: input.text.trim().slice(0, 40),
    status: input.status,
    issueCount: input.issues.length,
    issues: input.issues,
    insertedComment: input.insertedComment,
    errorMessage: input.errorMessage,
  };

  localStorage.setItem(
    HISTORY_STORAGE_KEY,
    JSON.stringify([entry, ...entries].slice(0, MAX_HISTORY_ENTRIES))
  );
  renderHistory();
}

function renderHistory() {
  const history = getElement("history");
  const entries = getHistoryEntries();

  if (entries.length === 0) {
    history.className = "history-empty";
    history.textContent = "暂无历史记录";
    return;
  }

  history.className = "history-list";
  history.innerHTML = entries
    .map(
      (entry) => `
        <button class="history-item" type="button" data-history-id="${escapeHtml(entry.id)}">
          <span class="history-title">${escapeHtml(entry.textPreview || "空文本")}</span>
          <span class="history-meta">${escapeHtml(formatHistoryMeta(entry))}</span>
        </button>
      `
    )
    .join("");

  history.querySelectorAll(".history-item").forEach((item) => {
    item.addEventListener("click", () => {
      const historyId = (item as HTMLElement).getAttribute("data-history-id");
      const entry = entries.find((candidate) => candidate.id === historyId);

      if (entry) {
        renderHistoryEntry(entry);
      }
    });
  });
}

function renderHistoryEntry(entry: ProofreadHistoryEntry) {
  resetProgress();
  appendProgressStatus({ stage: entry.status, message: formatHistoryMeta(entry) });
  renderResult(entry.issues, formatComment(entry.issues));
  showMessage(
    `已打开历史记录：${formatHistoryMeta(entry)}`,
    entry.status === "failed" ? "error" : "default"
  );
}

function getHistoryEntries(): ProofreadHistoryEntry[] {
  try {
    const parsed = JSON.parse(localStorage.getItem(HISTORY_STORAGE_KEY) || "[]");
    return Array.isArray(parsed) ? (parsed as ProofreadHistoryEntry[]) : [];
  } catch {
    return [];
  }
}

function formatHistoryMeta(entry: ProofreadHistoryEntry): string {
  const statusLabels: Record<TaskState, string> = {
    cancelled: "已停止",
    failed: "失败",
    idle: "未开始",
    running: "运行中",
    succeeded: "完成",
  };
  const createdAt = new Date(entry.createdAt).toLocaleString();
  const issueText = entry.issueCount > 0 ? `${entry.issueCount} 条问题` : "无问题";
  const commentText = entry.insertedComment ? "已插入批注" : "未插入批注";
  const errorText = entry.errorMessage ? `：${entry.errorMessage}` : "";

  return `${createdAt} / ${statusLabels[entry.status]} / ${issueText} / ${commentText}${errorText}`;
}

function createLocalId(): string {
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function setBusy(isBusy: boolean) {
  const button = getButton("proofread");
  button.disabled = false;
  button.querySelector(".ms-Button-label").textContent = isBusy ? "停止审校" : "AI 审校";
  getButton("new-conversation").disabled = isBusy;
}

function showMessage(message: string, type: "default" | "error" | "success" = "default") {
  const element = getElement("message");
  element.textContent = message;
  element.className = type === "default" ? "message" : `message is-${type}`;
}

function getElement(id: string): HTMLElement {
  return document.getElementById(id) as HTMLElement;
}

function getButton(id: string): HTMLButtonElement {
  return document.getElementById(id) as HTMLButtonElement;
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  return String(error);
}

function isAbortError(error: unknown): boolean {
  return (
    typeof error === "object" && error !== null && "name" in error && error.name === "AbortError"
  );
}

async function getResponseErrorMessage(response: Response): Promise<string> {
  try {
    const payload = (await response.json()) as { detail?: string };

    if (payload.detail) {
      return payload.detail;
    }
  } catch {
    // Fall back to the HTTP status below when the response is not JSON.
  }

  return `后端返回 HTTP ${response.status}`;
}

function escapeHtml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

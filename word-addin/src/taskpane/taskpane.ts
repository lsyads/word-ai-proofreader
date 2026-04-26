/* global AbortController, AbortSignal, Blob, URL, clearTimeout, document, Office, Word, HTMLElement, HTMLButtonElement, HTMLInputElement, HTMLSelectElement, HTMLTextAreaElement, Response, TextDecoder, fetch, localStorage, setTimeout */

interface ProofreadIssue {
  id: string;
  category: string;
  severity: "low" | "medium" | "high";
  original: string;
  replacement?: string | null;
  suggestion: string;
  start?: number | null;
  end?: number | null;
}

interface ProofreadResponse {
  issues: ProofreadIssue[];
}

interface ChunkedProofreadIssue extends ProofreadIssue {
  chunk_index: number;
  global_start?: number | null;
  global_end?: number | null;
}

interface ChunkedProofreadResponse {
  task_id?: string | null;
  scope: ProofreadScope;
  status: TaskState;
  total_chunks: number;
  completed_chunks: number;
  failed_chunks: number;
  issues: ChunkedProofreadIssue[];
  error_message?: string | null;
}

interface ProofreadStatusEvent {
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

interface SessionResponse {
  session_id: string;
  created_at: string;
}

interface BookInfo {
  title: string;
  introduction?: string | null;
}

type TaskState = "idle" | "queued" | "running" | "succeeded" | "failed" | "cancelled";
type ProviderAPI = "responses" | "chat";
type ProofreadMode = "fast" | "thinking";
type ApplicationMode = "comment" | "revision";
type ProofreadScope = "selection" | "document";

interface ProofreadHistoryEntry {
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
  errorMessage?: string;
}

interface IssueApplicationSummary {
  commentCount: number;
  revisionCount: number;
  fallbackCount: number;
}

const API_BASE_URL = "";
const HISTORY_STORAGE_KEY = "word-ai-proofreader-history-v1";
const PROVIDER_API_STORAGE_KEY = "word-ai-proofreader-provider-api-v1";
const PROOFREAD_MODE_STORAGE_KEY = "word-ai-proofreader-mode-v1";
const APPLICATION_MODE_STORAGE_KEY = "word-ai-proofreader-application-mode-v1";
const PROOFREAD_SCOPE_STORAGE_KEY = "word-ai-proofreader-scope-v1";
const BOOK_TITLE_STORAGE_KEY = "word-ai-proofreader-book-title-v1";
const BOOK_INTRODUCTION_STORAGE_KEY = "word-ai-proofreader-book-introduction-v1";
const MAX_HISTORY_ENTRIES = 20;
const CLEAR_HISTORY_CONFIRM_MS = 4000;
const SELECTION_CHUNK_THRESHOLD = 5000;
const DEFAULT_CHUNK_SIZE = 3000;

let currentSessionId: string | null = null;
let currentAbortController: AbortController | null = null;
let currentTaskId: string | null = null;
let clearHistoryConfirmTimer: number | null = null;
let isClearHistoryArmed = false;
let taskState: TaskState = "idle";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    getButton("proofread").onclick = proofreadSelection;
    getButton("new-conversation").onclick = newConversation;
    getButton("clear-history").onclick = clearHistory;
    getButton("export-history").onclick = exportHistory;
    getButton("import-history").onclick = () => getInput("history-file").click();
    getInput("history-file").onchange = importHistory;
    getInput("book-title").oninput = persistBookInfo;
    getTextArea("book-introduction").oninput = persistBookInfo;
    getSelect("provider-api").onchange = persistControls;
    getSelect("proofread-mode").onchange = persistControls;
    getSelect("application-mode").onchange = persistControls;
    getSelect("proofread-scope").onchange = persistControls;
    initializeControls();
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

  const book = getValidatedBookInfo();
  if (!book) {
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
  const scope = getProofreadScope();
  showMessage(scope === "document" ? "正在审校全书正文..." : "正在审校当前选区...", "default");
  resetProgress();
  renderEmptyResult("审校中...");

  let sourceText = "";

  try {
    sourceText = scope === "document" ? await getDocumentBodyText() : await getSelectedText();
    const useChunkedFlow = scope === "document" || sourceText.length > SELECTION_CHUNK_THRESHOLD;
    let proofreadResult: ProofreadResponse;
    let taskId: string | null = null;
    let totalChunks = 1;
    let completedChunks = 1;
    let failedChunks = 0;

    if (useChunkedFlow) {
      const chunkedResult = await requestChunkedProofreadTask(
        sourceText,
        book,
        scope,
        (status) => {
          showMessage(status.message, "default");
          appendProgressStatus(status);
          renderEmptyResult(formatProgressResult(status));
        },
        abortController.signal
      );
      taskId = chunkedResult.task_id || null;
      totalChunks = chunkedResult.total_chunks;
      completedChunks = chunkedResult.completed_chunks;
      failedChunks = chunkedResult.failed_chunks;
      proofreadResult = {
        issues: normalizeChunkedIssuesForScope(chunkedResult.issues),
      };
    } else {
      proofreadResult = await requestProofread(
        sourceText,
        book,
        (status) => {
          showMessage(status.message, "default");
          appendProgressStatus(status);
          renderEmptyResult(status.message);
        },
        abortController.signal
      );
    }

    const commentText = formatComment(proofreadResult.issues);
    let applicationSummary: IssueApplicationSummary = {
      commentCount: 0,
      revisionCount: 0,
      fallbackCount: 0,
    };

    if (proofreadResult.issues.length > 0) {
      applicationSummary = await applyIssuesToScope(sourceText, proofreadResult.issues, scope);
    }

    renderResult(proofreadResult.issues, commentText);
    taskState = "succeeded";
    saveHistoryEntry({
      status: taskState,
      text: sourceText,
      book,
      issues: proofreadResult.issues,
      insertedComment: applicationSummary.commentCount + applicationSummary.fallbackCount > 0,
      locatedIssueCount: applicationSummary.commentCount + applicationSummary.revisionCount,
      revisionCount: applicationSummary.revisionCount,
      fallbackCount: applicationSummary.fallbackCount,
      scope,
      taskId,
      totalChunks,
      completedChunks,
      failedChunks,
    });

    showMessage(
      proofreadResult.issues.length > 0
        ? formatCompletionMessage(applicationSummary)
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
        text: sourceText,
        book,
        issues: [],
        insertedComment: false,
        locatedIssueCount: 0,
        revisionCount: 0,
        fallbackCount: 0,
        scope,
        taskId: currentTaskId,
        totalChunks: 0,
        completedChunks: 0,
        failedChunks: 0,
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
      text: sourceText,
      book,
      issues: [],
      insertedComment: false,
      locatedIssueCount: 0,
      revisionCount: 0,
      fallbackCount: 0,
      scope,
      taskId: currentTaskId,
      totalChunks: 0,
      completedChunks: 0,
      failedChunks: 0,
      errorMessage: getErrorMessage(error),
    });
    showMessage(`审校失败：${getErrorMessage(error)}`, "error");
  } finally {
    currentAbortController = null;
    currentTaskId = null;
    setBusy(false);
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

async function getSelectedText(): Promise<string> {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const selectedText = (selection.text || "").trim();
    if (selectedText.length === 0) {
      throw new Error("请先在 Word 中选中一段文字。");
    }

    return selectedText;
  });
}

async function getDocumentBodyText(): Promise<string> {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const bodyText = (body.text || "").trim();
    if (bodyText.length === 0) {
      throw new Error("当前 Word 正文为空，无法进行全书审校。");
    }

    return bodyText;
  });
}

async function applyIssuesToScope(
  sourceText: string,
  issues: ProofreadIssue[],
  scope: ProofreadScope
): Promise<IssueApplicationSummary> {
  if (scope === "document") {
    return applyIssuesToDocument(sourceText, issues);
  }

  return applyIssuesToSelection(sourceText, issues);
}

async function applyIssuesToSelection(
  selectedText: string,
  issues: ProofreadIssue[]
): Promise<IssueApplicationSummary> {
  if (getApplicationMode() === "revision") {
    return applyRevisionsForIssues(selectedText, issues);
  }

  return insertCommentsForIssues(selectedText, issues);
}

async function applyIssuesToDocument(
  documentText: string,
  issues: ProofreadIssue[]
): Promise<IssueApplicationSummary> {
  if (getApplicationMode() === "revision") {
    return applyDocumentRevisionsForIssues(documentText, issues);
  }

  return insertDocumentCommentsForIssues(documentText, issues);
}

async function insertCommentsForIssues(
  selectedText: string,
  issues: ProofreadIssue[]
): Promise<IssueApplicationSummary> {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    const fallbackIssues: ProofreadIssue[] = [];
    const pendingSearches: Array<{
      issue: ProofreadIssue;
      occurrenceIndex: number;
      searchResults: Word.RangeCollection;
    }> = [];
    let commentCount = 0;

    for (const issue of issues) {
      if (!isIssueLocatable(selectedText, issue)) {
        fallbackIssues.push(issue);
        continue;
      }

      const occurrenceIndex = getOccurrenceIndexBeforeOffset(
        selectedText,
        issue.original,
        issue.start as number
      );
      const searchResults = selection.search(issue.original, {
        matchCase: true,
        matchWholeWord: false,
      });
      searchResults.load("items");
      pendingSearches.push({ issue, occurrenceIndex, searchResults });
    }

    await context.sync();

    pendingSearches.forEach(({ issue, occurrenceIndex, searchResults }) => {
      const targetRange = searchResults.items[occurrenceIndex];

      if (!targetRange) {
        fallbackIssues.push(issue);
        return;
      }

      targetRange.insertComment(formatIssueComment(issue));
      commentCount += 1;
    });

    if (fallbackIssues.length > 0) {
      selection.insertComment(formatFallbackComment(fallbackIssues));
    }

    await context.sync();
    return { commentCount, revisionCount: 0, fallbackCount: fallbackIssues.length };
  });
}

async function applyRevisionsForIssues(
  selectedText: string,
  issues: ProofreadIssue[]
): Promise<IssueApplicationSummary> {
  return Word.run(async (context) => {
    const document = context.document;
    const selection = document.getSelection();
    const fallbackIssues: ProofreadIssue[] = [];
    const pendingSearches: Array<{
      issue: ProofreadIssue;
      occurrenceIndex: number;
      searchResults: Word.RangeCollection;
    }> = [];
    let revisionCount = 0;

    document.load("changeTrackingMode");

    for (const issue of issues) {
      if (!isIssueLocatable(selectedText, issue) || !hasReplacement(issue)) {
        fallbackIssues.push(issue);
        continue;
      }

      const occurrenceIndex = getOccurrenceIndexBeforeOffset(
        selectedText,
        issue.original,
        issue.start as number
      );
      const searchResults = selection.search(issue.original, {
        matchCase: true,
        matchWholeWord: false,
      });
      searchResults.load("items");
      pendingSearches.push({ issue, occurrenceIndex, searchResults });
    }

    await context.sync();

    const originalTrackingMode = document.changeTrackingMode;

    try {
      if (pendingSearches.length > 0) {
        document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;

        pendingSearches.forEach(({ issue, occurrenceIndex, searchResults }) => {
          const targetRange = searchResults.items[occurrenceIndex];

          if (!targetRange || !issue.replacement) {
            fallbackIssues.push(issue);
            return;
          }

          targetRange.insertText(issue.replacement, Word.InsertLocation.replace);
          revisionCount += 1;
        });

        await context.sync();
      }
    } finally {
      document.changeTrackingMode = originalTrackingMode;
      await context.sync();
    }

    if (fallbackIssues.length > 0) {
      selection.insertComment(formatFallbackComment(fallbackIssues));
      await context.sync();
    }

    return { commentCount: 0, revisionCount, fallbackCount: fallbackIssues.length };
  });
}

async function insertDocumentCommentsForIssues(
  documentText: string,
  issues: ProofreadIssue[]
): Promise<IssueApplicationSummary> {
  return Word.run(async (context) => {
    const body = context.document.body;
    const fallbackIssues: ProofreadIssue[] = [];
    const pendingSearches: Array<{
      issue: ProofreadIssue;
      occurrenceIndex: number;
      searchResults: Word.RangeCollection;
    }> = [];
    let commentCount = 0;

    for (const issue of issues) {
      if (!isIssueLocatable(documentText, issue)) {
        fallbackIssues.push(issue);
        continue;
      }

      const occurrenceIndex = getOccurrenceIndexBeforeOffset(
        documentText,
        issue.original,
        issue.start as number
      );
      const searchResults = body.search(issue.original, {
        matchCase: true,
        matchWholeWord: false,
      });
      searchResults.load("items");
      pendingSearches.push({ issue, occurrenceIndex, searchResults });
    }

    await context.sync();

    pendingSearches.forEach(({ issue, occurrenceIndex, searchResults }) => {
      const targetRange = searchResults.items[occurrenceIndex];

      if (!targetRange) {
        fallbackIssues.push(issue);
        return;
      }

      targetRange.insertComment(formatIssueComment(issue));
      commentCount += 1;
    });

    if (fallbackIssues.length > 0) {
      body.getRange().insertComment(formatFallbackComment(fallbackIssues));
    }

    await context.sync();
    return { commentCount, revisionCount: 0, fallbackCount: fallbackIssues.length };
  });
}

async function applyDocumentRevisionsForIssues(
  documentText: string,
  issues: ProofreadIssue[]
): Promise<IssueApplicationSummary> {
  return Word.run(async (context) => {
    const document = context.document;
    const body = document.body;
    const fallbackIssues: ProofreadIssue[] = [];
    const pendingSearches: Array<{
      issue: ProofreadIssue;
      occurrenceIndex: number;
      searchResults: Word.RangeCollection;
    }> = [];
    let revisionCount = 0;

    document.load("changeTrackingMode");

    for (const issue of issues) {
      if (!isIssueLocatable(documentText, issue) || !hasReplacement(issue)) {
        fallbackIssues.push(issue);
        continue;
      }

      const occurrenceIndex = getOccurrenceIndexBeforeOffset(
        documentText,
        issue.original,
        issue.start as number
      );
      const searchResults = body.search(issue.original, {
        matchCase: true,
        matchWholeWord: false,
      });
      searchResults.load("items");
      pendingSearches.push({ issue, occurrenceIndex, searchResults });
    }

    await context.sync();

    const originalTrackingMode = document.changeTrackingMode;

    try {
      if (pendingSearches.length > 0) {
        document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;

        pendingSearches
          .sort((left, right) => (right.issue.start || 0) - (left.issue.start || 0))
          .forEach(({ issue, occurrenceIndex, searchResults }) => {
            const targetRange = searchResults.items[occurrenceIndex];

            if (!targetRange || !issue.replacement) {
              fallbackIssues.push(issue);
              return;
            }

            targetRange.insertText(issue.replacement, Word.InsertLocation.replace);
            revisionCount += 1;
          });

        await context.sync();
      }
    } finally {
      document.changeTrackingMode = originalTrackingMode;
      await context.sync();
    }

    if (fallbackIssues.length > 0) {
      body.getRange().insertComment(formatFallbackComment(fallbackIssues));
      await context.sync();
    }

    return { commentCount: 0, revisionCount, fallbackCount: fallbackIssues.length };
  });
}

async function requestProofread(
  text: string,
  book: BookInfo,
  onStatus: (status: ProofreadStatusEvent) => void,
  signal: AbortSignal
): Promise<ProofreadResponse> {
  if (getProviderApi() === "chat") {
    onStatus({ stage: "api", message: "正在请求 Chat Completions 接口" });
    return requestProofreadJson(text, book, signal);
  }

  try {
    onStatus({ stage: "api", message: "正在请求流式接口" });
    return await requestProofreadStream(text, book, onStatus, signal);
  } catch (error) {
    if (isAbortError(error)) {
      throw error;
    }

    onStatus({ stage: "fallback", message: "流式接口不可用，正在回退到普通接口" });
    return requestProofreadJson(text, book, signal);
  }
}

async function requestProofreadJson(
  text: string,
  book: BookInfo,
  signal: AbortSignal
): Promise<ProofreadResponse> {
  const response = await fetch(`${API_BASE_URL}/api/proofread`, {
    method: "POST",
    signal,
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      text,
      book,
      session_id: currentSessionId,
      provider_api: getProviderApi(),
      proofread_mode: getProofreadMode(),
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
  book: BookInfo,
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
      book,
      session_id: currentSessionId,
      provider_api: getProviderApi(),
      proofread_mode: getProofreadMode(),
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

async function requestChunkedProofreadTask(
  text: string,
  book: BookInfo,
  scope: ProofreadScope,
  onStatus: (status: ProofreadStatusEvent) => void,
  signal: AbortSignal
): Promise<ChunkedProofreadResponse> {
  onStatus({
    stage: "task",
    message:
      scope === "document"
        ? "正在创建全书分块审校任务。"
        : "当前选区超过 5000 字，正在创建分块审校任务。",
  });

  const createdTask = await createProofreadTask(text, book, scope, signal);
  currentTaskId = createdTask.task_id || null;
  onStatus(taskSnapshotToStatus("queued", createdTask, "审校任务已创建。"));

  try {
    await streamProofreadTaskEvents(createdTask.task_id as string, onStatus, signal);
  } catch (error) {
    if (isAbortError(error)) {
      throw error;
    }

    onStatus({ stage: "polling", message: "任务进度流不可用，正在切换到轮询查询。" });
    return pollProofreadTask(createdTask.task_id as string, onStatus, signal);
  }

  const finalTask = await getProofreadTask(createdTask.task_id as string, signal);
  if (finalTask.status === "failed") {
    throw new Error(finalTask.error_message || "分块审校任务失败。");
  }

  if (finalTask.status === "cancelled") {
    throw createAbortError();
  }

  return finalTask;
}

async function createProofreadTask(
  text: string,
  book: BookInfo,
  scope: ProofreadScope,
  signal: AbortSignal
): Promise<ChunkedProofreadResponse> {
  const response = await fetch(`${API_BASE_URL}/api/proofread/tasks`, {
    method: "POST",
    signal,
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      text,
      book,
      session_id: currentSessionId,
      provider_api: getProviderApi(),
      proofread_mode: getProofreadMode(),
      scope,
      chunk_size: DEFAULT_CHUNK_SIZE,
      context: {
        source: "word-addin",
        flow: "chunked-task",
      },
    }),
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as ChunkedProofreadResponse;
}

async function streamProofreadTaskEvents(
  taskId: string,
  onStatus: (status: ProofreadStatusEvent) => void,
  signal: AbortSignal
): Promise<void> {
  const response = await fetch(`${API_BASE_URL}/api/proofread/tasks/${taskId}/events`, {
    method: "GET",
    signal,
    headers: {
      Accept: "text/event-stream",
    },
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  if (!response.body) {
    throw new Error("当前 Word WebView 不支持任务进度流。");
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";

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

      const status = taskEventToStatus(event.name, event.data);
      onStatus(status);

      if (event.name === "error") {
        throw new Error(status.error_message || status.message || "分块审校任务失败。");
      }
    });

    if (done) {
      break;
    }
  }
}

async function pollProofreadTask(
  taskId: string,
  onStatus: (status: ProofreadStatusEvent) => void,
  signal: AbortSignal
): Promise<ChunkedProofreadResponse> {
  while (true) {
    const task = await getProofreadTask(taskId, signal);
    onStatus(taskSnapshotToStatus("polling", task, "正在查询分块审校进度。"));

    if (task.status === "succeeded") {
      return task;
    }

    if (task.status === "failed") {
      throw new Error(task.error_message || "分块审校任务失败。");
    }

    if (task.status === "cancelled") {
      throw createAbortError();
    }

    await delay(1000, signal);
  }
}

async function getProofreadTask(
  taskId: string,
  signal: AbortSignal
): Promise<ChunkedProofreadResponse> {
  const response = await fetch(`${API_BASE_URL}/api/proofread/tasks/${taskId}`, {
    method: "GET",
    signal,
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as ChunkedProofreadResponse;
}

async function cancelProofreadTask(taskId: string): Promise<void> {
  await fetch(`${API_BASE_URL}/api/proofread/tasks/${taskId}`, {
    method: "DELETE",
  });
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

function taskEventToStatus(stage: string, data: unknown): ProofreadStatusEvent {
  const payload = isRecord(data) ? data : {};
  const message = typeof payload.message === "string" ? payload.message : formatStage(stage);

  return {
    stage,
    message,
    task_id: typeof payload.task_id === "string" ? payload.task_id : undefined,
    status: isTaskState(payload.status) ? payload.status : undefined,
    scope: isProofreadScope(payload.scope) ? payload.scope : undefined,
    total_chunks: typeof payload.total_chunks === "number" ? payload.total_chunks : undefined,
    completed_chunks:
      typeof payload.completed_chunks === "number" ? payload.completed_chunks : undefined,
    failed_chunks: typeof payload.failed_chunks === "number" ? payload.failed_chunks : undefined,
    issue_count: typeof payload.issue_count === "number" ? payload.issue_count : undefined,
    chunk_index: typeof payload.chunk_index === "number" ? payload.chunk_index : undefined,
    chunk_start: typeof payload.chunk_start === "number" ? payload.chunk_start : undefined,
    chunk_end: typeof payload.chunk_end === "number" ? payload.chunk_end : undefined,
    error_message: typeof payload.error_message === "string" ? payload.error_message : undefined,
  };
}

function taskSnapshotToStatus(
  stage: string,
  task: ChunkedProofreadResponse,
  message: string
): ProofreadStatusEvent {
  return {
    stage,
    message,
    task_id: task.task_id || undefined,
    status: task.status,
    scope: task.scope,
    total_chunks: task.total_chunks,
    completed_chunks: task.completed_chunks,
    failed_chunks: task.failed_chunks,
    issue_count: task.issues.length,
    error_message: task.error_message || undefined,
  };
}

function formatProgressResult(status: ProofreadStatusEvent): string {
  if (typeof status.total_chunks !== "number") {
    return status.message;
  }

  const completed = status.completed_chunks || 0;
  const failed = status.failed_chunks || 0;
  const issueCount = status.issue_count || 0;
  return `${status.message} 进度 ${completed}/${status.total_chunks}，失败 ${failed} 块，累计问题 ${issueCount} 条。`;
}

function normalizeChunkedIssuesForScope(issues: ChunkedProofreadIssue[]): ProofreadIssue[] {
  return issues.map((issue) => ({
    ...issue,
    id: `chunk-${issue.chunk_index}-${issue.id}`,
    start: issue.global_start,
    end: issue.global_end,
  }));
}

function isIssueLocatable(selectedText: string, issue: ProofreadIssue): boolean {
  if (!issue.original || issue.start === null || issue.end === null) {
    return false;
  }

  if (typeof issue.start !== "number" || typeof issue.end !== "number") {
    return false;
  }

  return selectedText.slice(issue.start, issue.end) === issue.original;
}

function hasReplacement(issue: ProofreadIssue): boolean {
  return typeof issue.replacement === "string" && issue.replacement.trim().length > 0;
}

function getOccurrenceIndexBeforeOffset(
  text: string,
  original: string,
  targetStart: number
): number {
  let count = 0;
  let searchFrom = 0;

  while (searchFrom < targetStart) {
    const foundAt = text.indexOf(original, searchFrom);

    if (foundAt === -1 || foundAt >= targetStart) {
      break;
    }

    count += 1;
    searchFrom = foundAt + original.length;
  }

  return count;
}

function formatIssueComment(issue: ProofreadIssue): string {
  return [
    issue.replacement ? `替换为：${issue.replacement}` : "",
    `建议：${issue.suggestion || "未提供"}`,
    `类别：${issue.category} / ${issue.severity}`,
  ]
    .filter(Boolean)
    .join("\n");
}

function formatFallbackComment(issues: ProofreadIssue[]): string {
  return `AI 审校：以下 ${issues.length} 条建议未能定位到具体原文片段，已汇总到当前选区。\n\n${formatComment(issues)}`;
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
    if (issue.replacement) {
      lines.push(`替换为：${issue.replacement}`);
    }
    lines.push(`建议：${issue.suggestion || "未提供"}`);
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
          <p class="result-item-title">${index + 1}. ${escapeHtml(issue.category)} / ${escapeHtml(issue.severity)}
            <span class="location-status">${escapeHtml(formatLocationStatus(issue))}</span>
          </p>
          <p><b>原文：</b>${escapeHtml(issue.original || "未提供")}</p>
          <p><b>替换为：</b>${escapeHtml(issue.replacement || "无直接替换文本")}</p>
          <p><b>建议：</b>${escapeHtml(issue.suggestion || "未提供")}</p>
        </article>
      `
    )
    .join("");
}

function formatCompletionMessage(summary: IssueApplicationSummary): string {
  if (getApplicationMode() === "revision") {
    return `审校完成，已生成修订 ${summary.revisionCount} 条，回退批注 ${summary.fallbackCount} 条。`;
  }

  return `审校完成，已精准批注 ${summary.commentCount} 条，回退批注 ${summary.fallbackCount} 条。`;
}

function formatLocationStatus(issue: ProofreadIssue): string {
  if (typeof issue.start === "number" && typeof issue.end === "number") {
    return `已定位 ${issue.start}-${issue.end}`;
  }

  return "未定位";
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
    chunk_completed: "分块",
    chunk_failed: "分块",
    chunk_started: "分块",
    normalizing: "整理",
    polling: "轮询",
    queued: "排队",
    received: "接收",
    result: "结果",
    running: "运行",
    task: "任务",
  };

  return `[${stageLabels[stage] || stage}]`;
}

function saveHistoryEntry(input: {
  status: TaskState;
  text: string;
  book: BookInfo;
  issues: ProofreadIssue[];
  insertedComment: boolean;
  locatedIssueCount: number;
  revisionCount: number;
  fallbackCount: number;
  scope: ProofreadScope;
  taskId?: string | null;
  totalChunks: number;
  completedChunks: number;
  failedChunks: number;
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
    bookTitle: input.book.title,
    bookIntroductionPreview: (input.book.introduction || "").trim().slice(0, 40),
    status: input.status,
    issueCount: input.issues.length,
    locatedIssueCount: input.locatedIssueCount,
    revisionCount: input.revisionCount,
    fallbackCount: input.fallbackCount,
    providerApi: getProviderApi(),
    proofreadMode: getProofreadMode(),
    applicationMode: getApplicationMode(),
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
    return Array.isArray(parsed) ? normalizeHistoryEntries(parsed) : [];
  } catch {
    return [];
  }
}

function normalizeHistoryEntries(entries: unknown[]): ProofreadHistoryEntry[] {
  return entries
    .filter(
      (entry): entry is Partial<ProofreadHistoryEntry> =>
        typeof entry === "object" && entry !== null
    )
    .map((entry) => {
      const issues = Array.isArray(entry.issues) ? (entry.issues as ProofreadIssue[]) : [];
      return {
        id: typeof entry.id === "string" ? entry.id : createLocalId(),
        sessionId: typeof entry.sessionId === "string" ? entry.sessionId : "",
        createdAt: typeof entry.createdAt === "string" ? entry.createdAt : new Date().toISOString(),
        textPreview: typeof entry.textPreview === "string" ? entry.textPreview : "",
        bookTitle: typeof entry.bookTitle === "string" ? entry.bookTitle : "",
        bookIntroductionPreview:
          typeof entry.bookIntroductionPreview === "string" ? entry.bookIntroductionPreview : "",
        status: isTaskState(entry.status) ? entry.status : "succeeded",
        issueCount: typeof entry.issueCount === "number" ? entry.issueCount : issues.length,
        locatedIssueCount:
          typeof entry.locatedIssueCount === "number"
            ? entry.locatedIssueCount
            : issues.filter(
                (issue) => typeof issue.start === "number" && typeof issue.end === "number"
              ).length,
        revisionCount: typeof entry.revisionCount === "number" ? entry.revisionCount : 0,
        fallbackCount: typeof entry.fallbackCount === "number" ? entry.fallbackCount : 0,
        providerApi: isProviderApi(entry.providerApi) ? entry.providerApi : "responses",
        proofreadMode: isProofreadMode(entry.proofreadMode) ? entry.proofreadMode : "fast",
        applicationMode: isApplicationMode(entry.applicationMode)
          ? entry.applicationMode
          : "comment",
        scope: isProofreadScope(entry.scope) ? entry.scope : "selection",
        taskId: typeof entry.taskId === "string" ? entry.taskId : null,
        totalChunks: typeof entry.totalChunks === "number" ? entry.totalChunks : 1,
        completedChunks: typeof entry.completedChunks === "number" ? entry.completedChunks : 1,
        failedChunks: typeof entry.failedChunks === "number" ? entry.failedChunks : 0,
        globalLocatedIssueCount:
          typeof entry.globalLocatedIssueCount === "number"
            ? entry.globalLocatedIssueCount
            : issues.filter(
                (issue) => typeof issue.start === "number" && typeof issue.end === "number"
              ).length,
        issues,
        insertedComment: Boolean(entry.insertedComment),
        errorMessage: typeof entry.errorMessage === "string" ? entry.errorMessage : undefined,
      };
    })
    .slice(0, MAX_HISTORY_ENTRIES);
}

function formatHistoryMeta(entry: ProofreadHistoryEntry): string {
  const statusLabels: Record<TaskState, string> = {
    cancelled: "已停止",
    failed: "失败",
    idle: "未开始",
    queued: "排队中",
    running: "运行中",
    succeeded: "完成",
  };
  const createdAt = new Date(entry.createdAt).toLocaleString();
  const issueText = entry.issueCount > 0 ? `${entry.issueCount} 条问题` : "无问题";
  const locatedText =
    entry.issueCount > 0 ? `定位 ${entry.locatedIssueCount}/${entry.issueCount}` : "无需定位";
  const actionText =
    entry.applicationMode === "revision"
      ? `修订 ${entry.revisionCount} / 回退 ${entry.fallbackCount}`
      : `批注 / 回退 ${entry.fallbackCount}`;
  const modeText = `${formatProofreadScope(entry.scope)} / ${formatProofreadMode(entry.proofreadMode)} / ${formatProviderApi(entry.providerApi)} / ${formatApplicationMode(entry.applicationMode)}`;
  const chunkText =
    entry.totalChunks > 1
      ? `分块 ${entry.completedChunks}/${entry.totalChunks}，失败 ${entry.failedChunks}`
      : "单段";
  const bookText = entry.bookTitle ? `《${entry.bookTitle}》` : "未记录书名";
  const commentText = entry.insertedComment ? "已插入批注" : "未插入批注";
  const errorText = entry.errorMessage ? `：${entry.errorMessage}` : "";

  return `${createdAt} / ${bookText} / ${statusLabels[entry.status]} / ${modeText} / ${chunkText} / ${issueText} / ${locatedText} / ${actionText} / ${commentText}${errorText}`;
}

function initializeControls() {
  getInput("book-title").value = localStorage.getItem(BOOK_TITLE_STORAGE_KEY) || "";
  getTextArea("book-introduction").value =
    localStorage.getItem(BOOK_INTRODUCTION_STORAGE_KEY) || "";
  getSelect("provider-api").value = readStoredProviderApi();
  getSelect("proofread-mode").value = readStoredProofreadMode();
  getSelect("application-mode").value = readStoredApplicationMode();
  getSelect("proofread-scope").value = readStoredProofreadScope();
}

function persistControls() {
  localStorage.setItem(PROVIDER_API_STORAGE_KEY, getProviderApi());
  localStorage.setItem(PROOFREAD_MODE_STORAGE_KEY, getProofreadMode());
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

function getProviderApi(): ProviderAPI {
  const value = getSelect("provider-api").value;
  return isProviderApi(value) ? value : "responses";
}

function getProofreadMode(): ProofreadMode {
  const value = getSelect("proofread-mode").value;
  return isProofreadMode(value) ? value : "fast";
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

function isTaskState(value: unknown): value is TaskState {
  return (
    value === "idle" ||
    value === "queued" ||
    value === "running" ||
    value === "succeeded" ||
    value === "failed" ||
    value === "cancelled"
  );
}

function formatProviderApi(providerApi: ProviderAPI): string {
  return providerApi === "chat" ? "Chat" : "Responses";
}

function formatProofreadMode(proofreadMode: ProofreadMode): string {
  return proofreadMode === "thinking" ? "深度审校" : "快速审校";
}

function formatApplicationMode(applicationMode: ApplicationMode): string {
  return applicationMode === "revision" ? "修订模式" : "批注模式";
}

function formatProofreadScope(scope: ProofreadScope): string {
  return scope === "document" ? "全书正文" : "当前选区";
}

function clearHistory() {
  if (!isClearHistoryArmed) {
    armClearHistoryConfirmation();
    return;
  }

  resetClearHistoryConfirmation();
  localStorage.removeItem(HISTORY_STORAGE_KEY);
  renderHistory();
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
  const entries = getHistoryEntries();
  const blob = new Blob([JSON.stringify(entries, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = `word-ai-proofreader-history-${new Date().toISOString().slice(0, 10)}.json`;
  link.click();
  URL.revokeObjectURL(url);
  showMessage(`已导出 ${entries.length} 条历史记录。`, "success");
}

async function importHistory() {
  const input = getInput("history-file");
  const file = input.files && input.files[0];

  if (!file) {
    return;
  }

  try {
    const parsed = JSON.parse(await file.text());

    if (!Array.isArray(parsed)) {
      throw new Error("历史文件必须是数组格式。");
    }

    const imported = normalizeHistoryEntries(parsed);
    localStorage.setItem(
      HISTORY_STORAGE_KEY,
      JSON.stringify(imported.slice(0, MAX_HISTORY_ENTRIES))
    );
    renderHistory();
    showMessage(`已导入 ${imported.length} 条历史记录。`, "success");
  } catch (error) {
    showMessage(`导入历史失败：${getErrorMessage(error)}`, "error");
  } finally {
    input.value = "";
  }
}

function createLocalId(): string {
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function setBusy(isBusy: boolean) {
  const button = getButton("proofread");
  if (isBusy) {
    resetClearHistoryConfirmation();
  }

  button.disabled = false;
  button.querySelector(".ms-Button-label").textContent = isBusy ? "停止审校" : "AI 审校";
  getButton("new-conversation").disabled = isBusy;
  getButton("clear-history").disabled = isBusy;
  getButton("export-history").disabled = isBusy;
  getButton("import-history").disabled = isBusy;
  getInput("book-title").disabled = isBusy;
  getTextArea("book-introduction").disabled = isBusy;
  getSelect("provider-api").disabled = isBusy;
  getSelect("proofread-mode").disabled = isBusy;
  getSelect("application-mode").disabled = isBusy;
  getSelect("proofread-scope").disabled = isBusy;
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

function getSelect(id: string): HTMLSelectElement {
  return document.getElementById(id) as HTMLSelectElement;
}

function getInput(id: string): HTMLInputElement {
  return document.getElementById(id) as HTMLInputElement;
}

function getTextArea(id: string): HTMLTextAreaElement {
  return document.getElementById(id) as HTMLTextAreaElement;
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

function createAbortError(): Error {
  const error = new Error("Aborted");
  error.name = "AbortError";
  return error;
}

function delay(milliseconds: number, signal: AbortSignal): Promise<void> {
  if (signal.aborted) {
    return Promise.reject(createAbortError());
  }

  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      resolve();
    }, milliseconds);

    signal.addEventListener(
      "abort",
      () => {
        clearTimeout(timer);
        reject(createAbortError());
      },
      { once: true }
    );
  });
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
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

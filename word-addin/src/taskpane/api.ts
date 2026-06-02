/* global AbortSignal, File, Response, TextDecoder, URLSearchParams, clearTimeout, fetch, setTimeout */

import {
  ApplicationMode,
  AgentRunTrace,
  AIProfile,
  BookInfo,
  ChunkedProofreadIssue,
  ChunkedProofreadResponse,
  DEFAULT_CHUNK_SIZE,
  DocxProofreadResponse,
  ProofreadIssue,
  ProofreadMode,
  ProofreadResponse,
  ProofreadScope,
  ProofreadStatusEvent,
  ProviderAPI,
  SessionResponse,
  TaskState,
} from "./types";

const API_BASE_URL = "";

export async function createSession(): Promise<SessionResponse> {
  const response = await fetch(`${API_BASE_URL}/api/sessions`, {
    method: "POST",
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as SessionResponse;
}

export async function getAIProfiles(): Promise<AIProfile[]> {
  const response = await fetch(`${API_BASE_URL}/api/ai-profiles`, {
    method: "GET",
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as AIProfile[];
}

export async function requestProofread(
  text: string,
  book: BookInfo,
  sessionId: string,
  aiProfileId: string,
  providerApi: ProviderAPI,
  proofreadMode: ProofreadMode,
  reasoningEnabled: boolean,
  temperature: number,
  onStatus: (status: ProofreadStatusEvent) => void,
  signal: AbortSignal
): Promise<ProofreadResponse> {
  if (providerApi === "chat") {
    onStatus({ stage: "api", message: "正在请求 Chat Completions 接口" });
    return requestProofreadJson(
      text,
      book,
      sessionId,
      aiProfileId,
      providerApi,
      proofreadMode,
      reasoningEnabled,
      temperature,
      signal
    );
  }

  try {
    onStatus({ stage: "api", message: "正在请求流式接口" });
    return await requestProofreadStream(
      text,
      book,
      sessionId,
      aiProfileId,
      providerApi,
      proofreadMode,
      reasoningEnabled,
      temperature,
      onStatus,
      signal
    );
  } catch (error) {
    if (isAbortError(error)) {
      throw error;
    }

    onStatus({ stage: "fallback", message: "流式接口不可用，正在回退到普通接口" });
    return requestProofreadJson(
      text,
      book,
      sessionId,
      aiProfileId,
      providerApi,
      proofreadMode,
      reasoningEnabled,
      temperature,
      signal
    );
  }
}

export async function requestChunkedProofreadTask(
  text: string,
  book: BookInfo,
  scope: ProofreadScope,
  sessionId: string,
  aiProfileId: string,
  providerApi: ProviderAPI,
  proofreadMode: ProofreadMode,
  reasoningEnabled: boolean,
  temperature: number,
  onStatus: (status: ProofreadStatusEvent) => void,
  onTaskCreated: (taskId: string | null, runId: string | null) => void,
  signal: AbortSignal
): Promise<ChunkedProofreadResponse> {
  onStatus({
    stage: "task",
    message:
      scope === "document"
        ? "正在创建全书分块审校任务。"
        : "当前选区超过 7000 字，正在创建分块审校任务。",
  });

  const createdTask = await createProofreadTask(
    text,
    book,
    scope,
    sessionId,
    aiProfileId,
    providerApi,
    proofreadMode,
    reasoningEnabled,
    temperature,
    signal
  );
  onTaskCreated(createdTask.task_id || null, createdTask.run_id || null);
  onStatus(taskSnapshotToStatus("queued", createdTask, "审校任务已创建。"));

  return waitForProofreadTask(createdTask.task_id as string, onStatus, signal);
}

export async function cancelProofreadTask(taskId: string): Promise<void> {
  await fetch(`${API_BASE_URL}/api/proofread/tasks/${taskId}`, {
    method: "DELETE",
  });
}

export async function requestDocxProofreadTask(
  file: File,
  book: BookInfo,
  sessionId: string,
  aiProfileId: string,
  providerApi: ProviderAPI,
  proofreadMode: ProofreadMode,
  reasoningEnabled: boolean,
  temperature: number,
  applicationMode: ApplicationMode,
  fallbackSummaryTruncateEnabled: boolean,
  onStatus: (status: ProofreadStatusEvent) => void,
  onTaskCreated: (taskId: string | null, runId: string | null) => void,
  signal: AbortSignal
): Promise<DocxProofreadResponse> {
  onStatus({ stage: "task", message: "正在上传 DOCX 并创建全书审校任务。" });
  const createdTask = await createDocxProofreadTask(
    file,
    book,
    sessionId,
    aiProfileId,
    providerApi,
    proofreadMode,
    reasoningEnabled,
    temperature,
    applicationMode,
    fallbackSummaryTruncateEnabled,
    signal
  );
  onTaskCreated(createdTask.task_id || null, createdTask.run_id || null);
  onStatus(docxTaskSnapshotToStatus("queued", createdTask, "DOCX 审校任务已创建。"));
  return waitForDocxProofreadTask(createdTask.task_id, onStatus, signal);
}

export async function cancelDocxProofreadTask(taskId: string): Promise<void> {
  await fetch(`${API_BASE_URL}/api/proofread/docx/tasks/${taskId}`, {
    method: "DELETE",
  });
}

export async function retryCurrentProofreadChunk(
  taskId: string
): Promise<ChunkedProofreadResponse> {
  const response = await fetch(`${API_BASE_URL}/api/proofread/tasks/${taskId}/retry-current`, {
    method: "POST",
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as ChunkedProofreadResponse;
}

export async function retryCurrentDocxProofreadChunk(
  taskId: string
): Promise<DocxProofreadResponse> {
  const response = await fetch(`${API_BASE_URL}/api/proofread/docx/tasks/${taskId}/retry-current`, {
    method: "POST",
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as DocxProofreadResponse;
}

export async function retryFailedProofreadChunks(
  taskId: string,
  onStatus: (status: ProofreadStatusEvent) => void,
  signal: AbortSignal
): Promise<ChunkedProofreadResponse> {
  const response = await fetch(`${API_BASE_URL}/api/proofread/tasks/${taskId}/retry-failed`, {
    method: "POST",
    signal,
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  const task = (await response.json()) as ChunkedProofreadResponse;
  onStatus(taskSnapshotToStatus("retry_queued", task, "已创建失败分块重试任务。"));
  return waitForProofreadTask(taskId, onStatus, signal);
}

export async function retryFailedDocxProofreadChunks(
  taskId: string,
  onStatus: (status: ProofreadStatusEvent) => void,
  signal: AbortSignal
): Promise<DocxProofreadResponse> {
  const response = await fetch(`${API_BASE_URL}/api/proofread/docx/tasks/${taskId}/retry-failed`, {
    method: "POST",
    signal,
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  const task = (await response.json()) as DocxProofreadResponse;
  onStatus(docxTaskSnapshotToStatus("retry_queued", task, "已创建 DOCX 失败分块重试任务。"));
  return waitForDocxProofreadTask(taskId, onStatus, signal);
}

export function getDocxDownloadUrl(taskId: string): string {
  return `${API_BASE_URL}/api/proofread/docx/tasks/${taskId}/download`;
}

export async function getAgentRunTrace(
  runId: string,
  signal: AbortSignal
): Promise<AgentRunTrace> {
  const response = await fetch(`${API_BASE_URL}/api/agent/runs/${encodeURIComponent(runId)}/trace`, {
    method: "GET",
    signal,
  });

  if (!response.ok) {
    if (response.status === 404) {
      throw new Error("Trace 不存在或已清理。");
    }
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as AgentRunTrace;
}

export function normalizeChunkedIssuesForScope(issues: ChunkedProofreadIssue[]): ProofreadIssue[] {
  return issues.map((issue) => ({
    ...issue,
    id: `chunk-${issue.chunk_index}-${issue.id}`,
    start: issue.global_start,
    end: issue.global_end,
  }));
}

export function taskSnapshotToStatus(
  stage: string,
  task: ChunkedProofreadResponse,
  message: string
): ProofreadStatusEvent {
  return {
    stage,
    message,
    task_id: task.task_id || undefined,
    run_id: task.run_id || undefined,
    status: task.status,
    scope: task.scope,
    total_chunks: task.total_chunks,
    completed_chunks: task.completed_chunks,
    failed_chunks: task.failed_chunks,
    issue_count: task.issues.length,
    error_message: task.error_message || undefined,
  };
}

export function docxTaskSnapshotToStatus(
  stage: string,
  task: DocxProofreadResponse,
  message: string
): ProofreadStatusEvent {
  return {
    stage,
    message,
    task_id: task.task_id,
    run_id: task.run_id || undefined,
    status: task.status,
    scope: "document",
    total_chunks: task.total_chunks,
    completed_chunks: task.completed_chunks,
    failed_chunks: task.failed_chunks,
    issue_count: task.issue_count,
    source_filename: task.source_filename,
    output_filename: task.output_filename || undefined,
    download_url: task.download_url || undefined,
    expires_at: task.expires_at || undefined,
    retention_days: task.retention_days || undefined,
    error_message: task.error_message || undefined,
  };
}

export function formatProgressResult(status: ProofreadStatusEvent): string {
  if (typeof status.total_chunks !== "number") {
    return status.message;
  }

  const completed = status.completed_chunks || 0;
  const failed = status.failed_chunks || 0;
  const issueCount = status.issue_count || 0;
  const details = [
    `进度 ${completed}/${status.total_chunks}`,
    `失败 ${failed} 块`,
    `累计问题 ${issueCount} 条`,
  ];

  if (typeof status.elapsed_seconds === "number") {
    details.push(`当前块耗时 ${formatElapsedSeconds(status.elapsed_seconds)}`);
  }

  if (status.error_message) {
    details.push(`失败原因：${status.error_message}`);
  }

  return `${status.message} ${details.join("，")}。`;
}

export function isAbortError(error: unknown): boolean {
  return (
    typeof error === "object" && error !== null && "name" in error && error.name === "AbortError"
  );
}

async function requestProofreadJson(
  text: string,
  book: BookInfo,
  sessionId: string,
  aiProfileId: string,
  providerApi: ProviderAPI,
  proofreadMode: ProofreadMode,
  reasoningEnabled: boolean,
  temperature: number,
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
      session_id: sessionId,
      ai_profile_id: aiProfileId,
      provider_api: providerApi,
      proofread_mode: proofreadMode,
      reasoning_enabled: reasoningEnabled,
      temperature,
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
  sessionId: string,
  aiProfileId: string,
  providerApi: ProviderAPI,
  proofreadMode: ProofreadMode,
  reasoningEnabled: boolean,
  temperature: number,
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
      session_id: sessionId,
      ai_profile_id: aiProfileId,
      provider_api: providerApi,
      proofread_mode: proofreadMode,
      reasoning_enabled: reasoningEnabled,
      temperature,
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

async function createProofreadTask(
  text: string,
  book: BookInfo,
  scope: ProofreadScope,
  sessionId: string,
  aiProfileId: string,
  providerApi: ProviderAPI,
  proofreadMode: ProofreadMode,
  reasoningEnabled: boolean,
  temperature: number,
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
      session_id: sessionId,
      ai_profile_id: aiProfileId,
      provider_api: providerApi,
      proofread_mode: proofreadMode,
      reasoning_enabled: reasoningEnabled,
      temperature,
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

async function createDocxProofreadTask(
  file: File,
  book: BookInfo,
  sessionId: string,
  aiProfileId: string,
  providerApi: ProviderAPI,
  proofreadMode: ProofreadMode,
  reasoningEnabled: boolean,
  temperature: number,
  applicationMode: ApplicationMode,
  fallbackSummaryTruncateEnabled: boolean,
  signal: AbortSignal
): Promise<DocxProofreadResponse> {
  const params = new URLSearchParams({
    filename: file.name,
    book: JSON.stringify(book),
    session_id: sessionId,
    ai_profile_id: aiProfileId,
    provider_api: providerApi,
    proofread_mode: proofreadMode,
    reasoning_enabled: String(reasoningEnabled),
    temperature: String(temperature),
    application_mode: applicationMode,
    fallback_summary_truncate_enabled: String(fallbackSummaryTruncateEnabled),
  });
  const response = await fetch(`${API_BASE_URL}/api/proofread/docx/tasks?${params.toString()}`, {
    method: "POST",
    signal,
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    },
    body: file,
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as DocxProofreadResponse;
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
    });

    if (done) {
      break;
    }
  }
}

async function streamDocxProofreadTaskEvents(
  taskId: string,
  onStatus: (status: ProofreadStatusEvent) => void,
  signal: AbortSignal
): Promise<void> {
  const response = await fetch(`${API_BASE_URL}/api/proofread/docx/tasks/${taskId}/events`, {
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
    });

    if (done) {
      break;
    }
  }
}

async function waitForProofreadTask(
  taskId: string,
  onStatus: (status: ProofreadStatusEvent) => void,
  signal: AbortSignal
): Promise<ChunkedProofreadResponse> {
  try {
    await streamProofreadTaskEvents(taskId, onStatus, signal);
  } catch (error) {
    if (isAbortError(error)) {
      throw error;
    }

    onStatus({ stage: "polling", message: "任务进度流不可用，正在切换到轮询查询。" });
    return pollProofreadTask(taskId, onStatus, signal);
  }

  const finalTask = await getProofreadTask(taskId, signal);
  if (finalTask.status === "cancelled") {
    throw createAbortError();
  }

  return finalTask;
}

async function waitForDocxProofreadTask(
  taskId: string,
  onStatus: (status: ProofreadStatusEvent) => void,
  signal: AbortSignal
): Promise<DocxProofreadResponse> {
  try {
    await streamDocxProofreadTaskEvents(taskId, onStatus, signal);
  } catch (error) {
    if (isAbortError(error)) {
      throw error;
    }

    onStatus({ stage: "polling", message: "DOCX 任务进度流不可用，正在切换到轮询查询。" });
    return pollDocxProofreadTask(taskId, onStatus, signal);
  }

  const finalTask = await getDocxProofreadTask(taskId, signal);
  if (finalTask.status === "cancelled") {
    throw createAbortError();
  }

  return finalTask;
}

async function pollProofreadTask(
  taskId: string,
  onStatus: (status: ProofreadStatusEvent) => void,
  signal: AbortSignal
): Promise<ChunkedProofreadResponse> {
  while (true) {
    const task = await getProofreadTask(taskId, signal);
    onStatus(taskSnapshotToStatus("polling", task, "正在查询分块审校进度。"));

    if (task.status === "succeeded" || task.status === "partial_succeeded") {
      return task;
    }

    if (task.status === "failed" || task.status === "cancelled") {
      if (task.status === "failed") {
        return task;
      }
      throw createAbortError();
    }

    await delay(1000, signal);
  }
}

async function pollDocxProofreadTask(
  taskId: string,
  onStatus: (status: ProofreadStatusEvent) => void,
  signal: AbortSignal
): Promise<DocxProofreadResponse> {
  while (true) {
    const task = await getDocxProofreadTask(taskId, signal);
    onStatus(docxTaskSnapshotToStatus("polling", task, "正在查询 DOCX 分块审校进度。"));

    if (task.status === "succeeded" || task.status === "partial_succeeded") {
      return task;
    }

    if (task.status === "failed" || task.status === "cancelled") {
      if (task.status === "failed") {
        return task;
      }
      throw createAbortError();
    }

    await delay(1000, signal);
  }
}

export async function getProofreadTask(
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

export async function getDocxProofreadTask(
  taskId: string,
  signal: AbortSignal
): Promise<DocxProofreadResponse> {
  const response = await fetch(`${API_BASE_URL}/api/proofread/docx/tasks/${taskId}`, {
    method: "GET",
    signal,
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as DocxProofreadResponse;
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
  const message = typeof payload.message === "string" ? payload.message : formatStageMessage(stage);

  return {
    stage,
    message,
    task_id: typeof payload.task_id === "string" ? payload.task_id : undefined,
    run_id:
      typeof payload.run_id === "string"
        ? payload.run_id
        : payload.run_id === null
          ? null
          : undefined,
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
    chunk_len: typeof payload.chunk_len === "number" ? payload.chunk_len : undefined,
    elapsed_seconds:
      typeof payload.elapsed_seconds === "number" ? payload.elapsed_seconds : undefined,
    error_message: typeof payload.error_message === "string" ? payload.error_message : undefined,
    source_filename:
      typeof payload.source_filename === "string" ? payload.source_filename : undefined,
    output_filename:
      typeof payload.output_filename === "string" ? payload.output_filename : undefined,
    download_url: typeof payload.download_url === "string" ? payload.download_url : undefined,
    expires_at: typeof payload.expires_at === "string" ? payload.expires_at : undefined,
    retention_days: typeof payload.retention_days === "number" ? payload.retention_days : undefined,
  };
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

function formatStageMessage(stage: string): string {
  return stage;
}

function formatElapsedSeconds(seconds: number): string {
  if (seconds < 60) {
    return `${seconds.toFixed(1)} 秒`;
  }

  const minutes = Math.floor(seconds / 60);
  const remainingSeconds = Math.round(seconds % 60);
  return `${minutes} 分 ${remainingSeconds} 秒`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
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
    value === "partial_succeeded" ||
    value === "failed" ||
    value === "cancelled"
  );
}

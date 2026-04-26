/* global document, HTMLElement, HTMLButtonElement, HTMLInputElement, HTMLSelectElement, HTMLTextAreaElement */

import {
  ApplicationMode,
  IssueApplicationSummary,
  ProofreadHistoryEntry,
  ProofreadIssue,
  ProofreadMode,
  ProofreadScope,
  ProofreadStatusEvent,
  ProviderAPI,
  TaskState,
} from "./types";

export function renderResult(issues: ProofreadIssue[], commentText: string) {
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

export function formatComment(issues: ProofreadIssue[]): string {
  if (issues.length === 0) {
    return "AI 审校：未发现明显问题。";
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

export function renderEmptyResult(message: string) {
  const result = getElement("result");
  result.className = "result-empty";
  result.textContent = message;
}

export function resetProgress() {
  const progress = getElement("progress");
  progress.className = "progress-log";
  progress.innerHTML = "";
}

export function appendProgressStatus(status: ProofreadStatusEvent) {
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

export function renderHistory(
  entries: ProofreadHistoryEntry[],
  onOpen: (entry: ProofreadHistoryEntry) => void
) {
  const history = getElement("history");

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
        onOpen(entry);
      }
    });
  });
}

export function renderHistoryEntry(entry: ProofreadHistoryEntry) {
  resetProgress();
  appendProgressStatus({ stage: entry.status, message: formatHistoryMeta(entry) });
  renderResult(entry.issues, formatComment(entry.issues));
  showMessage(
    `已打开历史记录：${formatHistoryMeta(entry)}`,
    entry.status === "failed" ? "error" : "default"
  );
}

export function formatCompletionMessage(summary: IssueApplicationSummary): string {
  return `应用完成，已精准批注 ${summary.commentCount} 条，已生成修订 ${summary.revisionCount} 条，未定位 ${summary.fallbackCount} 条。`;
}

export function showMessage(message: string, type: "default" | "error" | "success" = "default") {
  const element = getElement("message");
  element.textContent = message;
  element.className = type === "default" ? "message" : `message is-${type}`;
}

export function getElement(id: string): HTMLElement {
  return document.getElementById(id) as HTMLElement;
}

export function getButton(id: string): HTMLButtonElement {
  return document.getElementById(id) as HTMLButtonElement;
}

export function getSelect(id: string): HTMLSelectElement {
  return document.getElementById(id) as HTMLSelectElement;
}

export function getInput(id: string): HTMLInputElement {
  return document.getElementById(id) as HTMLInputElement;
}

export function getTextArea(id: string): HTMLTextAreaElement {
  return document.getElementById(id) as HTMLTextAreaElement;
}

export function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  return String(error);
}

export function formatProviderApi(providerApi: ProviderAPI): string {
  return providerApi === "chat" ? "Chat" : "Responses";
}

export function formatProofreadMode(proofreadMode: ProofreadMode): string {
  return proofreadMode === "thinking" ? "深度审校" : "快速审校";
}

export function formatApplicationMode(applicationMode: ApplicationMode): string {
  return applicationMode === "revision" ? "修订模式" : "批注模式";
}

export function formatProofreadScope(scope: ProofreadScope): string {
  return scope === "document" ? "全书正文" : "当前选区";
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
    heartbeat: "心跳",
    normalizing: "整理",
    partial_succeeded: "部分完成",
    polling: "轮询",
    queued: "排队",
    received: "接收",
    result: "结果",
    running: "运行",
    task: "任务",
  };

  return `[${stageLabels[stage] || stage}]`;
}

function formatLocationStatus(issue: ProofreadIssue): string {
  if (typeof issue.start === "number" && typeof issue.end === "number") {
    return `已定位 ${issue.start}-${issue.end}`;
  }

  return "未定位";
}

function formatHistoryMeta(entry: ProofreadHistoryEntry): string {
  const statusLabels: Record<TaskState, string> = {
    cancelled: "已停止",
    failed: "失败",
    idle: "未开始",
    partial_succeeded: "部分完成",
    queued: "排队中",
    running: "运行中",
    succeeded: "完成",
  };
  const createdAt = new Date(entry.createdAt).toLocaleString();
  const issueText = entry.issueCount > 0 ? `${entry.issueCount} 条问题` : "无问题";
  const locatedText =
    entry.issueCount > 0 ? `定位 ${entry.locatedIssueCount}/${entry.issueCount}` : "无需定位";
  const actionText = entry.appliedToWord
    ? entry.applicationMode === "revision"
      ? `已应用修订 ${entry.revisionCount} / 未定位 ${entry.fallbackCount}`
      : `已应用批注 ${entry.locatedIssueCount} / 未定位 ${entry.fallbackCount}`
    : `未应用 / 未定位 ${entry.fallbackCount}`;
  const modeText = `${formatProofreadScope(entry.scope)} / ${formatProofreadMode(entry.proofreadMode)} / ${formatProviderApi(entry.providerApi)} / ${formatApplicationMode(entry.applicationMode)}`;
  const chunkText =
    entry.totalChunks > 1
      ? `分块 ${entry.completedChunks}/${entry.totalChunks}，失败 ${entry.failedChunks}`
      : "单段";
  const bookText = entry.bookTitle ? `《${entry.bookTitle}》` : "未记录书名";
  const errorText = entry.errorMessage ? `：${entry.errorMessage}` : "";

  return `${createdAt} / ${bookText} / ${statusLabels[entry.status]} / ${modeText} / ${chunkText} / ${issueText} / ${locatedText} / ${actionText}${errorText}`;
}

function escapeHtml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

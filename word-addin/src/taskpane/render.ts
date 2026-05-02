/* global document, HTMLElement, HTMLButtonElement, HTMLInputElement, HTMLSelectElement, HTMLTextAreaElement */

import { appendDebugLog } from "./debug";
import {
  ApplicationMode,
  IssueApplicationSummary,
  IssueFilterState,
  IssueReviewState,
  ProofreadHistoryEntry,
  ProofreadIssue,
  ProofreadMode,
  ProofreadScope,
  ProofreadStatusEvent,
  ProviderAPI,
  TaskState,
} from "./types";

export type BulkSelectionAction = "all" | "none" | "high-medium" | "replaceable";

export interface ResultRenderOptions {
  sourceText?: string;
  reviewState?: IssueReviewState;
  applicationMode?: ApplicationMode;
  readonly?: boolean;
  onToggleIssue?: (issueId: string, selected: boolean) => void;
  onBulkSelect?: (action: BulkSelectionAction) => void;
  onFilterChange?: (filter: IssueFilterState) => void;
  onLocateIssue?: (issue: ProofreadIssue) => void;
}

export function renderResult(
  issues: ProofreadIssue[],
  commentText: string,
  options: ResultRenderOptions = {}
) {
  const result = getElement("result");

  if (issues.length === 0) {
    result.className = "result-empty";
    result.textContent = commentText;
    return;
  }

  const reviewState = options.reviewState || {
    selectedIssueIds: issues.map((issue) => issue.id),
    filter: createDefaultFilterState(),
  };
  const selectedIssueIdSet = new Set(reviewState.selectedIssueIds);
  const orderedIssues = issues;
  const visibleIssues = orderedIssues.filter((issue) =>
    matchesFilter(issue, reviewState.filter, options.sourceText)
  );
  const selectedIssues = orderedIssues.filter((issue) => selectedIssueIdSet.has(issue.id));
  const summary = formatSelectionSummary(
    orderedIssues,
    selectedIssues,
    options.applicationMode || "comment",
    options.sourceText
  );

  result.className = "result-list";
  result.innerHTML = `
    ${renderReviewToolbar(orderedIssues, visibleIssues, reviewState, summary, Boolean(options.readonly))}
    <div class="result-items">
      ${visibleIssues
        .map((issue, index) =>
          renderIssueItem({
            issue,
            displayIndex: orderedIssues.indexOf(issue) + 1 || index + 1,
            selected: selectedIssueIdSet.has(issue.id),
            sourceText: options.sourceText,
            readonly: Boolean(options.readonly),
          })
        )
        .join("")}
    </div>
  `;

  bindResultEvents(result, visibleIssues, reviewState, options);
}

export function createDefaultFilterState(): IssueFilterState {
  return {
    severity: "all",
    category: "all",
    location: "all",
    replacement: "all",
  };
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

export function formatApplyButtonLabel(selectedCount: number): string {
  return selectedCount > 0 ? `应用 ${selectedCount} 条到 Word` : "应用到 Word";
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
  renderResult(entry.issues, formatComment(entry.issues), {
    reviewState: {
      selectedIssueIds: entry.selectedIssueIds,
      filter: createDefaultFilterState(),
    },
    applicationMode: entry.applicationMode,
    readonly: true,
  });
  showMessage(
    `已打开历史记录：${formatHistoryMeta(entry)}`,
    entry.status === "failed" ? "error" : "default"
  );
}

export function formatCompletionMessage(summary: IssueApplicationSummary): string {
  const successCount = summary.commentCount + summary.revisionCount + summary.fallbackCount;
  const failedText =
    summary.failedCount > 0 ? `，${summary.failedCount} 条因 Word 写入失败未应用` : "";
  const truncatedText =
    summary.truncatedFallbackCount > 0
      ? `，${summary.truncatedFallbackCount} 条因汇总批注截断未写入 Word`
      : "";
  const fallbackCommentText =
    summary.fallbackCommentCount > 1
      ? `，汇总批注已拆分为 ${summary.fallbackCommentCount} 条短批注`
      : "";

  if (successCount === 0 && summary.failedCount > 0 && summary.truncatedFallbackCount === 0) {
    return `应用失败，${summary.failedCount} 条均被 Word 拒绝写入。`;
  }

  return `应用完成，已写入批注 ${summary.commentCount} 条，已生成修订并附批注 ${summary.revisionCount} 条，未定位汇总 ${summary.fallbackCount} 条${fallbackCommentText}${truncatedText}${failedText}。`;
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
  return applicationMode === "revision" ? "修订+批注" : "批注模式";
}

export function formatProofreadScope(scope: ProofreadScope): string {
  return scope === "document" ? "全书正文" : "当前选区";
}

function renderReviewToolbar(
  issues: ProofreadIssue[],
  visibleIssues: ProofreadIssue[],
  reviewState: IssueReviewState,
  summary: string,
  readonly: boolean
): string {
  const categories = Array.from(new Set(issues.map((issue) => issue.category))).sort();

  return `
    <div class="result-toolbar">
      <div class="result-summary">${escapeHtml(summary)}</div>
      <div class="result-filters">
        <label>
          <span>严重程度</span>
          <select class="control-select result-filter" data-filter="severity" ${readonly ? "disabled" : ""}>
            ${renderFilterOption("all", "全部", reviewState.filter.severity)}
            ${renderFilterOption("high-medium", "高/中风险", reviewState.filter.severity)}
            ${renderFilterOption("high", "高风险", reviewState.filter.severity)}
            ${renderFilterOption("medium", "中风险", reviewState.filter.severity)}
            ${renderFilterOption("low", "低风险", reviewState.filter.severity)}
          </select>
        </label>
        <label>
          <span>类别</span>
          <select class="control-select result-filter" data-filter="category" ${readonly ? "disabled" : ""}>
            ${renderFilterOption("all", "全部", reviewState.filter.category)}
            ${categories
              .map((category) =>
                renderFilterOption(category, category, reviewState.filter.category)
              )
              .join("")}
          </select>
        </label>
        <label>
          <span>定位</span>
          <select class="control-select result-filter" data-filter="location" ${readonly ? "disabled" : ""}>
            ${renderFilterOption("all", "全部", reviewState.filter.location)}
            ${renderFilterOption("located", "已定位", reviewState.filter.location)}
            ${renderFilterOption("unlocated", "未定位", reviewState.filter.location)}
          </select>
        </label>
        <label>
          <span>替换</span>
          <select class="control-select result-filter" data-filter="replacement" ${readonly ? "disabled" : ""}>
            ${renderFilterOption("all", "全部", reviewState.filter.replacement)}
            ${renderFilterOption("with-replacement", "可直接替换", reviewState.filter.replacement)}
            ${renderFilterOption("needs-review", "需人工核查", reviewState.filter.replacement)}
          </select>
        </label>
      </div>
      <div class="result-bulk-actions">
        <button class="ms-Button result-bulk" type="button" data-action="all" ${readonly ? "disabled" : ""}>全选</button>
        <button class="ms-Button result-bulk" type="button" data-action="none" ${readonly ? "disabled" : ""}>全不选</button>
        <button class="ms-Button result-bulk" type="button" data-action="high-medium" ${readonly ? "disabled" : ""}>只选高/中风险</button>
        <button class="ms-Button result-bulk" type="button" data-action="replaceable" ${readonly ? "disabled" : ""}>只选可直接替换项</button>
      </div>
      <div class="result-visible-count">当前显示 ${visibleIssues.length}/${issues.length} 条</div>
    </div>
  `;
}

function renderIssueItem(input: {
  issue: ProofreadIssue;
  displayIndex: number;
  selected: boolean;
  sourceText?: string;
  readonly: boolean;
}): string {
  const issue = input.issue;
  const locatable = input.sourceText
    ? isIssueLocatable(input.sourceText, issue)
    : hasReplayableLocator(issue);
  const preciselyWritable = input.sourceText
    ? isPreciselyWritableIssue(input.sourceText, issue)
    : hasReplayableLocator(issue);
  const canLocate = Boolean(!input.readonly && locatable);
  const needsReview = !hasReplacement(issue);
  const statusClass = preciselyWritable ? "is-located" : "is-unlocated";
  const replacementClass = needsReview ? "needs-review" : "has-replacement";

  return `
    <article class="result-item ${statusClass} ${replacementClass}">
      <div class="result-item-heading">
        <label class="issue-select">
          <input type="checkbox" data-issue-id="${escapeHtml(issue.id)}" ${
            input.selected ? "checked" : ""
          } ${input.readonly ? "disabled" : ""} />
          <span>${input.displayIndex}. ${escapeHtml(issue.category)} / ${escapeHtml(issue.severity)}</span>
        </label>
        <span class="location-status">${escapeHtml(formatLocationStatus(issue, preciselyWritable))}</span>
      </div>
      <p><b>原文：</b>${escapeHtml(issue.original || "未提供")}</p>
      <p><b>替换为：</b>${escapeHtml(issue.replacement || "无直接替换文本，需人工核查")}</p>
      <p><b>建议：</b>${escapeHtml(issue.suggestion || "未提供")}</p>
      ${renderContextPreview(input.sourceText, issue)}
      <button class="ms-Button locate-issue" type="button" data-locate-id="${escapeHtml(issue.id)}" ${
        canLocate ? "" : "disabled"
      }>定位</button>
    </article>
  `;
}

function bindResultEvents(
  result: HTMLElement,
  visibleIssues: ProofreadIssue[],
  reviewState: IssueReviewState,
  options: ResultRenderOptions
) {
  result.querySelectorAll("[data-issue-id]").forEach((item) => {
    item.addEventListener("change", () => {
      const input = item as HTMLInputElement;
      options.onToggleIssue?.(input.getAttribute("data-issue-id") || "", input.checked);
    });
  });

  result.querySelectorAll(".result-bulk").forEach((item) => {
    item.addEventListener("click", () => {
      const action = (item as HTMLElement).getAttribute("data-action") as BulkSelectionAction;
      options.onBulkSelect?.(action);
    });
  });

  result.querySelectorAll(".result-filter").forEach((item) => {
    item.addEventListener("change", () => {
      const select = item as HTMLSelectElement;
      const filterKey = select.getAttribute("data-filter");
      const nextFilter = { ...reviewState.filter };

      if (filterKey === "severity") {
        nextFilter.severity = select.value as IssueFilterState["severity"];
      }
      if (filterKey === "category") {
        nextFilter.category = select.value;
      }
      if (filterKey === "location") {
        nextFilter.location = select.value as IssueFilterState["location"];
      }
      if (filterKey === "replacement") {
        nextFilter.replacement = select.value as IssueFilterState["replacement"];
      }

      options.onFilterChange?.(nextFilter);
    });
  });

  result.querySelectorAll("[data-locate-id]").forEach((item) => {
    item.addEventListener("click", () => {
      const issueId = (item as HTMLElement).getAttribute("data-locate-id");
      const target = visibleIssues.find((issue) => issue.id === issueId);

      appendDebugLog("info", "定位按钮已点击", {
        issueId,
        foundInVisibleIssues: Boolean(target),
      });

      if (target) {
        options.onLocateIssue?.(target);
      }
    });
  });
}

function matchesFilter(
  issue: ProofreadIssue,
  filter: IssueFilterState,
  sourceText?: string
): boolean {
  if (filter.severity === "high-medium" && issue.severity === "low") {
    return false;
  }

  if (
    filter.severity !== "all" &&
    filter.severity !== "high-medium" &&
    issue.severity !== filter.severity
  ) {
    return false;
  }

  if (filter.category !== "all" && issue.category !== filter.category) {
    return false;
  }

  const located = sourceText
    ? isPreciselyWritableIssue(sourceText, issue)
    : isLocatedIssue(issue) && Boolean(issue.locator);
  if (filter.location === "located" && !located) {
    return false;
  }
  if (filter.location === "unlocated" && located) {
    return false;
  }

  const replaceable = hasReplacement(issue);
  if (filter.replacement === "with-replacement" && !replaceable) {
    return false;
  }
  if (filter.replacement === "needs-review" && replaceable) {
    return false;
  }

  return true;
}

function renderFilterOption(value: string, label: string, selectedValue: string): string {
  return `<option value="${escapeHtml(value)}" ${value === selectedValue ? "selected" : ""}>${escapeHtml(label)}</option>`;
}

function renderContextPreview(sourceText: string | undefined, issue: ProofreadIssue): string {
  if (sourceText && typeof issue.start === "number" && typeof issue.end === "number") {
    const before = sourceText.slice(Math.max(0, issue.start - 24), issue.start);
    const original = sourceText.slice(issue.start, issue.end);
    const after = sourceText.slice(issue.end, Math.min(sourceText.length, issue.end + 24));

    return `<p class="context-preview"><b>上下文：</b>${escapeHtml(before)}<mark>${escapeHtml(original)}</mark>${escapeHtml(after)}</p>`;
  }

  if (!hasReplayableLocator(issue)) {
    return "";
  }

  const locator = issue.locator;
  const before = locator.key.slice(0, locator.original_start_in_key);
  const original = locator.key.slice(locator.original_start_in_key, locator.original_end_in_key);
  const after = locator.key.slice(locator.original_end_in_key);

  return `<p class="context-preview"><b>上下文：</b>${escapeHtml(before)}<mark>${escapeHtml(original)}</mark>${escapeHtml(after)}</p>`;
}

function formatSelectionSummary(
  allIssues: ProofreadIssue[],
  selectedIssues: ProofreadIssue[],
  applicationMode: ApplicationMode,
  sourceText?: string
): string {
  const skippedCount = allIssues.length - selectedIssues.length;
  const revisionCount =
    applicationMode === "revision"
      ? selectedIssues.filter(
          (issue) => isPreciselyWritableIssue(sourceText || "", issue) && hasReplacement(issue)
        ).length
      : 0;
  const commentCount = selectedIssues.filter((issue) => {
    if (!isPreciselyWritableIssue(sourceText || "", issue)) {
      return false;
    }

    return applicationMode === "comment" || !hasReplacement(issue);
  }).length;
  const fallbackCount = selectedIssues.filter(
    (issue) => !isPreciselyWritableIssue(sourceText || "", issue)
  ).length;
  const batchCount = Math.ceil((commentCount + revisionCount) / 16);
  const batchText = batchCount > 0 ? `；预计分 ${batchCount} 批应用` : "";

  return `已选 ${selectedIssues.length}/${allIssues.length} 条；预计精准批注 ${commentCount} 条，生成修订并附批注 ${revisionCount} 条，将汇总批注 ${fallbackCount} 条，跳过 ${skippedCount} 条${batchText}。`;
}

function isLocatedIssue(issue: ProofreadIssue): boolean {
  return typeof issue.start === "number" && typeof issue.end === "number";
}

function isIssueLocatable(sourceText: string, issue: ProofreadIssue): boolean {
  if (!sourceText || !isLocatedIssue(issue)) {
    return false;
  }

  return sourceText.slice(issue.start as number, issue.end as number) === issue.original;
}

function isPreciselyWritableIssue(sourceText: string, issue: ProofreadIssue): boolean {
  if (!sourceText) {
    return hasReplayableLocator(issue);
  }

  if (!isIssueLocatable(sourceText, issue)) {
    return false;
  }

  if (issue.locator) {
    return true;
  }

  return canBuildClientLocator(sourceText, issue);
}

function hasReplayableLocator(issue: ProofreadIssue): issue is ProofreadIssue & {
  locator: NonNullable<ProofreadIssue["locator"]>;
} {
  return (
    Boolean(issue.locator?.key) &&
    typeof issue.locator?.key_occurrence_index === "number" &&
    typeof issue.locator.original_start_in_key === "number" &&
    typeof issue.locator.original_end_in_key === "number"
  );
}

function canBuildClientLocator(sourceText: string, issue: ProofreadIssue): boolean {
  if (issue.original.length >= 6 && getOccurrenceCount(sourceText, issue.original) <= 3) {
    return true;
  }

  for (const windowSize of [16, 32, 64]) {
    const keyStart = Math.max(0, (issue.start as number) - windowSize);
    const keyEnd = Math.min(sourceText.length, (issue.end as number) + windowSize);
    const key = sourceText.slice(keyStart, keyEnd);

    if (key && getOccurrenceCount(sourceText, key) <= 3) {
      return true;
    }
  }

  return false;
}

function getOccurrenceCount(text: string, target: string): number {
  if (!target) {
    return 0;
  }

  let count = 0;
  let searchFrom = 0;

  while (searchFrom < text.length) {
    const foundAt = text.indexOf(target, searchFrom);

    if (foundAt === -1) {
      break;
    }

    count += 1;
    searchFrom = foundAt + 1;
  }

  return count;
}

function hasReplacement(issue: ProofreadIssue): boolean {
  return typeof issue.replacement === "string" && issue.replacement.trim().length > 0;
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
    retry_queued: "重试",
    chunk_retry_requested: "重试",
    chunk_retrying: "重试",
    result: "结果",
    running: "运行",
    task: "任务",
  };

  return `[${stageLabels[stage] || stage}]`;
}

function formatLocationStatus(issue: ProofreadIssue, preciselyWritable: boolean): string {
  if (preciselyWritable && typeof issue.start === "number" && typeof issue.end === "number") {
    return `可精准写回 ${issue.start}-${issue.end}`;
  }

  return "将汇总批注";
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
  const skippedText = entry.skippedIssueCount > 0 ? ` / 跳过 ${entry.skippedIssueCount}` : "";
  const failedText =
    entry.failedCount && entry.failedCount > 0 ? ` / 写入失败 ${entry.failedCount}` : "";
  const actionText = entry.appliedToWord
    ? entry.applicationMode === "revision"
      ? `已应用修订+批注 ${entry.revisionCount} / 未定位 ${entry.fallbackCount}${failedText}${skippedText}`
      : `已应用批注 ${entry.locatedIssueCount} / 未定位 ${entry.fallbackCount}${failedText}${skippedText}`
    : `未应用 / 未定位 ${entry.fallbackCount}${failedText}${skippedText}`;
  const reasoningText = entry.reasoningEnabled ? "深度思考开" : "深度思考关";
  const modeText = `${formatProofreadScope(entry.scope)} / ${formatProofreadMode(entry.proofreadMode)} / ${reasoningText} / ${formatProviderApi(entry.providerApi)} / ${formatApplicationMode(entry.applicationMode)}`;
  const chunkText =
    entry.totalChunks > 1
      ? `分块 ${entry.completedChunks}/${entry.totalChunks}，失败 ${entry.failedChunks}`
      : "单段";
  const bookText = entry.bookTitle ? `《${entry.bookTitle}》` : "未记录书名";
  const fileText = entry.sourceFilename
    ? ` / 文件 ${entry.sourceFilename}${entry.outputFilename ? ` -> ${entry.outputFilename}` : ""}`
    : "";
  const errorText = entry.errorMessage ? `：${entry.errorMessage}` : "";

  return `${createdAt} / ${bookText}${fileText} / ${statusLabels[entry.status]} / ${modeText} / ${chunkText} / ${issueText} / ${locatedText} / ${actionText}${errorText}`;
}

function escapeHtml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

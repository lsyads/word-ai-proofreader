/* global AbortController, AbortSignal, File, HTMLButtonElement, HTMLElement, HTMLInputElement, HTMLSelectElement, HTMLTextAreaElement, Office, document, window */

import {
  createSession,
  createV2Project,
  createV2SelectionProject,
  decideV2Candidates,
  getAIProfiles,
  getV2Candidates,
  getV2DocumentMap,
  getV2DownloadUrl,
  getV2Project,
  getV2ReviewReport,
  getV2RunTrace,
  markV2CandidatesWritten,
  runV2Project,
  writebackV2Project,
} from "./api";
import { appendDebugLog, clearDebugLog } from "./debug";
import {
  AIProfile,
  ApplicationMode,
  BookInfo,
  DEFAULT_TEMPERATURE,
  ProofreadIssue,
  ProviderAPI,
  V2ApprovalDecision,
  V2CandidateIssue,
  V2CandidateStatus,
  V2DocumentMap,
  V2Project,
  V2ReviewReport,
  V2Run,
  V2RunTrace,
} from "./types";
import {
  applyIssuesToScope,
  clearTrackedSelectionRange,
  ensureWordCommentSupport,
  getSelectedText,
} from "./word";

type SourceType = "selection" | "docx";

const DEFAULT_REVIEW_GOAL = "完成出版审校，找出明显错别字、语病、体例问题和上下文一致性风险。";

let currentSessionId: string | null = null;
let currentAbortController: AbortController | null = null;
let aiProfiles: AIProfile[] = [];
let currentProject: V2Project | null = null;
let currentDocumentMap: V2DocumentMap | null = null;
let currentRun: V2Run | null = null;
let currentTrace: V2RunTrace | null = null;
let currentCandidates: V2CandidateIssue[] = [];
let currentReport: V2ReviewReport | null = null;
let currentSelectionText = "";
let isBusy = false;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    getButton("create-project").disabled = true;
    showMessage("请在 Microsoft Word 任务窗格中使用此插件。", "error");
    return;
  }

  bindEvents();
  initializeDefaults();
  initializeSession();
  initializeAIProfiles();
  renderWorkspace();
});

function bindEvents() {
  getButton("create-project").onclick = createProject;
  getButton("run-agent").onclick = runAgent;
  getButton("refresh-workspace").onclick = refreshWorkspace;
  getButton("refresh-trace").onclick = refreshTrace;
  getButton("approve-all").onclick = () => decidePendingCandidates("approved");
  getButton("reject-all").onclick = () => decidePendingCandidates("rejected");
  getButton("writeback").onclick = writebackApproved;
  getButton("download-docx").onclick = downloadDocx;
  getButton("refresh-report").onclick = refreshReport;
  getButton("clear-debug-log").onclick = clearDebugLog;
  getSelect("source-type").onchange = handleSourceTypeChange;
  getSelect("candidate-filter").onchange = renderCandidates;
  getInput("docx-file").onchange = updateDocxFileSummary;
}

function initializeDefaults() {
  getTextArea("review-goal").value = DEFAULT_REVIEW_GOAL;
  getInput("temperature").value = String(DEFAULT_TEMPERATURE);
  getInput("fallback-summary-truncate-enabled").checked = true;
  handleSourceTypeChange();
}

async function initializeSession() {
  try {
    const session = await createSession();
    currentSessionId = session.session_id;
    showMessage("已创建 V2 审校会话。", "default");
  } catch (error) {
    showMessage(`创建会话失败：${getErrorMessage(error)}`, "error");
  }
}

async function initializeAIProfiles() {
  try {
    aiProfiles = await getAIProfiles();
    const select = getSelect("ai-profile");
    select.innerHTML = aiProfiles
      .map(
        (profile) =>
          `<option value="${escapeHtml(profile.id)}">${escapeHtml(profile.label)} · ${escapeHtml(profile.model)}</option>`
      )
      .join("");
    const firstConfigured = aiProfiles.find((profile) => profile.configured) || aiProfiles[0];
    if (firstConfigured) {
      select.value = firstConfigured.id;
      getSelect("provider-api").value = firstConfigured.default_api;
    }
  } catch (error) {
    appendDebugLog("warn", "加载 AI 配置失败", { error: getErrorMessage(error) });
  }
}

async function createProject() {
  const book = getValidatedBookInfo();
  if (!book) {
    return;
  }
  const reviewGoal = getTextArea("review-goal").value.trim() || DEFAULT_REVIEW_GOAL;
  const sourceType = getSourceType();
  const abortController = startBusy();

  try {
    resetProjectState();
    if (sourceType === "selection") {
      currentSelectionText = await getSelectedText();
      currentProject = await createV2SelectionProject(
        currentSelectionText,
        book,
        reviewGoal,
        currentSessionId,
        abortController.signal
      );
    } else {
      await clearTrackedSelectionRange();
      const file = getDocxFile();
      if (!file) {
        throw new Error("请先选择一个 .docx 文件。");
      }
      currentProject = await createV2Project(file, book, reviewGoal, abortController.signal);
      currentSelectionText = "";
    }
    currentDocumentMap = await getV2DocumentMap(currentProject.project_id, abortController.signal);
    showMessage("V2 审校项目已创建。", "success");
  } catch (error) {
    showMessage(`创建项目失败：${getErrorMessage(error)}`, "error");
    appendDebugLog("error", "创建 V2 项目失败", { error: getErrorMessage(error) });
  } finally {
    stopBusy();
    renderWorkspace();
  }
}

async function runAgent() {
  if (!currentProject) {
    return;
  }
  const controls = getRunControls();
  const abortController = startBusy();

  try {
    currentRun = await runV2Project(
      currentProject.project_id,
      currentSessionId || "",
      controls.aiProfileId,
      controls.providerApi,
      controls.proofreadMode,
      controls.reasoningEnabled,
      controls.temperature,
      abortController.signal
    );
    await refreshWorkspaceData(abortController.signal);
    showMessage("Agent 已完成审校，候选问题已进入确认队列。", "success");
  } catch (error) {
    showMessage(`运行 Agent 失败：${getErrorMessage(error)}`, "error");
    appendDebugLog("error", "运行 V2 Agent 失败", { error: getErrorMessage(error) });
  } finally {
    stopBusy();
    renderWorkspace();
  }
}

async function refreshWorkspace() {
  if (!currentProject) {
    return;
  }
  const abortController = startBusy();
  try {
    await refreshWorkspaceData(abortController.signal);
    showMessage("工作台已刷新。", "success");
  } catch (error) {
    showMessage(`刷新失败：${getErrorMessage(error)}`, "error");
  } finally {
    stopBusy();
    renderWorkspace();
  }
}

async function refreshWorkspaceData(signal: AbortSignal) {
  if (!currentProject) {
    return;
  }
  currentProject = await getV2Project(currentProject.project_id, signal);
  currentDocumentMap = await getV2DocumentMap(currentProject.project_id, signal);
  currentCandidates = (await getV2Candidates(currentProject.project_id, signal)).candidates;
  if (currentRun) {
    currentTrace = await getV2RunTrace(currentProject.project_id, currentRun.run_id, signal);
  }
  currentReport = await getV2ReviewReport(currentProject.project_id, signal);
}

async function refreshTrace() {
  if (!currentProject || !currentRun) {
    return;
  }
  const abortController = startBusy();
  try {
    currentTrace = await getV2RunTrace(
      currentProject.project_id,
      currentRun.run_id,
      abortController.signal
    );
    showMessage("Trace 已刷新。", "success");
  } catch (error) {
    showMessage(`刷新 Trace 失败：${getErrorMessage(error)}`, "error");
  } finally {
    stopBusy();
    renderWorkspace();
  }
}

async function refreshReport() {
  if (!currentProject) {
    return;
  }
  const abortController = startBusy();
  try {
    currentReport = await getV2ReviewReport(currentProject.project_id, abortController.signal);
    showMessage("报告已刷新。", "success");
  } catch (error) {
    showMessage(`刷新报告失败：${getErrorMessage(error)}`, "error");
  } finally {
    stopBusy();
    renderWorkspace();
  }
}

async function decidePendingCandidates(status: "approved" | "rejected") {
  const pending = currentCandidates.filter((candidate) => candidate.status === "pending");
  await decideCandidates(
    pending.map((candidate) => ({ candidate_id: candidate.candidate_id, status }))
  );
}

async function decideCandidate(candidateId: string, status: "approved" | "rejected" | "deferred") {
  await decideCandidates([{ candidate_id: candidateId, status }]);
}

async function decideCandidates(decisions: V2ApprovalDecision[]) {
  if (!currentProject || decisions.length === 0) {
    return;
  }
  const abortController = startBusy();
  try {
    const response = await decideV2Candidates(
      currentProject.project_id,
      decisions,
      abortController.signal
    );
    currentCandidates = response.candidates;
    currentReport = await getV2ReviewReport(currentProject.project_id, abortController.signal);
    showMessage("候选问题决策已更新。", "success");
  } catch (error) {
    showMessage(`更新候选问题失败：${getErrorMessage(error)}`, "error");
  } finally {
    stopBusy();
    renderWorkspace();
  }
}

async function writebackApproved() {
  if (!currentProject) {
    return;
  }
  const approved = currentCandidates.filter((candidate) => candidate.status === "approved");
  if (approved.length === 0) {
    showMessage("没有已批准的候选问题可写回。", "error");
    return;
  }
  const abortController = startBusy();
  try {
    if (currentProject.source_type === "selection") {
      await writebackSelection(approved, abortController.signal);
    } else {
      await writebackDocx(abortController.signal);
    }
    await refreshWorkspaceData(abortController.signal);
    showMessage("已写回已批准候选问题。", "success");
  } catch (error) {
    showMessage(`写回失败：${getErrorMessage(error)}`, "error");
    appendDebugLog("error", "V2 写回失败", { error: getErrorMessage(error) });
  } finally {
    stopBusy();
    renderWorkspace();
  }
}

async function writebackSelection(approved: V2CandidateIssue[], signal: AbortSignal) {
  if (!currentProject) {
    return;
  }
  if (!ensureWordCommentSupport()) {
    throw new Error("当前 Word 环境不支持批注 API。");
  }
  if (!currentSelectionText) {
    throw new Error("当前选区项目缺少缓存文本，请重新创建当前选区项目。");
  }
  const issues = approved.map(candidateToProofreadIssue);
  await applyIssuesToScope(currentSelectionText, issues, "selection", getApplicationMode(), {
    fallbackSummaryTruncateEnabled: getInput("fallback-summary-truncate-enabled").checked,
  });
  await markV2CandidatesWritten(
    currentProject.project_id,
    approved.map((candidate) => candidate.candidate_id),
    signal
  );
}

async function writebackDocx(signal: AbortSignal) {
  if (!currentProject) {
    return;
  }
  await writebackV2Project(
    currentProject.project_id,
    getApplicationMode(),
    getInput("fallback-summary-truncate-enabled").checked,
    signal
  );
}

function downloadDocx() {
  if (!currentProject || currentProject.source_type !== "docx" || !currentProject.output_filename) {
    return;
  }
  window.location.href = getV2DownloadUrl(currentProject.project_id);
}

function renderWorkspace() {
  renderProjectBadge();
  renderProjectSummary();
  renderDocumentMap();
  renderReviewPlan();
  renderTrace();
  renderCandidates();
  renderReport();
  updateButtons();
}

function renderProjectBadge() {
  const badge = getElement("project-badge");
  const status = currentProject?.status || "未创建";
  badge.textContent = translateStatus(status);
  badge.className = "status-badge";
  if (status === "running" || status === "waiting_for_approval") {
    badge.classList.add("is-running");
  } else if (status === "written") {
    badge.classList.add("is-success");
  } else if (status === "failed" || status === "cancelled") {
    badge.classList.add("is-error");
  }
}

function renderProjectSummary() {
  const container = getElement("project-summary");
  if (!currentProject) {
    renderEmpty(container, "暂无项目。");
    return;
  }
  container.className = "workspace-summary";
  container.innerHTML = `
    <div class="summary-grid">
      ${summaryItem("来源", currentProject.source_type === "selection" ? "当前选区" : "全书 DOCX")}
      ${summaryItem("项目状态", translateStatus(currentProject.status))}
      ${summaryItem("书名", currentProject.book.title)}
      ${summaryItem("候选问题", String(currentProject.candidate_count))}
      ${summaryItem("待确认", String(currentProject.pending_count))}
      ${summaryItem("已批准", String(currentProject.approved_count))}
      ${summaryItem("文件", currentProject.source_filename || currentProject.text_preview || "-")}
      ${summaryItem("Run", currentRun?.run_id || "-")}
    </div>
  `;
}

function renderDocumentMap() {
  const container = getElement("document-map");
  if (!currentDocumentMap) {
    renderEmpty(container, "创建项目后显示文档地图。");
    return;
  }
  container.className = "workspace-list";
  const chunks = currentDocumentMap.chunks.slice(0, 8);
  container.innerHTML = `
    <div class="summary-line">全文 ${currentDocumentMap.text_len} 字，${currentDocumentMap.block_count} 个块，${currentDocumentMap.chunk_count} 个审校分块。</div>
    ${chunks
      .map(
        (chunk) => `
          <div class="map-item">
            <div class="item-title">Chunk ${chunk.index + 1} · ${chunk.chunk_len} 字</div>
            <div class="item-meta">${chunk.start} - ${chunk.end}</div>
            <div class="item-body">${escapeHtml(chunk.text_preview || "-")}</div>
          </div>
        `
      )
      .join("")}
  `;
}

function renderReviewPlan() {
  const container = getElement("review-plan");
  if (!currentRun && !currentProject) {
    renderEmpty(container, "运行 Agent 后显示审校计划。");
    return;
  }
  container.className = "workspace-list";
  const steps = [
    ["建立文档地图", "build_document_map"],
    ["生成审校计划", "create_review_plan"],
    ["分块审校", "proofread_document_chunk"],
    ["候选问题归并与自检", "evaluate_candidate_issues"],
    ["等待编辑确认", "human_approval_queue"],
  ];
  container.innerHTML = steps
    .map(
      ([title, toolName], index) => `
        <div class="plan-item">
          <div class="item-title">${index + 1}. ${title}</div>
          <div class="item-meta">${toolName}</div>
        </div>
      `
    )
    .join("");
}

function renderTrace() {
  const container = getElement("run-trace");
  if (!currentTrace || currentTrace.events.length === 0) {
    renderEmpty(container, "暂无运行记录。");
    return;
  }
  container.className = "workspace-list";
  container.innerHTML = currentTrace.events
    .map(
      (event) => `
        <div class="event-item">
          <div class="item-title">${escapeHtml(event.event)}</div>
          <div class="item-meta">${escapeHtml(event.created_at)}</div>
          <div class="item-body">${escapeHtml(JSON.stringify(event.data))}</div>
        </div>
      `
    )
    .join("");
}

function renderCandidates() {
  const container = getElement("candidates");
  const filter = getSelect("candidate-filter").value as V2CandidateStatus | "all";
  const visible =
    filter === "all"
      ? currentCandidates
      : currentCandidates.filter((candidate) => candidate.status === filter);
  getElement("candidate-summary").textContent = formatCandidateSummary(currentCandidates);
  if (visible.length === 0) {
    renderEmpty(
      container,
      currentCandidates.length === 0 ? "运行 Agent 后显示候选问题。" : "当前筛选下没有候选问题。"
    );
    return;
  }
  container.className = "workspace-list";
  container.innerHTML = visible.map(renderCandidateCard).join("");
  visible.forEach((candidate) => {
    bindCandidateButton(candidate.candidate_id, "approved");
    bindCandidateButton(candidate.candidate_id, "rejected");
    bindCandidateButton(candidate.candidate_id, "deferred");
  });
}

function renderCandidateCard(candidate: V2CandidateIssue): string {
  return `
    <div class="candidate-item">
      <div class="candidate-header">
        <div class="item-title">${escapeHtml(candidate.category)} · ${translateStatus(candidate.status)}</div>
        <span class="severity ${candidate.severity}">${candidate.severity}</span>
      </div>
      <div class="item-body"><strong>原文：</strong>${escapeHtml(candidate.original || "-")}</div>
      <div class="item-body"><strong>替换：</strong>${escapeHtml(candidate.replacement || "需人工判断")}</div>
      <div class="item-body"><strong>建议：</strong>${escapeHtml(candidate.suggestion || "-")}</div>
      <div class="item-body"><strong>证据：</strong>${escapeHtml(candidate.evidence || "-")}</div>
      <div class="item-meta">${escapeHtml(candidate.self_check || "")}</div>
      <div class="candidate-actions">
        <button class="ms-Button" data-candidate-id="${escapeHtml(candidate.candidate_id)}" data-decision="approved" type="button">批准</button>
        <button class="ms-Button" data-candidate-id="${escapeHtml(candidate.candidate_id)}" data-decision="rejected" type="button">拒绝</button>
        <button class="ms-Button" data-candidate-id="${escapeHtml(candidate.candidate_id)}" data-decision="deferred" type="button">暂缓</button>
      </div>
    </div>
  `;
}

function bindCandidateButton(candidateId: string, status: "approved" | "rejected" | "deferred") {
  const selector = `[data-candidate-id="${candidateId}"][data-decision="${status}"]`;
  const button = getElement("candidates").querySelector(selector);
  if (button) {
    button.addEventListener("click", () => decideCandidate(candidateId, status));
  }
}

function renderReport() {
  const container = getElement("report");
  if (!currentReport) {
    renderEmpty(container, "暂无审校报告。");
    return;
  }
  container.className = "workspace-summary";
  container.innerHTML = `
    <div class="summary-grid">
      ${summaryItem("问题总数", String(currentReport.issue_count))}
      ${summaryItem("待确认", String(currentReport.pending_count))}
      ${summaryItem("已批准", String(currentReport.approved_count))}
      ${summaryItem("已拒绝", String(currentReport.rejected_count))}
      ${summaryItem("暂缓", String(currentReport.deferred_count))}
      ${summaryItem("已写回", String(currentReport.written_count))}
    </div>
    <div class="report-item">
      <div class="item-title">严重程度</div>
      <div class="item-body">${escapeHtml(JSON.stringify(currentReport.severity_counts))}</div>
    </div>
    <div class="report-item">
      <div class="item-title">类别分布</div>
      <div class="item-body">${escapeHtml(JSON.stringify(currentReport.category_counts))}</div>
    </div>
    <div class="report-item">
      <div class="item-title">未处理事项</div>
      <div class="item-body">${escapeHtml(currentReport.unresolved_items.join("；") || "无")}</div>
    </div>
  `;
}

function updateButtons() {
  const hasProject = Boolean(currentProject);
  const hasRun = Boolean(currentRun);
  const approvedCount = currentCandidates.filter(
    (candidate) => candidate.status === "approved"
  ).length;
  const pendingCount = currentCandidates.filter(
    (candidate) => candidate.status === "pending"
  ).length;
  getButton("create-project").disabled = isBusy;
  getButton("run-agent").disabled = isBusy || !hasProject;
  getButton("refresh-workspace").disabled = isBusy || !hasProject;
  getButton("refresh-trace").disabled = isBusy || !hasProject || !hasRun;
  getButton("approve-all").disabled = isBusy || pendingCount === 0;
  getButton("reject-all").disabled = isBusy || pendingCount === 0;
  getButton("writeback").disabled = isBusy || approvedCount === 0;
  getButton("refresh-report").disabled = isBusy || !hasProject;
  getButton("download-docx").disabled =
    isBusy ||
    !currentProject ||
    currentProject.source_type !== "docx" ||
    !currentProject.output_filename;
}

function getValidatedBookInfo(): BookInfo | null {
  const title = getInput("book-title").value.trim();
  if (!title) {
    showMessage("请先填写书名。", "error");
    return null;
  }
  const introduction = getTextArea("book-introduction").value.trim();
  return { title, introduction: introduction || null };
}

function getRunControls() {
  return {
    aiProfileId: getSelect("ai-profile").value || "default",
    providerApi: getSelect("provider-api").value as ProviderAPI,
    proofreadMode: getSelect("proofread-mode").value as "fast" | "thinking",
    reasoningEnabled: getInput("reasoning-enabled").checked,
    temperature: Number(getInput("temperature").value || DEFAULT_TEMPERATURE),
  };
}

function getApplicationMode(): ApplicationMode {
  return getSelect("application-mode").value as ApplicationMode;
}

function getSourceType(): SourceType {
  return getSelect("source-type").value as SourceType;
}

function getDocxFile(): File | null {
  return getInput("docx-file").files?.[0] || null;
}

function handleSourceTypeChange() {
  getElement("docx-source-panel").hidden = getSourceType() !== "docx";
}

function updateDocxFileSummary() {
  const file = getDocxFile();
  getElement("docx-file-summary").textContent = file
    ? `${file.name} · ${file.size} bytes`
    : "尚未选择 DOCX。";
}

function startBusy(): AbortController {
  currentAbortController?.abort();
  currentAbortController = new AbortController();
  isBusy = true;
  updateButtons();
  return currentAbortController;
}

function stopBusy() {
  isBusy = false;
  currentAbortController = null;
  updateButtons();
}

function resetProjectState() {
  currentProject = null;
  currentDocumentMap = null;
  currentRun = null;
  currentTrace = null;
  currentCandidates = [];
  currentReport = null;
}

function candidateToProofreadIssue(candidate: V2CandidateIssue): ProofreadIssue {
  return {
    id: candidate.candidate_id,
    category: candidate.category,
    severity: candidate.severity,
    original: candidate.original,
    replacement: candidate.replacement,
    suggestion: candidate.suggestion,
    start: candidate.global_start,
    end: candidate.global_end,
    locator: candidate.locator,
  };
}

function formatCandidateSummary(candidates: V2CandidateIssue[]): string {
  const count = (status: V2CandidateStatus) =>
    candidates.filter((candidate) => candidate.status === status).length;
  return `总数 ${candidates.length}，待确认 ${count("pending")}，已批准 ${count("approved")}，已拒绝 ${count("rejected")}，暂缓 ${count("deferred")}，已写回 ${count("written")}`;
}

function summaryItem(label: string, value: string): string {
  return `
    <div class="summary-item">
      <span class="summary-label">${escapeHtml(label)}</span>
      <span class="summary-value">${escapeHtml(value)}</span>
    </div>
  `;
}

function renderEmpty(element: HTMLElement, message: string) {
  element.className = "workspace-empty";
  element.textContent = message;
}

function translateStatus(status: string): string {
  const labels: Record<string, string> = {
    approved: "已批准",
    cancelled: "已取消",
    created: "已创建",
    deferred: "暂缓",
    docx: "DOCX",
    failed: "失败",
    partial_succeeded: "部分成功",
    pending: "待确认",
    rejected: "已拒绝",
    running: "运行中",
    selection: "当前选区",
    succeeded: "成功",
    waiting_for_approval: "等待确认",
    written: "已写回",
  };
  return labels[status] || status;
}

function showMessage(message: string, type: "default" | "error" | "success" = "default") {
  const element = getElement("message");
  element.textContent = message;
  element.className = type === "default" ? "message" : `message is-${type}`;
}

function getButton(id: string): HTMLButtonElement {
  return document.getElementById(id) as HTMLButtonElement;
}

function getElement(id: string): HTMLElement {
  return document.getElementById(id) as HTMLElement;
}

function getInput(id: string): HTMLInputElement {
  return document.getElementById(id) as HTMLInputElement;
}

function getSelect(id: string): HTMLSelectElement {
  return document.getElementById(id) as HTMLSelectElement;
}

function getTextArea(id: string): HTMLTextAreaElement {
  return document.getElementById(id) as HTMLTextAreaElement;
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

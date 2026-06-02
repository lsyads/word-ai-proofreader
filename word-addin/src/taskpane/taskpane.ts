/* global AbortController, AbortSignal, File, HTMLButtonElement, HTMLElement, HTMLInputElement, HTMLSelectElement, HTMLTextAreaElement, Office, document, localStorage, window */

import {
  createSession,
  createV2Project,
  createV2SelectionProject,
  decideV2Candidates,
  deleteV2Project,
  getAIProfiles,
  getV2Candidates,
  getV2DocumentMap,
  getV2DownloadUrl,
  getV2Memory,
  getV2Project,
  getV2ReviewPlan,
  getV2ReviewReport,
  getV2Run,
  getV2RunTrace,
  listV2Projects,
  markV2CandidatesWritten,
  runV2Project,
  writebackV2Project,
} from "./api";
import { appendDebugLog, clearDebugLog } from "./debug";
import {
  AIProfile,
  ApplicationMode,
  BookInfo,
  ProofreadIssue,
  ProviderAPI,
  V2ApprovalDecision,
  V2CandidateIssue,
  V2CandidateStatus,
  V2DocumentMap,
  V2MemoryItem,
  V2Project,
  V2ReviewPlan,
  V2ReviewReport,
  V2Run,
  V2RunTrace,
} from "./types";
import {
  applyIssuesToScope,
  clearTrackedSelectionRange,
  ensureWordCommentSupport,
  getSelectedText,
  selectIssueInScope,
} from "./word";

type SourceType = "selection" | "docx";

const DEFAULT_REVIEW_GOAL = "完成出版审校，找出明显错别字、语病、体例问题和上下文一致性风险。";
const DEFAULT_V2_TEMPERATURE = 0.6;
const PREFERRED_AI_PROFILE = "mimo-v2.5-pro";
const LAST_PROJECT_ID_KEY = "docpilot_v2_last_project_id";
const TERMINAL_RUN_STATUSES = new Set([
  "waiting_for_approval",
  "succeeded",
  "partial_succeeded",
  "failed",
  "cancelled",
]);
const INTERNAL_MEMORY_KEYS = new Set([
  "observed_categories",
  "observed_passes",
  "observed_severities",
]);

let currentSessionId: string | null = null;
let currentAbortController: AbortController | null = null;
let aiProfiles: AIProfile[] = [];
let recentProjects: V2Project[] = [];
let currentProject: V2Project | null = null;
let currentDocumentMap: V2DocumentMap | null = null;
let currentPlan: V2ReviewPlan | null = null;
let currentRun: V2Run | null = null;
let currentTrace: V2RunTrace | null = null;
let currentCandidates: V2CandidateIssue[] = [];
let currentMemory: V2MemoryItem[] = [];
let currentReport: V2ReviewReport | null = null;
let currentSelectionText = "";
let isBusy = false;
let confirmDialogResolver: ((confirmed: boolean) => void) | null = null;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    getButton("start-review").disabled = true;
    showMessage("请在 Microsoft Word 任务窗格中使用此插件。", "error");
    return;
  }

  bindEvents();
  initializeDefaults();
  initializeSession();
  initializeAIProfiles();
  initializeProjects();
  renderWorkspace();
});

function bindEvents() {
  getButton("start-review").onclick = startReview;
  getButton("refresh-projects").onclick = refreshProjects;
  getButton("refresh-trace").onclick = refreshTrace;
  getButton("approve-all").onclick = () => decidePendingCandidates("approved");
  getButton("reject-all").onclick = () => decidePendingCandidates("rejected");
  getButton("writeback").onclick = writebackApproved;
  getButton("download-docx").onclick = downloadDocx;
  getButton("clear-debug-log").onclick = clearDebugLog;
  getSelect("source-type").onchange = handleSourceTypeChange;
  getSelect("candidate-filter").onchange = renderCandidates;
  getSelect("candidate-pass-filter").onchange = renderCandidates;
  getInput("docx-file").onchange = updateDocxFileSummary;
  getButton("confirm-dialog-cancel").onclick = () => resolveConfirmDialog(false);
  getButton("confirm-dialog-submit").onclick = () => resolveConfirmDialog(true);
  getElement("confirm-dialog").addEventListener("click", (event) => {
    if (event.target === event.currentTarget) {
      resolveConfirmDialog(false);
    }
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !getElement("confirm-dialog").hidden) {
      resolveConfirmDialog(false);
    }
  });
}

function initializeDefaults() {
  getTextArea("review-goal").value = DEFAULT_REVIEW_GOAL;
  getSelect("proofread-mode").value = "thinking";
  getInput("temperature").value = String(DEFAULT_V2_TEMPERATURE);
  getSelect("application-mode").value = "revision";
  getInput("fallback-summary-truncate-enabled").checked = false;
  handleSourceTypeChange();
}

async function initializeSession() {
  try {
    const session = await createSession();
    currentSessionId = session.session_id;
    showMessage("已准备好审校会话。", "default");
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
    const preferred = findPreferredAIProfile(aiProfiles);
    if (preferred) {
      select.value = preferred.id;
      getSelect("provider-api").value = preferred.default_api;
    }
  } catch (error) {
    appendDebugLog("warn", "加载 AI 配置失败", { error: getErrorMessage(error) });
  }
}

function findPreferredAIProfile(profiles: AIProfile[]): AIProfile | undefined {
  const normalized = PREFERRED_AI_PROFILE.toLowerCase();
  return (
    profiles.find((profile) => profile.id.toLowerCase().includes(normalized)) ||
    profiles.find((profile) => profile.model.toLowerCase().includes(normalized)) ||
    profiles.find((profile) => profile.label.toLowerCase().includes(normalized)) ||
    profiles.find((profile) => profile.configured) ||
    profiles[0]
  );
}

async function initializeProjects() {
  const abortController = new AbortController();
  try {
    await loadRecentProjects(abortController.signal);
    const lastProjectId = localStorage.getItem(LAST_PROJECT_ID_KEY);
    if (lastProjectId && recentProjects.some((project) => project.project_id === lastProjectId)) {
      await loadProject(lastProjectId, abortController.signal);
      showMessage("已恢复最近审校项目。", "success");
    }
  } catch (error) {
    appendDebugLog("warn", "加载最近项目失败", { error: getErrorMessage(error) });
  } finally {
    renderWorkspace();
  }
}

async function startReview() {
  const abortController = startBusy();
  try {
    if (!currentProject) {
      await createProjectForCurrentInputs(abortController.signal);
    }
    if (!currentProject) {
      throw new Error("审校项目创建失败。");
    }
    await runCurrentProject(abortController.signal);
    showMessage("审校完成，请确认候选问题。", "success");
  } catch (error) {
    showMessage(`审校失败：${getErrorMessage(error)}`, "error");
    appendDebugLog("error", "V2.2 审校失败", { error: getErrorMessage(error) });
  } finally {
    stopBusy();
    renderWorkspace();
  }
}

async function createProjectForCurrentInputs(signal: AbortSignal) {
  const book = getValidatedBookInfo();
  if (!book) {
    throw new Error("请先填写书名。");
  }
  const reviewGoal = getTextArea("review-goal").value.trim() || DEFAULT_REVIEW_GOAL;
  const sourceType = getSourceType();
  resetProjectState();
  if (sourceType === "selection") {
    currentSelectionText = await getSelectedText();
    currentProject = await createV2SelectionProject(
      currentSelectionText,
      book,
      reviewGoal,
      currentSessionId,
      signal
    );
  } else {
    await clearTrackedSelectionRange();
    const file = getDocxFile();
    if (!file) {
      throw new Error("请先选择一个 .docx 文件。");
    }
    currentProject = await createV2Project(file, book, reviewGoal, signal);
    currentSelectionText = "";
  }
  currentDocumentMap = await getV2DocumentMap(currentProject.project_id, signal);
  currentPlan = await getV2ReviewPlan(currentProject.project_id, signal);
  localStorage.setItem(LAST_PROJECT_ID_KEY, currentProject.project_id);
  await loadRecentProjects(signal);
}

async function runCurrentProject(signal: AbortSignal) {
  if (!currentProject) {
    return;
  }
  const controls = getRunControls();
  currentRun = await runV2Project(
    currentProject.project_id,
    currentSessionId || "",
    controls.aiProfileId,
    controls.providerApi,
    controls.proofreadMode,
    controls.reasoningEnabled,
    controls.temperature,
    signal
  );
  showMessage("正在审校，请稍候。", "default");
  renderWorkspace();
  await pollRunUntilFinished(currentRun.run_id, signal);
  await refreshWorkspaceData(signal);
}

async function refreshProjects() {
  const abortController = startBusy();
  try {
    await loadRecentProjects(abortController.signal);
    showMessage("最近项目已刷新。", "success");
  } catch (error) {
    showMessage(`刷新项目失败：${getErrorMessage(error)}`, "error");
  } finally {
    stopBusy();
    renderWorkspace();
  }
}

async function openRecentProject(projectId: string) {
  const abortController = startBusy();
  try {
    await loadProject(projectId, abortController.signal);
    showMessage("已打开审校项目。", "success");
  } catch (error) {
    showMessage(`打开项目失败：${getErrorMessage(error)}`, "error");
  } finally {
    stopBusy();
    renderWorkspace();
  }
}

async function deleteProject(projectId: string) {
  const confirmed = await confirmAction(
    "确认删除项目",
    "删除后会清理该项目的候选问题、报告和生成的 DOCX 文件。这个操作不能撤销。"
  );
  if (!confirmed) {
    return;
  }
  const abortController = startBusy();
  try {
    await deleteV2Project(projectId, abortController.signal);
    if (currentProject?.project_id === projectId) {
      resetProjectState();
      localStorage.removeItem(LAST_PROJECT_ID_KEY);
      await clearTrackedSelectionRange();
    }
    await loadRecentProjects(abortController.signal);
    showMessage("项目已删除。", "success");
  } catch (error) {
    showMessage(`删除项目失败：${getErrorMessage(error)}`, "error");
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
  if (!currentRun && currentProject.latest_run_id) {
    currentRun = await getV2Run(currentProject.project_id, currentProject.latest_run_id, signal);
  } else if (currentRun) {
    currentRun = await getV2Run(currentProject.project_id, currentRun.run_id, signal);
  }
  currentDocumentMap = await getV2DocumentMap(currentProject.project_id, signal);
  currentPlan = await getV2ReviewPlan(currentProject.project_id, signal);
  currentCandidates = (await getV2Candidates(currentProject.project_id, signal)).candidates;
  if (currentRun) {
    currentTrace = await getV2RunTrace(currentProject.project_id, currentRun.run_id, signal);
  }
  currentMemory = (await getV2Memory(currentProject.project_id, signal)).memory;
  currentReport = await getV2ReviewReport(currentProject.project_id, signal);
  await loadRecentProjects(signal);
}

async function loadRecentProjects(signal: AbortSignal) {
  recentProjects = (await listV2Projects(signal)).projects;
}

async function loadProject(projectId: string, signal: AbortSignal) {
  resetProjectState();
  currentProject = await getV2Project(projectId, signal);
  if (currentProject.latest_run_id) {
    currentRun = await getV2Run(currentProject.project_id, currentProject.latest_run_id, signal);
  }
  await refreshWorkspaceData(signal);
  localStorage.setItem(LAST_PROJECT_ID_KEY, currentProject.project_id);
}

async function pollRunUntilFinished(runId: string, signal: AbortSignal) {
  if (!currentProject) {
    return;
  }
  for (let attempt = 0; attempt < 180; attempt += 1) {
    currentRun = await getV2Run(currentProject.project_id, runId, signal);
    currentTrace = await getV2RunTrace(currentProject.project_id, runId, signal);
    currentProject = await getV2Project(currentProject.project_id, signal);
    currentCandidates = (await getV2Candidates(currentProject.project_id, signal)).candidates;
    renderWorkspace();
    if (TERMINAL_RUN_STATUSES.has(currentRun.status)) {
      return;
    }
    await delay(1000);
  }
  throw new Error("审校仍在运行，请稍后刷新。");
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
    showMessage("运行记录已刷新。", "success");
  } catch (error) {
    showMessage(`刷新运行记录失败：${getErrorMessage(error)}`, "error");
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

async function decideCandidate(candidateId: string, status: "approved" | "rejected") {
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
    currentProject = await getV2Project(currentProject.project_id, abortController.signal);
    showMessage("候选问题已更新。", "success");
  } catch (error) {
    showMessage(`更新候选问题失败：${getErrorMessage(error)}`, "error");
  } finally {
    stopBusy();
    renderWorkspace();
  }
}

async function locateCandidate(candidateId: string) {
  if (!currentProject) {
    return;
  }
  if (currentProject.source_type !== "selection") {
    showMessage("DOCX 项目请写回后下载结果文件查看。", "default");
    return;
  }
  const candidate = currentCandidates.find((item) => item.candidate_id === candidateId);
  if (!candidate) {
    return;
  }
  if (!currentSelectionText) {
    showMessage("当前选区缓存已丢失，请重新开始当前选区审校后再定位。", "error");
    return;
  }
  startBusy();
  try {
    await selectIssueInScope(
      currentSelectionText,
      candidateToProofreadIssue(candidate),
      "selection"
    );
    showMessage("已定位到 Word 原文。", "success");
  } catch (error) {
    showMessage(`定位失败：${getErrorMessage(error)}`, "error");
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
    throw new Error("当前选区项目缺少缓存文本，请重新开始当前选区审校。");
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
  renderRecentProjects();
  renderProjectSummary();
  renderDocumentMap();
  renderReviewPlan();
  renderTrace();
  renderCandidates();
  renderMemory();
  renderReport();
  updateButtons();
}

function renderProjectBadge() {
  const badge = getElement("project-badge");
  const status = getDisplayStatus(
    currentRun?.status || currentProject?.status || "not_started",
    currentCandidates.length
  );
  badge.textContent = translateStatus(status);
  badge.className = "status-badge";
  if (status === "queued" || status === "running" || status === "waiting_for_approval") {
    badge.classList.add("is-running");
  } else if (status === "written" || status === "succeeded" || status === "no_candidates") {
    badge.classList.add("is-success");
  } else if (status === "failed" || status === "cancelled") {
    badge.classList.add("is-error");
  }
}

function renderProjectSummary() {
  const container = getElement("project-summary");
  if (!currentProject) {
    renderEmpty(container, "暂无审校项目。");
    return;
  }
  container.className = "compact-status";
  const approvedCount = currentCandidates.filter(
    (candidate) => candidate.status === "approved"
  ).length;
  const writtenCount = currentCandidates.filter(
    (candidate) => candidate.status === "written"
  ).length;
  container.innerHTML = `
    <span>${currentProject.source_type === "selection" ? "当前选区" : "全书 DOCX"}</span>
    <span>${escapeHtml(currentProject.book.title)}</span>
    <span>${translateStatus(
      getDisplayStatus(currentRun?.status || currentProject.status, currentCandidates.length)
    )}</span>
    <span>候选 ${currentCandidates.length}</span>
    <span>已批准 ${approvedCount}</span>
    <span>${writtenCount > 0 ? `已写回 ${writtenCount}` : "未写回"}</span>
  `;
}

function renderRecentProjects() {
  const container = getElement("recent-projects");
  if (recentProjects.length === 0) {
    renderEmpty(container, "暂无最近项目。");
    return;
  }
  container.className = "workspace-list";
  container.innerHTML = recentProjects
    .map(
      (project) => `
        <div class="project-item">
          <button class="project-open-button" data-open-project-id="${escapeHtml(project.project_id)}" type="button">
            <span class="item-title">${escapeHtml(project.book.title || project.source_filename)}</span>
            <span class="item-meta">${translateStatus(project.source_type)} · ${translateStatus(
              getDisplayStatus(project.status, project.candidate_count)
            )} · 候选 ${project.candidate_count}</span>
          </button>
          <button class="ms-Button danger-button" data-delete-project-id="${escapeHtml(project.project_id)}" type="button">删除</button>
        </div>
      `
    )
    .join("");
  recentProjects.forEach((project) => {
    const openButton = container.querySelector(`[data-open-project-id="${project.project_id}"]`);
    const deleteButton = container.querySelector(
      `[data-delete-project-id="${project.project_id}"]`
    );
    openButton?.addEventListener("click", () => openRecentProject(project.project_id));
    deleteButton?.addEventListener("click", () => deleteProject(project.project_id));
  });
}

function renderDocumentMap() {
  const container = getElement("document-map");
  if (!currentDocumentMap) {
    renderEmpty(container, "暂无文档地图。");
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
  if (!currentPlan) {
    renderEmpty(container, "暂无审校计划。");
    return;
  }
  container.className = "workspace-list";
  container.innerHTML = currentPlan.steps
    .map((step, index) => {
      const status = getPlanStepDisplayStatus(step.step_id, step.status);
      return `
        <div class="plan-item ${step.enabled ? "" : "is-disabled"}">
          <div class="item-title">${index + 1}. ${escapeHtml(step.title)} · ${translateStatus(status)}</div>
          <div class="item-meta">${escapeHtml(step.tool_name)} · ${step.enabled ? "已启用" : "未启用"}</div>
          <div class="item-body">${escapeHtml(step.description)}</div>
          <div class="item-meta">${escapeHtml(step.reason || "")}</div>
        </div>
      `;
    })
    .join("");
}

function getPlanStepDisplayStatus(stepId: string, fallback: string): string {
  if (!currentRun && !currentTrace) {
    return fallback;
  }
  if (currentRun?.stage === stepId) {
    return "running";
  }
  const events = currentTrace?.events || [];
  if (stepId === "plan_review" && events.some((event) => event.event === "plan_created")) {
    return "succeeded";
  }
  if (
    stepId === "human_approval" &&
    currentRun &&
    TERMINAL_RUN_STATUSES.has(currentRun.status) &&
    currentRun.status !== "failed" &&
    currentCandidates.length > 0
  ) {
    return "running";
  }
  if (
    events.some(
      (event) => event.event === "pass_completed" && String(event.data.pass_name || "") === stepId
    )
  ) {
    return "succeeded";
  }
  return fallback;
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
  updateCandidatePassFilterOptions();
  const passFilter = getSelect("candidate-pass-filter").value;
  const statusVisible =
    filter === "all"
      ? currentCandidates
      : currentCandidates.filter((candidate) => candidate.status === filter);
  const visible =
    passFilter === "all"
      ? statusVisible
      : statusVisible.filter((candidate) => candidate.pass_name === passFilter);
  getElement("candidate-summary").textContent = formatCandidateSummary(currentCandidates);
  if (isRunInProgress()) {
    renderEmpty(
      container,
      `正在审校：${translateStatus(currentRun?.stage || currentRun?.status || "running")}`
    );
    return;
  }
  if (visible.length === 0) {
    renderEmpty(
      container,
      currentCandidates.length === 0 && isCompletedWithoutCandidates()
        ? "审校完成，未发现需要确认的问题。"
        : currentCandidates.length === 0
          ? "开始审校后显示候选问题。"
          : "当前筛选下没有候选问题。"
    );
    return;
  }
  container.className = "workspace-list";
  container.innerHTML = visible.map(renderCandidateCard).join("");
  visible.forEach((candidate) => {
    bindCandidateButton(candidate.candidate_id, "approved");
    bindCandidateButton(candidate.candidate_id, "rejected");
    bindLocateButton(candidate.candidate_id);
  });
}

function renderCandidateCard(candidate: V2CandidateIssue): string {
  const showLocate = currentProject?.source_type === "selection";
  const locateHint =
    currentProject?.source_type === "docx" ? "DOCX 项目请写回后下载结果文件查看。" : "";
  return `
    <div class="candidate-item">
      <div class="candidate-header">
        <div class="item-title">${escapeHtml(candidate.category)} · ${translateStatus(candidate.status)}</div>
        <span class="severity ${candidate.severity}">${candidate.severity}</span>
      </div>
      <div class="item-body"><strong>原文：</strong>${escapeHtml(candidate.original || "-")}</div>
      <div class="item-body"><strong>建议：</strong>${escapeHtml(candidate.suggestion || "-")}</div>
      <div class="item-body"><strong>替换：</strong>${escapeHtml(candidate.replacement || "需人工判断")}</div>
      <div class="item-body"><strong>证据：</strong>${escapeHtml(candidate.evidence || "-")}</div>
      ${locateHint ? `<div class="item-meta">${locateHint}</div>` : ""}
      <details class="candidate-detail">
        <summary>查看详情</summary>
        <div class="item-meta">${escapeHtml(candidate.pass_name)} · 置信度 ${Math.round(candidate.confidence * 100)}% · ${escapeHtml(candidate.rule_id || candidate.evidence_kind)}</div>
        <div class="item-meta">${escapeHtml(candidate.evaluation_note || candidate.self_check || "")}</div>
      </details>
      <div class="candidate-actions">
        ${showLocate ? `<button class="ms-Button" data-locate-candidate-id="${escapeHtml(candidate.candidate_id)}" type="button">定位</button>` : ""}
        <button class="ms-Button" data-candidate-id="${escapeHtml(candidate.candidate_id)}" data-decision="approved" type="button">批准</button>
        <button class="ms-Button" data-candidate-id="${escapeHtml(candidate.candidate_id)}" data-decision="rejected" type="button">拒绝</button>
      </div>
    </div>
  `;
}

function bindCandidateButton(candidateId: string, status: "approved" | "rejected") {
  const selector = `[data-candidate-id="${candidateId}"][data-decision="${status}"]`;
  const button = getElement("candidates").querySelector(selector);
  button?.addEventListener("click", () => decideCandidate(candidateId, status));
}

function bindLocateButton(candidateId: string) {
  const button = getElement("candidates").querySelector(
    `[data-locate-candidate-id="${candidateId}"]`
  );
  button?.addEventListener("click", () => locateCandidate(candidateId));
}

function updateCandidatePassFilterOptions() {
  const select = getSelect("candidate-pass-filter");
  const currentValue = select.value || "all";
  const passNames = Array.from(
    new Set(currentCandidates.map((candidate) => candidate.pass_name))
  ).sort();
  select.innerHTML = [
    `<option value="all">全部阶段</option>`,
    ...passNames.map(
      (passName) => `<option value="${escapeHtml(passName)}">${escapeHtml(passName)}</option>`
    ),
  ].join("");
  select.value = passNames.includes(currentValue) ? currentValue : "all";
}

function renderMemory() {
  const container = getElement("project-memory");
  const visibleMemory = currentMemory.filter((item) => !INTERNAL_MEMORY_KEYS.has(item.key));
  if (!currentProject || visibleMemory.length === 0) {
    renderEmpty(container, "暂无本书规则。");
    return;
  }
  container.className = "workspace-list";
  container.innerHTML = visibleMemory
    .map(
      (item) => `
        <div class="memory-item">
          <div class="item-title">${translateMemoryKind(item.kind)} · ${escapeHtml(item.key)}</div>
          <div class="item-body">${escapeHtml(item.value)}</div>
          <div class="item-meta">${escapeHtml(item.source)} · 置信度 ${Math.round(item.confidence * 100)}%</div>
        </div>
      `
    )
    .join("");
}

function renderReport() {
  const container = getElement("report");
  if (!currentReport || currentReport.written_count === 0) {
    renderEmpty(container, "写回后显示简短报告。");
    return;
  }
  container.className = "workspace-summary compact-report";
  container.innerHTML = `
    <div class="summary-grid">
      ${summaryItem("已写回", String(currentReport.written_count))}
      ${summaryItem("待确认", String(currentReport.pending_count))}
      ${summaryItem("已拒绝", String(currentReport.rejected_count))}
      ${summaryItem("问题总数", String(currentReport.issue_count))}
    </div>
    <details class="details-panel">
      <summary>查看完整报告</summary>
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
    </details>
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
  getButton("start-review").disabled = isBusy;
  getButton("start-review").textContent = hasProject ? "重新审校" : "开始审校";
  getButton("refresh-projects").disabled = isBusy;
  getButton("refresh-trace").disabled = isBusy || !hasProject || !hasRun;
  getButton("approve-all").disabled = isBusy || pendingCount === 0;
  getButton("reject-all").disabled = isBusy || pendingCount === 0;
  getButton("writeback").disabled = isBusy || approvedCount === 0;
  getButton("download-docx").disabled =
    isBusy ||
    !currentProject ||
    currentProject.source_type !== "docx" ||
    !currentProject.output_filename;
}

function getValidatedBookInfo(): BookInfo | null {
  const title = getInput("book-title").value.trim();
  if (!title) {
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
    temperature: Number(getInput("temperature").value || DEFAULT_V2_TEMPERATURE),
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
  if (currentProject) {
    resetProjectState();
    renderWorkspace();
  }
}

function updateDocxFileSummary() {
  const file = getDocxFile();
  getElement("docx-file-summary").textContent = file
    ? `${file.name} · ${file.size} bytes`
    : "尚未选择 DOCX。";
}

function isRunInProgress(): boolean {
  return Boolean(currentRun && !TERMINAL_RUN_STATUSES.has(currentRun.status));
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
  currentPlan = null;
  currentRun = null;
  currentTrace = null;
  currentCandidates = [];
  currentMemory = [];
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
  if (candidates.length === 0 && isCompletedWithoutCandidates()) {
    return "总数 0，未发现需要确认的问题。";
  }
  const count = (status: V2CandidateStatus) =>
    candidates.filter((candidate) => candidate.status === status).length;
  return `总数 ${candidates.length}，待确认 ${count("pending")}，已批准 ${count("approved")}，已拒绝 ${count("rejected")}，已写回 ${count("written")}`;
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
    no_candidates: "未发现问题",
    not_started: "未开始",
    partial_succeeded: "部分成功",
    pending: "待确认",
    queued: "排队中",
    rejected: "已拒绝",
    running: "审校中",
    selection: "当前选区",
    succeeded: "成功",
    waiting_for_approval: "等待确认",
    written: "已写回",
  };
  return labels[status] || status;
}

function getDisplayStatus(status: string, candidateCount: number): string {
  if (
    candidateCount === 0 &&
    (status === "waiting_for_approval" || status === "succeeded" || status === "partial_succeeded")
  ) {
    return "no_candidates";
  }
  return status;
}

function isCompletedWithoutCandidates(): boolean {
  const status = currentRun?.status || currentProject?.status;
  return Boolean(status && getDisplayStatus(status, currentCandidates.length) === "no_candidates");
}

function translateMemoryKind(kind: string): string {
  const labels: Record<string, string> = {
    book_convention: "本书约定",
    observation: "审校记录",
    preference: "编辑偏好",
    style_rule: "体例规则",
    terminology: "术语",
  };
  return labels[kind] || kind;
}

function showMessage(message: string, type: "default" | "error" | "success" = "default") {
  const element = getElement("message");
  element.textContent = message;
  element.className = type === "default" ? "message" : `message is-${type}`;
}

function confirmAction(
  title: string,
  description: string,
  confirmLabel = "确认"
): Promise<boolean> {
  if (confirmDialogResolver) {
    confirmDialogResolver(false);
  }
  getElement("confirm-dialog-title").textContent = title;
  getElement("confirm-dialog-description").textContent = description;
  getButton("confirm-dialog-submit").textContent = confirmLabel;
  const dialog = getElement("confirm-dialog");
  dialog.hidden = false;
  getButton("confirm-dialog-submit").focus();
  return new Promise((resolve) => {
    confirmDialogResolver = resolve;
  });
}

function resolveConfirmDialog(confirmed: boolean) {
  const resolver = confirmDialogResolver;
  confirmDialogResolver = null;
  getElement("confirm-dialog").hidden = true;
  if (resolver) {
    resolver(confirmed);
  }
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

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => {
    window.setTimeout(resolve, ms);
  });
}

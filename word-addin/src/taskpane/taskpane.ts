/* global AbortController, AbortSignal, File, HTMLButtonElement, HTMLElement, HTMLInputElement, HTMLSelectElement, HTMLTextAreaElement, Office, document, window */

import {
  createSession,
  createV2Project,
  createV2SelectionProject,
  decideAllPendingV2Candidates,
  decideV2Candidates,
  deleteV2Project,
  downloadV2ProjectDocx,
  getAIProfiles,
  getV2Candidates,
  getV2DocumentMap,
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
type RunPollResult = "completed" | "timed_out" | "aborted";

const DEFAULT_REVIEW_GOAL = "完成出版审校，找出明显错别字、语病、体例问题和上下文一致性风险。";
const DEFAULT_V2_TEMPERATURE = 0.6;
const PREFERRED_AI_PROFILE = "mimo-v2.5-pro";
const RUN_POLL_INTERVAL_MS = 2000;
const RUN_POLL_TIMEOUT_MS = 30 * 60 * 1000;
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
let isDownloadingDocx = false;
let isRunPolling = false;
let runPollingTimedOut = false;
let runPollAbortController: AbortController | null = null;
let runPollingProjectId: string | null = null;
let runPollingRunId: string | null = null;
let candidatePage = 1;
const candidatePageSize = 20;
let candidateTotal = 0;
let candidateTotalPages = 0;
let candidateHasPrevious = false;
let candidateHasNext = false;
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
  getButton("refresh-current-project").onclick = refreshCurrentProject;
  getButton("continue-waiting").onclick = continueWaitingForRun;
  getButton("refresh-trace").onclick = refreshCurrentProject;
  getButton("approve-all").onclick = () => decidePendingCandidates("approved");
  getButton("reject-all").onclick = () => decidePendingCandidates("rejected");
  getButton("writeback").onclick = writebackApproved;
  getButton("download-docx").onclick = downloadDocx;
  getButton("clear-debug-log").onclick = clearDebugLog;
  getSelect("source-type").onchange = handleSourceTypeChange;
  getSelect("candidate-filter").onchange = handleCandidateFilterChange;
  getSelect("candidate-pass-filter").onchange = handleCandidateFilterChange;
  getButton("candidate-prev-page").onclick = () => changeCandidatePage(candidatePage - 1);
  getButton("candidate-next-page").onclick = () => changeCandidatePage(candidatePage + 1);
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
  } catch (error) {
    appendDebugLog("warn", "加载最近项目失败", { error: getErrorMessage(error) });
  } finally {
    renderWorkspace();
  }
}

async function startReview() {
  const abortController = startBusy();
  try {
    if (isRunInProgress()) {
      showMessage("本次审校仍在运行，请刷新进度查看最新结果。", "default");
      return;
    }
    if (!currentProject) {
      await createProjectForCurrentInputs(abortController.signal);
    }
    if (!currentProject) {
      throw new Error("本次审校创建失败。");
    }
    await runCurrentProject(abortController.signal);
  } catch (error) {
    showMessage(`启动审校失败：${getErrorMessage(error)}`, "error");
    appendDebugLog("error", "V2.2 审校启动失败", { error: getErrorMessage(error) });
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
  showMessage("已开始审校，完成后会显示可处理的建议。", "default");
  renderWorkspace();
  startRunPolling(currentProject.project_id, currentRun.run_id);
}

async function refreshProjects() {
  const abortController = startBusy();
  try {
    await loadRecentProjects(abortController.signal);
    showMessage("最近审校已刷新。", "success");
  } catch (error) {
    showMessage(`刷新最近审校失败：${getErrorMessage(error)}`, "error");
  } finally {
    stopBusy();
    renderWorkspace();
  }
}

async function openRecentProject(projectId: string) {
  const abortController = startBusy();
  try {
    await loadProject(projectId, abortController.signal);
    showMessage("已打开这次审校。", "success");
  } catch (error) {
    showMessage(`打开审校失败：${getErrorMessage(error)}`, "error");
  } finally {
    stopBusy();
    renderWorkspace();
  }
}

async function deleteProject(projectId: string) {
  if (currentProject?.project_id === projectId && isRunInProgress()) {
    showMessage("这次审校仍在运行，请等待完成后再删除。", "default");
    return;
  }
  const confirmed = await confirmAction(
    "确认删除这次审校",
    "删除后会清理本次建议、报告和生成的 DOCX 文件。这个操作不能撤销。",
    "删除"
  );
  if (!confirmed) {
    return;
  }
  const abortController = startBusy();
  try {
    await deleteV2Project(projectId, abortController.signal);
    if (currentProject?.project_id === projectId) {
      resetProjectState();
      await clearTrackedSelectionRange();
    }
    await loadRecentProjects(abortController.signal);
    showMessage("这次审校已删除。", "success");
  } catch (error) {
    showMessage(`删除审校失败：${getErrorMessage(error)}`, "error");
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
  await loadCandidates(signal);
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
}

async function pollRunUntilFinished(
  projectId: string,
  runId: string,
  signal: AbortSignal
): Promise<RunPollResult> {
  const startedAt = Date.now();
  while (Date.now() - startedAt < RUN_POLL_TIMEOUT_MS) {
    if (signal.aborted) {
      return "aborted";
    }
    try {
      const nextRun = await getV2Run(projectId, runId, signal);
      const nextTrace = await getV2RunTrace(projectId, runId, signal);
      const nextProject = await getV2Project(projectId, signal);
      if (currentProject?.project_id === projectId) {
        currentRun = nextRun;
        currentTrace = nextTrace;
        currentProject = nextProject;
        renderWorkspace();
      }
      if (TERMINAL_RUN_STATUSES.has(nextRun.status)) {
        if (currentProject?.project_id === projectId) {
          await refreshWorkspaceData(signal);
        }
        return "completed";
      }
    } catch (error) {
      if (signal.aborted || isAbortError(error)) {
        return "aborted";
      }
      appendDebugLog("warn", "自动刷新运行状态失败", {
        error: getErrorMessage(error),
        project_id: projectId,
        run_id: runId,
      });
    }
    await delay(RUN_POLL_INTERVAL_MS);
  }
  return "timed_out";
}

function startRunPolling(projectId: string, runId: string) {
  stopRunPolling();
  const abortController = new AbortController();
  runPollAbortController = abortController;
  runPollingProjectId = projectId;
  runPollingRunId = runId;
  isRunPolling = true;
  runPollingTimedOut = false;
  updateButtons();

  void pollRunUntilFinished(projectId, runId, abortController.signal)
    .then((result) => {
      if (runPollingProjectId !== projectId || runPollingRunId !== runId) {
        return;
      }
      runPollAbortController = null;
      isRunPolling = false;
      if (result === "completed") {
        clearRunPollingState();
        showRunCompletionMessage();
      } else if (result === "timed_out") {
        runPollingTimedOut = true;
        showMessage(
          "已等待 30 分钟，审校可能仍在后台运行。你可以稍后刷新进度查看结果，或点击继续等待。",
          "default"
        );
        appendDebugLog("warn", "前端轮询已到 30 分钟上限，后台任务未判定失败。", {
          project_id: projectId,
          run_id: runId,
          status: currentRun?.status,
          stage: currentRun?.stage,
        });
      }
      renderWorkspace();
    })
    .catch((error) => {
      if (isAbortError(error) || runPollingProjectId !== projectId || runPollingRunId !== runId) {
        return;
      }
      runPollAbortController = null;
      isRunPolling = false;
      showMessage("自动刷新中断，请刷新进度查看结果。", "default");
      appendDebugLog("warn", "自动刷新运行状态中断", {
        error: getErrorMessage(error),
        project_id: projectId,
        run_id: runId,
      });
      renderWorkspace();
    });
}

function stopRunPolling() {
  runPollAbortController?.abort();
  clearRunPollingState();
}

function clearRunPollingState() {
  runPollAbortController = null;
  isRunPolling = false;
  runPollingTimedOut = false;
  runPollingProjectId = null;
  runPollingRunId = null;
  updateButtons();
}

function showRunCompletionMessage() {
  if (!currentRun) {
    return;
  }
  if (currentRun.status === "failed") {
    showMessage(`审校失败：${currentRun.error_message || "请查看运行记录。"}`, "error");
    appendDebugLog("error", "V2.2 审校失败", {
      run_id: currentRun.run_id,
      error: currentRun.error_message || "unknown",
    });
  } else if (currentRun.status === "waiting_for_approval") {
    showMessage("审校完成，请查看并处理建议。", "success");
  } else if (isCompletedWithoutCandidates()) {
    showMessage("审校完成，未发现需要处理的建议。", "success");
  } else if (currentRun.status === "partial_succeeded") {
    showMessage("审校部分完成，请查看建议；排障信息里可查看运行记录。", "success");
  } else if (currentRun.status === "cancelled") {
    showMessage("审校已取消。", "default");
  } else {
    showMessage("审校完成。", "success");
  }
}

async function loadCandidates(signal: AbortSignal) {
  if (!currentProject) {
    currentCandidates = [];
    candidateTotal = 0;
    candidateTotalPages = 0;
    candidateHasPrevious = false;
    candidateHasNext = false;
    return;
  }
  const filter = getCandidateStatusFilter();
  const passFilter = getSelect("candidate-pass-filter").value || "all";
  const response = await getV2Candidates(currentProject.project_id, signal, {
    page: candidatePage,
    pageSize: candidatePageSize,
    status: filter,
    passName: passFilter,
  });
  currentCandidates = response.candidates;
  candidatePage = response.page;
  candidateTotal = response.total;
  candidateTotalPages = response.total_pages;
  candidateHasPrevious = response.has_previous;
  candidateHasNext = response.has_next;
}

async function loadAllCandidatesByStatus(
  status: V2CandidateStatus,
  signal: AbortSignal
): Promise<V2CandidateIssue[]> {
  if (!currentProject) {
    return [];
  }
  const allCandidates: V2CandidateIssue[] = [];
  let page = 1;
  while (true) {
    const response = await getV2Candidates(currentProject.project_id, signal, {
      page,
      pageSize: 100,
      status,
    });
    allCandidates.push(...response.candidates);
    if (!response.has_next) {
      return allCandidates;
    }
    page += 1;
  }
}

async function refreshCurrentProject() {
  if (!currentProject) {
    return;
  }
  const abortController = startBusy();
  try {
    await refreshWorkspaceData(abortController.signal);
    if (currentRun && TERMINAL_RUN_STATUSES.has(currentRun.status)) {
      stopRunPolling();
      showRunCompletionMessage();
    } else {
      showMessage("进度已刷新。", "success");
    }
  } catch (error) {
    showMessage(`刷新进度失败：${getErrorMessage(error)}`, "error");
  } finally {
    stopBusy();
    renderWorkspace();
  }
}

function continueWaitingForRun() {
  if (!currentProject || !currentRun || !isRunInProgress()) {
    return;
  }
  startRunPolling(currentProject.project_id, currentRun.run_id);
  showMessage("继续等待审校结果，完成后会自动刷新建议。", "default");
  renderWorkspace();
}

async function decidePendingCandidates(status: "approved" | "rejected") {
  if (!currentProject) {
    return;
  }
  const actionLabel = status === "approved" ? "接受" : "忽略";
  const confirmed = await confirmAction(
    `确认${actionLabel}全部待处理建议`,
    `这会${actionLabel}本次审校中所有待处理建议，不只处理当前页。写回前仍可在筛选中查看已${actionLabel}的建议。`,
    `${actionLabel}全部`
  );
  if (!confirmed) {
    return;
  }
  const abortController = startBusy();
  try {
    await decideAllPendingV2Candidates(currentProject.project_id, status, abortController.signal);
    candidatePage = 1;
    await refreshWorkspaceData(abortController.signal);
    showMessage(`所有待处理建议已${actionLabel}。`, "success");
  } catch (error) {
    showMessage(`批量处理建议失败：${getErrorMessage(error)}`, "error");
  } finally {
    stopBusy();
    renderWorkspace();
  }
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
    if (response.updated_count > 0 && currentCandidates.length === 1 && candidatePage > 1) {
      candidatePage -= 1;
    }
    await loadCandidates(abortController.signal);
    currentReport = await getV2ReviewReport(currentProject.project_id, abortController.signal);
    currentProject = await getV2Project(currentProject.project_id, abortController.signal);
    showMessage(formatDecisionMessage(decisions), "success");
  } catch (error) {
    showMessage(`处理建议失败：${getErrorMessage(error)}`, "error");
  } finally {
    stopBusy();
    renderWorkspace();
  }
}

async function locateCandidate(candidateId: string) {
  if (!currentProject) {
    return;
  }
  const candidate = currentCandidates.find((item) => item.candidate_id === candidateId);
  if (!candidate) {
    return;
  }
  if (!canLocateCandidate(candidate)) {
    showMessage("这条建议缺少可定位原文；DOCX 写回后可在审校后文件中查看。", "default");
    return;
  }
  if (currentProject.source_type === "selection" && !currentSelectionText) {
    showMessage("当前选区缓存已丢失，请重新开始当前选区审校后再定位。", "error");
    return;
  }
  startBusy();
  try {
    await selectIssueInScope(
      currentProject.source_type === "selection" ? currentSelectionText : "",
      candidateToProofreadIssue(candidate),
      currentProject.source_type === "selection" ? "selection" : "document"
    );
    showMessage("已定位到 Word 原文。", "success");
  } catch (error) {
    const suffix =
      currentProject.source_type === "docx" ? "。可写回后下载审校后文件查看批注位置。" : "";
    showMessage(`定位失败：${getErrorMessage(error)}${suffix}`, "error");
  } finally {
    stopBusy();
    renderWorkspace();
  }
}

async function writebackApproved() {
  if (!currentProject) {
    return;
  }
  const abortController = startBusy();
  try {
    const approved = await loadAllCandidatesByStatus("approved", abortController.signal);
    if (approved.length === 0) {
      showMessage("没有已接受的建议可写回。请先接受需要写入 Word 的建议。", "error");
      return;
    }
    const isDocxProject = currentProject.source_type === "docx";
    if (currentProject.source_type === "selection") {
      await writebackSelection(approved, abortController.signal);
    } else {
      await writebackDocx(abortController.signal);
    }
    await refreshWorkspaceData(abortController.signal);
    showMessage(
      isDocxProject ? "已生成审校后文件，可点击下载。" : "已将接受的建议写回 Word。",
      "success"
    );
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
    throw new Error("当前选区缓存已丢失，请重新开始当前选区审校。");
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

async function downloadDocx() {
  if (!currentProject || currentProject.source_type !== "docx" || !currentProject.output_filename) {
    return;
  }
  if (isDownloadingDocx) {
    return;
  }
  const abortController = startBusy();
  isDownloadingDocx = true;
  updateButtons();
  try {
    await downloadV2ProjectDocx(
      currentProject.project_id,
      currentProject.output_filename,
      abortController.signal
    );
    showMessage(`已开始下载审校后文件：${currentProject.output_filename}`, "success");
  } catch (error) {
    showMessage(`下载失败：${getErrorMessage(error)}`, "error");
  } finally {
    isDownloadingDocx = false;
    stopBusy();
    renderWorkspace();
  }
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
    currentProject?.candidate_count || candidateTotal
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
    renderEmpty(container, "暂无本次审校。");
    return;
  }
  container.className = "compact-status";
  const approvedCount = currentProject.approved_count;
  const writtenCount = currentReport?.written_count || 0;
  container.innerHTML = `
    <span>${currentProject.source_type === "selection" ? "当前选区" : "全书 DOCX"}</span>
    <span>${escapeHtml(currentProject.book.title)}</span>
    <span>${translateStatus(
      getDisplayStatus(currentRun?.status || currentProject.status, currentProject.candidate_count)
    )}</span>
    <span>建议 ${currentProject.candidate_count}</span>
    <span>已接受 ${approvedCount}</span>
    <span>${writtenCount > 0 ? `已写入 ${writtenCount}` : "未写入"}</span>
  `;
}

function renderRecentProjects() {
  const container = getElement("recent-projects");
  if (recentProjects.length === 0) {
    renderEmpty(container, "暂无最近审校。");
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
            )} · 建议 ${project.candidate_count}</span>
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
    (currentProject?.candidate_count || candidateTotal) > 0
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
  updateCandidatePassFilterOptions();
  renderCandidatePagination();
  getElement("candidate-summary").textContent = formatCandidateSummary();
  if (runPollingTimedOut && isRunInProgress()) {
    renderEmpty(
      container,
      "审校时间较长，已停止自动刷新。你可以刷新进度查看结果，或点击继续等待。"
    );
    return;
  }
  if (isRunInProgress()) {
    renderEmpty(
      container,
      `仍在审校：${translateStage(currentRun?.stage || currentRun?.status || "running")}`
    );
    return;
  }
  if (currentCandidates.length === 0) {
    renderEmpty(
      container,
      candidateTotal === 0 && isCompletedWithoutCandidates()
        ? "审校完成，未发现需要处理的建议。"
        : candidateTotal === 0
          ? "开始审校后显示建议。"
          : "当前筛选下没有建议。"
    );
    return;
  }
  container.className = "workspace-list";
  container.innerHTML = currentCandidates.map(renderCandidateCard).join("");
  currentCandidates.forEach((candidate) => {
    bindCandidateButton(candidate.candidate_id, "approved");
    bindCandidateButton(candidate.candidate_id, "rejected");
    bindLocateButton(candidate.candidate_id);
  });
}

function renderCandidateCard(candidate: V2CandidateIssue): string {
  const canLocate = canLocateCandidate(candidate);
  const locateHint = canLocate
    ? ""
    : currentProject?.source_type === "docx"
      ? "缺少可定位原文；写回后可在审校后文件中查看。"
      : "缺少可定位原文。";
  const replacement = candidate.replacement || "不直接替换，请按修改说明判断";
  const riskHint = candidate.needs_human_review
    ? "需要人工重点判断。"
    : "置信度较高，仍需接受后才会写回。";
  return `
    <div class="candidate-item">
      <div class="candidate-header">
        <div class="item-title">${escapeHtml(translateCategory(candidate.category))} · ${translateStatus(candidate.status)}</div>
        <span class="severity ${candidate.severity}">${translateSeverity(candidate.severity)}</span>
      </div>
      <div class="item-body"><strong>原文：</strong>${escapeHtml(candidate.original || "-")}</div>
      <div class="item-body"><strong>建议改为：</strong>${escapeHtml(replacement)}</div>
      <div class="item-body"><strong>修改说明：</strong>${escapeHtml(candidate.suggestion || "-")}</div>
      <div class="item-body"><strong>依据：</strong>${escapeHtml(candidate.evidence || "-")}</div>
      <div class="item-body"><strong>风险提示：</strong>${escapeHtml(riskHint)}</div>
      ${locateHint ? `<div class="item-meta">${locateHint}</div>` : ""}
      <details class="candidate-detail">
        <summary>查看技术详情</summary>
        <div class="item-meta">${escapeHtml(translatePassName(candidate.pass_name))} · 置信度 ${Math.round(candidate.confidence * 100)}% · ${escapeHtml(candidate.rule_id || candidate.evidence_kind)}</div>
        <div class="item-meta">${escapeHtml(candidate.evaluation_note || candidate.self_check || "")}</div>
      </details>
      <div class="candidate-actions">
        <button class="ms-Button" data-locate-candidate-id="${escapeHtml(candidate.candidate_id)}" type="button" ${canLocate ? "" : "disabled"}>定位原文</button>
        <button class="ms-Button" data-candidate-id="${escapeHtml(candidate.candidate_id)}" data-decision="approved" type="button">接受</button>
        <button class="ms-Button" data-candidate-id="${escapeHtml(candidate.candidate_id)}" data-decision="rejected" type="button">忽略</button>
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
  const optionPassNames =
    currentValue !== "all" && !passNames.includes(currentValue)
      ? [currentValue, ...passNames]
      : passNames;
  select.hidden = optionPassNames.length <= 1 && currentValue === "all";
  select.innerHTML = [
    `<option value="all">全部来源</option>`,
    ...optionPassNames.map(
      (passName) =>
        `<option value="${escapeHtml(passName)}">${escapeHtml(translatePassName(passName))}</option>`
    ),
  ].join("");
  select.value = currentValue;
}

function renderCandidatePagination() {
  getElement("candidate-page-summary").textContent =
    candidateTotalPages > 0
      ? `第 ${candidatePage}/${candidateTotalPages} 页 · 当前 ${currentCandidates.length} 条`
      : "第 0/0 页";
  getButton("candidate-prev-page").disabled = isBusy || !candidateHasPrevious;
  getButton("candidate-next-page").disabled = isBusy || !candidateHasNext;
}

async function handleCandidateFilterChange() {
  candidatePage = 1;
  if (!currentProject) {
    renderWorkspace();
    return;
  }
  const abortController = startBusy();
  try {
    await loadCandidates(abortController.signal);
  } catch (error) {
    showMessage(`刷新建议失败：${getErrorMessage(error)}`, "error");
  } finally {
    stopBusy();
    renderWorkspace();
  }
}

async function changeCandidatePage(nextPage: number) {
  if (
    !currentProject ||
    nextPage < 1 ||
    (candidateTotalPages > 0 && nextPage > candidateTotalPages)
  ) {
    return;
  }
  candidatePage = nextPage;
  const abortController = startBusy();
  try {
    await loadCandidates(abortController.signal);
  } catch (error) {
    showMessage(`切换建议分页失败：${getErrorMessage(error)}`, "error");
  } finally {
    stopBusy();
    renderWorkspace();
  }
}

function getCandidateStatusFilter(): V2CandidateStatus | "all" {
  const value = getSelect("candidate-filter").value;
  return isCandidateStatus(value) ? value : "all";
}

function canLocateCandidate(candidate: V2CandidateIssue): boolean {
  return Boolean(
    candidate.original ||
    candidate.locator?.key ||
    (typeof candidate.global_start === "number" && typeof candidate.global_end === "number")
  );
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
      ${summaryItem("已写入", String(currentReport.written_count))}
      ${summaryItem("待处理", String(currentReport.pending_count))}
      ${summaryItem("已忽略", String(currentReport.rejected_count))}
      ${summaryItem("建议总数", String(currentReport.issue_count))}
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
  const runInProgress = isRunInProgress();
  const approvedCount = currentProject?.approved_count || 0;
  const pendingCount = currentProject?.pending_count || 0;
  getButton("start-review").disabled = isBusy || runInProgress;
  setButtonLabel(
    "start-review",
    runInProgress ? "审校进行中" : hasProject ? "重新审校" : "开始审校"
  );
  getButton("refresh-projects").disabled = isBusy;
  getButton("refresh-current-project").disabled = isBusy || !hasProject;
  getButton("refresh-trace").disabled = isBusy || !hasProject;
  getButton("continue-waiting").hidden = !runInProgress || isRunPolling;
  getButton("continue-waiting").disabled = isBusy || !hasProject || !hasRun || isRunPolling;
  getButton("approve-all").disabled = isBusy || pendingCount === 0;
  getButton("reject-all").disabled = isBusy || pendingCount === 0;
  getButton("writeback").disabled = isBusy || approvedCount === 0;
  setButtonLabel(
    "writeback",
    approvedCount > 0 ? `写回 ${approvedCount} 条已接受建议` : "先接受建议后写回"
  );
  getButton("download-docx").disabled =
    isBusy ||
    isDownloadingDocx ||
    !currentProject ||
    currentProject.source_type !== "docx" ||
    !currentProject.output_filename;
  setButtonLabel("download-docx", isDownloadingDocx ? "下载中" : "下载审校后文件");
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
    ? `${file.name} · ${formatBytes(file.size)}`
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
  stopRunPolling();
  currentProject = null;
  currentDocumentMap = null;
  currentPlan = null;
  currentRun = null;
  currentTrace = null;
  currentCandidates = [];
  currentMemory = [];
  currentReport = null;
  candidatePage = 1;
  candidateTotal = 0;
  candidateTotalPages = 0;
  candidateHasPrevious = false;
  candidateHasNext = false;
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

function formatCandidateSummary(): string {
  if ((currentProject?.candidate_count || 0) === 0 && isCompletedWithoutCandidates()) {
    return "共 0 条建议，未发现需要处理的内容。";
  }
  const rejectedCount = currentReport?.rejected_count ?? "-";
  const writtenCount = currentReport?.written_count ?? "-";
  return `共 ${currentProject?.candidate_count || 0} 条建议，待处理 ${currentProject?.pending_count || 0}，已接受 ${currentProject?.approved_count || 0}，已忽略 ${rejectedCount}，已写入 ${writtenCount}；当前筛选 ${candidateTotal} 条。`;
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
    approved: "已接受",
    cancelled: "已取消",
    created: "已创建",
    deferred: "稍后处理",
    docx: "DOCX",
    failed: "失败",
    no_candidates: "未发现建议",
    not_started: "未开始",
    partial_succeeded: "部分完成",
    pending: "待处理",
    queued: "排队中",
    rejected: "已忽略",
    running: "审校中",
    selection: "当前选区",
    succeeded: "已完成",
    waiting_for_approval: "等待处理",
    written: "已写入",
  };
  return labels[status] || status;
}

function translateStage(stage: string): string {
  const labels: Record<string, string> = {
    consistency_pass: "正在检查前后一致性",
    evaluate_candidates: "正在整理建议",
    failed: "审校失败",
    merge_candidates: "正在合并重复建议",
    partial_succeeded: "部分完成",
    plan_review: "正在准备审校",
    proofread_pass: "正在检查正文",
    queued: "排队中",
    running: "审校中",
    style_rule_pass: "正在检查体例",
    succeeded: "已完成",
    terminology_pass: "正在检查术语",
    waiting_for_approval: "等待处理建议",
  };
  return labels[stage] || translateStatus(stage);
}

function translateSeverity(severity: string): string {
  const labels: Record<string, string> = {
    high: "高",
    low: "低",
    medium: "中",
  };
  return labels[severity] || severity;
}

function translateCategory(category: string): string {
  const labels: Record<string, string> = {
    fact: "事实核查",
    grammar: "语句问题",
    punctuation: "标点体例",
    style: "体例建议",
    terminology: "术语问题",
    typo: "错别字",
  };
  return labels[category] || category;
}

function translatePassName(passName: string): string {
  const labels: Record<string, string> = {
    consistency_pass: "前后一致性",
    proofread_pass: "基础审校",
    style_rule_pass: "体例规则",
    terminology_pass: "术语检查",
  };
  return labels[passName] || passName;
}

function formatDecisionMessage(decisions: V2ApprovalDecision[]): string {
  const statuses = Array.from(new Set(decisions.map((decision) => decision.status)));
  if (statuses.length === 1) {
    if (statuses[0] === "approved") {
      return "建议已接受。";
    }
    if (statuses[0] === "rejected") {
      return "建议已忽略。";
    }
  }
  return "建议已更新。";
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
  return Boolean(
    status &&
    getDisplayStatus(status, currentProject?.candidate_count || candidateTotal) === "no_candidates"
  );
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

function isCandidateStatus(value: string): value is V2CandidateStatus {
  return (
    value === "pending" ||
    value === "approved" ||
    value === "rejected" ||
    value === "deferred" ||
    value === "written"
  );
}

function showMessage(message: string, type: "default" | "error" | "success" = "default") {
  const element = getElement("message");
  element.textContent = message;
  element.className = type === "default" ? "message" : `message is-${type}`;
}

function formatBytes(bytes: number): string {
  if (bytes >= 1024 * 1024) {
    return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  }
  if (bytes >= 1024) {
    return `${Math.round(bytes / 1024)} KB`;
  }
  return `${bytes} bytes`;
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

function setButtonLabel(id: string, label: string) {
  const button = getButton(id);
  const labelElement = button.querySelector(".ms-Button-label");
  if (labelElement) {
    labelElement.textContent = label;
  } else {
    button.textContent = label;
  }
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

function isAbortError(error: unknown): boolean {
  return error instanceof Error && error.name === "AbortError";
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

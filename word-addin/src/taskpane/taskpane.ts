/* global document, Office, Word, HTMLElement, HTMLButtonElement, fetch */

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

const API_BASE_URL = "";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    getButton("proofread").onclick = proofreadSelection;
    return;
  }

  getButton("proofread").disabled = true;
  showMessage("请在 Microsoft Word 任务窗格中使用此插件。", "error");
});

export async function proofreadSelection() {
  if (!Office.context.requirements.isSetSupported("WordApi", "1.4")) {
    showMessage("当前 Word 环境不支持批注 API，无法完成审校。", "error");
    return;
  }

  setBusy(true);
  showMessage("正在审校当前选区...", "default");
  renderEmptyResult("审校中...");

  try {
    const selectedText = await getSelectedText();
    const proofreadResult = await requestProofread(selectedText);
    const commentText = formatComment(proofreadResult.issues);

    await insertCommentToSelection(commentText);
    renderResult(proofreadResult.issues, commentText);

    showMessage("审校完成，已在当前选区插入批注。", "success");
  } catch (error) {
    renderEmptyResult("审校失败");
    showMessage(`审校失败：${getErrorMessage(error)}`, "error");
  } finally {
    setBusy(false);
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

async function requestProofread(text: string): Promise<ProofreadResponse> {
  const response = await fetch(`${API_BASE_URL}/api/proofread`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      text,
      context: {
        source: "word-addin",
      },
    }),
  });

  if (!response.ok) {
    throw new Error(`后端返回 HTTP ${response.status}`);
  }

  return (await response.json()) as ProofreadResponse;
}

function formatComment(issues: ProofreadIssue[]): string {
  if (issues.length === 0) {
    return "AI 审校：未发现明显问题。";
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

function setBusy(isBusy: boolean) {
  const button = getButton("proofread");
  button.disabled = isBusy;
  button.querySelector(".ms-Button-label").textContent = isBusy ? "审校中..." : "AI 审校";
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

function escapeHtml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

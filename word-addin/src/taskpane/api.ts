/* global AbortSignal, Blob, File, Response, URL, URLSearchParams, document, fetch, setTimeout */

import {
  AIProfile,
  ApplicationMode,
  BookInfo,
  ProofreadMode,
  ProviderAPI,
  SessionResponse,
  V2ApprovalDecision,
  V2ApprovalDecisionResponse,
  V2CandidateList,
  V2DocumentMap,
  V2MemoryList,
  V2MarkWrittenResponse,
  V2Project,
  V2ProjectDeleteResponse,
  V2ProjectList,
  V2ReviewPlan,
  V2ReviewReport,
  V2Run,
  V2RunTrace,
  V2WritebackResponse,
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

export async function createV2Project(
  file: File,
  book: BookInfo,
  signal: AbortSignal
): Promise<V2Project> {
  const params = new URLSearchParams({
    filename: file.name,
    book: JSON.stringify(book),
  });
  const response = await fetch(`${API_BASE_URL}/api/v2/projects?${params.toString()}`, {
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

  return (await response.json()) as V2Project;
}

export async function createV2SelectionProject(
  text: string,
  book: BookInfo,
  sessionId: string | null,
  signal: AbortSignal
): Promise<V2Project> {
  const response = await fetch(`${API_BASE_URL}/api/v2/projects/selection`, {
    method: "POST",
    signal,
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      text,
      book,
      session_id: sessionId,
    }),
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as V2Project;
}

export async function listV2Projects(signal: AbortSignal): Promise<V2ProjectList> {
  const response = await fetch(`${API_BASE_URL}/api/v2/projects`, {
    method: "GET",
    signal,
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as V2ProjectList;
}

export async function deleteV2Project(
  projectId: string,
  signal: AbortSignal
): Promise<V2ProjectDeleteResponse> {
  const response = await fetch(`${API_BASE_URL}/api/v2/projects/${encodeURIComponent(projectId)}`, {
    method: "DELETE",
    signal,
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as V2ProjectDeleteResponse;
}

export async function getV2Project(projectId: string, signal: AbortSignal): Promise<V2Project> {
  const response = await fetch(`${API_BASE_URL}/api/v2/projects/${encodeURIComponent(projectId)}`, {
    method: "GET",
    signal,
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as V2Project;
}

export async function getV2ReviewPlan(
  projectId: string,
  signal: AbortSignal
): Promise<V2ReviewPlan> {
  const response = await fetch(
    `${API_BASE_URL}/api/v2/projects/${encodeURIComponent(projectId)}/plan`,
    {
      method: "GET",
      signal,
    }
  );

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as V2ReviewPlan;
}

export async function getV2DocumentMap(
  projectId: string,
  signal: AbortSignal
): Promise<V2DocumentMap> {
  const response = await fetch(
    `${API_BASE_URL}/api/v2/projects/${encodeURIComponent(projectId)}/document-map`,
    {
      method: "GET",
      signal,
    }
  );

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as V2DocumentMap;
}

export async function runV2Project(
  projectId: string,
  sessionId: string,
  aiProfileId: string,
  providerApi: ProviderAPI,
  proofreadMode: ProofreadMode,
  reasoningEnabled: boolean,
  temperature: number,
  signal: AbortSignal
): Promise<V2Run> {
  const response = await fetch(
    `${API_BASE_URL}/api/v2/projects/${encodeURIComponent(projectId)}/runs`,
    {
      method: "POST",
      signal,
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        session_id: sessionId,
        ai_profile_id: aiProfileId,
        provider_api: providerApi,
        proofread_mode: proofreadMode,
        reasoning_enabled: reasoningEnabled,
        temperature,
      }),
    }
  );

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as V2Run;
}

export async function getV2Run(
  projectId: string,
  runId: string,
  signal: AbortSignal
): Promise<V2Run> {
  const response = await fetch(
    `${API_BASE_URL}/api/v2/projects/${encodeURIComponent(projectId)}/runs/${encodeURIComponent(runId)}`,
    {
      method: "GET",
      signal,
    }
  );

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as V2Run;
}

export async function getV2RunTrace(
  projectId: string,
  runId: string,
  signal: AbortSignal
): Promise<V2RunTrace> {
  const response = await fetch(
    `${API_BASE_URL}/api/v2/projects/${encodeURIComponent(projectId)}/runs/${encodeURIComponent(runId)}/trace`,
    {
      method: "GET",
      signal,
    }
  );

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as V2RunTrace;
}

export async function getV2Memory(projectId: string, signal: AbortSignal): Promise<V2MemoryList> {
  const response = await fetch(
    `${API_BASE_URL}/api/v2/projects/${encodeURIComponent(projectId)}/memory`,
    {
      method: "GET",
      signal,
    }
  );

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as V2MemoryList;
}

export async function deleteV2Memory(
  projectId: string,
  memoryId: string,
  signal: AbortSignal
): Promise<V2MemoryList> {
  const response = await fetch(
    `${API_BASE_URL}/api/v2/projects/${encodeURIComponent(projectId)}/memory/${encodeURIComponent(memoryId)}`,
    {
      method: "DELETE",
      signal,
    }
  );

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as V2MemoryList;
}

export async function getV2Candidates(
  projectId: string,
  signal: AbortSignal,
  options: {
    page?: number;
    pageSize?: number;
    status?: string;
    passName?: string;
  } = {}
): Promise<V2CandidateList> {
  const params = new URLSearchParams({
    page: String(options.page || 1),
    page_size: String(options.pageSize || 20),
  });
  if (options.status && options.status !== "all") {
    params.set("status", options.status);
  }
  if (options.passName && options.passName !== "all") {
    params.set("pass_name", options.passName);
  }
  const response = await fetch(
    `${API_BASE_URL}/api/v2/projects/${encodeURIComponent(projectId)}/candidates?${params.toString()}`,
    {
      method: "GET",
      signal,
    }
  );

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as V2CandidateList;
}

export async function decideAllPendingV2Candidates(
  projectId: string,
  status: "approved" | "rejected" | "deferred",
  signal: AbortSignal
): Promise<V2ApprovalDecisionResponse> {
  const response = await fetch(
    `${API_BASE_URL}/api/v2/projects/${encodeURIComponent(projectId)}/candidates/bulk-decisions`,
    {
      method: "POST",
      signal,
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ status }),
    }
  );

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as V2ApprovalDecisionResponse;
}

export async function decideV2Candidates(
  projectId: string,
  decisions: V2ApprovalDecision[],
  signal: AbortSignal
): Promise<V2ApprovalDecisionResponse> {
  const response = await fetch(
    `${API_BASE_URL}/api/v2/projects/${encodeURIComponent(projectId)}/candidates/decisions`,
    {
      method: "POST",
      signal,
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ decisions }),
    }
  );

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as V2ApprovalDecisionResponse;
}

export async function markV2CandidatesWritten(
  projectId: string,
  candidateIds: string[],
  signal: AbortSignal
): Promise<V2MarkWrittenResponse> {
  const response = await fetch(
    `${API_BASE_URL}/api/v2/projects/${encodeURIComponent(projectId)}/candidates/mark-written`,
    {
      method: "POST",
      signal,
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ candidate_ids: candidateIds }),
    }
  );

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as V2MarkWrittenResponse;
}

export async function writebackV2Project(
  projectId: string,
  applicationMode: ApplicationMode,
  fallbackSummaryTruncateEnabled: boolean,
  signal: AbortSignal
): Promise<V2WritebackResponse> {
  const response = await fetch(
    `${API_BASE_URL}/api/v2/projects/${encodeURIComponent(projectId)}/writeback`,
    {
      method: "POST",
      signal,
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        application_mode: applicationMode,
        fallback_summary_truncate_enabled: fallbackSummaryTruncateEnabled,
      }),
    }
  );

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as V2WritebackResponse;
}

export async function getV2ReviewReport(
  projectId: string,
  signal: AbortSignal
): Promise<V2ReviewReport> {
  const response = await fetch(
    `${API_BASE_URL}/api/v2/projects/${encodeURIComponent(projectId)}/report`,
    {
      method: "GET",
      signal,
    }
  );

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  return (await response.json()) as V2ReviewReport;
}

export function getV2DownloadUrl(projectId: string): string {
  return `${API_BASE_URL}/api/v2/projects/${encodeURIComponent(projectId)}/download`;
}

export async function downloadV2ProjectDocx(
  projectId: string,
  filename: string,
  signal?: AbortSignal
): Promise<void> {
  const response = await fetch(getV2DownloadUrl(projectId), {
    method: "GET",
    signal,
  });

  if (!response.ok) {
    throw new Error(await getResponseErrorMessage(response));
  }

  const blob = await response.blob();
  triggerBlobDownload(blob, filename);
}

function triggerBlobDownload(blob: Blob, filename: string) {
  const objectUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = objectUrl;
  link.download = filename;
  link.style.display = "none";
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
}

async function getResponseErrorMessage(response: Response): Promise<string> {
  try {
    const payload = await response.json();
    if (typeof payload.detail === "string") {
      return payload.detail;
    }
    return JSON.stringify(payload.detail || payload);
  } catch {
    return response.statusText || `HTTP ${response.status}`;
  }
}

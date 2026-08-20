import type {
  ConversationMessage,
  DocumentConflict,
  DocumentSource,
  DocumentStructurePlan,
  EditOperation,
  EventEnvelope,
  ExportJob,
  Project,
  ProjectHistory,
  User,
  WorkflowMode,
  WorkSession,
} from "../shared/models.js";

const API_BASE = import.meta.env.VITE_API_BASE || "";
let memoryToken: string | null = null;

export interface ApiError extends Error {
  status?: number;
}

export function getToken(): string | null {
  return memoryToken;
}

export function setToken(token: string): void {
  memoryToken = token;
}

export function clearToken(): void {
  memoryToken = null;
}

async function request<T>(path: string, init: RequestInit = {}): Promise<T> {
  const headers = new Headers(init.headers);
  if (init.body && !(init.body instanceof FormData))
    headers.set("Content-Type", "application/json");
  const token = getToken();
  if (token) headers.set("Authorization", `Bearer ${token}`);
  const response = await fetch(`${API_BASE}${path}`, {
    ...init,
    headers,
    credentials: "include",
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(
      data?.error?.message || `Request failed with ${response.status}`,
    ) as ApiError;
    error.status = response.status;
    throw error;
  }
  return data as T;
}

export async function login(
  email: string,
  name: string,
  accessCode?: string,
): Promise<{ token: string; user: User }> {
  const result = await request<{ token: string; user: User }>(
    "/api/v1/auth/login",
    { method: "POST", body: JSON.stringify({ email, name, accessCode }) },
  );
  setToken(result.token);
  return result;
}

export async function me(): Promise<User | null> {
  return (await request<{ user: User | null }>("/api/v1/auth/me")).user;
}

export async function logout(): Promise<void> {
  await request("/api/v1/auth/logout", { method: "POST" });
  clearToken();
}

export async function listProjects(includeArchived = true): Promise<Project[]> {
  return (
    await request<{ projects: Project[] }>(
      `/api/v1/projects?includeArchived=${includeArchived}`,
    )
  ).projects;
}

export async function createProject(input: {
  name: string;
  slidesMarkdown?: string;
}): Promise<Project> {
  return (
    await request<{ project: Project }>("/api/v1/projects", {
      method: "POST",
      body: JSON.stringify(input),
    })
  ).project;
}

export async function loadProject(projectId: string): Promise<Project> {
  return (await request<{ project: Project }>(`/api/v1/projects/${projectId}`))
    .project;
}

export async function updateProjectSettings(
  projectId: string,
  input: { sensitiveMode: boolean },
): Promise<Project> {
  return (
    await request<{ project: Project }>(
      `/api/v2/projects/${projectId}/settings`,
      { method: "PATCH", body: JSON.stringify(input) },
    )
  ).project;
}

export async function createWorkSession(
  projectId: string,
  workflowMode: WorkflowMode,
  intent: "create" | "modify" | "reference",
): Promise<WorkSession> {
  return (
    await request<{ session: WorkSession }>(
      `/api/v2/projects/${projectId}/work-sessions`,
      { method: "POST", body: JSON.stringify({ workflowMode, intent }) },
    )
  ).session;
}

export async function listProjectWorkSessions(
  projectId: string,
): Promise<WorkSession[]> {
  return (
    await request<{ sessions: WorkSession[] }>(
      `/api/v2/projects/${projectId}/work-sessions`,
    )
  ).sessions;
}

export async function updateWorkSession(
  projectId: string,
  sessionId: string,
  input: {
    intent?: "create" | "modify" | "reference";
    pageBudget?: number;
    inheritTheme?: boolean;
  } = {},
): Promise<WorkSession> {
  return (
    await request<{ session: WorkSession }>(
      `/api/v2/projects/${projectId}/work-sessions/${sessionId}`,
      { method: "PATCH", body: JSON.stringify(input) },
    )
  ).session;
}

export async function createDocumentStructurePlan(
  projectId: string,
  sessionId: string,
  input: { pageBudget: number; inheritTheme: boolean },
): Promise<{
  session: WorkSession;
  structurePlan: DocumentStructurePlan;
  operation: EditOperation;
}> {
  return request(
    `/api/v2/projects/${projectId}/work-sessions/${sessionId}/structure-plan`,
    { method: "POST", body: JSON.stringify(input) },
  );
}

export async function listProjectDocuments(
  projectId: string,
  sessionId?: string,
): Promise<DocumentSource[]> {
  const query = sessionId ? `?sessionId=${encodeURIComponent(sessionId)}` : "";
  return (
    await request<{ documents: DocumentSource[] }>(
      `/api/v2/projects/${projectId}/documents${query}`,
    )
  ).documents;
}

export async function uploadProjectDocument(
  projectId: string,
  sessionId: string,
  file: File,
): Promise<DocumentSource> {
  const form = new FormData();
  form.append("file", file, file.name);
  return (
    await request<{ document: DocumentSource }>(
      `/api/v2/projects/${projectId}/documents/upload?sessionId=${encodeURIComponent(sessionId)}`,
      { method: "POST", body: form },
    )
  ).document;
}

export async function parseProjectDocument(
  projectId: string,
  documentId: string,
): Promise<DocumentSource> {
  return (
    await request<{ document: DocumentSource }>(
      `/api/v2/projects/${projectId}/documents/${documentId}/parse`,
      { method: "POST" },
    )
  ).document;
}

export async function updateProjectDocumentContext(
  projectId: string,
  documentId: string,
  includedInContext: boolean,
): Promise<DocumentSource> {
  return (
    await request<{ document: DocumentSource }>(
      `/api/v2/projects/${projectId}/documents/${documentId}/context`,
      { method: "PATCH", body: JSON.stringify({ includedInContext }) },
    )
  ).document;
}

export async function listDocumentConflicts(
  projectId: string,
  sessionId: string,
): Promise<DocumentConflict[]> {
  return (
    await request<{ conflicts: DocumentConflict[] }>(
      `/api/v2/projects/${projectId}/conflicts?sessionId=${encodeURIComponent(sessionId)}`,
    )
  ).conflicts;
}

export async function resolveDocumentConflict(
  projectId: string,
  conflictId: string,
  resolution: {
    action: "prefer_source" | "keep_both" | "ignore";
    selectedValue?: string;
    selectedDocumentIds?: string[];
    note?: string;
  },
): Promise<DocumentConflict> {
  return (
    await request<{ conflict: DocumentConflict }>(
      `/api/v2/projects/${projectId}/conflicts/${conflictId}/resolve`,
      { method: "POST", body: JSON.stringify(resolution) },
    )
  ).conflict;
}

export async function createOfficePreview(
  projectId: string,
  exportId: string,
): Promise<{
  grantId: string;
  expiresAt: string;
  publicUrl: string;
  officeViewerUrl: string;
}> {
  return request(
    `/api/v2/projects/${projectId}/exports/${exportId}/office-preview`,
    { method: "POST" },
  );
}

export async function copyProject(projectId: string): Promise<Project> {
  return (
    await request<{ project: Project }>(`/api/v1/projects/${projectId}/copy`, {
      method: "POST",
    })
  ).project;
}

export async function archiveProject(projectId: string): Promise<Project> {
  return (
    await request<{ project: Project }>(
      `/api/v1/projects/${projectId}/archive`,
      { method: "POST" },
    )
  ).project;
}

export async function restoreProject(projectId: string): Promise<Project> {
  return (
    await request<{ project: Project }>(
      `/api/v1/projects/${projectId}/restore`,
      { method: "POST" },
    )
  ).project;
}

export async function importSlides(
  projectId: string,
  file: File,
): Promise<Project> {
  const form = new FormData();
  form.append("file", file, file.name);
  return (
    await request<{ project: Project }>(
      `/api/v1/projects/${projectId}/import`,
      { method: "POST", body: form },
    )
  ).project;
}

export async function loadMessages(
  projectId: string,
): Promise<ConversationMessage[]> {
  return (
    await request<{ messages: ConversationMessage[] }>(
      `/api/v1/projects/${projectId}/messages`,
    )
  ).messages;
}

export async function loadProjectHistory(
  projectId: string,
): Promise<ProjectHistory> {
  return request<ProjectHistory>(`/api/v1/projects/${projectId}/history`);
}

export async function sendTurn(
  projectId: string,
  input: {
    deckRevisionId: string;
    target: { mode: string; pageIds: string[] };
    message: string;
    clientRevision: number;
    conversationId?: string;
    sessionId?: string;
  },
): Promise<{ operation: any; messages: ConversationMessage[] }> {
  return request(`/api/v1/projects/${projectId}/chat/turns`, {
    method: "POST",
    body: JSON.stringify(input),
  });
}

export async function confirmOperation(
  projectId: string,
  operationId: string,
): Promise<any> {
  return request(
    `/api/v1/projects/${projectId}/edit-operations/${operationId}/confirm`,
    { method: "POST", body: JSON.stringify({ operationId }) },
  );
}

export async function cancelOperation(
  projectId: string,
  operationId: string,
): Promise<any> {
  return request(
    `/api/v1/projects/${projectId}/edit-operations/${operationId}/cancel`,
    { method: "POST" },
  );
}

export async function rollbackOperation(
  projectId: string,
  operationId: string,
): Promise<any> {
  return request(
    `/api/v1/projects/${projectId}/edit-operations/${operationId}/rollback`,
    { method: "POST" },
  );
}

export async function retryFailedOperation(
  projectId: string,
  operationId: string,
): Promise<any> {
  return request(
    `/api/v1/projects/${projectId}/edit-operations/${operationId}/retry-failed`,
    { method: "POST" },
  );
}

export async function rollbackVersion(
  projectId: string,
  versionId: string,
): Promise<Project> {
  return (
    await request<{ page: any }>(
      `/api/v1/projects/${projectId}/versions/${versionId}/rollback`,
      { method: "POST" },
    )
  ).page;
}

export async function startExport(projectId: string): Promise<ExportJob> {
  return (
    await request<{ export: ExportJob }>(
      `/api/v1/projects/${projectId}/render`,
      { method: "POST" },
    )
  ).export;
}

export async function getExport(
  projectId: string,
  exportId: string,
): Promise<ExportJob> {
  return (
    await request<{ export: ExportJob }>(
      `/api/v1/projects/${projectId}/exports/${exportId}`,
    )
  ).export;
}

export function exportDownloadUrl(projectId: string, exportId: string): string {
  return `${API_BASE}/api/v1/projects/${projectId}/exports/${exportId}/download`;
}

export function authoritativeRenderUrl(
  projectId: string,
  artifactId: string,
): string {
  return `${API_BASE}/api/v1/projects/${projectId}/artifacts/${artifactId}/content`;
}

export function visualPreviewUrl(
  projectId: string,
  artifactId: string,
): string {
  return `${API_BASE}/api/v2/projects/${projectId}/visual-previews/${artifactId}/content`;
}

export interface ProjectSocket {
  close(): void;
}

export async function openProjectSocket(
  projectId: string,
  onEvent: (event: EventEnvelope) => void,
): Promise<ProjectSocket | null> {
  const configured = import.meta.env.VITE_WS_URL;
  const url =
    configured ||
    `${window.location.protocol === "https:" ? "wss" : "ws"}://${window.location.hostname}:${import.meta.env.VITE_WS_PORT || "8788"}`;
  let closed = false;
  let retry = 0;
  let lastSeq = 0;
  let socket: WebSocket | null = null;
  let retryTimer: number | undefined;
  let connecting = false;
  const scheduleReconnect = (delay: number): void => {
    if (closed || retryTimer || window.navigator.onLine === false) return;
    retryTimer = window.setTimeout(() => {
      retryTimer = undefined;
      void connect();
    }, delay);
  };
  const connect = async (): Promise<void> => {
    if (closed || connecting || socket) return;
    connecting = true;
    try {
      const result = await request<{ ticket: string }>(
        "/api/v1/auth/ws-ticket",
        { method: "POST" },
      );
      if (closed) {
        connecting = false;
        return;
      }
      socket = new WebSocket(
        `${url}/?projectId=${encodeURIComponent(projectId)}&afterSeq=${lastSeq}`,
        [`fastppt-ticket.${result.ticket}`],
      );
      connecting = false;
      socket.onopen = () => {
        retry = 0;
      };
      socket.onmessage = (event) => {
        try {
          const parsed = JSON.parse(event.data) as EventEnvelope;
          if (parsed.seq > lastSeq) lastSeq = parsed.seq;
          onEvent(parsed);
        } catch {
          /* Ignore malformed event payloads. */
        }
      };
      socket.onclose = () => {
        socket = null;
        if (!closed) {
          retry += 1;
          scheduleReconnect(
            Math.min(10_000, 400 * 2 ** Math.min(retry, 5)),
          );
        }
      };
    } catch {
      connecting = false;
      if (!closed) {
        retry += 1;
        scheduleReconnect(
          Math.min(10_000, 400 * 2 ** Math.min(retry, 5)),
        );
      }
    }
  };
  const handleOffline = (): void => {
    socket?.close();
  };
  const handleOnline = (): void => {
    if (closed) return;
    if (retryTimer) {
      window.clearTimeout(retryTimer);
      retryTimer = undefined;
    }
    retry = 0;
    void connect();
  };
  window.addEventListener("offline", handleOffline);
  window.addEventListener("online", handleOnline);
  void connect();
  return {
    close: () => {
      closed = true;
      if (retryTimer) window.clearTimeout(retryTimer);
      window.removeEventListener("offline", handleOffline);
      window.removeEventListener("online", handleOnline);
      socket?.close();
    },
  };
}

import type { ConversationMessage, EventEnvelope, ExportJob, Project, User } from '../shared/models.js';

const API_BASE = import.meta.env.VITE_API_BASE || '';
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
  if (init.body && !(init.body instanceof FormData)) headers.set('Content-Type', 'application/json');
  const token = getToken();
  if (token) headers.set('Authorization', `Bearer ${token}`);
  const response = await fetch(`${API_BASE}${path}`, { ...init, headers, credentials: 'include' });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(data?.error?.message || `Request failed with ${response.status}`) as ApiError;
    error.status = response.status;
    throw error;
  }
  return data as T;
}

export async function login(email: string, name: string, accessCode?: string): Promise<{ token: string; user: User }> {
  const result = await request<{ token: string; user: User }>('/api/v1/auth/login', { method: 'POST', body: JSON.stringify({ email, name, accessCode }) });
  setToken(result.token);
  return result;
}

export async function me(): Promise<User | null> {
  return (await request<{ user: User | null }>('/api/v1/auth/me')).user;
}

export async function logout(): Promise<void> {
  await request('/api/v1/auth/logout', { method: 'POST' });
  clearToken();
}

export async function listProjects(includeArchived = true): Promise<Project[]> {
  return (await request<{ projects: Project[] }>(`/api/v1/projects?includeArchived=${includeArchived}`)).projects;
}

export async function createProject(input: { name: string; slidesMarkdown?: string }): Promise<Project> {
  return (await request<{ project: Project }>('/api/v1/projects', { method: 'POST', body: JSON.stringify(input) })).project;
}

export async function loadProject(projectId: string): Promise<Project> {
  return (await request<{ project: Project }>(`/api/v1/projects/${projectId}`)).project;
}

export async function copyProject(projectId: string): Promise<Project> {
  return (await request<{ project: Project }>(`/api/v1/projects/${projectId}/copy`, { method: 'POST' })).project;
}

export async function archiveProject(projectId: string): Promise<Project> {
  return (await request<{ project: Project }>(`/api/v1/projects/${projectId}/archive`, { method: 'POST' })).project;
}

export async function restoreProject(projectId: string): Promise<Project> {
  return (await request<{ project: Project }>(`/api/v1/projects/${projectId}/restore`, { method: 'POST' })).project;
}

export async function importSlides(projectId: string, file: File): Promise<Project> {
  const form = new FormData();
  form.append('file', file, file.name);
  return (await request<{ project: Project }>(`/api/v1/projects/${projectId}/import`, { method: 'POST', body: form })).project;
}

export async function loadMessages(projectId: string): Promise<ConversationMessage[]> {
  return (await request<{ messages: ConversationMessage[] }>(`/api/v1/projects/${projectId}/messages`)).messages;
}

export async function sendTurn(projectId: string, input: { deckRevisionId: string; target: { mode: string; pageIds: string[] }; message: string; clientRevision: number; conversationId?: string }): Promise<{ operation: any; messages: ConversationMessage[] }> {
  return request(`/api/v1/projects/${projectId}/chat/turns`, { method: 'POST', body: JSON.stringify(input) });
}

export async function confirmOperation(projectId: string, operationId: string): Promise<any> {
  return request(`/api/v1/projects/${projectId}/edit-operations/${operationId}/confirm`, { method: 'POST', body: JSON.stringify({ operationId }) });
}

export async function cancelOperation(projectId: string, operationId: string): Promise<any> {
  return request(`/api/v1/projects/${projectId}/edit-operations/${operationId}/cancel`, { method: 'POST' });
}

export async function rollbackOperation(projectId: string, operationId: string): Promise<any> {
  return request(`/api/v1/projects/${projectId}/edit-operations/${operationId}/rollback`, { method: 'POST' });
}

export async function retryFailedOperation(projectId: string, operationId: string): Promise<any> {
  return request(`/api/v1/projects/${projectId}/edit-operations/${operationId}/retry-failed`, { method: 'POST' });
}

export async function rollbackVersion(projectId: string, versionId: string): Promise<Project> {
  return (await request<{ page: any }>(`/api/v1/projects/${projectId}/versions/${versionId}/rollback`, { method: 'POST' })).page;
}

export async function startExport(projectId: string): Promise<ExportJob> {
  return (await request<{ export: ExportJob }>(`/api/v1/projects/${projectId}/render`, { method: 'POST' })).export;
}

export async function getExport(projectId: string, exportId: string): Promise<ExportJob> {
  return (await request<{ export: ExportJob }>(`/api/v1/projects/${projectId}/exports/${exportId}`)).export;
}

export function exportDownloadUrl(projectId: string, exportId: string): string {
  return `${API_BASE}/api/v1/projects/${projectId}/exports/${exportId}/download`;
}

export interface ProjectSocket {
  close(): void;
}

export async function openProjectSocket(projectId: string, onEvent: (event: EventEnvelope) => void): Promise<ProjectSocket | null> {
  const configured = import.meta.env.VITE_WS_URL;
  const url = configured || `${window.location.protocol === 'https:' ? 'wss' : 'ws'}://${window.location.hostname}:8788`;
  let closed = false;
  let retry = 0;
  let lastSeq = 0;
  let socket: WebSocket | null = null;
  let retryTimer: number | undefined;
  const connect = async (): Promise<void> => {
    if (closed) return;
    try {
      const result = await request<{ ticket: string }>('/api/v1/auth/ws-ticket', { method: 'POST' });
      if (closed) return;
      socket = new WebSocket(`${url}/?projectId=${encodeURIComponent(projectId)}&afterSeq=${lastSeq}`, [`fastppt-ticket.${result.ticket}`]);
      socket.onopen = () => { retry = 0; };
      socket.onmessage = (event) => {
        try {
          const parsed = JSON.parse(event.data) as EventEnvelope;
          if (parsed.seq > lastSeq) lastSeq = parsed.seq;
          onEvent(parsed);
        } catch { /* Ignore malformed event payloads. */ }
      };
      socket.onclose = () => {
        socket = null;
        if (!closed) {
          retry += 1;
          retryTimer = window.setTimeout(() => void connect(), Math.min(10_000, 400 * 2 ** Math.min(retry, 5)));
        }
      };
    } catch {
      if (!closed) {
        retry += 1;
        retryTimer = window.setTimeout(() => void connect(), Math.min(10_000, 400 * 2 ** Math.min(retry, 5)));
      }
    }
  };
  void connect();
  return {
    close: () => {
      closed = true;
      if (retryTimer) window.clearTimeout(retryTimer);
      socket?.close();
    },
  };
}

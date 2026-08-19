import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { JSX } from 'react';
import {
  Archive,
  ArrowLeft,
  Check,
  ChevronDown,
  CircleAlert,
  Clock3,
  Copy,
  Download,
  Expand,
  FilePlus2,
  FileText,
  History,
  Layers3,
  LayoutPanelLeft,
  LoaderCircle,
  LogOut,
  MessageSquareText,
  MoreHorizontal,
  PanelRight,
  Play,
  Plus,
  RotateCcw,
  Send,
  Sparkles,
  SquarePen,
  Undo2,
  Upload,
  WandSparkles,
  X,
} from 'lucide-react';
import type { ConversationMessage, EditOperation, EventEnvelope, ExportJob, Page, PageVersion, Project, User } from '../shared/models.js';
import {
  archiveProject,
  authoritativeRenderUrl,
  cancelOperation,
  clearToken,
  copyProject,
  confirmOperation,
  createProject,
  exportDownloadUrl,
  getExport,
  importSlides,
  listProjects,
  loadMessages,
  loadProject,
  login,
  me,
  openProjectSocket,
  rollbackVersion,
  rollbackOperation,
  retryFailedOperation,
  sendTurn,
  startExport,
  logout,
  restoreProject,
} from './client.js';
import type { ProjectSocket } from './client.js';

type Mode = 'single' | 'multi' | 'global';

const statusText: Record<Page['status'], string> = {
  untouched: '未修改',
  quick_preview: '快速预览',
  generating: '生成中',
  authoritative: 'PPTX 权威渲染',
  svg_fallback: 'SVG 回退',
  failed: '失败',
  rolled_back: '已回滚',
};

const statusTone: Record<Page['status'], string> = {
  untouched: 'neutral',
  quick_preview: 'blue',
  generating: 'orange',
  authoritative: 'green',
  svg_fallback: 'purple',
  failed: 'red',
  rolled_back: 'orange',
};

function formatTime(value: string): string {
  return new Intl.DateTimeFormat('zh-CN', { hour: '2-digit', minute: '2-digit' }).format(new Date(value));
}

function currentVersion(page: Page): PageVersion {
  return page.versions.find((version) => version.versionId === page.currentVersionId) || page.versions[page.versions.length - 1];
}

function versionDisplayStatus(version: PageVersion): Page['status'] {
  if (version.status === 'failed' || version.status === 'rejected') return 'failed';
  if (version.status === 'rendering') return 'generating';
  if (version.previewKind === 'pptx_authoritative') return 'authoritative';
  if (version.previewKind === 'svg_fallback') return 'svg_fallback';
  return 'quick_preview';
}

function PagePreview({ projectId, version }: { projectId: string; version: PageVersion }): JSX.Element {
  const [authorityFailed, setAuthorityFailed] = useState(false);
  const artifactId = version.previewKind === 'pptx_authoritative' ? version.pptxPageRenderId : null;
  useEffect(() => setAuthorityFailed(false), [artifactId]);
  if (artifactId && !authorityFailed) {
    return <div className="slide-render authoritative-render"><img src={authoritativeRenderUrl(projectId, artifactId)} alt={`PowerPoint 权威渲染 ${version.versionId}`} data-artifact-id={artifactId} onError={() => setAuthorityFailed(true)} /></div>;
  }
  if (authorityFailed) {
    return <div className="slide-render authority-load-fallback" data-preview-kind="authority-load-fallback"><div className="authority-fallback-note"><CircleAlert size={13} />权威 PNG 加载失败，正在显示同版本 SVG 回退</div><div dangerouslySetInnerHTML={{ __html: version.previewSvg }} /></div>;
  }
  return <div className="slide-render" data-preview-kind={version.previewKind} dangerouslySetInnerHTML={{ __html: version.previewSvg }} />;
}

function App(): JSX.Element {
  const [user, setUser] = useState<User | null>(null);
  const [projects, setProjects] = useState<Project[]>([]);
  const [project, setProject] = useState<Project | null>(null);
  const [messages, setMessages] = useState<ConversationMessage[]>([]);
  const [currentPageId, setCurrentPageId] = useState('');
  const [selectedPageIds, setSelectedPageIds] = useState<string[]>([]);
  const [mode, setMode] = useState<Mode>('single');
  const [input, setInput] = useState('');
  const [pendingOperation, setPendingOperation] = useState<EditOperation | null>(null);
  const [busy, setBusy] = useState(false);
  const [loading, setLoading] = useState(true);
  const [loginForm, setLoginForm] = useState({ email: 'demo@fastppt.local', name: 'Demo Editor', accessCode: '' });
  const [error, setError] = useState('');
  const [eventNote, setEventNote] = useState('');
  const [showHistory, setShowHistory] = useState(false);
  const [compareVersionId, setCompareVersionId] = useState('');
  const [exportJob, setExportJob] = useState<ExportJob | null>(null);
  const socketRef = useRef<ProjectSocket | null>(null);
  const importRef = useRef<HTMLInputElement | null>(null);
  const [projectMenuOpen, setProjectMenuOpen] = useState(false);

  const visiblePages = useMemo(() => project?.pages.filter((page) => !page.archived) || [], [project]);
  const activePage = visiblePages.find((page) => page.pageId === currentPageId) || visiblePages[0];
  const activeVersion = activePage ? currentVersion(activePage) : null;
  const compareVersion = activePage?.versions.find((version) => version.versionId === compareVersionId);
  const displayedVersion = compareVersion || activeVersion;
  const displayedStatus = compareVersion ? versionDisplayStatus(compareVersion) : activePage?.status || 'untouched';

  const refreshProject = useCallback(async (projectId: string) => {
    const [nextProject, nextMessages] = await Promise.all([loadProject(projectId), loadMessages(projectId)]);
    setProject(nextProject);
    setMessages(nextMessages);
    const visible = nextProject.pages.filter((page) => !page.archived);
    setCurrentPageId((previous) => visible.some((page) => page.pageId === previous) ? previous : visible[0]?.pageId || '');
  }, []);

  const connectProject = useCallback(async (nextProject: Project) => {
    socketRef.current?.close();
    const socket = await openProjectSocket(nextProject.projectId, (event: EventEnvelope) => {
      setEventNote(event.type.replaceAll('.', ' · '));
      if (['page.version.created', 'preview.quick.ready', 'preview.pptx.ready', 'render.warning', 'edit.failed', 'edit.completed', 'edit.rolled_back', 'page.archived', 'page.restored', 'page.reordered', 'page.split'].includes(event.type)) {
        void refreshProject(nextProject.projectId);
      }
    });
    socketRef.current = socket;
  }, [refreshProject]);

  const selectProject = useCallback(async (nextProject: Project) => {
    setProject(nextProject);
    setCurrentPageId(nextProject.pages.find((page) => !page.archived)?.pageId || '');
    setSelectedPageIds([]);
    setMode('single');
    await refreshProject(nextProject.projectId);
    await connectProject(nextProject);
  }, [connectProject, refreshProject]);

  useEffect(() => {
    let cancelled = false;
    async function bootstrap(): Promise<void> {
      try {
        const nextUser = await me();
        if (cancelled) return;
        if (!nextUser) return;
        setUser(nextUser);
        const nextProjects = await listProjects();
        setProjects(nextProjects);
        const firstActive = nextProjects.find((item) => item.status !== 'archived');
        if (firstActive) await selectProject(firstActive);
      } catch {
        clearToken();
      } finally {
        if (!cancelled) setLoading(false);
      }
    }
    void bootstrap();
    return () => { cancelled = true; socketRef.current?.close(); };
  }, [selectProject]);

  useEffect(() => () => socketRef.current?.close(), []);

  async function onLogin(event: React.FormEvent): Promise<void> {
    event.preventDefault();
    setBusy(true); setError('');
    try {
      const result = await login(loginForm.email, loginForm.name, loginForm.accessCode);
      setUser(result.user);
      const nextProjects = await listProjects();
      setProjects(nextProjects);
      const firstActive = nextProjects.find((item) => item.status !== 'archived');
      if (firstActive) await selectProject(firstActive);
    } catch (caught) {
      setError((caught as Error).message);
    } finally { setBusy(false); setLoading(false); }
  }

  async function onCreateProject(): Promise<void> {
    const name = window.prompt('项目名称', '我的 FastPPT 项目');
    if (!name?.trim()) return;
    setBusy(true); setError('');
    try {
      const next = await createProject({ name: name.trim() });
      setProjects((current) => [next, ...current]);
      await selectProject(next);
    } catch (caught) { setError((caught as Error).message); } finally { setBusy(false); }
  }

  async function onCopyProject(): Promise<void> {
    if (!project) return;
    setBusy(true); setError('');
    try {
      const next = await copyProject(project.projectId);
      setProjects((current) => [next, ...current]);
      await selectProject(next);
      setProjectMenuOpen(false);
    } catch (caught) { setError((caught as Error).message); } finally { setBusy(false); }
  }

  async function onArchiveProject(): Promise<void> {
    if (!project) return;
    setBusy(true); setError('');
    try {
      const archived = await archiveProject(project.projectId);
      setProjects((current) => current.map((item) => item.projectId === archived.projectId ? archived : item));
      const next = projects.find((item) => item.projectId !== archived.projectId && item.status !== 'archived');
      if (next) await selectProject(next); else setProject(null);
      setProjectMenuOpen(false);
    } catch (caught) { setError((caught as Error).message); } finally { setBusy(false); }
  }

  async function onRestoreProject(projectId: string): Promise<void> {
    setBusy(true); setError('');
    try {
      const restored = await restoreProject(projectId);
      setProjects((current) => current.map((item) => item.projectId === restored.projectId ? restored : item));
      await selectProject(restored);
      setProjectMenuOpen(false);
    } catch (caught) { setError((caught as Error).message); } finally { setBusy(false); }
  }

  async function onImportFile(event: React.ChangeEvent<HTMLInputElement>): Promise<void> {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file || !project) return;
    setBusy(true); setError('');
    try {
      const next = await importSlides(project.projectId, file);
      setProjects((current) => current.map((item) => item.projectId === next.projectId ? next : item));
      await selectProject(next);
      setEventNote('slides.md 已导入，页面合同和稳定 ID 已刷新');
    } catch (caught) { setError((caught as Error).message); } finally { setBusy(false); }
  }

  async function onLogout(): Promise<void> {
    try { await logout(); } catch { clearToken(); }
    socketRef.current?.close();
    setUser(null);
    setProject(null);
  }

  function togglePageSelection(pageId: string): void {
    setSelectedPageIds((current) => current.includes(pageId) ? current.filter((id) => id !== pageId) : [...current, pageId]);
    setMode('multi');
  }

  async function onSend(event: React.FormEvent): Promise<void> {
    event.preventDefault();
    if (!project || !input.trim() || busy) return;
    const pageIds = mode === 'single' ? (activePage ? [activePage.pageId] : []) : mode === 'multi' ? selectedPageIds : [];
    if (mode === 'multi' && pageIds.length === 0) { setError('请先勾选要修改的页面。'); return; }
    setBusy(true); setError('');
    try {
      const result = await sendTurn(project.projectId, {
        deckRevisionId: project.currentDeckRevisionId,
        target: { mode, pageIds },
        message: input.trim(),
        clientRevision: project.pages.reduce((sum, page) => sum + page.versions.length, 0),
      });
      setInput('');
      setMessages((current) => [...current, ...result.messages.filter((message) => !current.some((existing) => existing.messageId === message.messageId))]);
      setPendingOperation(result.operation.confirmationRequired ? result.operation : null);
      await refreshProject(project.projectId);
    } catch (caught) { setError((caught as Error).message); } finally { setBusy(false); }
  }

  async function onConfirm(): Promise<void> {
    if (!project || !pendingOperation) return;
    setBusy(true); setError('');
    try { await confirmOperation(project.projectId, pendingOperation.operationId); setPendingOperation(null); await refreshProject(project.projectId); }
    catch (caught) { setError((caught as Error).message); } finally { setBusy(false); }
  }

  async function onCancel(): Promise<void> {
    if (!project || !pendingOperation) return;
    setBusy(true);
    try { await cancelOperation(project.projectId, pendingOperation.operationId); setPendingOperation(null); await refreshProject(project.projectId); }
    catch (caught) { setError((caught as Error).message); } finally { setBusy(false); }
  }

  async function onRollback(versionId: string): Promise<void> {
    if (!project) return;
    setBusy(true); setError('');
    try { await rollbackVersion(project.projectId, versionId); setShowHistory(false); setCompareVersionId(''); await refreshProject(project.projectId); }
    catch (caught) { setError((caught as Error).message); } finally { setBusy(false); }
  }

  async function onRollbackOperation(operationId: string): Promise<void> {
    if (!project) return;
    setBusy(true); setError('');
    try {
      await rollbackOperation(project.projectId, operationId);
      setEventNote('本次操作涉及的页面已整组回滚');
      await refreshProject(project.projectId);
    } catch (caught) { setError((caught as Error).message); } finally { setBusy(false); }
  }

  async function onRetryOperation(operationId: string): Promise<void> {
    if (!project) return;
    setBusy(true); setError('');
    try {
      await retryFailedOperation(project.projectId, operationId);
      setEventNote('失败页面已创建独立重试任务');
      await refreshProject(project.projectId);
    } catch (caught) { setError((caught as Error).message); } finally { setBusy(false); }
  }

  async function onExport(): Promise<void> {
    if (!project || busy) return;
    setBusy(true); setError('');
    try {
      const job = await startExport(project.projectId);
      setExportJob(job);
      const poll = async (): Promise<void> => {
        const next = await getExport(project.projectId, job.exportId);
        setExportJob(next);
        if (next.status === 'running' || next.status === 'queued') window.setTimeout(() => void poll(), 700);
      };
      window.setTimeout(() => void poll(), 500);
    } catch (caught) { setError((caught as Error).message); } finally { setBusy(false); }
  }

  const targetDescription = mode === 'single'
    ? `当前第 ${(activePage?.orderIndex ?? 0) + 1} 页 · ${activePage?.pageId || '未选择'}`
    : mode === 'multi'
      ? `已选 ${selectedPageIds.length} 页 · 执行前确认`
      : '全局相似页 · 先看候选，再确认';

  if (loading) return <div className="loading-screen"><LoaderCircle className="spin" size={24} /> 正在恢复工作区</div>;
  if (!user) return <LoginScreen form={loginForm} setForm={setLoginForm} onSubmit={onLogin} busy={busy} error={error} />;
  if (!project) return <EmptyWorkspace user={user} onCreate={onCreateProject} onLogout={() => void onLogout()} busy={busy} error={error} />;

  return <div className="app-shell">
    <aside className="sidebar">
      <div className="brand-row"><div className="brand-mark">F</div><div><strong>FastPPT</strong><span>ONLINE WORKBENCH</span></div><button className="icon-button subtle" title="更多项目操作"><MoreHorizontal size={18} /></button></div>
      <div className="project-switcher"><div className="eyebrow">当前项目</div><button className="project-name" onClick={() => setProjectMenuOpen((value) => !value)}><FileText size={15} />{project.name}<ChevronDown size={14} /></button>{projectMenuOpen && <div className="project-menu">{projects.filter((item) => item.status !== 'archived').map((item) => <button key={item.projectId} className={item.projectId === project.projectId ? 'selected' : ''} onClick={() => { void selectProject(item); setProjectMenuOpen(false); }}><FileText size={14} />{item.name}</button>)}{projects.filter((item) => item.status === 'archived').map((item) => <button key={item.projectId} onClick={() => void onRestoreProject(item.projectId)}><RotateCcw size={14} />恢复 {item.name}</button>)}<button onClick={() => void onCopyProject()}><Copy size={14} />复制当前项目</button><button onClick={() => void onArchiveProject()}><Archive size={14} />归档当前项目</button></div>}<div className="project-meta"><span className="live-dot" />{visiblePages.length} 页 · {project.themeId}</div></div>
      <div className="sidebar-heading"><span>页面</span><span className="muted-count">{visiblePages.length}</span><button className="icon-button" title="新建项目" onClick={() => void onCreateProject()}><Plus size={16} /></button></div>
      <div className="page-list">{visiblePages.map((page) => <PageRow key={page.pageId} page={page} active={page.pageId === activePage?.pageId} selected={selectedPageIds.includes(page.pageId)} onClick={() => { setCurrentPageId(page.pageId); if (mode === 'single') setSelectedPageIds([]); }} onToggle={() => togglePageSelection(page.pageId)} />)}</div>
      <div className="sidebar-footer"><div className="user-chip"><div className="avatar">{user.name.slice(0, 1).toUpperCase()}</div><div><strong>{user.name}</strong><span>{user.email}</span></div><button className="icon-button" title="退出登录" onClick={() => void onLogout()}><LogOut size={15} /></button></div><button className="outline-button full" onClick={() => void onCreateProject()}><FilePlus2 size={15} />新建项目</button></div>
    </aside>

    <main className="main-stage">
      <header className="topbar"><div className="breadcrumbs"><span>工作台</span><ArrowLeft size={14} /><strong>{project.name}</strong></div><div className="top-actions"><span className="saved-state"><span className="save-dot" />已保存</span><button className="outline-button" onClick={() => void onExport()} disabled={busy}><Download size={15} />导出 PPTX</button><button className="icon-button" title="页面全屏" onClick={() => document.documentElement.requestFullscreen?.()}><Expand size={17} /></button></div></header>
      <section className="canvas-toolbar"><div className="toolbar-left"><button className="tool-button active"><LayoutPanelLeft size={16} />预览</button><button className="tool-button" onClick={() => setEventNote('合同视图已在右侧消息元数据中保留')}><Layers3 size={16} />合同</button></div><div className="toolbar-center"><span className={`status-pill ${statusTone[displayedStatus]}`} data-page-status={displayedStatus}><span className="status-dot" />{compareVersion && '对比 · '}{statusText[displayedStatus]}</span><span className="version-label" data-version-id={displayedVersion?.versionId || ''}>{displayedVersion?.versionId || 'ver_pending'}</span></div><div className="toolbar-right"><button className="icon-button" title="版本历史" onClick={() => setShowHistory((value) => !value)}><History size={17} /></button><button className="icon-button" title="撤销到上一版本" onClick={() => activePage && activePage.versions.length > 1 && void onRollback(activePage.versions[activePage.versions.length - 2].versionId)} disabled={!activePage || activePage.versions.length < 2}><Undo2 size={17} /></button></div></section>
      <section className="preview-workspace"><div className="preview-wrap"><div className="preview-frame"><div className="preview-chrome"><span className="chrome-dot red" /><span className="chrome-dot amber" /><span className="chrome-dot green" /><span className="preview-url">fastppt.online / preview / {activePage?.pageId}</span><button className="icon-button subtle" title="刷新预览" onClick={() => project && void refreshProject(project.projectId)}><RotateCcw size={14} /></button></div>{displayedVersion ? <PagePreview projectId={project.projectId} version={displayedVersion} /> : <div className="empty-slide">选择一个页面开始预览</div>}<div className="preview-legend"><span><span className="legend-line quick" />快速预览</span><span><span className="legend-line authority" />PPTX 权威</span><span><span className="legend-line fallback" />SVG 回退</span></div></div><div className="preview-caption"><div><strong>第 {(activePage?.orderIndex ?? 0) + 1} 页</strong><span>{displayedVersion?.title || activePage?.title}</span></div><div className="caption-right"><span className="editable-tag"><Sparkles size={13} />{displayedVersion?.editableLevel || activePage?.editableLevel}</span>{displayedVersion?.previewKind === 'svg_fallback' && <span className="warning-caption"><CircleAlert size={13} />可能与 PPTX 有差异</span>}</div></div></div>{showHistory && activePage && <HistoryPanel page={activePage} compareId={compareVersionId} onCompare={setCompareVersionId} onRollback={onRollback} onClose={() => setShowHistory(false)} />}</section>
      {exportJob && <ExportBanner job={exportJob} project={project} onDismiss={() => setExportJob(null)} />}
    </main>

    <aside className="chat-panel"><div className="chat-header"><div><div className="eyebrow">AI 精修助手</div><h1>页面对话</h1></div><div className="assistant-state"><span className="live-dot" />在线</div></div><div className="target-strip"><div className="target-icon"><PanelRight size={16} /></div><div><span>编辑目标</span><strong>{targetDescription}</strong></div></div><div className="mode-tabs"><button className={mode === 'single' ? 'active' : ''} onClick={() => setMode('single')}><SquarePen size={14} />当前页</button><button className={mode === 'multi' ? 'active' : ''} onClick={() => setMode('multi')}><Layers3 size={14} />多页</button><button className={mode === 'global' ? 'active' : ''} onClick={() => setMode('global')}><WandSparkles size={14} />全局</button></div><div className="chat-scroll">{messages.length === 0 ? <WelcomeMessage /> : messages.map((message) => <ChatMessage key={message.messageId} message={message} busy={busy} onRollbackOperation={onRollbackOperation} onRetryOperation={onRetryOperation} />)}{pendingOperation && <ConfirmationCard operation={pendingOperation} onConfirm={() => void onConfirm()} onCancel={() => void onCancel()} busy={busy} />}{eventNote && <div className="event-note"><Play size={12} />{eventNote}</div>}</div><div className="composer-wrap"><div className="composer-hint"><span>事实锚点已锁定</span><span>·</span><span>{activePage?.factAnchors.length || 0} 个数字</span></div><form className="composer" onSubmit={onSend}><textarea value={input} onChange={(event) => setInput(event.target.value)} placeholder={mode === 'single' ? '描述当前页要怎么改…' : mode === 'multi' ? '描述选中页面的共同修改…' : '描述要查找并修改的相似页…'} rows={3} disabled={busy} /><div className="composer-actions"><button type="button" className="icon-button subtle" title="上传 slides.md" onClick={() => importRef.current?.click()}><Upload size={16} /></button><input ref={importRef} type="file" accept=".md,text/markdown" hidden onChange={(event) => void onImportFile(event)} /><span className="composer-model">{busy ? <><LoaderCircle className="spin" size={14} />处理中</> : 'FastPPT planner'}</span><button type="submit" className="send-button" disabled={busy || !input.trim()} title="发送消息"><Send size={16} /></button></div></form>{error && <div className="error-inline"><CircleAlert size={14} />{error}</div>}</div></aside>
  </div>;
}

function PageRow({ page, active, selected, onClick, onToggle }: { page: Page; active: boolean; selected: boolean; onClick: () => void; onToggle: () => void }): JSX.Element {
  const version = currentVersion(page);
  return <div className={`page-row ${active ? 'active' : ''}`} onClick={onClick}><button className={`check-box ${selected ? 'checked' : ''}`} onClick={(event) => { event.stopPropagation(); onToggle(); }} aria-label={`选择第 ${page.orderIndex + 1} 页`}>{selected && <Check size={12} />}</button><div className="thumb"><div dangerouslySetInnerHTML={{ __html: version.previewSvg }} /></div><div className="page-row-info"><div className="page-row-top"><strong>{String(page.orderIndex + 1).padStart(2, '0')}</strong><span className={`mini-status ${statusTone[page.status]}`} /></div><div className="page-row-title">{page.title}</div><div className="page-row-meta"><span>{version.versionId.slice(0, 13)}</span><span>{statusText[page.status]}</span></div></div></div>;
}

function ChatMessage({ message, busy, onRollbackOperation, onRetryOperation }: { message: ConversationMessage; busy: boolean; onRollbackOperation: (operationId: string) => Promise<void>; onRetryOperation: (operationId: string) => Promise<void> }): JSX.Element {
  const versionIds = Array.isArray(message.meta?.versionIds) ? message.meta.versionIds as string[] : [];
  const failedPageIds = Array.isArray(message.meta?.failedPageIds) ? message.meta.failedPageIds as string[] : [];
  const rolledBack = message.meta?.rolledBack === true;
  const retryOperationId = typeof message.meta?.retryOperationId === 'string' ? message.meta.retryOperationId : '';
  return <div className={`chat-message ${message.role}`}><div className="message-meta"><span>{message.role === 'user' ? '你' : message.role === 'assistant' ? 'FastPPT' : '系统'}</span><time>{formatTime(message.createdAt)}</time></div><div className="message-bubble">{message.text}</div>{versionIds.length > 0 && <div className="message-foot"><span><Check size={12} />版本已登记 · {versionIds.length} 页</span>{message.operationId && (rolledBack ? <span>已整组回滚</span> : <button type="button" onClick={() => void onRollbackOperation(message.operationId!)} disabled={busy}><Undo2 size={12} />撤销本次操作</button>)}</div>}{failedPageIds.length > 0 && <div className="message-foot"><span><CircleAlert size={12} />{failedPageIds.length} 页失败，可单独重试</span>{retryOperationId ? <span>重试任务已创建</span> : message.operationId && <button type="button" onClick={() => void onRetryOperation(message.operationId!)} disabled={busy}><RotateCcw size={12} />重试失败页</button>}</div>}</div>;
}

function ConfirmationCard({ operation, onConfirm, onCancel, busy }: { operation: EditOperation; onConfirm: () => void; onCancel: () => void; busy: boolean }): JSX.Element {
  const plan = operation.structuredPlan;
  return <div className="confirmation-card"><div className="confirmation-title"><CircleAlert size={16} /><strong>执行前确认</strong><span>{operation.mode === 'global' ? '全局相似页' : `${operation.resolvedPageIds.length} 页`}</span></div><p>{plan.summary}</p><div className="candidate-list">{operation.resolvedPageIds.map((pageId) => <div key={pageId}><span className="candidate-index">{pageId.slice(-3)}</span><span>{plan.candidateReasons?.[pageId] || '合同与请求目标匹配'}</span></div>)}</div><div className="confirmation-cost"><span>预计成本</span><strong>${plan.estimatedCost.amount.toFixed(2)}</strong><span>· {plan.estimatedCost.imageUnits} image unit</span></div>{plan.unsupported.length > 0 && <div className="unsupported-note">{plan.unsupported.join(' ')}</div>}<div className="confirmation-actions"><button className="ghost-button" onClick={onCancel} disabled={busy}><X size={14} />取消</button><button className="primary-button" onClick={onConfirm} disabled={busy}><Check size={14} />确认执行</button></div></div>;
}

function HistoryPanel({ page, compareId, onCompare, onRollback, onClose }: { page: Page; compareId: string; onCompare: (id: string) => void; onRollback: (id: string) => Promise<void>; onClose: () => void }): JSX.Element {
  return <div className="history-panel"><div className="history-heading"><div><span className="eyebrow">IMMUTABLE VERSIONS</span><strong>{page.versions.length} 个版本</strong></div><button className="icon-button" title="关闭历史" onClick={onClose}><X size={15} /></button></div><div className="history-list">{[...page.versions].reverse().map((version, index) => {
    const isCurrent = version.versionId === page.currentVersionId;
    return <div className={`history-row ${isCurrent ? 'current' : ''}`} data-version-id={version.versionId} data-current={isCurrent ? 'true' : 'false'} key={version.versionId}><div className="history-marker">{isCurrent ? <Check size={12} /> : <Clock3 size={12} />}</div><div className="history-info"><strong>{isCurrent ? '当前版本' : `版本 ${page.versions.length - index}`}</strong><span>{version.versionId} · {formatTime(version.createdAt)}</span><small>{version.message}</small></div><div className="history-actions"><button className="icon-button subtle" title="对比版本" onClick={() => onCompare(compareId === version.versionId ? '' : version.versionId)}><Copy size={14} /></button>{!isCurrent && <button className="icon-button subtle" title="恢复此版本" onClick={() => void onRollback(version.versionId)}><RotateCcw size={14} /></button>}</div></div>;
  })}</div>{compareId && <div className="compare-note"><Copy size={13} />预览区正在显示对比版本。再次点击可返回当前版本。</div>}</div>;
}

function ExportBanner({ job, project, onDismiss }: { job: ExportJob; project: Project; onDismiss: () => void }): JSX.Element {
  const complete = job.status === 'completed';
  return <div className="export-banner"><div className={`export-icon ${complete ? 'done' : 'pending'}`}>{complete ? <Check size={18} /> : <LoaderCircle className="spin" size={18} />}</div><div><strong>{complete ? 'PPTX 导出完成' : '正在生成可编辑 PPTX'}</strong><span>{job.renderMode === 'svg_fallback' ? '静态结构已生成，PowerPoint 权威渲染待补齐' : '正在运行交付 QA'}</span></div><div className="export-actions">{complete && <a className="outline-button" href={exportDownloadUrl(project.projectId, job.exportId)}><Download size={14} />下载</a>}<button className="icon-button" title="关闭导出状态" onClick={onDismiss}><X size={15} /></button></div></div>;
}

function WelcomeMessage(): JSX.Element {
  return <div className="welcome-message"><div className="welcome-icon"><MessageSquareText size={20} /></div><strong>从当前页开始</strong><p>说出要改什么，计划通过合同校验后会直接生成新版本。</p><div className="suggestion-list"><span>“标题改短，保留所有数字”</span><span>“右侧改为两列”</span><span>“所有来源统一格式”</span></div></div>;
}

function LoginScreen({ form, setForm, onSubmit, busy, error }: { form: { email: string; name: string; accessCode: string }; setForm: (value: { email: string; name: string; accessCode: string }) => void; onSubmit: (event: React.FormEvent) => Promise<void>; busy: boolean; error: string }): JSX.Element {
  return <div className="auth-screen"><div className="auth-panel"><div className="brand-mark large">F</div><div className="eyebrow">FASTPPT ONLINE</div><h1>把每一页，聊到可交付。</h1><p>浏览器内完成页面选择、聊天精修、版本审计和可编辑 PPTX 导出。</p><form onSubmit={(event) => void onSubmit(event)}><label>邮箱<input type="email" value={form.email} onChange={(event) => setForm({ ...form, email: event.target.value })} required /></label><label>显示名称<input value={form.name} onChange={(event) => setForm({ ...form, name: event.target.value })} required /></label><label>访问码<input type="password" value={form.accessCode} onChange={(event) => setForm({ ...form, accessCode: event.target.value })} autoComplete="current-password" /></label><button type="submit" className="primary-button full" disabled={busy}>{busy ? <LoaderCircle className="spin" size={16} /> : <ArrowLeft size={16} />}进入工作台</button></form>{error && <div className="error-inline"><CircleAlert size={14} />{error}</div>}<small>生产环境仅允许白名单账号；模型密钥和 PowerPoint 凭证不会进入浏览器。</small></div><div className="auth-aside"><div className="auth-grid"><div><span>01</span><strong>Page-scoped</strong><p>稳定 page_id 贯穿聊天、版本与成本。</p></div><div><span>02</span><strong>Image-first</strong><p>视觉预览先行，再进入可编辑重建。</p></div><div><span>03</span><strong>QA-aware</strong><p>明确区分快速预览与 PPTX 权威状态。</p></div></div></div></div>;
}

function EmptyWorkspace({ user, onCreate, onLogout, busy, error }: { user: User; onCreate: () => Promise<void>; onLogout: () => void; busy: boolean; error: string }): JSX.Element {
  return <div className="empty-workspace"><div className="empty-card"><div className="brand-mark">F</div><span className="eyebrow">欢迎回来，{user.name}</span><h1>创建第一个 FastPPT 项目</h1><p>上传 slides.md 或从一个空白项目开始，浏览器会为每页建立合同、版本和预览状态。</p><button className="primary-button" onClick={() => void onCreate()} disabled={busy}><Plus size={16} />创建项目</button>{error && <div className="error-inline"><CircleAlert size={14} />{error}</div>}<button className="text-button" onClick={onLogout}><LogOut size={14} />退出登录</button></div></div>;
}

export default App;

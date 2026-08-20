import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type {
  ChangeEvent,
  FormEvent,
  JSX,
  PointerEvent as ReactPointerEvent,
} from "react";
import {
  Archive,
  ArrowDown,
  ArrowRight,
  ArrowUp,
  Check,
  ChevronDown,
  CircleAlert,
  Clock3,
  Copy,
  Download,
  Expand,
  FileText,
  Files,
  FolderKanban,
  GripVertical,
  History,
  LayoutDashboard,
  LoaderCircle,
  LogOut,
  MessageSquareText,
  PanelLeftClose,
  PanelLeftOpen,
  PanelRightClose,
  PanelRightOpen,
  Palette,
  PanelsTopLeft,
  Plus,
  RotateCcw,
  Send,
  Settings,
  ShieldCheck,
  Sparkles,
  Upload,
  WandSparkles,
  X,
} from "lucide-react";
import type {
  ConversationMessage,
  DocumentConflict,
  DocumentSource,
  DocumentStructurePlan,
  EditOperation,
  EventEnvelope,
  ExportJob,
  Page,
  PageVersion,
  Project,
  ProjectHistory,
  User,
  WorkflowMode,
} from "../shared/models.js";
import {
  archiveProject,
  authoritativeRenderUrl,
  cancelOperation,
  clearToken,
  confirmOperation,
  copyProject,
  createDocumentStructurePlan,
  createOfficePreview,
  createProject,
  createWorkSession,
  exportDownloadUrl,
  getExport,
  importSlides,
  listDocumentConflicts,
  listProjectDocuments,
  listProjectWorkSessions,
  listProjects,
  loadProjectHistory,
  loadProject,
  login,
  logout,
  me,
  openProjectSocket,
  restoreProject,
  parseProjectDocument,
  resolveDocumentConflict,
  retryFailedOperation,
  rollbackOperation,
  rollbackVersion,
  sendTurn,
  startExport,
  updateProjectDocumentContext,
  updateProjectSettings,
  updateWorkSession,
  uploadProjectDocument,
  visualPreviewUrl,
} from "./client.js";
import type { ProjectSocket } from "./client.js";

type ProductionMode = "import" | "pages" | "beautify";
type Scope = "single" | "multi" | "global";
type DrawerTab = "conversation" | "versions";
type DocumentState = "queued" | "parsing" | "ready" | "conflict" | "failed";
interface LocalDocument {
  id: string;
  name: string;
  kind: string;
  size: number;
  status: DocumentState;
  source: string;
  sha256: string;
  sourceVersionId: string;
  uploadedBy: string;
  createdAt: string;
  includedInContext: boolean;
  facts: string[];
  structure: string;
}
interface PageDraft {
  id: string;
  title: string;
  body: string;
  locked: boolean;
}

interface WorkspaceProjectPreference {
  currentPageId: string;
  productionMode: ProductionMode | null;
}

interface WorkspacePreferences {
  version: 2;
  lastProjectId: string;
  navCollapsed: boolean;
  chatPosition: "center" | "left" | "right";
  chatExpanded: boolean;
  projects: Record<string, WorkspaceProjectPreference>;
}

function defaultWorkspacePreferences(): WorkspacePreferences {
  return {
    version: 2,
    lastProjectId: "",
    navCollapsed: false,
    chatPosition: "center",
    chatExpanded: false,
    projects: {},
  };
}

function workspaceStorageKey(userId: string): string {
  return `fastppt.workspace.v2.${userId}`;
}

function readWorkspacePreferences(userId: string): WorkspacePreferences {
  try {
    const parsed = JSON.parse(
      window.localStorage.getItem(workspaceStorageKey(userId)) || "null",
    ) as Partial<WorkspacePreferences> | null;
    if (!parsed || parsed.version !== 2) return defaultWorkspacePreferences();
    return {
      ...defaultWorkspacePreferences(),
      ...parsed,
      chatPosition: ["center", "left", "right"].includes(
        parsed.chatPosition || "",
      )
        ? parsed.chatPosition!
        : "center",
      projects: parsed.projects || {},
    };
  } catch {
    return defaultWorkspacePreferences();
  }
}

const statusText: Record<Page["status"], string> = {
  untouched: "未修改",
  quick_preview: "快速预览",
  generating: "生成中",
  authoritative: "PPTX 权威渲染",
  svg_fallback: "SVG 回退",
  failed: "失败",
  rolled_back: "已回滚",
};
const statusTone: Record<Page["status"], string> = {
  untouched: "neutral",
  quick_preview: "blue",
  generating: "orange",
  authoritative: "green",
  svg_fallback: "purple",
  failed: "red",
  rolled_back: "orange",
};
const modeMeta: Record<ProductionMode, { label: string; description: string }> =
  {
    import: {
      label: "导入文档",
      description: "从 Markdown、Word 或 PDF 规划一套 PPT",
    },
    pages: { label: "按页录入", description: "逐页整理标题、正文和视觉层级" },
    beautify: {
      label: "PPT 美化",
      description: "保留结构，先预览视觉方向再重建",
    },
  };
function formatTime(value: string): string {
  return new Intl.DateTimeFormat("zh-CN", {
    hour: "2-digit",
    minute: "2-digit",
  }).format(new Date(value));
}
function currentVersion(page: Page): PageVersion {
  return (
    page.versions.find(
      (version) => version.versionId === page.currentVersionId,
    ) || page.versions[page.versions.length - 1]
  );
}
function versionDisplayStatus(version: PageVersion): Page["status"] {
  if (version.status === "failed" || version.status === "rejected")
    return "failed";
  if (version.status === "rendering") return "generating";
  if (version.previewKind === "pptx_authoritative") return "authoritative";
  if (version.previewKind === "svg_fallback") return "svg_fallback";
  return "quick_preview";
}
function localDocumentFromSource(document: DocumentSource): LocalDocument {
  const status: DocumentState =
    document.parseStatus === "failed" ||
    document.parseStatus === "blocked" ||
    document.parseStatus === "rejected"
      ? "failed"
      : document.parseStatus === "ready" ||
          document.parseStatus === "warning" ||
          document.parseStatus === "partial"
        ? "ready"
        : document.parseStatus === "parsing"
          ? "parsing"
          : "queued";
  return {
    id: document.documentId,
    name: document.fileName,
    kind: document.fileType.toUpperCase(),
    size: document.sizeBytes,
    status,
    source: document.warnings.length
      ? `服务端解析 · ${document.warnings[0]}`
      : "服务端解析",
    sha256: document.sha256,
    sourceVersionId:
      document.sourceVersionId || `docver_${document.documentId}`,
    uploadedBy: document.uploadedBy || "legacy-source",
    createdAt: document.createdAt || new Date().toISOString(),
    includedInContext: document.includedInContext,
    facts: document.facts.map((fact) => `${fact.value} · ${fact.location}`),
    structure:
      document.error ||
      document.warnings[0] ||
      (document.structure
        ? `${document.structure.headings.length} 个标题 · ${document.structure.characterCount} 字符`
        : "等待服务端解析"),
  };
}

function markConflictDocuments(
  sourceDocuments: LocalDocument[],
  conflicts: DocumentConflict[],
): LocalDocument[] {
  const conflictedIds = new Set(
    conflicts
      .filter(
        (conflict) =>
          conflict.status === "open" && conflict.severity === "blocking",
      )
      .flatMap((conflict) => conflict.values.map((value) => value.documentId)),
  );
  return sourceDocuments.map((document) => ({
    ...document,
    status: conflictedIds.has(document.id)
      ? "conflict"
      : document.status === "conflict"
        ? "ready"
        : document.status,
  }));
}

function documentsForMode(
  documents: DocumentSource[],
  mode: ProductionMode,
): DocumentSource[] {
  if (mode === "pages") return [];
  return documents.filter((document) =>
    mode === "beautify"
      ? document.fileType === "pptx"
      : document.fileType !== "pptx",
  );
}

function PagePreview({
  projectId,
  version,
}: {
  projectId: string;
  version: PageVersion;
}): JSX.Element {
  const [authorityFailed, setAuthorityFailed] = useState(false);
  const artifactId =
    version.previewKind === "pptx_authoritative"
      ? version.pptxPageRenderId
      : null;
  useEffect(() => setAuthorityFailed(false), [artifactId]);
  if (artifactId && !authorityFailed)
    return (
      <div className="slide-render authoritative-render">
        <img
          src={authoritativeRenderUrl(projectId, artifactId)}
          alt={`PowerPoint 权威渲染 ${version.versionId}`}
          data-artifact-id={artifactId}
          onError={() => setAuthorityFailed(true)}
        />
      </div>
    );
  if (authorityFailed)
    return (
      <div
        className="slide-render authority-load-fallback"
        data-preview-kind="authority-load-fallback"
      >
        <div className="authority-fallback-note">
          <CircleAlert size={14} />
          权威 PNG 加载失败，显示同版本 SVG 回退
        </div>
        <div dangerouslySetInnerHTML={{ __html: version.previewSvg }} />
      </div>
    );
  return (
    <div
      className="slide-render"
      data-preview-kind={version.previewKind}
      dangerouslySetInnerHTML={{ __html: version.previewSvg }}
    />
  );
}

function App(): JSX.Element {
  const [user, setUser] = useState<User | null>(null);
  const [projects, setProjects] = useState<Project[]>([]);
  const [project, setProject] = useState<Project | null>(null);
  const [messages, setMessages] = useState<ConversationMessage[]>([]);
  const [projectHistory, setProjectHistory] = useState<ProjectHistory>({
    messages: [],
    operations: [],
    sessions: [],
  });
  const [currentPageId, setCurrentPageId] = useState("");
  const [selectedPageIds, setSelectedPageIds] = useState<string[]>([]);
  const [scope, setScope] = useState<Scope>("single");
  const [productionMode, setProductionMode] = useState<ProductionMode | null>(
    null,
  );
  const [documents, setDocuments] = useState<LocalDocument[]>([]);
  const [documentIntent, setDocumentIntent] = useState<
    "create" | "modify" | "reference"
  >("modify");
  const [modeSessions, setModeSessions] = useState<
    Partial<Record<ProductionMode, string>>
  >({});
  const [documentConflicts, setDocumentConflicts] = useState<
    DocumentConflict[]
  >([]);
  const [pageBudget, setPageBudget] = useState(8);
  const [inheritTheme, setInheritTheme] = useState(true);
  const [structurePlan, setStructurePlan] =
    useState<DocumentStructurePlan | null>(null);
  const [draftPages, setDraftPages] = useState<PageDraft[]>([]);
  const [input, setInput] = useState("");
  const [pendingOperation, setPendingOperation] =
    useState<EditOperation | null>(null);
  const [busy, setBusy] = useState(false);
  const [loading, setLoading] = useState(true);
  const [loginForm, setLoginForm] = useState({
    email: "demo@fastppt.local",
    name: "Demo Editor",
    accessCode: "",
  });
  const [error, setError] = useState("");
  const [eventNote, setEventNote] = useState("");
  const [navCollapsed, setNavCollapsed] = useState(false);
  const [showHistory, setShowHistory] = useState(false);
  const [drawerTab, setDrawerTab] = useState<DrawerTab>("versions");
  const [compareVersionId, setCompareVersionId] = useState("");
  const [exportJob, setExportJob] = useState<ExportJob | null>(null);
  const [projectMenuOpen, setProjectMenuOpen] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [sensitiveMode, setSensitiveMode] = useState(false);
  const [chatExpanded, setChatExpanded] = useState(false);
  const [chatPosition, setChatPosition] = useState<"center" | "left" | "right">(
    "center",
  );
  const [dragging, setDragging] = useState(false);
  const [preferencesReady, setPreferencesReady] = useState(false);
  const socketRef = useRef<ProjectSocket | null>(null);
  const activeProjectIdRef = useRef("");
  const uploadRef = useRef<HTMLInputElement | null>(null);
  const pageUploadRef = useRef<HTMLInputElement | null>(null);
  const dragOriginRef = useRef<{ x: number; y: number } | null>(null);
  const sessionCreationRef = useRef<
    Partial<Record<ProductionMode, Promise<string>>>
  >({});
  const workspacePreferencesRef = useRef<WorkspacePreferences>(
    defaultWorkspacePreferences(),
  );
  const workflowModeByProduction: Record<ProductionMode, WorkflowMode> = {
    import: "document_import",
    pages: "page_entry",
    beautify: "ppt_beautify",
  };
  const visiblePages = useMemo(
    () =>
      project?.pages
        .filter((page) => !page.archived)
        .sort((a, b) => a.orderIndex - b.orderIndex) || [],
    [project],
  );
  const activePage =
    visiblePages.find((page) => page.pageId === currentPageId) ||
    visiblePages[0];
  const activeVersion = activePage ? currentVersion(activePage) : null;
  const compareVersion = activePage?.versions.find(
    (version) => version.versionId === compareVersionId,
  );
  const displayedVersion = compareVersion || activeVersion;
  const displayedStatus = compareVersion
    ? versionDisplayStatus(compareVersion)
    : activePage?.status || "untouched";
  const refreshProject = useCallback(
    async (projectId: string, preferredPageId?: string) => {
    if (
      activeProjectIdRef.current &&
      activeProjectIdRef.current !== projectId
    )
      return;
    const [nextProject, nextHistory] = await Promise.all([
      loadProject(projectId),
      loadProjectHistory(projectId),
    ]);
    if (activeProjectIdRef.current !== projectId) return;
    setProject(nextProject);
    setMessages(nextHistory.messages);
    setProjectHistory(nextHistory);
    const visible = nextProject.pages
      .filter((page) => !page.archived)
      .sort((a, b) => a.orderIndex - b.orderIndex);
      setCurrentPageId((previous) => {
        const requested = preferredPageId || previous;
        return visible.some((page) => page.pageId === requested)
          ? requested
          : visible[0]?.pageId || "";
      });
    },
    [],
  );
  const restoreModeState = useCallback(
    async (
      projectId: string,
      sessions: Awaited<ReturnType<typeof listProjectWorkSessions>>,
      preferredMode?: ProductionMode | null,
    ): Promise<void> => {
      const workflowToMode: Record<WorkflowMode, ProductionMode> = {
        document_import: "import",
        page_entry: "pages",
        ppt_beautify: "beautify",
      };
      const latestByMode = new Map<ProductionMode, (typeof sessions)[number]>();
      for (const session of sessions) {
        const mode = workflowToMode[session.workflowMode];
        if (!latestByMode.has(mode)) latestByMode.set(mode, session);
      }
      const nextModeSessions: Partial<Record<ProductionMode, string>> = {};
      latestByMode.forEach((session, mode) => {
        nextModeSessions[mode] = session.sessionId;
      });
      setModeSessions(nextModeSessions);
      const mode =
        preferredMode ||
        (sessions[0] ? workflowToMode[sessions[0].workflowMode] : null);
      setProductionMode(mode);
      if (!mode) {
        setDocuments([]);
        setDocumentConflicts([]);
        return;
      }
      const session = latestByMode.get(mode);
      if (!session) {
        setDocuments([]);
        setDocumentConflicts([]);
        return;
      }
      setDocumentIntent(session.intent);
      setPageBudget(session.pageBudget || 8);
      setInheritTheme(session.inheritTheme !== false);
      setStructurePlan(session.structurePlan || null);
      const [sources, conflicts] = await Promise.all([
        listProjectDocuments(projectId),
        listDocumentConflicts(projectId, session.sessionId),
      ]);
      setDocuments(
        markConflictDocuments(
          documentsForMode(sources, mode).map(localDocumentFromSource),
          conflicts,
        ),
      );
      setDocumentConflicts(conflicts);
    },
    [],
  );
  const ensureModeSession = useCallback(
    async (mode: ProductionMode): Promise<string> => {
      if (!project) throw new Error("请先选择项目。");
      const existing = modeSessions[mode];
      if (existing) return existing;
      const pending = sessionCreationRef.current[mode];
      if (pending) return pending;
      const promise = createWorkSession(
        project.projectId,
        workflowModeByProduction[mode],
        documentIntent,
      ).then((session) => {
        setModeSessions((current) => ({
          ...current,
          [mode]: session.sessionId,
        }));
        return session.sessionId;
      });
      sessionCreationRef.current[mode] = promise;
      try {
        return await promise;
      } finally {
        delete sessionCreationRef.current[mode];
      }
    },
    [documentIntent, modeSessions, project],
  );
  const connectProject = useCallback(
    async (nextProject: Project) => {
      socketRef.current?.close();
      const socket = await openProjectSocket(
        nextProject.projectId,
        (event: EventEnvelope) => {
          setEventNote(event.type.replaceAll(".", " · "));
          if (
            [
              "page.version.created",
              "preview.quick.ready",
              "preview.pptx.ready",
              "render.warning",
              "edit.failed",
              "edit.completed",
              "edit.rolled_back",
              "page.archived",
              "page.restored",
              "page.reordered",
              "page.split",
            ].includes(event.type)
          )
            void refreshProject(nextProject.projectId);
        },
      );
      if (!socket) return;
      if (activeProjectIdRef.current !== nextProject.projectId) {
        socket.close();
        return;
      }
      socketRef.current = socket;
    },
    [refreshProject],
  );
  const selectProject = useCallback(
    async (
      nextProject: Project,
      explicitPreference?: WorkspaceProjectPreference,
    ) => {
      activeProjectIdRef.current = nextProject.projectId;
      socketRef.current?.close();
      socketRef.current = null;
      setModeSessions({});
      const preference =
        explicitPreference ||
        workspacePreferencesRef.current.projects[nextProject.projectId];
      setProject(nextProject);
      setSensitiveMode(nextProject.settings?.sensitiveMode === true);
      setCurrentPageId(
        nextProject.pages.some(
          (page) =>
            !page.archived && page.pageId === preference?.currentPageId,
        )
          ? preference!.currentPageId
          : nextProject.pages.find((page) => !page.archived)?.pageId || "",
      );
      setSelectedPageIds([]);
      setScope("single");
      sessionCreationRef.current = {};
      await refreshProject(nextProject.projectId, preference?.currentPageId);
      if (activeProjectIdRef.current !== nextProject.projectId) return;
      const sessions = await listProjectWorkSessions(nextProject.projectId);
      if (activeProjectIdRef.current !== nextProject.projectId) return;
      await restoreModeState(
        nextProject.projectId,
        sessions,
        preference?.productionMode,
      );
      if (activeProjectIdRef.current !== nextProject.projectId) return;
      setDraftPages(
        nextProject.pages
          .filter((page) => !page.archived)
          .sort((a, b) => a.orderIndex - b.orderIndex)
          .map((page) => ({
            id: page.pageId,
            title: page.title,
            body: page.body,
            locked: page.locked,
          })),
      );
      setCompareVersionId("");
      setPendingOperation(null);
      await connectProject(nextProject);
    },
    [connectProject, refreshProject, restoreModeState],
  );
  useEffect(() => {
    let cancelled = false;
    async function bootstrap(): Promise<void> {
      try {
        const nextUser = await me();
        if (cancelled || !nextUser) return;
        setUser(nextUser);
        const preferences = readWorkspacePreferences(nextUser.userId);
        workspacePreferencesRef.current = preferences;
        setNavCollapsed(preferences.navCollapsed);
        setChatPosition(preferences.chatPosition);
        setChatExpanded(preferences.chatExpanded);
        const nextProjects = await listProjects();
        setProjects(nextProjects);
        const selected =
          nextProjects.find(
            (item) =>
              item.projectId === preferences.lastProjectId &&
              item.status !== "archived",
          ) || nextProjects.find((item) => item.status !== "archived");
        if (selected)
          await selectProject(selected, preferences.projects[selected.projectId]);
        if (!cancelled) setPreferencesReady(true);
      } catch {
        clearToken();
      } finally {
        if (!cancelled) setLoading(false);
      }
    }
    void bootstrap();
    return () => {
      cancelled = true;
      socketRef.current?.close();
    };
  }, [selectProject]);
  useEffect(() => () => socketRef.current?.close(), []);
  useEffect(() => {
    if (!preferencesReady || !user || !project) return;
    const previous = workspacePreferencesRef.current;
    const next: WorkspacePreferences = {
      ...previous,
      version: 2,
      lastProjectId: project.projectId,
      navCollapsed,
      chatPosition,
      chatExpanded,
      projects: {
        ...previous.projects,
        [project.projectId]: {
          currentPageId,
          productionMode,
        },
      },
    };
    workspacePreferencesRef.current = next;
    try {
      window.localStorage.setItem(
        workspaceStorageKey(user.userId),
        JSON.stringify(next),
      );
    } catch {
      /* The workspace remains usable when browser storage is unavailable. */
    }
  }, [
    chatExpanded,
    chatPosition,
    currentPageId,
    navCollapsed,
    preferencesReady,
    productionMode,
    project,
    user,
  ]);

  async function onLogin(event: FormEvent): Promise<void> {
    event.preventDefault();
    setBusy(true);
    setError("");
    try {
      const result = await login(
        loginForm.email,
        loginForm.name,
        loginForm.accessCode,
      );
      setUser(result.user);
      const preferences = readWorkspacePreferences(result.user.userId);
      workspacePreferencesRef.current = preferences;
      setNavCollapsed(preferences.navCollapsed);
      setChatPosition(preferences.chatPosition);
      setChatExpanded(preferences.chatExpanded);
      const nextProjects = await listProjects();
      setProjects(nextProjects);
      const selected =
        nextProjects.find(
          (item) =>
            item.projectId === preferences.lastProjectId &&
            item.status !== "archived",
        ) || nextProjects.find((item) => item.status !== "archived");
      if (selected)
        await selectProject(selected, preferences.projects[selected.projectId]);
      setPreferencesReady(true);
    } catch (caught) {
      setError((caught as Error).message);
    } finally {
      setBusy(false);
      setLoading(false);
    }
  }
  async function onCreateProject(): Promise<void> {
    const name = window.prompt("项目名称", "我的 FastPPT 项目");
    if (!name?.trim()) return;
    setBusy(true);
    setError("");
    try {
      const next = await createProject({ name: name.trim() });
      setProjects((current) => [next, ...current]);
      await selectProject(next);
    } catch (caught) {
      setError((caught as Error).message);
    } finally {
      setBusy(false);
    }
  }
  async function onCopyProject(): Promise<void> {
    if (!project) return;
    setBusy(true);
    setError("");
    try {
      const next = await copyProject(project.projectId);
      setProjects((current) => [next, ...current]);
      await selectProject(next);
      setProjectMenuOpen(false);
    } catch (caught) {
      setError((caught as Error).message);
    } finally {
      setBusy(false);
    }
  }
  async function onArchiveProject(): Promise<void> {
    if (!project) return;
    setBusy(true);
    setError("");
    try {
      const archived = await archiveProject(project.projectId);
      setProjects((current) =>
        current.map((item) =>
          item.projectId === archived.projectId ? archived : item,
        ),
      );
      const next = projects.find(
        (item) =>
          item.projectId !== archived.projectId && item.status !== "archived",
      );
      if (next) await selectProject(next);
      else setProject(null);
      setProjectMenuOpen(false);
    } catch (caught) {
      setError((caught as Error).message);
    } finally {
      setBusy(false);
    }
  }
  async function onRestoreProject(projectId: string): Promise<void> {
    setBusy(true);
    try {
      const restored = await restoreProject(projectId);
      setProjects((current) =>
        current.map((item) =>
          item.projectId === restored.projectId ? restored : item,
        ),
      );
      await selectProject(restored);
      setProjectMenuOpen(false);
    } catch (caught) {
      setError((caught as Error).message);
    } finally {
      setBusy(false);
    }
  }
  async function onLogout(): Promise<void> {
    try {
      await logout();
    } catch {
      clearToken();
    }
    socketRef.current?.close();
    activeProjectIdRef.current = "";
    setPreferencesReady(false);
    setUser(null);
    setProject(null);
  }
  async function onImportFiles(
    event: ChangeEvent<HTMLInputElement>,
  ): Promise<void> {
    const files = Array.from(event.target.files || []);
    event.target.value = "";
    if (!files.length || !project || !productionMode) return;
    const allowed =
      productionMode === "beautify" ? [".pptx"] : [".md", ".docx", ".pdf"];
    const accepted = files.filter((file) =>
      allowed.some((ext) => file.name.toLowerCase().endsWith(ext)),
    );
    const rejected = files.filter((file) => !accepted.includes(file));
    if (rejected.length)
      setError(
        `已忽略不支持的文件：${rejected.map((file) => file.name).join("、")}`,
      );
    if (!accepted.length) return;
    setBusy(true);
    try {
      const sessionId = await ensureModeSession(productionMode);
      for (const file of accepted) {
        const pending: LocalDocument = {
          id: `${file.name}-${file.lastModified}`,
          name: file.name,
          kind: file.name.split(".").pop()?.toUpperCase() || "FILE",
          size: file.size,
          status: "parsing",
          source: "uploading",
          sha256: "",
          sourceVersionId: "pending",
          uploadedBy: user?.userId || "pending",
          createdAt: new Date().toISOString(),
          includedInContext: true,
          facts: [],
          structure: "服务端正在安全解析",
        };
        setDocuments((current) => [
          ...current.filter((document) => document.name !== file.name),
          pending,
        ]);
        const uploaded = await uploadProjectDocument(
          project.projectId,
          sessionId,
          file,
        );
        const parsed = await parseProjectDocument(
          project.projectId,
          uploaded.documentId,
        );
        setDocuments((current) => [
          ...current.filter(
            (document) =>
              document.name !== file.name &&
              document.id !== uploaded.documentId,
          ),
          localDocumentFromSource(parsed),
        ]);
      }
      const conflicts = await listDocumentConflicts(
        project.projectId,
        sessionId,
      );
      setDocuments((current) => markConflictDocuments(current, conflicts));
      setDocumentConflicts(conflicts);
      setEventNote("document · ready · 来源、结构和事实已服务端登记");
    } catch (caught) {
      setError((caught as Error).message);
    } finally {
      setBusy(false);
    }
  }
  async function onToggleDocumentContext(
    document: LocalDocument,
  ): Promise<void> {
    if (!project) return;
    setBusy(true);
    try {
      const updated = await updateProjectDocumentContext(
        project.projectId,
        document.id,
        !document.includedInContext,
      );
      setDocuments((current) =>
        current.map((item) =>
          item.id === document.id ? localDocumentFromSource(updated) : item,
        ),
      );
      const sessionId = productionMode
        ? modeSessions[productionMode]
        : undefined;
      if (sessionId) {
        const conflicts = await listDocumentConflicts(
          project.projectId,
          sessionId,
        );
        setDocuments((current) => markConflictDocuments(current, conflicts));
        setDocumentConflicts(conflicts);
      }
    } catch (caught) {
      setError((caught as Error).message);
    } finally {
      setBusy(false);
    }
  }
  async function onResolveDocumentConflict(
    conflict: DocumentConflict,
    resolution: {
      action: "prefer_source" | "keep_both" | "ignore";
      selectedValue?: string;
      selectedDocumentIds?: string[];
    },
  ): Promise<void> {
    if (!project) return;
    setBusy(true);
    try {
      const resolved = await resolveDocumentConflict(
        project.projectId,
        conflict.conflictId,
        resolution,
      );
      setDocumentConflicts((current) =>
        current.map((item) =>
          item.conflictId === conflict.conflictId ? resolved : item,
        ),
      );
    } catch (caught) {
      setError((caught as Error).message);
    } finally {
      setBusy(false);
    }
  }
  function onChangeDocumentIntent(
    intent: "create" | "modify" | "reference",
  ): void {
    setDocumentIntent(intent);
    if (intent !== "create") setPageBudget(Math.max(1, visiblePages.length));
    setStructurePlan(null);
    if (!project || !productionMode) return;
    const sessionId = modeSessions[productionMode];
    if (!sessionId) return;
    void updateWorkSession(project.projectId, sessionId, { intent }).catch(
      (caught: unknown) => setError((caught as Error).message),
    );
  }
  async function onCreateStructurePlan(): Promise<void> {
    if (!project) return;
    setBusy(true);
    setError("");
    try {
      const sessionId = await ensureModeSession("import");
      await updateWorkSession(project.projectId, sessionId, {
        intent: documentIntent,
        pageBudget,
        inheritTheme,
      });
      const result = await createDocumentStructurePlan(
        project.projectId,
        sessionId,
        { pageBudget, inheritTheme },
      );
      setStructurePlan(result.structurePlan);
      setEventNote("plan · created · 文档结构计划已登记");
      if (result.operation.status === "planned") {
        setPendingOperation(result.operation);
        setChatExpanded(true);
      }
    } catch (caught) {
      setError((caught as Error).message);
    } finally {
      setBusy(false);
    }
  }
  function startMode(mode: ProductionMode): void {
    setProductionMode(mode);
    setError("");
    if (mode === "pages" && !draftPages.length)
      setDraftPages(
        visiblePages.map((page) => ({
          id: page.pageId,
          title: page.title,
          body: page.body,
          locked: page.locked,
        })),
      );
    if (!project) return;
    void ensureModeSession(mode)
      .then(async (sessionId) => {
        await updateWorkSession(project.projectId, sessionId);
        const [sources, conflicts] = await Promise.all([
          listProjectDocuments(project.projectId),
          listDocumentConflicts(project.projectId, sessionId),
        ]);
        setDocuments(
          markConflictDocuments(
            documentsForMode(sources, mode).map(localDocumentFromSource),
            conflicts,
          ),
        );
        setDocumentConflicts(conflicts);
      })
      .catch((caught: unknown) => setError((caught as Error).message));
  }
  function updateDraft(
    id: string,
    field: "title" | "body",
    value: string,
  ): void {
    setDraftPages((current) =>
      current.map((page) =>
        page.id === id ? { ...page, [field]: value } : page,
      ),
    );
  }
  function addDraftPage(): void {
    setDraftPages((current) => [
      ...current,
      { id: `draft_${Date.now()}`, title: "新页面", body: "", locked: false },
    ]);
  }
  function removeDraftPage(id: string): void {
    setDraftPages((current) => current.filter((page) => page.id !== id));
  }
  function moveDraftPage(id: string, offset: -1 | 1): void {
    setDraftPages((current) => {
      const index = current.findIndex((page) => page.id === id);
      const target = index + offset;
      if (index < 0 || target < 0 || target >= current.length) return current;
      const next = [...current];
      [next[index], next[target]] = [next[target], next[index]];
      return next;
    });
  }
  function reorderDraftPage(sourceId: string, targetId: string): void {
    if (sourceId === targetId) return;
    setDraftPages((current) => {
      const sourceIndex = current.findIndex((page) => page.id === sourceId);
      const targetIndex = current.findIndex((page) => page.id === targetId);
      if (sourceIndex < 0 || targetIndex < 0) return current;
      const next = [...current];
      const [moved] = next.splice(sourceIndex, 1);
      next.splice(targetIndex, 0, moved);
      return next;
    });
  }
  function pasteDraftPages(): void {
    const value = window.prompt(
      "粘贴多页内容。使用 --- 分隔页面，每页第一行作为标题。",
      "第一页标题\n第一页正文\n---\n第二页标题\n第二页正文",
    );
    if (!value?.trim()) return;
    const additions = value.split(/\n\s*---\s*\n/g).map((block, index) => {
      const [title, ...body] = block.trim().split(/\r?\n/);
      return {
        id: `draft_${Date.now()}_${index}`,
        title: title?.trim() || `第 ${index + 1} 页`,
        body: body.join("\n").trim(),
        locked: false,
      };
    });
    setDraftPages((current) => [...current, ...additions]);
  }
  async function applyDraftPages(): Promise<void> {
    if (!project || !draftPages.length) return;
    setBusy(true);
    setError("");
    try {
      const markdown = draftPages
        .map(
          (draft) =>
            `# ${draft.title.trim() || "未命名页面"}\n\n${draft.body.trim()}`,
        )
        .join("\n\n---\n\n");
      const file = new File([markdown], "page-drafts.md", {
        type: "text/markdown;charset=utf-8",
      });
      const next = await importSlides(project.projectId, file);
      setProjects((current) =>
        current.map((item) =>
          item.projectId === next.projectId ? next : item,
        ),
      );
      await selectProject(next);
      setEventNote(`page_entry · applied · 已确认生成 ${draftPages.length} 页`);
    } catch (caught) {
      setError((caught as Error).message);
    } finally {
      setBusy(false);
    }
  }
  function togglePageSelection(pageId: string): void {
    setSelectedPageIds((current) =>
      current.includes(pageId)
        ? current.filter((id) => id !== pageId)
        : [...current, pageId],
    );
    setScope("multi");
  }
  async function onSend(event: FormEvent): Promise<void> {
    event.preventDefault();
    if (!project || !input.trim() || busy) return;
    const pageIds =
      scope === "single"
        ? activePage
          ? [activePage.pageId]
          : []
        : scope === "multi"
          ? selectedPageIds
          : [];
    if (scope === "multi" && pageIds.length === 0) {
      setError("请先勾选要修改的页面。");
      return;
    }
    setBusy(true);
    setError("");
    try {
      const result = await sendTurn(project.projectId, {
        deckRevisionId: project.currentDeckRevisionId,
        target: { mode: scope, pageIds },
        message: input.trim(),
        clientRevision: project.pages.reduce(
          (sum, page) => sum + page.versions.length,
          0,
        ),
        sessionId: productionMode ? modeSessions[productionMode] : undefined,
      });
      setInput("");
      setMessages((current) => [
        ...current,
        ...result.messages.filter(
          (message) =>
            !current.some(
              (existing) => existing.messageId === message.messageId,
            ),
        ),
      ]);
      setPendingOperation(
        result.operation.confirmationRequired ? result.operation : null,
      );
      if (result.operation.confirmationRequired) setChatExpanded(true);
      await refreshProject(project.projectId);
    } catch (caught) {
      setError((caught as Error).message);
    } finally {
      setBusy(false);
    }
  }
  async function onConfirm(): Promise<void> {
    if (!project || !pendingOperation) return;
    setBusy(true);
    setError("");
    try {
      await confirmOperation(project.projectId, pendingOperation.operationId);
      setPendingOperation(null);
      await refreshProject(project.projectId);
    } catch (caught) {
      setError((caught as Error).message);
    } finally {
      setBusy(false);
    }
  }
  async function onCancel(): Promise<void> {
    if (!project || !pendingOperation) return;
    setBusy(true);
    try {
      await cancelOperation(project.projectId, pendingOperation.operationId);
      setPendingOperation(null);
      await refreshProject(project.projectId);
    } catch (caught) {
      setError((caught as Error).message);
    } finally {
      setBusy(false);
    }
  }
  async function onRollback(versionId: string): Promise<void> {
    if (!project) return;
    setBusy(true);
    setError("");
    try {
      await rollbackVersion(project.projectId, versionId);
      setCompareVersionId("");
      await refreshProject(project.projectId);
    } catch (caught) {
      setError((caught as Error).message);
    } finally {
      setBusy(false);
    }
  }
  async function onRollbackOperation(operationId: string): Promise<void> {
    if (!project) return;
    setBusy(true);
    try {
      await rollbackOperation(project.projectId, operationId);
      setEventNote("edit · rolled back · 本次操作已整组回滚");
      await refreshProject(project.projectId);
    } catch (caught) {
      setError((caught as Error).message);
    } finally {
      setBusy(false);
    }
  }
  async function onRetryOperation(operationId: string): Promise<void> {
    if (!project) return;
    setBusy(true);
    try {
      await retryFailedOperation(project.projectId, operationId);
      setEventNote("失败页面已创建独立重试任务");
      await refreshProject(project.projectId);
    } catch (caught) {
      setError((caught as Error).message);
    } finally {
      setBusy(false);
    }
  }
  async function onExport(): Promise<void> {
    if (!project || busy) return;
    setBusy(true);
    setError("");
    try {
      const job = await startExport(project.projectId);
      setExportJob(job);
      const poll = async (): Promise<void> => {
        const next = await getExport(project.projectId, job.exportId);
        setExportJob(next);
        if (next.status === "running" || next.status === "queued")
          window.setTimeout(() => void poll(), 700);
      };
      window.setTimeout(() => void poll(), 500);
    } catch (caught) {
      setError((caught as Error).message);
    } finally {
      setBusy(false);
    }
  }
  async function onSensitiveChange(value: boolean): Promise<void> {
    if (!project) return;
    setBusy(true);
    setError("");
    try {
      const updated = await updateProjectSettings(project.projectId, {
        sensitiveMode: value,
      });
      setProject(updated);
      setSensitiveMode(updated.settings?.sensitiveMode === true);
      setEventNote(
        value
          ? "settings · sensitive_mode.enabled"
          : "settings · sensitive_mode.disabled",
      );
    } catch (caught) {
      setError((caught as Error).message);
    } finally {
      setBusy(false);
    }
  }
  async function onOfficePreview(job: ExportJob): Promise<void> {
    if (!project) return;
    try {
      const grant = await createOfficePreview(project.projectId, job.exportId);
      window.open(grant.officeViewerUrl, "_blank", "noopener,noreferrer");
    } catch (caught) {
      setError((caught as Error).message);
    }
  }
  function onChatPointerDown(event: ReactPointerEvent<HTMLDivElement>): void {
    dragOriginRef.current = { x: event.clientX, y: event.clientY };
    event.currentTarget.setPointerCapture(event.pointerId);
    setDragging(true);
    event.preventDefault();
  }
  function onChatPointerMove(event: ReactPointerEvent<HTMLDivElement>): void {
    if (!dragging || !dragOriginRef.current) return;
    const dx = event.clientX - dragOriginRef.current.x;
    if (Math.abs(dx) > 70) {
      setChatPosition(dx < 0 ? "left" : "right");
      dragOriginRef.current = { x: event.clientX, y: event.clientY };
    }
  }
  function onChatPointerUp(event: ReactPointerEvent<HTMLDivElement>): void {
    setDragging(false);
    dragOriginRef.current = null;
    event.currentTarget.releasePointerCapture?.(event.pointerId);
  }

  if (loading)
    return (
      <div className="loading-screen">
        <LoaderCircle className="spin" size={24} />
        正在恢复工作区
      </div>
    );
  if (!user)
    return (
      <LoginScreen
        form={loginForm}
        setForm={setLoginForm}
        onSubmit={onLogin}
        busy={busy}
        error={error}
      />
    );
  if (!project)
    return (
      <EmptyWorkspace
        user={user}
        onCreate={onCreateProject}
        onLogout={() => void onLogout()}
        busy={busy}
        error={error}
      />
    );
  const targetDescription =
    scope === "single"
      ? `当前第 ${(activePage?.orderIndex ?? 0) + 1} 页`
      : scope === "multi"
        ? `已选 ${selectedPageIds.length} 页`
        : "全局相似页";
  return (
    <div className={`app-shell ${navCollapsed ? "nav-collapsed" : ""}`}>
      <aside className="mode-sidebar">
        <div className="brand-row">
          <div className="brand-mark">F</div>
          {!navCollapsed && (
            <div className="brand-copy">
              <strong>FastPPT</strong>
              <span>ONLINE WORKBENCH</span>
            </div>
          )}
          <button
            className="icon-button sidebar-toggle"
            title={navCollapsed ? "展开导航" : "收起导航"}
            aria-label={navCollapsed ? "展开导航" : "收起导航"}
            onClick={() => setNavCollapsed((value) => !value)}
          >
            {navCollapsed ? (
              <PanelLeftOpen size={18} />
            ) : (
              <PanelLeftClose size={18} />
            )}
          </button>
        </div>
        <nav className="mode-nav" aria-label="生产模式与工作台导航">
          {(["import", "pages", "beautify"] as ProductionMode[]).map((mode) => (
            <button
              key={mode}
              className={`nav-item ${productionMode === mode ? "active" : ""}`}
              onClick={() => startMode(mode)}
              title={modeMeta[mode].label}
            >
              <span className="nav-icon">
                {mode === "import" ? (
                  <Files size={18} />
                ) : mode === "pages" ? (
                  <PanelsTopLeft size={18} />
                ) : (
                  <WandSparkles size={18} />
                )}
              </span>
              {!navCollapsed && <span>{modeMeta[mode].label}</span>}
            </button>
          ))}
          <div className="nav-divider" />
          <button
            className="nav-item"
            onClick={() => setProjectMenuOpen((value) => !value)}
            title="我的项目"
          >
            <span className="nav-icon">
              <FolderKanban size={18} />
            </span>
            {!navCollapsed && <span>我的项目</span>}
          </button>
          <button
            className="nav-item"
            onClick={() => {
              setSettingsOpen(false);
              setProductionMode(null);
            }}
            title="设计风格"
          >
            <span className="nav-icon">
              <Palette size={18} />
            </span>
            {!navCollapsed && <span>设计风格</span>}
          </button>
          <button
            className={`nav-item ${settingsOpen ? "active" : ""}`}
            onClick={() => setSettingsOpen((value) => !value)}
            title="设置"
          >
            <span className="nav-icon">
              <Settings size={18} />
            </span>
            {!navCollapsed && <span>设置</span>}
          </button>
        </nav>
        {!navCollapsed && (
          <div className="sidebar-project">
            <span className="eyebrow">当前项目</span>
            <button
              className="project-name"
              onClick={() => setProjectMenuOpen((value) => !value)}
            >
              <FileText size={14} />
              {project.name}
              <ChevronDown size={14} />
            </button>
            <span className="project-meta">
              <span className="live-dot" />
              {visiblePages.length} 页 · {project.themeId}
            </span>
            {projectMenuOpen && (
              <ProjectMenu
                projects={projects}
                project={project}
                onSelect={(next) => {
                  void selectProject(next);
                  setProjectMenuOpen(false);
                }}
                onCopy={() => void onCopyProject()}
                onArchive={() => void onArchiveProject()}
                onRestore={(id) => void onRestoreProject(id)}
              />
            )}
          </div>
        )}
        <div className="sidebar-footer">
          {!navCollapsed && (
            <div className="user-chip">
              <div className="avatar">
                {user.name.slice(0, 1).toUpperCase()}
              </div>
              <div>
                <strong>{user.name}</strong>
                <span>{user.email}</span>
              </div>
            </div>
          )}
          <button
            className="icon-button logout-button"
            title="退出登录"
            onClick={() => void onLogout()}
          >
            <LogOut size={16} />
          </button>
        </div>
      </aside>
      <main className="main-stage">
        <header className="topbar">
          <div className="breadcrumbs">
            <span>工作台</span>
            <ArrowRight size={14} />
            <strong>{project.name}</strong>
          </div>
          <div className="top-actions">
            <span className="saved-state">
              <span className="save-dot" />
              已保存
            </span>
            <button
              className="outline-button"
              onClick={() => void onExport()}
              disabled={busy}
            >
              <Download size={15} />
              导出 PPTX
            </button>
            <button
              className="icon-button"
              title="页面全屏"
              onClick={() => document.documentElement.requestFullscreen?.()}
            >
              <Expand size={17} />
            </button>
          </div>
        </header>
        <section className="workspace-header">
          <div>
            <span className="eyebrow">
              {productionMode ? modeMeta[productionMode].label : "页面预览"}
            </span>
            <h1>
              {productionMode
                ? modeMeta[productionMode].description
                : activePage?.title || "选择一个页面开始"}
            </h1>
            <p>
              {productionMode
                ? "输入资料、确认影响范围，再开始生成。"
                : "当前页面保持稳定 page_id，版本和事实锚点由服务端保护。"}
            </p>
          </div>
          <div className="workspace-actions">
            <button
              className="history-trigger"
              onClick={() => setShowHistory(true)}
            >
              <History size={16} />
              历史
            </button>
            <span
              className={`status-pill ${statusTone[displayedStatus]}`}
              data-page-status={displayedStatus}
            >
              <span className="status-dot" />
              {compareVersion && "对比 · "}
              {statusText[displayedStatus]}
            </span>
          </div>
        </section>
        {settingsOpen && (
          <SettingsPanel
            sensitiveMode={sensitiveMode}
            onSensitiveChange={(value) => void onSensitiveChange(value)}
          />
        )}
        {productionMode && (
          <ModeWorkspace
            mode={productionMode}
            documents={documents}
            onUpload={() =>
              productionMode === "beautify"
                ? pageUploadRef.current?.click()
                : uploadRef.current?.click()
            }
            uploadRef={uploadRef}
            pageUploadRef={pageUploadRef}
            onFiles={onImportFiles}
            draftPages={draftPages}
            updateDraft={updateDraft}
            addDraftPage={addDraftPage}
            removeDraftPage={removeDraftPage}
            moveDraftPage={moveDraftPage}
            reorderDraftPage={reorderDraftPage}
            pasteDraftPages={pasteDraftPages}
            applyDraftPages={() => void applyDraftPages()}
            documentIntent={documentIntent}
            onIntentChange={onChangeDocumentIntent}
            pageBudget={pageBudget}
            onPageBudgetChange={setPageBudget}
            inheritTheme={inheritTheme}
            onInheritThemeChange={setInheritTheme}
            structurePlan={structurePlan}
            onCreateStructurePlan={() => void onCreateStructurePlan()}
            busy={busy}
            conflicts={documentConflicts}
            onResolveConflict={(conflict, value) =>
              void onResolveDocumentConflict(conflict, value)
            }
            onToggleContext={(document) =>
              void onToggleDocumentContext(document)
            }
          />
        )}
        <section
          className={`preview-workspace ${productionMode ? "preview-with-mode" : ""}`}
        >
          <div className="page-rail">
            <div className="rail-heading">
              <span>页面</span>
              <span className="muted-count">{visiblePages.length}</span>
              <button
                className="icon-button"
                title="新建项目"
                onClick={() => void onCreateProject()}
              >
                <Plus size={16} />
              </button>
            </div>
            <div className="page-list">
              {visiblePages.map((page) => (
                <PageRow
                  key={page.pageId}
                  page={page}
                  active={page.pageId === activePage?.pageId}
                  selected={selectedPageIds.includes(page.pageId)}
                  onClick={() => {
                    setCurrentPageId(page.pageId);
                    if (scope === "single") setSelectedPageIds([]);
                  }}
                  onToggle={() => togglePageSelection(page.pageId)}
                />
              ))}
            </div>
          </div>
          <div className="preview-column">
            <div className="canvas-toolbar">
              <div className="toolbar-left">
                <button className="tool-button active">
                  <LayoutDashboard size={16} />
                  预览
                </button>
                <button
                  className="tool-button"
                  onClick={() => setEventNote("页面合同已绑定当前版本")}
                >
                  <ShieldCheck size={16} />
                  合同
                </button>
              </div>
              <div className="toolbar-center">
                <span
                  className="version-label"
                  data-version-id={displayedVersion?.versionId || ""}
                >
                  {displayedVersion?.versionId || "ver_pending"}
                </span>
              </div>
              <div className="toolbar-right">
                <button
                  className="icon-button"
                  title="版本历史"
                  onClick={() => setShowHistory(true)}
                >
                  <History size={17} />
                </button>
                <button
                  className="icon-button"
                  title="撤销到上一版本"
                  onClick={() =>
                    activePage &&
                    activePage.versions.length > 1 &&
                    void onRollback(
                      activePage.versions[activePage.versions.length - 2]
                        .versionId,
                    )
                  }
                  disabled={!activePage || activePage.versions.length < 2}
                >
                  <RotateCcw size={17} />
                </button>
              </div>
            </div>
            <div className="preview-wrap">
              <div className="preview-frame">
                <div className="preview-chrome">
                  <span className="chrome-dot red" />
                  <span className="chrome-dot amber" />
                  <span className="chrome-dot green" />
                  <span className="preview-url">
                    fastppt.online / preview / {activePage?.pageId}
                  </span>
                  <button
                    className="icon-button subtle"
                    title="刷新预览"
                    onClick={() => void refreshProject(project.projectId)}
                  >
                    <RotateCcw size={14} />
                  </button>
                </div>
                {displayedVersion ? (
                  <PagePreview
                    projectId={project.projectId}
                    version={displayedVersion}
                  />
                ) : (
                  <div className="empty-slide">选择一个页面开始预览</div>
                )}
                <div className="preview-legend">
                  <span>
                    <span className="legend-line quick" />
                    快速预览
                  </span>
                  <span>
                    <span className="legend-line authority" />
                    PPTX 权威
                  </span>
                  <span>
                    <span className="legend-line fallback" />
                    SVG 回退
                  </span>
                </div>
              </div>
              <div className="preview-caption">
                <div>
                  <strong>第 {(activePage?.orderIndex ?? 0) + 1} 页</strong>
                  <span>{displayedVersion?.title || activePage?.title}</span>
                </div>
                <div className="caption-right">
                  <span className="editable-tag">
                    <Sparkles size={13} />
                    {displayedVersion?.editableLevel ||
                      activePage?.editableLevel}
                  </span>
                  {displayedVersion?.previewKind === "svg_fallback" && (
                    <span className="warning-caption">
                      <CircleAlert size={13} />
                      可能与 PPTX 有差异
                    </span>
                  )}
                </div>
              </div>
            </div>
          </div>
        </section>
        {exportJob && (
          <ExportBanner
            job={exportJob}
            project={project}
            sensitiveMode={sensitiveMode}
            onOfficePreview={() => void onOfficePreview(exportJob)}
            onDismiss={() => setExportJob(null)}
          />
        )}
      </main>
      <div
        className={`chat-capsule-wrap ${chatPosition} ${chatExpanded ? "expanded" : ""}`}
        onContextMenu={(event) => event.preventDefault()}
      >
        <div
          className="chat-panel"
          onPointerDown={onChatPointerDown}
          onPointerMove={onChatPointerMove}
          onPointerUp={onChatPointerUp}
        >
          <div className="chat-header">
            <div className="chat-drag-handle">
              <MessageSquareText size={16} />
              <div>
                <strong>AI 精修助手</strong>
                <span>
                  <span className="live-dot" />
                  在线 · {targetDescription}
                </span>
              </div>
            </div>
            <div className="chat-header-actions">
              <button
                className="icon-button"
                title={chatExpanded ? "收起聊天面板" : "展开聊天面板"}
                onPointerDown={(event) => event.stopPropagation()}
                onClick={() => setChatExpanded((value) => !value)}
              >
                {chatExpanded ? (
                  <ArrowDown size={16} />
                ) : (
                  <ArrowRight size={16} />
                )}
              </button>
              <button
                className="icon-button"
                title="打开历史抽屉"
                onPointerDown={(event) => event.stopPropagation()}
                onClick={() => setShowHistory(true)}
              >
                <PanelRightOpen size={16} />
              </button>
              <button
                className="icon-button chat-reset-position"
                title="恢复聊天默认位置"
                aria-label="恢复聊天默认位置"
                onPointerDown={(event) => event.stopPropagation()}
                onClick={() => setChatPosition("center")}
                disabled={chatPosition === "center"}
              >
                <RotateCcw size={16} />
              </button>
            </div>
          </div>
          <div
            className="chat-body"
            onPointerDown={(event) => event.stopPropagation()}
          >
            <div className="mode-tabs" role="tablist" aria-label="聊天操作范围">
              <button
                className={scope === "single" ? "active" : ""}
                onClick={() => setScope("single")}
              >
                当前页
              </button>
              <button
                className={scope === "multi" ? "active" : ""}
                onClick={() => setScope("multi")}
              >
                多页
              </button>
              <button
                className={scope === "global" ? "active" : ""}
                onClick={() => setScope("global")}
              >
                全局
              </button>
            </div>
            {chatExpanded && (
              <div className="chat-scroll">
                {messages.length === 0 ? (
                  <WelcomeMessage />
                ) : (
                  messages
                    .slice(-8)
                    .map((message) => (
                      <ChatMessage
                        key={message.messageId}
                        message={message}
                        busy={busy}
                        onRollbackOperation={onRollbackOperation}
                        onRetryOperation={onRetryOperation}
                      />
                    ))
                )}
                {pendingOperation && (
                  <ConfirmationCard
                    operation={pendingOperation}
                    onConfirm={() => void onConfirm()}
                    onCancel={() => void onCancel()}
                    busy={busy}
                  />
                )}
                {eventNote && (
                  <div className="event-note">
                    <Check size={12} />
                    {eventNote}
                  </div>
                )}
              </div>
            )}
            <div className="composer-wrap">
              <div className="composer-hint">
                <span>事实锚点已锁定</span>
                <span>·</span>
                <span>{activePage?.factAnchors.length || 0} 个数字</span>
              </div>
              <form className="composer" onSubmit={onSend}>
                <textarea
                  value={input}
                  onChange={(event) => setInput(event.target.value)}
                  placeholder={
                    scope === "single"
                      ? "请输入要执行的修改……"
                      : scope === "multi"
                        ? "描述选中页面的共同修改……"
                        : "描述要查找并修改的相似页……"
                  }
                  rows={2}
                  disabled={busy}
                />
                <div className="composer-actions">
                  <button
                    type="button"
                    className="icon-button subtle"
                    title="上传资料"
                    onClick={() =>
                      productionMode
                        ? productionMode === "beautify"
                          ? pageUploadRef.current?.click()
                          : uploadRef.current?.click()
                        : setError("请先选择生产模式，再上传资料。")
                    }
                  >
                    <Upload size={16} />
                  </button>
                  <span className="composer-model">
                    {busy ? (
                      <>
                        <LoaderCircle className="spin" size={14} />
                        处理中
                      </>
                    ) : (
                      "FastPPT planner"
                    )}
                  </span>
                  <button
                    type="submit"
                    className="send-button"
                    disabled={busy || !input.trim()}
                    title="发送消息"
                  >
                    <Send size={16} />
                  </button>
                </div>
              </form>
              {error && (
                <div className="error-inline">
                  <CircleAlert size={14} />
                  {error}
                </div>
              )}
            </div>
          </div>
        </div>
        <div className="chat-drag-note">拖动胶囊 · 松开自动吸附</div>
      </div>
      {eventNote && !chatExpanded && (
        <div
          className="event-note detached"
          style={{
            position: "fixed",
            zIndex: 34,
            left: "50%",
            bottom: 202,
            transform: "translateX(-50%)",
            padding: "7px 10px",
            border: "1px solid #2a5b53",
            borderRadius: 7,
            background: "#12312f",
            boxShadow: "0 8px 22px #02091288",
            whiteSpace: "nowrap",
          }}
        >
          <Check size={12} />
          {eventNote}
        </div>
      )}
      {showHistory && project && activePage && (
        <HistoryDrawer
          project={project}
          page={activePage}
          history={projectHistory}
          tab={drawerTab}
          onTab={setDrawerTab}
          compareId={compareVersionId}
          onCompare={setCompareVersionId}
          onRollback={onRollback}
          onRollbackOperation={onRollbackOperation}
          onLocate={(pageId, versionId) => {
            const targetPage = project.pages.find(
              (candidate) => candidate.pageId === pageId,
            );
            setCurrentPageId(pageId);
            setCompareVersionId(
              versionId && versionId !== targetPage?.currentVersionId
                ? versionId
                : "",
            );
          }}
          onClose={() => setShowHistory(false)}
        />
      )}
    </div>
  );
}

function ProjectMenu({
  projects,
  project,
  onSelect,
  onCopy,
  onArchive,
  onRestore,
}: {
  projects: Project[];
  project: Project;
  onSelect: (project: Project) => void;
  onCopy: () => void;
  onArchive: () => void;
  onRestore: (id: string) => void;
}): JSX.Element {
  return (
    <div className="project-menu">
      {projects
        .filter((item) => item.status !== "archived")
        .map((item) => (
          <button
            key={item.projectId}
            className={item.projectId === project.projectId ? "selected" : ""}
            onClick={() => onSelect(item)}
          >
            <FileText size={14} />
            {item.name}
          </button>
        ))}
      {projects
        .filter((item) => item.status === "archived")
        .map((item) => (
          <button
            key={item.projectId}
            onClick={() => onRestore(item.projectId)}
          >
            <RotateCcw size={14} />
            恢复 {item.name}
          </button>
        ))}
      <button onClick={onCopy}>
        <Copy size={14} />
        复制当前项目
      </button>
      <button onClick={onArchive}>
        <Archive size={14} />
        归档当前项目
      </button>
    </div>
  );
}
function ModeWorkspace({
  mode,
  documents,
  onUpload,
  uploadRef,
  pageUploadRef,
  onFiles,
  draftPages,
  updateDraft,
  addDraftPage,
  removeDraftPage,
  moveDraftPage,
  reorderDraftPage,
  pasteDraftPages,
  applyDraftPages,
  documentIntent,
  onIntentChange,
  pageBudget,
  onPageBudgetChange,
  inheritTheme,
  onInheritThemeChange,
  structurePlan,
  onCreateStructurePlan,
  busy,
  conflicts,
  onResolveConflict,
  onToggleContext,
}: {
  mode: ProductionMode;
  documents: LocalDocument[];
  onUpload: () => void;
  uploadRef: React.RefObject<HTMLInputElement | null>;
  pageUploadRef: React.RefObject<HTMLInputElement | null>;
  onFiles: (event: ChangeEvent<HTMLInputElement>) => void;
  draftPages: PageDraft[];
  updateDraft: (id: string, field: "title" | "body", value: string) => void;
  addDraftPage: () => void;
  removeDraftPage: (id: string) => void;
  moveDraftPage: (id: string, offset: -1 | 1) => void;
  reorderDraftPage: (sourceId: string, targetId: string) => void;
  pasteDraftPages: () => void;
  applyDraftPages: () => void;
  documentIntent: "create" | "modify" | "reference";
  onIntentChange: (intent: "create" | "modify" | "reference") => void;
  pageBudget: number;
  onPageBudgetChange: (value: number) => void;
  inheritTheme: boolean;
  onInheritThemeChange: (value: boolean) => void;
  structurePlan: DocumentStructurePlan | null;
  onCreateStructurePlan: () => void;
  busy: boolean;
  conflicts: DocumentConflict[];
  onResolveConflict: (
    conflict: DocumentConflict,
    resolution: {
      action: "prefer_source" | "keep_both" | "ignore";
      selectedValue?: string;
      selectedDocumentIds?: string[];
    },
  ) => void;
  onToggleContext: (document: LocalDocument) => void;
}): JSX.Element {
  const [draggedDraftId, setDraggedDraftId] = useState<string | null>(null);
  const draggedDraftRef = useRef<string | null>(null);
  const draggedDraftTargetRef = useRef<string | null>(null);
  const moveDraggedDraftAtPoint = (clientX: number, clientY: number): void => {
    const sourceId = draggedDraftRef.current;
    if (!sourceId) return;
    const target = document
      .elementFromPoint(clientX, clientY)
      ?.closest<HTMLElement>("[data-draft-id]");
    const targetId = target?.dataset.draftId;
    if (
      targetId &&
      targetId !== sourceId &&
      targetId !== draggedDraftTargetRef.current
    ) {
      draggedDraftTargetRef.current = targetId;
      reorderDraftPage(sourceId, targetId);
    }
  };
  const endDraftDrag = (): void => {
    draggedDraftRef.current = null;
    draggedDraftTargetRef.current = null;
    setDraggedDraftId(null);
  };
  useEffect(() => {
    if (!draggedDraftId) return;
    const onPointerMove = (event: PointerEvent): void =>
      moveDraggedDraftAtPoint(event.clientX, event.clientY);
    const onPointerEnd = (): void => endDraftDrag();
    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerEnd, { once: true });
    window.addEventListener("pointercancel", onPointerEnd, { once: true });
    return () => {
      window.removeEventListener("pointermove", onPointerMove);
      window.removeEventListener("pointerup", onPointerEnd);
      window.removeEventListener("pointercancel", onPointerEnd);
    };
  }, [draggedDraftId]);
  const onDraftPointerDown = (
    event: ReactPointerEvent<HTMLButtonElement>,
    pageId: string,
  ): void => {
    draggedDraftRef.current = pageId;
    draggedDraftTargetRef.current = null;
    setDraggedDraftId(pageId);
    event.currentTarget.setPointerCapture(event.pointerId);
    event.preventDefault();
  };
  const onDraftPointerMove = (
    event: ReactPointerEvent<HTMLButtonElement>,
  ): void => {
    moveDraggedDraftAtPoint(event.clientX, event.clientY);
  };
  const onDraftPointerEnd = (
    event: ReactPointerEvent<HTMLButtonElement>,
  ): void => {
    endDraftDrag();
    if (event.currentTarget.hasPointerCapture(event.pointerId))
      event.currentTarget.releasePointerCapture(event.pointerId);
  };
  return (
    <section className="mode-workspace">
      <input
        ref={uploadRef}
        type="file"
        hidden
        multiple
        accept=".md,.docx,.pdf"
        onChange={onFiles}
      />
      <input
        ref={pageUploadRef}
        type="file"
        hidden
        accept=".pptx"
        onChange={onFiles}
      />
      <div className="mode-workspace-head">
        <div>
          <span className="eyebrow">{modeMeta[mode].label}</span>
          <h2>
            {mode === "import"
              ? "资料池"
              : mode === "pages"
                ? "逐页内容草稿"
                : "原始 PPTX 结构"}
          </h2>
          <p>{modeMeta[mode].description}</p>
        </div>
        <div className="mode-workspace-actions">
          {mode === "pages" ? (
            <>
              <button className="outline-button" onClick={pasteDraftPages}>
                <Copy size={15} />
                粘贴多页
              </button>
              <button className="primary-button" onClick={applyDraftPages}>
                <Check size={15} />
                确认应用 {draftPages.length} 页
              </button>
            </>
          ) : (
            <button className="primary-button" onClick={onUpload}>
              <Upload size={16} />
              {mode === "import" ? "选择多个资料" : "上传 PPTX"}
            </button>
          )}
        </div>
      </div>
      {mode !== "pages" && (
        <div className="intent-selector" aria-label="资料用途">
          <span>资料用途</span>
          {(["create", "modify", "reference"] as const).map((intent) => (
            <button
              key={intent}
              className={documentIntent === intent ? "active" : ""}
              onClick={() => onIntentChange(intent)}
            >
              {intent === "create"
                ? "创建新 PPT"
                : intent === "modify"
                  ? "修改当前项目"
                  : "仅作为参考"}
            </button>
          ))}
        </div>
      )}
      {mode === "import" && (
        <div className="structure-controls">
          <label>
            <span>页面预算</span>
            <input
              type="number"
              min={documentIntent === "create" ? 3 : 1}
              max={100}
              value={pageBudget}
              onChange={(event) =>
                onPageBudgetChange(
                  Math.min(100, Math.max(1, Number(event.target.value) || 1)),
                )
              }
              disabled={documentIntent !== "create"}
            />
          </label>
          <label className="inherit-theme-toggle">
            <input
              type="checkbox"
              checked={inheritTheme}
              onChange={(event) => onInheritThemeChange(event.target.checked)}
            />
            <span>继承当前主题</span>
          </label>
          <button
            className="outline-button"
            onClick={onCreateStructurePlan}
            disabled={
              busy ||
              !documents.some(
                (document) =>
                  document.includedInContext &&
                  ["ready", "conflict"].includes(document.status),
              ) ||
              conflicts.some(
                (conflict) =>
                  conflict.status === "open" &&
                  conflict.severity === "blocking",
              )
            }
          >
            <LayoutDashboard size={15} />
            生成结构计划
          </button>
        </div>
      )}
      {mode === "import" && structurePlan && (
        <div className="structure-plan" data-operation-id={structurePlan.operationId}>
          <div className="structure-plan-head">
            <div>
              <strong>{structurePlan.totalPages} 页结构计划</strong>
              <span>
                {structurePlan.preservesPageCount
                  ? "保持当前页数"
                  : "确认后创建新 Deck"}
              </span>
            </div>
            <small>{structurePlan.inheritTheme ? "继承主题" : "使用默认主题"}</small>
          </div>
          <div className="structure-page-list">
            {structurePlan.pages.map((page, index) => (
              <div key={page.pageId}>
                <span>{String(index + 1).padStart(2, "0")}</span>
                <div>
                  <strong>{page.title}</strong>
                  <small>
                    {page.kind === "cover"
                      ? "封面"
                      : page.kind === "transition"
                        ? "目录 / 过渡"
                        : "内容页"}
                  </small>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
      {conflicts.some((conflict) => conflict.status === "open") && (
        <div className="conflict-list">
          {conflicts
            .filter((conflict) => conflict.status === "open")
            .map((conflict) => (
              <div className="conflict-card" key={conflict.conflictId}>
                <div>
                  <CircleAlert size={16} />
                  <strong>
                    {conflict.severity === "blocking"
                      ? "严重事实冲突，已阻断执行"
                      : "资料表述存在差异，请核对"}
                  </strong>
                  <span>{conflict.key}</span>
                </div>
                <div className="conflict-values">
                  {conflict.values.map((entry) => (
                    <div
                      className="conflict-source"
                      key={`${entry.documentId}-${entry.value}`}
                    >
                      <div>
                        <strong>{entry.value}</strong>
                        <span>
                          {entry.fileName} · {entry.sourceLocator} · 可靠性 {Math.round(entry.confidence * 100)}%
                        </span>
                        <p>{entry.originalText}</p>
                        <small>
                          影响页面：{entry.affectedPageIds.length ? entry.affectedPageIds.join("、") : "待计划确定"}
                        </small>
                      </div>
                      <button
                        onClick={() =>
                          onResolveConflict(conflict, {
                            action: "prefer_source",
                            selectedValue: entry.value,
                            selectedDocumentIds: [entry.documentId],
                          })
                        }
                      >
                        采用此来源
                      </button>
                    </div>
                  ))}
                  <div className="conflict-resolution-actions">
                    <button
                      onClick={() =>
                        onResolveConflict(conflict, { action: "keep_both" })
                      }
                    >
                      保留两种结论并明确区分
                    </button>
                    <button
                      onClick={() =>
                        onResolveConflict(conflict, { action: "ignore" })
                      }
                    >
                      忽略该字段
                    </button>
                  </div>
                </div>
              </div>
            ))}
        </div>
      )}
      {mode === "import" && (
        <div className="document-pool">
          {documents.length === 0 ? (
            <div className="dropzone" onClick={onUpload}>
              <Files size={24} />
              <strong>拖入 .md、.docx 或 .pdf</strong>
              <span>PDF 只提取文字和结构，不作为页面参考图</span>
            </div>
          ) : (
            documents.map((document) => (
              <DocumentCard
                key={document.id}
                document={document}
                onToggleContext={onToggleContext}
              />
            ))
          )}
        </div>
      )}
      {mode === "beautify" && (
        <div className="beautify-flow">
          <div className="flow-step done">
            <span>01</span>
            <strong>上传 PPTX</strong>
            <small>提取文本、形状和主题</small>
          </div>
          <ArrowRight size={16} />
          <div className="flow-step active">
            <span>02</span>
            <strong>视觉预览</strong>
            <small>确认方向后再重建</small>
          </div>
          <ArrowRight size={16} />
          <div className="flow-step">
            <span>03</span>
            <strong>可编辑重建</strong>
            <small>登记不可编辑局部</small>
          </div>
          <ArrowRight size={16} />
          <div className="flow-step">
            <span>04</span>
            <strong>权威 QA</strong>
            <small>PowerPoint PNG</small>
          </div>
          {documents
            .filter((document) => document.kind === "PPTX")
            .map((document) => (
              <DocumentCard
                key={document.id}
                document={document}
                onToggleContext={onToggleContext}
              />
            ))}
        </div>
      )}
      {mode === "pages" && (
        <div className="draft-grid">
          {draftPages.map((page, index) => (
            <div
              className={`draft-card ${draggedDraftId === page.id ? "dragging" : ""}`}
              data-draft-id={page.id}
              key={page.id}
            >
              <div className="draft-card-head">
                <span>
                  <button
                    type="button"
                    className="draft-drag-handle"
                    aria-label={`拖动第 ${index + 1} 页排序`}
                    title="拖动排序"
                    onPointerDown={(event) =>
                      onDraftPointerDown(event, page.id)
                    }
                    onPointerMove={onDraftPointerMove}
                    onPointerUp={onDraftPointerEnd}
                    onPointerCancel={onDraftPointerEnd}
                  >
                    <GripVertical size={15} />
                  </button>
                  第 {index + 1} 页
                </span>
                <div>
                  <button
                    className="icon-button"
                    title="上移页面"
                    disabled={index === 0}
                    onClick={() => moveDraftPage(page.id, -1)}
                  >
                    <ArrowUp size={14} />
                  </button>
                  <button
                    className="icon-button"
                    title="下移页面"
                    disabled={index === draftPages.length - 1}
                    onClick={() => moveDraftPage(page.id, 1)}
                  >
                    <ArrowDown size={14} />
                  </button>
                  {!page.locked && (
                    <button
                      className="icon-button"
                      title="删除页面"
                      onClick={() => removeDraftPage(page.id)}
                    >
                      <X size={14} />
                    </button>
                  )}
                </div>
              </div>
              <input
                value={page.title}
                aria-label={`第 ${index + 1} 页标题`}
                onChange={(event) =>
                  updateDraft(page.id, "title", event.target.value)
                }
              />
              <textarea
                value={page.body}
                aria-label={`第 ${index + 1} 页正文`}
                onChange={(event) =>
                  updateDraft(page.id, "body", event.target.value)
                }
                rows={4}
              />
            </div>
          ))}
          <button className="draft-add" onClick={addDraftPage}>
            <Plus size={16} />
            明确添加新页面
          </button>
        </div>
      )}
    </section>
  );
}
function DocumentCard({
  document,
  onToggleContext,
}: {
  document: LocalDocument;
  onToggleContext: (document: LocalDocument) => void;
}): JSX.Element {
  const tone =
    document.status === "ready"
      ? "green"
      : document.status === "failed"
        ? "red"
        : "orange";
  return (
    <div className="document-card">
      <div className={`document-icon ${tone}`}>
        {document.kind === "PDF"
          ? "PDF"
          : document.kind === "DOCX"
            ? "DOC"
            : document.kind}
      </div>
      <div className="document-info">
        <strong>{document.name}</strong>
        <span>
          {(document.size / 1024).toFixed(0)} KB · {document.source}
        </span>
        <span title={document.sha256}>SHA {document.sha256.slice(0, 12)}…</span>
        <span title={document.sourceVersionId}>
          {document.facts.length} 个事实 ·{" "}
          {document.sourceVersionId.slice(0, 13)} ·{" "}
          {formatTime(document.createdAt)}
        </span>
        <small>{document.structure}</small>
      </div>
      <span className={`document-status ${tone}`}>
        {document.status === "parsing"
          ? "解析中"
          : document.status === "ready"
            ? "已就绪"
            : document.status === "conflict"
              ? "冲突阻断"
              : document.status}
      </span>
      <label className="document-context-toggle">
        <input
          type="checkbox"
          checked={document.includedInContext}
          onChange={() => onToggleContext(document)}
          aria-label={`${document.name} 纳入当前任务上下文`}
        />
        <span>上下文</span>
      </label>
    </div>
  );
}
function PageRow({
  page,
  active,
  selected,
  onClick,
  onToggle,
}: {
  page: Page;
  active: boolean;
  selected: boolean;
  onClick: () => void;
  onToggle: () => void;
}): JSX.Element {
  const version = currentVersion(page);
  return (
    <div className={`page-row ${active ? "active" : ""}`} onClick={onClick}>
      <button
        className={`check-box ${selected ? "checked" : ""}`}
        onClick={(event) => {
          event.stopPropagation();
          onToggle();
        }}
        aria-label={`选择第 ${page.orderIndex + 1} 页`}
      >
        {selected && <Check size={12} />}
      </button>
      <div className="thumb">
        <div dangerouslySetInnerHTML={{ __html: version.previewSvg }} />
      </div>
      <div className="page-row-info">
        <div className="page-row-top">
          <strong>{String(page.orderIndex + 1).padStart(2, "0")}</strong>
          <span className={`mini-status ${statusTone[page.status]}`} />
        </div>
        <div className="page-row-title">{page.title}</div>
        <div className="page-row-meta">
          <span>{version.versionId.slice(0, 13)}</span>
          <span>{statusText[page.status]}</span>
        </div>
      </div>
    </div>
  );
}
function ChatMessage({
  message,
  busy,
  onRollbackOperation,
  onRetryOperation,
}: {
  message: ConversationMessage;
  busy: boolean;
  onRollbackOperation: (operationId: string) => Promise<void>;
  onRetryOperation: (operationId: string) => Promise<void>;
}): JSX.Element {
  const versionIds = Array.isArray(message.meta?.versionIds)
    ? (message.meta.versionIds as string[])
    : [];
  const failedPageIds = Array.isArray(message.meta?.failedPageIds)
    ? (message.meta.failedPageIds as string[])
    : [];
  const rolledBack = message.meta?.rolledBack === true;
  const retryOperationId =
    typeof message.meta?.retryOperationId === "string"
      ? message.meta.retryOperationId
      : "";
  return (
    <div className={`chat-message ${message.role}`}>
      <div className="message-meta">
        <span>
          {message.role === "user"
            ? "你"
            : message.role === "assistant"
              ? "FastPPT"
              : "系统"}
        </span>
        <time>{formatTime(message.createdAt)}</time>
      </div>
      <div className="message-bubble">{message.text}</div>
      {versionIds.length > 0 && (
        <div className="message-foot">
          <span>
            <Check size={12} />
            版本已登记 · {versionIds.length} 页
          </span>
          {message.operationId &&
            (rolledBack ? (
              <span>已整组回滚</span>
            ) : (
              <button
                type="button"
                onClick={() => void onRollbackOperation(message.operationId!)}
                disabled={busy}
              >
                <RotateCcw size={12} />
                撤销本次操作
              </button>
            ))}
        </div>
      )}
      {failedPageIds.length > 0 && (
        <div className="message-foot">
          <span>
            <CircleAlert size={12} />
            {failedPageIds.length} 页失败
          </span>
          {retryOperationId ? (
            <span>重试任务已创建</span>
          ) : (
            message.operationId && (
              <button
                type="button"
                onClick={() => void onRetryOperation(message.operationId!)}
                disabled={busy}
              >
                <RotateCcw size={12} />
                重试失败页
              </button>
            )
          )}
        </div>
      )}
    </div>
  );
}
function ConfirmationCard({
  operation,
  onConfirm,
  onCancel,
  busy,
}: {
  operation: EditOperation;
  onConfirm: () => void;
  onCancel: () => void;
  busy: boolean;
}): JSX.Element {
  const plan = operation.structuredPlan;
  const visualPreviewsReady = (operation.visualPreviews || []).every(
    (preview) => preview.status === "ready",
  );
  return (
    <div className="confirmation-card">
      <div className="confirmation-title">
        <CircleAlert size={16} />
        <strong>执行前确认</strong>
        <span>
          {operation.mode === "global"
            ? "全局相似页"
            : `${operation.resolvedPageIds.length} 页`}
        </span>
      </div>
      <p>{plan.summary}</p>
      <div className="candidate-list">
        {operation.resolvedPageIds.map((pageId) => (
          <div key={pageId}>
            <span className="candidate-index">{pageId.slice(-3)}</span>
            <span>
              {plan.candidateReasons?.[pageId] || "合同与请求目标匹配"}
            </span>
          </div>
        ))}
      </div>
      {operation.visualPreviews && operation.visualPreviews.length > 0 && (
        <div className="visual-candidate-grid">
          {operation.visualPreviews.map((preview) => (
            <div className="visual-candidate" key={preview.pageId}>
              <div className="visual-candidate-head">
                <strong>{preview.pageId.slice(-3)}</strong>
                <span>
                  {preview.status === "ready"
                    ? preview.source === "relay_image"
                      ? "视觉候选"
                      : "布局候选"
                    : preview.status === "failed"
                      ? "预览失败"
                      : "生成中"}
                </span>
              </div>
              {preview.artifactId && preview.status === "ready" ? (
                <img
                  src={visualPreviewUrl(operation.projectId, preview.artifactId)}
                  alt={`页面 ${preview.pageId} 的视觉候选`}
                />
              ) : (
                <div className="visual-candidate-empty">
                  {preview.error || "等待预览结果"}
                </div>
              )}
            </div>
          ))}
        </div>
      )}
      <div className="confirmation-cost">
        <span>预计成本</span>
        <strong>${plan.estimatedCost.amount.toFixed(2)}</strong>
        <span>· {plan.estimatedCost.imageUnits} image unit</span>
      </div>
      {plan.unsupported.length > 0 && (
        <div className="unsupported-note">{plan.unsupported.join(" ")}</div>
      )}
      <div className="confirmation-actions">
        <button className="ghost-button" onClick={onCancel} disabled={busy}>
          <X size={14} />
          取消
        </button>
        <button
          className="primary-button"
          onClick={onConfirm}
          disabled={busy || !visualPreviewsReady}
        >
          <Check size={14} />
          确认执行
        </button>
      </div>
    </div>
  );
}
function HistoryDrawer({
  project,
  page,
  history,
  tab,
  onTab,
  compareId,
  onCompare,
  onRollback,
  onRollbackOperation,
  onLocate,
  onClose,
}: {
  project: Project;
  page: Page;
  history: ProjectHistory;
  tab: DrawerTab;
  onTab: (tab: DrawerTab) => void;
  compareId: string;
  onCompare: (id: string) => void;
  onRollback: (id: string) => Promise<void>;
  onRollbackOperation: (id: string) => Promise<void>;
  onLocate: (pageId: string, versionId: string | null) => void;
  onClose: () => void;
}): JSX.Element {
  const [conversationFilter, setConversationFilter] = useState("all");
  const [pageFilter, setPageFilter] = useState("all");
  const [statusFilter, setStatusFilter] = useState("all");
  const operationStatus: Record<EditOperation["status"], string> = {
    planned: "待确认",
    confirmed: "已确认",
    applying: "执行中",
    completed: "已完成",
    failed: "失败",
    rolled_back: "已回滚",
  };
  const qaStatus = {
    passed: "QA 通过",
    warning: "QA 有警告",
    failed: "QA 失败",
    pending: "QA 待完成",
  } as const;
  const conversations = [
    ...new Set(history.operations.map((operation) => operation.conversationId)),
  ];
  const filteredOperations = history.operations.filter(
    (operation) =>
      (conversationFilter === "all" ||
        operation.conversationId === conversationFilter) &&
      (pageFilter === "all" || operation.pageIds.includes(pageFilter)) &&
      (statusFilter === "all" || operation.status === statusFilter),
  );
  const unlinkedMessages = history.messages.filter(
    (message) =>
      !message.operationId &&
      pageFilter === "all" &&
      statusFilter === "all" &&
      (conversationFilter === "all" ||
        message.conversationId === conversationFilter),
  );
  const formatDuration = (durationMs: number | null): string => {
    if (durationMs === null) return "进行中";
    if (durationMs < 1_000) return `${durationMs} ms`;
    return `${(durationMs / 1_000).toFixed(1)} 秒`;
  };
  return (
    <aside className="history-drawer">
      <div className="drawer-header">
        <div>
          <span className="eyebrow">PROJECT HISTORY</span>
          <h2>历史抽屉</h2>
        </div>
        <button className="icon-button" title="关闭历史抽屉" onClick={onClose}>
          <PanelRightClose size={17} />
        </button>
      </div>
      <div className="drawer-tabs">
        <button
          className={tab === "conversation" ? "active" : ""}
          onClick={() => onTab("conversation")}
        >
          <MessageSquareText size={15} />
          对话记录<span>{history.operations.length}</span>
        </button>
        <button
          className={tab === "versions" ? "active" : ""}
          onClick={() => onTab("versions")}
        >
          <History size={15} />
          版本历史<span>{page.versions.length}</span>
        </button>
      </div>
      {tab === "conversation" ? (
        <div className="drawer-scroll">
          <div className="history-filters" aria-label="历史筛选">
            <label>
              <span>会话</span>
              <select
                aria-label="按会话筛选"
                value={conversationFilter}
                onChange={(event) => setConversationFilter(event.target.value)}
              >
                <option value="all">全部会话</option>
                {conversations.map((conversationId, index) => (
                  <option value={conversationId} key={conversationId}>
                    会话 {index + 1} · {conversationId.slice(-8)}
                  </option>
                ))}
              </select>
            </label>
            <label>
              <span>页面</span>
              <select
                aria-label="按页面筛选"
                value={pageFilter}
                onChange={(event) => setPageFilter(event.target.value)}
              >
                <option value="all">全部页面</option>
                {[...project.pages]
                  .sort((a, b) => a.orderIndex - b.orderIndex)
                  .map((candidate, index) => (
                    <option value={candidate.pageId} key={candidate.pageId}>
                      P{index + 1} · {candidate.title}
                    </option>
                  ))}
              </select>
            </label>
            <label>
              <span>状态</span>
              <select
                aria-label="按状态筛选"
                value={statusFilter}
                onChange={(event) => setStatusFilter(event.target.value)}
              >
                <option value="all">全部状态</option>
                {Object.entries(operationStatus).map(([status, label]) => (
                  <option value={status} key={status}>
                    {label}
                  </option>
                ))}
              </select>
            </label>
          </div>
          {filteredOperations.length === 0 && unlinkedMessages.length === 0 ? (
            <WelcomeMessage />
          ) : (
            <>
              {filteredOperations.map((operation) => {
                const operationMessages = history.messages.filter(
                  (message) => message.operationId === operation.operationId,
                );
                return (
                  <article
                    className="history-operation"
                    data-operation-id={operation.operationId}
                    data-operation-status={operation.status}
                    key={operation.operationId}
                  >
                    <div className="history-operation-head">
                      <div>
                        <strong>{operationStatus[operation.status]}</strong>
                        <time>{formatTime(operation.createdAt)}</time>
                      </div>
                      <span className={`qa-chip ${operation.qaStatus}`}>
                        {qaStatus[operation.qaStatus]}
                      </span>
                    </div>
                    <div className="history-conversation">
                      {operationMessages.map((message) => (
                        <div className={message.role} key={message.messageId}>
                          <span>{message.role === "user" ? "你" : "FastPPT"}</span>
                          <p>{message.text}</p>
                        </div>
                      ))}
                    </div>
                    <dl className="history-audit-grid">
                      <div>
                        <dt>操作 ID</dt>
                        <dd>{operation.operationId}</dd>
                      </div>
                      <div>
                        <dt>会话</dt>
                        <dd>{operation.conversationId}</dd>
                      </div>
                      <div>
                        <dt>耗时</dt>
                        <dd>{formatDuration(operation.durationMs)}</dd>
                      </div>
                      <div>
                        <dt>成本</dt>
                        <dd>
                          {operation.currency} {operation.actualCost.toFixed(2)}
                          {operation.actualCost === 0 && operation.estimatedCost > 0
                            ? `（预估 ${operation.estimatedCost.toFixed(2)}）`
                            : ""}
                        </dd>
                      </div>
                      <div>
                        <dt>模型</dt>
                        <dd>{operation.models.join("、") || "本地确定性引擎"}</dd>
                      </div>
                      <div>
                        <dt>工作会话</dt>
                        <dd>{operation.sessionId || "默认会话"}</dd>
                      </div>
                    </dl>
                    <div className="history-page-links">
                      {operation.pages.map((pageAudit) => (
                        <button
                          type="button"
                          key={pageAudit.pageId}
                          onClick={() =>
                            onLocate(pageAudit.pageId, pageAudit.versionId)
                          }
                          title="定位到页面和版本"
                        >
                          <PanelsTopLeft size={13} />
                          <span>{pageAudit.pageTitle}</span>
                          <small>{pageAudit.versionId || "尚未生成版本"}</small>
                        </button>
                      ))}
                    </div>
                    {operation.warnings.length > 0 && (
                      <div className="history-warnings">
                        <CircleAlert size={13} />
                        <span>{operation.warnings.join("；")}</span>
                      </div>
                    )}
                    {operation.status === "completed" && (
                      <button
                        className="text-button history-rollback"
                        onClick={() =>
                          void onRollbackOperation(operation.operationId)
                        }
                      >
                        <RotateCcw size={13} />
                        整组回滚
                      </button>
                    )}
                  </article>
                );
              })}
              {unlinkedMessages.map((message) => (
                <div className="drawer-message" key={message.messageId}>
                  <div className="message-meta">
                    <span>{message.role === "user" ? "你" : "FastPPT"}</span>
                    <time>{formatTime(message.createdAt)}</time>
                  </div>
                  <p>{message.text}</p>
                </div>
              ))}
            </>
          )}
        </div>
      ) : (
        <div className="drawer-scroll">
          <div className="filter-row">
            <span>不可变页面版本</span>
            <span className="muted-count">按 version_id 判断当前</span>
          </div>
          {[...page.versions].reverse().map((version, index) => {
            const isCurrent = version.versionId === page.currentVersionId;
            const operation = history.operations.find(
              (candidate) =>
                candidate.operationId === version.editOperationId,
            );
            const pageAudit = operation?.pages.find(
              (candidate) => candidate.versionId === version.versionId,
            );
            return (
              <div
                className={`history-row ${isCurrent ? "current" : ""}`}
                data-version-id={version.versionId}
                data-current={isCurrent ? "true" : "false"}
                key={version.versionId}
              >
                <div className="history-marker">
                  {isCurrent ? <Check size={12} /> : <Clock3 size={12} />}
                </div>
                <div className="history-info">
                  <strong>
                    {isCurrent
                      ? "当前版本"
                      : `版本 ${page.versions.length - index}`}
                  </strong>
                  <span>
                    {version.versionId} · {formatTime(version.createdAt)}
                  </span>
                  <small>{version.message}</small>
                  <small className="version-audit">
                    {pageAudit?.model || "本地确定性引擎"} · {operation?.currency || "USD"}{" "}
                    {(pageAudit?.cost || 0).toFixed(2)} · {qaStatus[pageAudit?.qaStatus || "pending"]}
                  </small>
                  {version.qaWarnings.length > 0 && (
                    <small className="version-warning">
                      {version.qaWarnings.join("；")}
                    </small>
                  )}
                </div>
                <div className="history-actions">
                  <button
                    className="icon-button subtle"
                    title="对比版本"
                    onClick={() =>
                      onCompare(
                        compareId === version.versionId
                          ? ""
                          : version.versionId,
                      )
                    }
                  >
                    <Copy size={14} />
                  </button>
                  {!isCurrent && (
                    <button
                      className="icon-button subtle"
                      title="恢复此版本"
                      onClick={() => void onRollback(version.versionId)}
                    >
                      <RotateCcw size={14} />
                    </button>
                  )}
                </div>
              </div>
            );
          })}
          {compareId && (
            <div className="compare-note">
              <Copy size={13} />
              预览区正在显示对比版本
            </div>
          )}
        </div>
      )}
    </aside>
  );
}
function ExportBanner({
  job,
  project,
  sensitiveMode,
  onOfficePreview,
  onDismiss,
}: {
  job: ExportJob;
  project: Project;
  sensitiveMode: boolean;
  onOfficePreview: () => void;
  onDismiss: () => void;
}): JSX.Element {
  const complete = job.status === "completed";
  return (
    <div className="export-banner">
      <div className={`export-icon ${complete ? "done" : "pending"}`}>
        {complete ? (
          <Check size={18} />
        ) : (
          <LoaderCircle className="spin" size={18} />
        )}
      </div>
      <div>
        <strong>{complete ? "PPTX 导出完成" : "正在生成可编辑 PPTX"}</strong>
        <span>
          {sensitiveMode
            ? "敏感模式：仅提供 FastPPT 权威 PNG 和下载"
            : job.renderMode === "svg_fallback"
              ? "SVG 回退已登记，PowerPoint 权威渲染待补齐"
              : "正在运行交付 QA"}
        </span>
      </div>
      <div className="export-actions">
        {complete && (
          <a
            className="outline-button"
            href={exportDownloadUrl(project.projectId, job.exportId)}
          >
            <Download size={14} />
            下载
          </a>
        )}
        {complete && !sensitiveMode && (
          <button className="outline-button" onClick={onOfficePreview}>
            <Expand size={14} />
            PowerPoint 网页预览
          </button>
        )}
        <button
          className="icon-button"
          title="关闭导出状态"
          onClick={onDismiss}
        >
          <X size={15} />
        </button>
      </div>
    </div>
  );
}
function WelcomeMessage(): JSX.Element {
  return (
    <div className="welcome-message">
      <div className="welcome-icon">
        <MessageSquareText size={20} />
      </div>
      <strong>从当前页开始</strong>
      <p>说出要改什么，计划通过合同校验后会生成新版本。</p>
      <div className="suggestion-list">
        <span>标题改短，保留所有数字</span>
        <span>右侧改为两列</span>
        <span>所有来源统一格式</span>
      </div>
    </div>
  );
}
function SettingsPanel({
  sensitiveMode,
  onSensitiveChange,
}: {
  sensitiveMode: boolean;
  onSensitiveChange: (value: boolean) => void;
}): JSX.Element {
  return (
    <section className="settings-panel">
      <div>
        <span className="eyebrow">PROJECT SETTINGS</span>
        <h2>安全与预览</h2>
        <p>敏感模式会在服务端阻止 Office Viewer 和其他第三方文件预览服务。</p>
      </div>
      <label className="switch-row">
        <span>
          <ShieldCheck size={17} />
          敏感模式
        </span>
        <input
          type="checkbox"
          checked={sensitiveMode}
          onChange={(event) => onSensitiveChange(event.target.checked)}
        />
        <span className="switch" />
      </label>
    </section>
  );
}
function LoginScreen({
  form,
  setForm,
  onSubmit,
  busy,
  error,
}: {
  form: { email: string; name: string; accessCode: string };
  setForm: (value: { email: string; name: string; accessCode: string }) => void;
  onSubmit: (event: FormEvent) => Promise<void>;
  busy: boolean;
  error: string;
}): JSX.Element {
  return (
    <div className="auth-screen">
      <div className="auth-panel">
        <div className="brand-mark large">F</div>
        <div className="eyebrow">FASTPPT ONLINE</div>
        <h1>把每一页，聊到可交付。</h1>
        <p>
          浏览器内完成模式选择、资料解析、聊天精修、版本审计和可编辑 PPTX 导出。
        </p>
        <form onSubmit={(event) => void onSubmit(event)}>
          <label>
            邮箱
            <input
              type="email"
              value={form.email}
              onChange={(event) =>
                setForm({ ...form, email: event.target.value })
              }
              required
            />
          </label>
          <label>
            显示名称
            <input
              value={form.name}
              onChange={(event) =>
                setForm({ ...form, name: event.target.value })
              }
              required
            />
          </label>
          <label>
            访问码
            <input
              type="password"
              value={form.accessCode}
              onChange={(event) =>
                setForm({ ...form, accessCode: event.target.value })
              }
              autoComplete="current-password"
            />
          </label>
          <button type="submit" className="primary-button full" disabled={busy}>
            {busy ? (
              <LoaderCircle className="spin" size={16} />
            ) : (
              <ArrowRight size={16} />
            )}
            进入工作台
          </button>
        </form>
        {error && (
          <div className="error-inline">
            <CircleAlert size={14} />
            {error}
          </div>
        )}
        <small>
          生产环境仅允许白名单账号；模型密钥和 PowerPoint 凭证不会进入浏览器。
        </small>
      </div>
      <div className="auth-aside">
        <div className="auth-grid">
          <div>
            <span>01</span>
            <strong>Mode-first</strong>
            <p>先选择生产模式，再进入资料池或逐页录入。</p>
          </div>
          <div>
            <span>02</span>
            <strong>Image-first</strong>
            <p>视觉预览先行，再进入可编辑重建。</p>
          </div>
          <div>
            <span>03</span>
            <strong>Audit-ready</strong>
            <p>事实、版本、成本和导出状态可追溯。</p>
          </div>
        </div>
      </div>
    </div>
  );
}
function EmptyWorkspace({
  user,
  onCreate,
  onLogout,
  busy,
  error,
}: {
  user: User;
  onCreate: () => Promise<void>;
  onLogout: () => void;
  busy: boolean;
  error: string;
}): JSX.Element {
  return (
    <div className="empty-workspace">
      <div className="empty-card">
        <div className="brand-mark">F</div>
        <span className="eyebrow">欢迎回来，{user.name}</span>
        <h1>创建第一个 FastPPT 项目</h1>
        <p>从一个空白项目开始，之后可以选择导入文档、按页录入或 PPT 美化。</p>
        <button
          className="primary-button"
          onClick={() => void onCreate()}
          disabled={busy}
        >
          <Plus size={16} />
          创建项目
        </button>
        {error && (
          <div className="error-inline">
            <CircleAlert size={14} />
            {error}
          </div>
        )}
        <button className="text-button" onClick={onLogout}>
          <LogOut size={14} />
          退出登录
        </button>
      </div>
    </div>
  );
}
export default App;

import fs from "node:fs/promises";
import path from "node:path";
import type {
  ConversationMessage,
  ConflictResolution,
  DocumentConflict,
  DocumentSource,
  DocumentStructurePlan,
  DocumentVersion,
  DurableJob,
  EditOperation,
  ExportJob,
  Page,
  PageVersion,
  ProjectHistory,
  Project,
  PromptSnapshot,
  UsageLedgerEntry,
  PreviewArtifact,
  WorkflowMode,
  WorkSession,
} from "../shared/models.js";
import {
  editPlanSchema,
  type ChatTurnInput,
  type CreateProjectInput,
} from "../shared/protocol.js";
import { parseSlidesMarkdown, sourceHash } from "./contracts.js";
import { EventBus } from "./events.js";
import { HttpError } from "./errors.js";
import {
  applyFactChanges,
  buildEditPlan,
  composePrompt,
  factChanges,
  previewTextChange,
  resolveSimilarPages,
  validateEditPlan,
} from "./plans.js";
import { makePreviewSvg } from "./preview.js";
import type { StateStore } from "./store.js";
import { makeId, nowIso, safeFileName, sha256 } from "./utils.js";
import {
  runPowerPointRender,
  runPptxExport,
  type PptxVisualAsset,
} from "./workerBridge.js";
import { RelayModelAdapter } from "./relay.js";
import { SlidevQuickPreviewWorker } from "./quickPreview.js";
import { ObjectStorage, type ObjectStoreStatus } from "./objectStore.js";
import { runContentQa } from "./contentQa.js";
import { DurableJobQueue } from "./jobQueue.js";
import { parseControlledDocument } from "./documentParser.js";

const DEFAULT_MARKDOWN = `# FastPPT Online

选择页面并通过聊天精修，所有版本都保留稳定 page_id。

---

# 从快速预览到权威渲染

快速预览用于即时反馈，最终真相来自同版本的 PPTX PowerPoint 渲染。

---

# 事实、成本与可编辑交付

事实锚点默认锁定。图片生成先预留成本，最终 PPTX 不使用整页截图冒充可编辑内容。`;

const DOCUMENT_MEDIA_TYPES: Record<DocumentSource["fileType"], string> = {
  md: "text/markdown",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  pdf: "application/pdf",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};

function initialVersion(
  projectId: string,
  pageId: string,
  pageNumber: number,
  projectName: string,
  slide: ReturnType<typeof parseSlidesMarkdown>[number],
): PageVersion {
  const versionId = makeId("ver");
  return {
    versionId,
    pageId,
    parentVersionId: null,
    sourceRevision: `deckrev_${sha256(`${projectId}:${pageId}`).slice(0, 12)}`,
    pageContractPath: `contracts/${pageId}.json`,
    slidesSourceHash: sourceHash(slide.sourceMarkdown),
    previewArtifactId: makeId("artifact"),
    visualArtifactId: null,
    svgArtifactId: makeId("svg"),
    pptxPageRenderId: null,
    promptSnapshotId: makeId("prompt"),
    editOperationId: `import_${projectId}`,
    qualityReportId: makeId("qa"),
    status: "ready",
    createdAt: nowIso(),
    message: "Initial import",
    title: slide.title,
    body: slide.body,
    layout: slide.layout,
    sourceMarkdown: slide.sourceMarkdown,
    contract: {
      ...slide.contract,
      evidence: [...slide.contract.evidence],
      mustKeep: [...slide.contract.mustKeep],
      canTrim: [...slide.contract.canTrim],
      components: [...slide.contract.components],
    },
    previewKind: "svg_fallback",
    previewSvg: makePreviewSvg(
      pageNumber,
      slide.title,
      slide.body,
      slide.layout,
      projectName,
    ),
    editableLevel: "native_partial",
    nonEditableRegions: [],
    qaWarnings: ["PowerPoint authority has not rendered this initial version."],
    renderNote: "Quick SVG preview. The editable PPTX authority is pending.",
    factAnchors: slide.facts.map((fact) => ({ ...fact })),
  };
}

function currentVersion(page: Page): PageVersion {
  return (
    page.versions.find(
      (version) => version.versionId === page.currentVersionId,
    ) || page.versions[page.versions.length - 1]
  );
}

function createDurableJob(input: {
  jobId: string;
  projectId: string;
  operationId?: string | null;
  kind: DurableJob["kind"];
  payload: Record<string, unknown>;
}): DurableJob {
  const timestamp = nowIso();
  const configuredAttempts = Number(process.env.JOB_MAX_ATTEMPTS || 3);
  return {
    ...input,
    idempotencyKey: `${input.kind}:${input.jobId}`,
    status: "queued",
    attemptCount: 0,
    maxAttempts:
      Number.isFinite(configuredAttempts) && configuredAttempts > 0
        ? Math.min(10, Math.floor(configuredAttempts))
        : 3,
    availableAt: timestamp,
    leaseOwner: null,
    leaseExpiresAt: null,
    lastError: null,
    startedAt: null,
    createdAt: timestamp,
    updatedAt: timestamp,
    completedAt: null,
  };
}

export class OnlineService {
  private readonly relay = new RelayModelAdapter();
  private readonly quickPreview = new SlidevQuickPreviewWorker();
  private readonly objectStorage = new ObjectStorage();
  private readonly jobs: DurableJobQueue;

  constructor(
    private readonly store: StateStore,
    private readonly events: EventBus,
  ) {
    this.jobs = new DurableJobQueue(store);
  }

  objectStorageStatus(): ObjectStoreStatus {
    return this.objectStorage.status;
  }

  private async exportVisualAssets(
    project: Project,
  ): Promise<Record<string, PptxVisualAsset>> {
    const assets: Record<string, PptxVisualAsset> = {};
    for (const page of project.pages.filter(
      (candidate) => !candidate.archived,
    )) {
      const version = currentVersion(page);
      if (!version.visualArtifactId) continue;
      const artifact = this.store.state.artifacts.find(
        (candidate) => candidate.artifactId === version.visualArtifactId,
      );
      if (!artifact?.assetPath || artifact.source !== "relay_image") {
        throw new Error(
          `The current visual asset for page ${page.pageId} is missing or invalid.`,
        );
      }
      const extension = path.extname(artifact.assetPath).toLowerCase();
      if (![".png", ".jpg", ".jpeg", ".webp"].includes(extension)) {
        throw new Error(
          `The current visual asset for page ${page.pageId} has an unsupported image format.`,
        );
      }
      assets[page.pageId] = {
        bytes: await this.objectStorage.getBytes(artifact.assetPath),
        extension: extension as PptxVisualAsset["extension"],
      };
    }
    return assets;
  }

  private async persistGeneratedImage(
    ownerId: string,
    projectId: string,
    pageId: string,
    versionId: string,
    bytes: Buffer,
    mimeType: string,
  ): Promise<string> {
    if (bytes.length === 0 || bytes.length > 20 * 1024 * 1024)
      throw new Error("Relay image output exceeded the 20 MB artifact limit.");
    const extension =
      mimeType.includes("jpeg") || mimeType.includes("jpg")
        ? "jpg"
        : mimeType.includes("webp")
          ? "webp"
          : "png";
    const directory = path.resolve(
      process.env.DATA_DIR || "./data",
      "artifacts",
      ownerId,
      projectId,
      pageId,
    );
    await fs.mkdir(directory, { recursive: true });
    const outputPath = path.resolve(directory, `${versionId}.${extension}`);
    await fs.writeFile(outputPath, bytes, { flag: "wx" });
    const stored = await this.objectStorage.putFile(
      outputPath,
      `projects/${ownerId}/${projectId}/pages/${pageId}/${versionId}.${extension}`,
      mimeType,
    );
    return stored.objectKey;
  }

  private operationNeedsVisualPreview(operation: EditOperation): boolean {
    if (operation.structuredPlan.intent === "page_split") return false;
    return (
      operation.structuredPlan.workflowMode === "pptx_beautify" ||
      operation.structuredPlan.changes.some((change) =>
        ["layout_change", "style_change", "image_replace"].includes(
          change.kind,
        ),
      )
    );
  }

  private promptContext(operation: EditOperation) {
    const documents = operation.structuredPlan.sourceDocumentIds
      .map((documentId) =>
        this.store.state.documents.find(
          (document) =>
            document.documentId === documentId &&
            document.projectId === operation.projectId &&
            document.includedInContext,
        ),
      )
      .filter((document): document is DocumentSource => Boolean(document))
      .slice(0, 8)
      .map((document) => ({
        documentId: document.documentId,
        sourceVersionId: document.sourceVersionId,
        fileName: document.fileName,
        sha256: document.sha256,
        headings: document.structure?.headings.slice(0, 30) || [],
        facts: document.facts.slice(0, 100).map((fact) => ({
          key: fact.key,
          value: fact.value,
          sourceLocator: fact.sourceLocator,
          confidence: fact.confidence,
        })),
        excerpt: document.parsedText.slice(0, 6000),
      }));
    const conflictResolutions = operation.sessionId
      ? this.store.state.conflictResolutions
          .filter(
            (resolution) => resolution.sessionId === operation.sessionId,
          )
          .slice(-100)
          .map((resolution) => ({
            resolutionId: resolution.resolutionId,
            conflictId: resolution.conflictId,
            sourceRevision: resolution.sourceRevision,
            action: resolution.action,
            selectedValue: resolution.selectedValue,
            selectedDocumentIds: resolution.selectedDocumentIds,
            note: resolution.note,
          }))
      : [];
    return {
      sources: documents,
      conflictResolutions,
      model: process.env.RELAY_MODEL || "deterministic-local-planner",
      image: { width: 1600, height: 900, quality: "high" },
    };
  }

  private async prepareOperationVisualPreviews(
    ownerId: string,
    operation: EditOperation,
    pages: Page[],
  ): Promise<void> {
    if (!this.operationNeedsVisualPreview(operation)) return;
    const project = this.getProject(ownerId, operation.projectId);
    const relayStatus = this.relay.status();
    const imageChange = operation.structuredPlan.changes.some(
      (change) => change.kind === "image_replace",
    );
    operation.visualPreviews ||= [];
    for (const page of pages) {
      const existing = operation.visualPreviews.find(
        (preview) => preview.pageId === page.pageId,
      );
      if (existing?.status === "ready") continue;
      const prompt = composePrompt(
        project,
        page,
        operation.structuredPlan,
        operation.message,
        this.promptContext(operation),
      );
      const promptSnapshotId = `prompt_${sha256(`${operation.operationId}:${page.pageId}:${prompt.hash}`).slice(0, 24)}`;
      const preview = existing || {
        pageId: page.pageId,
        artifactId: null,
        promptSnapshotId,
        source: imageChange ? ("relay_image" as const) : ("slidev_svg" as const),
        status: "generating" as const,
        ledgerId: null,
        error: null,
      };
      const reserve: UsageLedgerEntry | null = imageChange
        ? {
            ledgerId: makeId("ledger"),
            projectId: project.projectId,
            operationId: operation.operationId,
            pageId: page.pageId,
            versionId: null,
            model: relayStatus.imageModel,
            unitPrice: relayStatus.priceSnapshot.imageUnit,
            reservedAmount: relayStatus.priceSnapshot.imageUnit,
            settledAmount: 0,
            refundedAmount: 0,
            status: "reserved",
            createdAt: nowIso(),
          }
        : null;
      preview.status = "generating";
      preview.error = null;
      preview.ledgerId = reserve?.ledgerId || null;
      await this.store.update((state) => {
        if (!existing) operation.visualPreviews!.push(preview);
        if (
          !state.promptSnapshots.some(
            (snapshot) => snapshot.promptSnapshotId === promptSnapshotId,
          )
        )
          state.promptSnapshots.push({
            promptSnapshotId,
            projectId: project.projectId,
            pageId: page.pageId,
            operationId: operation.operationId,
            prompt: prompt.prompt,
            hash: prompt.hash,
            model: relayStatus.model,
            createdAt: nowIso(),
          });
        if (reserve) state.ledger.push(reserve);
      });
      try {
        const artifactId = makeId("artifact");
        let assetPath: string;
        let model: string;
        let width: number;
        let height: number;
        let source: PreviewArtifact["source"];
        let provenance: Record<string, unknown>;
        if (imageChange) {
          if (!relayStatus.configured)
            throw new Error(
              "视觉预览需要已配置的 Relay 图像模型，当前未生成任何占位图片。",
            );
          const generated = await this.relay.generateImage({
            prompt: `${prompt.prompt}\nGenerate a presentation visual for this page. Preserve every locked fact and do not render text inside the image.`,
            width: 1600,
            height: 900,
            quality: "high",
          });
          assetPath = await this.persistGeneratedImage(
            ownerId,
            project.projectId,
            page.pageId,
            artifactId,
            generated.bytes,
            generated.mimeType,
          );
          model = generated.model;
          width = generated.width;
          height = generated.height;
          source = "relay_image";
          provenance = {
            stage: "preconfirmation_visual_preview",
            operationId: operation.operationId,
            promptHash: prompt.hash,
            relayRequestId: generated.requestId,
            sha256: sha256(generated.bytes),
          };
        } else {
          const edited = previewTextChange(page, operation.message);
          const layout =
            operation.structuredPlan.changes.find(
              (change) => change.kind === "layout_change",
            )?.value || page.layout;
          const rendered = await this.quickPreview.render(
            project,
            page,
            edited.title,
            edited.body,
            layout,
          );
          const bytes = Buffer.from(rendered.svg, "utf8");
          const stored = await this.objectStorage.putBuffer(
            bytes,
            `projects/${ownerId}/${project.projectId}/operations/${operation.operationId}/previews/${artifactId}.svg`,
            "image/svg+xml",
          );
          assetPath = stored.objectKey;
          model = rendered.engine;
          width = 960;
          height = 540;
          source = "deterministic_svg";
          provenance = {
            stage: "preconfirmation_visual_preview",
            operationId: operation.operationId,
            promptHash: prompt.hash,
            engine: rendered.engine,
            note: rendered.note,
            sha256: sha256(bytes),
          };
        }
        const artifact: PreviewArtifact = {
          artifactId,
          projectId: project.projectId,
          pageId: page.pageId,
          versionId: null,
          kind: "visual_preview",
          promptSnapshotId,
          model,
          width,
          height,
          quality: imageChange ? "high" : "quick",
          source,
          provenance,
          createdAt: nowIso(),
          assetPath,
        };
        await this.store.update((state) => {
          state.artifacts.push(artifact);
          preview.artifactId = artifactId;
          preview.status = "ready";
          if (reserve) {
            reserve.status = "settled";
            reserve.settledAmount = reserve.reservedAmount;
          }
          state.auditLogs.push({
            auditId: makeId("audit"),
            ownerId,
            projectId: project.projectId,
            action: "preview.visual.ready",
            payload: {
              operationId: operation.operationId,
              pageId: page.pageId,
              artifactId,
              source,
            },
            createdAt: nowIso(),
          });
        });
        await this.events.publish({
          type: "preview.visual.ready",
          projectId: project.projectId,
          sessionId: operation.sessionId || undefined,
          pageId: page.pageId,
          operationId: operation.operationId,
          payload: { artifactId, source, preconfirmation: true },
        });
      } catch (error) {
        await this.store.update(() => {
          preview.status = "failed";
          preview.error = (error as Error).message;
          if (reserve) {
            reserve.status = "refunded";
            reserve.refundedAmount = reserve.reservedAmount;
          }
        });
      }
    }
  }

  listProjects(ownerId: string, includeArchived = false): Project[] {
    return this.store.state.projects
      .filter(
        (project) =>
          project.ownerId === ownerId &&
          (includeArchived || project.status !== "archived"),
      )
      .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
  }

  getProject(ownerId: string, projectId: string): Project {
    const project = this.store.state.projects.find(
      (candidate) => candidate.projectId === projectId,
    );
    if (!project || project.ownerId !== ownerId)
      throw new HttpError(404, "Project not found.");
    return project;
  }

  async updateProjectSettings(
    ownerId: string,
    projectId: string,
    settings: { sensitiveMode?: boolean },
  ): Promise<Project> {
    const project = this.getProject(ownerId, projectId);
    const sensitiveMode = settings.sensitiveMode === true;
    await this.store.update((state) => {
      project.settings = {
        ...(project.settings || { sensitiveMode: false }),
        sensitiveMode,
      };
      project.updatedAt = nowIso();
      state.officePreviewGrants
        .filter(
          (grant) =>
            grant.projectId === projectId && sensitiveMode && !grant.revokedAt,
        )
        .forEach((grant) => {
          grant.revokedAt = nowIso();
        });
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId,
        action: "project.settings.updated",
        payload: { sensitiveMode },
        createdAt: nowIso(),
      });
    });
    if (sensitiveMode)
      await this.events.publish({
        type: "office_preview.revoked",
        projectId,
        payload: { reason: "sensitive_mode" },
      });
    return project;
  }

  createWorkSession(
    ownerId: string,
    projectId: string,
    input: {
      workflowMode: WorkflowMode;
      intent?: "create" | "modify" | "reference";
      pageBudget?: number;
      inheritTheme?: boolean;
    },
  ): Promise<WorkSession> {
    this.getProject(ownerId, projectId);
    const now = nowIso();
    const session: WorkSession = {
      sessionId: makeId("session"),
      projectId,
      workflowMode: input.workflowMode,
      intent: input.intent || "modify",
      sourceDocumentIds: [],
      sourceRevision: makeId("source"),
      planId: null,
      pageBudget: Math.min(100, Math.max(1, input.pageBudget || 8)),
      inheritTheme: input.inheritTheme !== false,
      structurePlan: null,
      createdBy: ownerId,
      status: "collecting",
      createdAt: now,
      updatedAt: now,
    };
    return this.store.update((state) => {
      state.workSessions.push(session);
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId,
        action: "work_session.created",
        payload: {
          sessionId: session.sessionId,
          workflowMode: session.workflowMode,
          intent: session.intent,
        },
        createdAt: now,
      });
      return session;
    });
  }

  listWorkSessions(ownerId: string, projectId: string): WorkSession[] {
    this.getProject(ownerId, projectId);
    return this.store.state.workSessions
      .filter((session) => session.projectId === projectId)
      .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
  }

  async updateWorkSession(
    ownerId: string,
    projectId: string,
    sessionId: string,
    input: {
      intent?: "create" | "modify" | "reference";
      pageBudget?: number;
      inheritTheme?: boolean;
    },
  ): Promise<WorkSession> {
    const session = this.sessionForProject(ownerId, projectId, sessionId);
    await this.store.update((state) => {
      if (input.intent) session.intent = input.intent;
      if (input.pageBudget !== undefined)
        session.pageBudget = Math.min(100, Math.max(1, input.pageBudget));
      if (input.inheritTheme !== undefined)
        session.inheritTheme = input.inheritTheme;
      session.structurePlan = null;
      session.updatedAt = nowIso();
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId,
        action: "work_session.updated",
        payload: {
          sessionId,
          intent: session.intent,
          pageBudget: session.pageBudget,
          inheritTheme: session.inheritTheme,
        },
        createdAt: session.updatedAt,
      });
    });
    return session;
  }

  private sessionForProject(
    ownerId: string,
    projectId: string,
    sessionId: string,
  ): WorkSession {
    this.getProject(ownerId, projectId);
    const session = this.store.state.workSessions.find(
      (candidate) =>
        candidate.sessionId === sessionId && candidate.projectId === projectId,
    );
    if (!session) throw new HttpError(404, "Work session not found.");
    session.sourceDocumentIds ||= [];
    session.sourceRevision ||= makeId("source");
    session.planId ??= null;
    session.pageBudget ||= 8;
    session.inheritTheme ??= true;
    session.structurePlan ??= null;
    session.createdBy ||= ownerId;
    return session;
  }

  async createDocumentStructurePlan(
    ownerId: string,
    projectId: string,
    sessionId: string,
    input: { pageBudget?: number; inheritTheme?: boolean } = {},
  ): Promise<{
    session: WorkSession;
    structurePlan: DocumentStructurePlan;
    operation: EditOperation;
  }> {
    const project = this.getProject(ownerId, projectId);
    const session = this.sessionForProject(ownerId, projectId, sessionId);
    if (session.workflowMode !== "document_import")
      throw new HttpError(
        409,
        "Structure plans are only available for document import sessions.",
      );
    const documents = this.store.state.documents.filter(
      (document) =>
        document.projectId === projectId &&
        document.sessionId === sessionId &&
        document.includedInContext &&
        ["ready", "warning"].includes(document.parseStatus),
    );
    if (!documents.length)
      throw new HttpError(
        409,
        "Parse at least one included document before creating a structure plan.",
      );
    const blockingConflicts = this.store.state.documentConflicts.filter(
      (conflict) =>
        conflict.projectId === projectId &&
        conflict.sessionId === sessionId &&
        conflict.status === "open" &&
        conflict.severity === "blocking",
    );
    if (blockingConflicts.length)
      throw new HttpError(
        409,
        "Resolve blocking document conflicts before creating a structure plan.",
      );
    if (input.pageBudget !== undefined)
      session.pageBudget = Math.min(100, Math.max(1, input.pageBudget));
    if (input.inheritTheme !== undefined)
      session.inheritTheme = input.inheritTheme;
    if (session.intent === "create" && session.pageBudget < 3)
      throw new HttpError(
        400,
        "Creating a new deck requires at least 3 pages for cover, outline, and content.",
      );
    const currentPages = [...project.pages]
      .filter((page) => !page.archived)
      .sort((a, b) => a.orderIndex - b.orderIndex);
    const headings = [
      ...new Set(
        documents.flatMap((document) => document.structure?.headings || []),
      ),
    ].filter(Boolean);
    const factLines = documents.flatMap((document) =>
      document.facts.slice(0, 30).map((fact) => fact.context),
    );
    const operationId = makeId("op");
    const pageCount =
      session.intent === "create"
        ? session.pageBudget
        : Math.max(1, currentPages.length);
    const plannedPages: DocumentStructurePlan["pages"] = [];
    if (session.intent === "create") {
      const contentTitles = Array.from(
        { length: Math.max(1, pageCount - 2) },
        (_, index) => headings[index] || `核心内容 ${index + 1}`,
      );
      plannedPages.push({
        pageId: makeId("page"),
        kind: "cover",
        title: headings[0] || project.name,
        bodyPreview: `基于 ${documents.map((document) => document.fileName).join("、")} 生成`,
        sourceDocumentIds: documents.map((document) => document.documentId),
      });
      plannedPages.push({
        pageId: makeId("page"),
        kind: "transition",
        title: "目录与内容路线",
        bodyPreview: contentTitles.join("\n"),
        sourceDocumentIds: documents.map((document) => document.documentId),
      });
      contentTitles.forEach((title, index) => {
        const document = documents[index % documents.length];
        plannedPages.push({
          pageId: makeId("page"),
          kind: "content",
          title,
          bodyPreview:
            factLines[index] ||
            document.parsedText.split(/\r?\n/).filter(Boolean).slice(0, 4).join("\n") ||
            "内容将在页面合同与来源约束下生成。",
          sourceDocumentIds: [document.documentId],
        });
      });
    } else {
      currentPages.forEach((page, index) => {
        const document = documents[index % documents.length];
        plannedPages.push({
          pageId: page.pageId,
          kind: index === 0 && page.pageType === "cover" ? "cover" : "content",
          title: page.title,
          bodyPreview:
            session.intent === "reference"
              ? `保持当前页面，仅登记参考来源：${document.fileName}`
              : headings[index] || page.body.slice(0, 240),
          sourceDocumentIds: [document.documentId],
        });
      });
    }
    const structurePlan: DocumentStructurePlan = {
      operationId,
      intent: session.intent,
      pageBudget: session.pageBudget,
      inheritTheme: session.inheritTheme,
      preservesPageCount: session.intent !== "create",
      totalPages: plannedPages.length,
      pages: plannedPages,
      createdAt: nowIso(),
    };
    const createDeck = session.intent === "create";
    const referenceOnly = session.intent === "reference";
    const conflictIds = this.store.state.documentConflicts
      .filter((conflict) => conflict.sessionId === sessionId)
      .map((conflict) => conflict.conflictId);
    const operation: EditOperation = {
      operationId,
      projectId,
      conversationId: makeId("conv"),
      sessionId,
      mode: "multi",
      requestedPageIds: plannedPages.map((page) => page.pageId),
      resolvedPageIds: plannedPages.map((page) => page.pageId),
      message: createDeck
        ? `按 ${plannedPages.length} 页结构计划创建新 PPT`
        : session.intent === "modify"
          ? "按资料结构计划修改当前项目并保持页数"
          : "登记资料为项目参考，不修改页面",
      structuredPlan: {
        workflowMode: "import_document",
        targetScope: "multi",
        intent: createDeck
          ? "create_deck_from_documents"
          : "document_structure_plan",
        affectedPageIds: plannedPages.map((page) => page.pageId),
        pageDelta: {
          add: createDeck ? plannedPages.map((page) => page.pageId) : [],
          remove: createDeck ? currentPages.map((page) => page.pageId) : [],
          split: [],
          merge: [],
        },
        changes: documents.flatMap((document) =>
          document.facts.slice(0, 100).map((fact) => ({
            kind: "preserve_fact" as const,
            target: "document_fact",
            value: fact.value,
          })),
        ),
        factImpact: { added: [], removed: [], changed: [] },
        sourceDocumentIds: documents.map((document) => document.documentId),
        conflictIds,
        unsupported: [],
        requiresConfirmation: createDeck,
        confirmationReasons: createDeck
          ? [
              `将归档当前 ${currentPages.length} 页并创建 ${plannedPages.length} 个新页面，必须确认页数和大纲。`,
            ]
          : [],
        estimatedCost: { imageUnits: 0, amount: 0, currency: "USD" },
        summary: createDeck
          ? `结构计划包含封面、目录/过渡和 ${Math.max(1, plannedPages.length - 2)} 个内容页，共 ${plannedPages.length} 页。`
          : `结构计划保持当前 ${plannedPages.length} 页不变。`,
      },
      factImpact: { added: [], removed: [], changed: [] },
      unsupportedItems: [],
      confirmationRequired: !referenceOnly,
      confirmedAt: referenceOnly ? nowIso() : null,
      resultVersionIds: [],
      status: referenceOnly ? "completed" : "planned",
      createdAt: nowIso(),
      completedAt: referenceOnly ? nowIso() : null,
      failedPageIds: [],
      estimatedCost: 0,
      visualPreviews: [],
    };
    await this.store.update((state) => {
      session.structurePlan = structurePlan;
      session.planId = operationId;
      session.status = referenceOnly ? "completed" : "planned";
      session.updatedAt = nowIso();
      state.operations.push(operation);
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId,
        action: "document.structure_plan.created",
        payload: { sessionId, operationId, structurePlan },
        createdAt: nowIso(),
      });
    });
    await this.events.publish({
      type: "plan.created",
      projectId,
      sessionId,
      operationId,
      payload: { plan: operation.structuredPlan, structurePlan },
    });
    if (!referenceOnly)
      await this.events.publish({
        type: "edit.confirmation.required",
        projectId,
        sessionId,
        operationId,
        payload: { pageIds: operation.resolvedPageIds, structurePlan },
      });
    return { session, structurePlan, operation };
  }

  async uploadDocument(
    ownerId: string,
    projectId: string,
    sessionId: string,
    input: { fileName: string; contentType: string; bytes: Buffer },
  ): Promise<{ document: DocumentSource; duplicate: boolean }> {
    const session = this.sessionForProject(ownerId, projectId, sessionId);
    const extension = path
      .extname(input.fileName)
      .toLowerCase()
      .slice(1) as DocumentSource["fileType"];
    const allowed =
      session.workflowMode === "ppt_beautify"
        ? ["pptx"]
        : ["md", "docx", "pdf"];
    if (!allowed.includes(extension))
      throw new HttpError(
        400,
        `File type .${extension || "unknown"} is not allowed for this workflow mode.`,
      );
    const maxBytes = Number(process.env.MAX_DOCUMENT_BYTES || 25 * 1024 * 1024);
    if (!input.bytes.length || input.bytes.length > maxBytes)
      throw new HttpError(
        413,
        `Document must be between 1 and ${maxBytes} bytes.`,
      );
    const digest = sha256(input.bytes);
    const duplicate = this.store.state.documents.find(
      (candidate) =>
        candidate.projectId === projectId && candidate.sha256 === digest,
    );
    if (duplicate) {
      await this.store.update(() => {
        if (!session.sourceDocumentIds.includes(duplicate.documentId))
          session.sourceDocumentIds.push(duplicate.documentId);
        session.sourceRevision = makeId("source");
        session.updatedAt = nowIso();
      });
      return { document: duplicate, duplicate: true };
    }
    const maxFiles = Number(process.env.MAX_DOCUMENT_FILES || 20);
    const maxBatchBytes = Number(
      process.env.MAX_DOCUMENT_BATCH_BYTES || 100 * 1024 * 1024,
    );
    const sessionDocuments = this.store.state.documents.filter(
      (document) =>
        document.projectId === projectId && document.sessionId === sessionId,
    );
    if (sessionDocuments.length >= maxFiles)
      throw new HttpError(
        413,
        `A work session cannot contain more than ${maxFiles} documents.`,
      );
    const sessionBytes = sessionDocuments.reduce(
      (total, document) => total + document.sizeBytes,
      0,
    );
    if (sessionBytes + input.bytes.length > maxBatchBytes)
      throw new HttpError(
        413,
        `The work session document budget is ${maxBatchBytes} bytes.`,
      );
    const documentId = makeId("doc");
    const objectKey = `projects/${ownerId}/${projectId}/documents/${documentId}/source.${extension}`;
    await this.objectStorage.putBuffer(
      input.bytes,
      objectKey,
      input.contentType || "application/octet-stream",
    );
    const now = nowIso();
    const document: DocumentSource = {
      documentId,
      projectId,
      sessionId,
      fileName: path.basename(input.fileName).slice(0, 240),
      fileType: extension,
      sizeBytes: input.bytes.length,
      sha256: digest,
      sourceVersionId: makeId("docver"),
      uploadedBy: ownerId,
      objectKey,
      parseStatus: "queued",
      includedInContext: true,
      parsedText: "",
      structure: null,
      facts: [],
      warnings: [],
      error: null,
      createdAt: now,
      updatedAt: now,
    };
    const version: DocumentVersion = {
      sourceVersionId: document.sourceVersionId,
      documentId,
      projectId,
      sha256: digest,
      sizeBytes: input.bytes.length,
      mediaType: DOCUMENT_MEDIA_TYPES[extension],
      objectKey,
      createdBy: ownerId,
      createdAt: now,
    };
    await this.store.update((state) => {
      state.documents.push(document);
      state.documentVersions.push(version);
      session.sourceDocumentIds.push(document.documentId);
      session.sourceRevision = makeId("source");
      session.updatedAt = now;
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId,
        action: "document.uploaded",
        payload: {
          sessionId,
          documentId,
          fileType: extension,
          sizeBytes: input.bytes.length,
          sha256: digest,
        },
        createdAt: now,
      });
    });
    await this.events.publish({
      type: "document.uploaded",
      projectId,
      sessionId,
      payload: { sessionId, documentId, fileType: extension },
    });
    return { document, duplicate: false };
  }

  listDocuments(
    ownerId: string,
    projectId: string,
    sessionId?: string,
  ): DocumentSource[] {
    this.getProject(ownerId, projectId);
    return this.store.state.documents
      .filter(
        (document) =>
          document.projectId === projectId &&
          (!sessionId || document.sessionId === sessionId),
      )
      .sort((a, b) => a.createdAt.localeCompare(b.createdAt));
  }

  getDocument(
    ownerId: string,
    projectId: string,
    documentId: string,
  ): DocumentSource {
    this.getProject(ownerId, projectId);
    const document = this.store.state.documents.find(
      (candidate) =>
        candidate.documentId === documentId &&
        candidate.projectId === projectId,
    );
    if (!document) throw new HttpError(404, "Document not found.");
    return document;
  }

  async setDocumentContext(
    ownerId: string,
    projectId: string,
    documentId: string,
    includedInContext: boolean,
  ): Promise<DocumentSource> {
    const document = this.getDocument(ownerId, projectId, documentId);
    await this.store.update((state) => {
      document.includedInContext = includedInContext;
      document.updatedAt = nowIso();
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId,
        action: "document.context.updated",
        payload: { documentId, includedInContext },
        createdAt: nowIso(),
      });
    });
    await this.refreshDocumentConflicts(ownerId, projectId, document.sessionId);
    return document;
  }

  async parseDocument(
    ownerId: string,
    projectId: string,
    documentId: string,
  ): Promise<DocumentSource> {
    const document = this.getDocument(ownerId, projectId, documentId);
    const project = this.getProject(ownerId, projectId);
    const session = this.sessionForProject(
      ownerId,
      projectId,
      document.sessionId,
    );
    await this.store.update(() => {
      document.parseStatus = "parsing";
      document.error = null;
      document.updatedAt = nowIso();
      session.status = "parsing";
      session.updatedAt = nowIso();
    });
    await this.events.publish({
      type: "document.parsing",
      projectId,
      sessionId: session.sessionId,
      payload: { sessionId: session.sessionId, documentId },
    });
    const stagingRoot = path.resolve(
      process.env.DOCUMENT_STAGING_DIR ||
        path.resolve(process.env.DATA_DIR || "./data", "document-staging"),
    );
    const stagingPath = path.resolve(
      stagingRoot,
      `${document.documentId}.${document.fileType}`,
    );
    if (!stagingPath.startsWith(`${stagingRoot}${path.sep}`))
      throw new HttpError(500, "Document staging path escaped its root.");
    try {
      await fs.mkdir(stagingRoot, { recursive: true });
      await fs.writeFile(
        stagingPath,
        await this.objectStorage.getBytes(document.objectKey),
      );
      const parsed = await parseControlledDocument(stagingPath);
      const maxTextChars = Number(
        process.env.MAX_DOCUMENT_TEXT_CHARS || 2_000_000,
      );
      if (parsed.text.length > maxTextChars)
        throw new Error(
          `Parsed document text exceeds the ${maxTextChars}-character limit.`,
        );
      await this.store.update((state) => {
        document.parsedText = parsed.text;
        document.structure = parsed.structure;
        document.facts = parsed.facts.map((fact) => ({
          ...fact,
          factId: fact.factId || makeId("fact"),
          normalizedValue:
            fact.normalizedValue ||
            fact.value.replaceAll(",", "").toLowerCase(),
          sourceDocumentId: document.documentId,
          sourceLocator: fact.sourceLocator || fact.location,
          confidence:
            document.fileType === "pdf" && parsed.warnings.length ? 0.65 : 0.95,
          locked: false,
        }));
        document.warnings = parsed.warnings;
        document.parseStatus = parsed.warnings.length ? "warning" : "ready";
        document.updatedAt = nowIso();
        state.auditLogs.push({
          auditId: makeId("audit"),
          ownerId,
          projectId,
          action: "document.parsed",
          payload: {
            documentId,
            factCount: parsed.facts.length,
            warningCount: parsed.warnings.length,
          },
          createdAt: nowIso(),
        });
      });
      if (document.fileType === "pptx" && document.structure?.slides?.length) {
        const sourceMarkdown = document.structure.slides
          .map((slide) => `# ${slide.title}\n\n${slide.body}`)
          .join("\n---\n");
        await this.importSlides(ownerId, projectId, sourceMarkdown);
        const orderedPages = project.pages
          .filter((page) => !page.archived)
          .sort((a, b) => a.orderIndex - b.orderIndex);
        await this.store.update(() => {
          document.structure?.slides?.forEach((slide, index) => {
            const page = orderedPages[index];
            if (!page) return;
            const version = currentVersion(page);
            page.editableLevel = slide.editableLevel;
            version.editableLevel = slide.editableLevel;
            version.nonEditableRegions = [...slide.nonEditableRegions];
            if (slide.nonEditableRegions.length)
              version.qaWarnings = [
                ...new Set([
                  ...version.qaWarnings,
                  "原始 PPTX 含登记为不可编辑的图片区域。",
                ]),
              ];
          });
        });
      }
      await this.refreshDocumentConflicts(
        ownerId,
        projectId,
        session.sessionId,
      );
      await this.events.publish({
        type: "document.ready",
        projectId,
        sessionId: session.sessionId,
        payload: {
          sessionId: session.sessionId,
          documentId,
          parseStatus: document.parseStatus,
          factCount: document.facts.length,
        },
      });
      return document;
    } catch (error) {
      await this.store.update((state) => {
        document.parseStatus = "failed";
        document.error = (error as Error).message.slice(0, 1000);
        document.updatedAt = nowIso();
        session.status = "failed";
        session.updatedAt = nowIso();
        state.auditLogs.push({
          auditId: makeId("audit"),
          ownerId,
          projectId,
          action: "document.parse_failed",
          payload: { documentId, message: document.error },
          createdAt: nowIso(),
        });
      });
      await this.events.publish({
        type: "document.failed",
        projectId,
        sessionId: session.sessionId,
        payload: { sessionId: session.sessionId, documentId },
      });
      throw new HttpError(422, `Document parsing failed: ${document.error}`);
    } finally {
      await fs.rm(stagingPath, { force: true }).catch(() => undefined);
    }
  }

  private async refreshDocumentConflicts(
    ownerId: string,
    projectId: string,
    sessionId: string,
  ): Promise<void> {
    const session = this.sessionForProject(ownerId, projectId, sessionId);
    const documents = this.store.state.documents.filter(
      (document) =>
        document.projectId === projectId &&
        document.sessionId === sessionId &&
        document.includedInContext &&
        ["ready", "warning", "partial"].includes(document.parseStatus),
    );
    const groups = new Map<
      string,
      Array<{
        value: string;
        normalizedValue: string;
        documentId: string;
        fileName: string;
        location: string;
        sourceLocator: string;
        originalText: string;
        confidence: number;
        affectedPageIds: string[];
      }>
    >();
    for (const document of documents)
      for (const fact of document.facts) {
        const values = groups.get(fact.key) || [];
        if (
          !values.some(
            (entry) =>
              entry.value === fact.value &&
              entry.documentId === document.documentId,
          )
        )
          values.push({
            value: fact.value,
            normalizedValue: fact.normalizedValue,
            documentId: document.documentId,
            fileName: document.fileName,
            location: fact.location,
            sourceLocator: fact.sourceLocator,
            originalText: fact.context,
            confidence: fact.confidence,
            affectedPageIds: this.getProject(ownerId, projectId).pages
              .filter(
                (page) =>
                  !page.archived &&
                  (page.body.includes(fact.value) ||
                    page.title.includes(fact.value) ||
                    page.factAnchors.some(
                      (anchor) => anchor.value === fact.value,
                    )),
              )
              .map((page) => page.pageId),
          });
        groups.set(fact.key, values);
      }
    const resolvedKeys = new Set(
      this.store.state.documentConflicts
        .filter(
          (conflict) =>
            conflict.projectId === projectId &&
            conflict.sessionId === sessionId &&
            conflict.sourceRevision === session.sourceRevision &&
            conflict.status === "resolved",
        )
        .map((conflict) => conflict.key),
    );
    const created: DocumentConflict[] = [];
    await this.store.update((state) => {
      state.documentConflicts = state.documentConflicts.filter(
        (conflict) =>
          !(
            conflict.projectId === projectId &&
            conflict.sessionId === sessionId &&
            conflict.status === "open"
          ),
      );
      for (const [key, values] of groups) {
        if (
          resolvedKeys.has(key) ||
          new Set(values.map((entry) => entry.value)).size < 2 ||
          new Set(values.map((entry) => entry.documentId)).size < 2
        )
          continue;
        const numericValues = values.map((entry) =>
          Number(entry.normalizedValue.replace("%", "")),
        );
        const formattingOnly =
          numericValues.every(Number.isFinite) &&
          new Set(numericValues).size === 1;
        const conflict: DocumentConflict = {
          conflictId: makeId("conflict"),
          projectId,
          sessionId,
          sourceRevision: session.sourceRevision,
          key,
          values: values.map(({ normalizedValue: _normalizedValue, ...value }) =>
            value,
          ),
          severity: formattingOnly ? "warning" : "blocking",
          status: "open",
          resolution: null,
          createdAt: nowIso(),
          resolvedAt: null,
        };
        state.documentConflicts.push(conflict);
        created.push(conflict);
      }
      session.status = created.some(
        (conflict) => conflict.severity === "blocking",
      )
        ? "blocked"
        : documents.length
          ? "planned"
          : "collecting";
      session.updatedAt = nowIso();
    });
    for (const conflict of created)
      await this.events.publish({
        type: "document.conflict.detected",
        projectId,
        sessionId,
        payload: {
          sessionId,
          conflictId: conflict.conflictId,
          severity: conflict.severity,
        },
      });
  }

  listDocumentConflicts(
    ownerId: string,
    projectId: string,
    sessionId?: string,
  ): DocumentConflict[] {
    this.getProject(ownerId, projectId);
    return this.store.state.documentConflicts.filter(
      (conflict) =>
        conflict.projectId === projectId &&
        (!sessionId || conflict.sessionId === sessionId),
    );
  }

  async resolveDocumentConflict(
    ownerId: string,
    projectId: string,
    conflictId: string,
    input: {
      action?: "prefer_source" | "keep_both" | "ignore";
      selectedValue?: string;
      selectedDocumentIds?: string[];
      note?: string;
    },
  ): Promise<DocumentConflict> {
    this.getProject(ownerId, projectId);
    const conflict = this.store.state.documentConflicts.find(
      (candidate) =>
        candidate.conflictId === conflictId &&
        candidate.projectId === projectId,
    );
    if (!conflict) throw new HttpError(404, "Document conflict not found.");
    const action = input.action || "prefer_source";
    const selectedDocumentIds = [
      ...new Set(
        input.selectedDocumentIds?.filter((documentId) =>
          conflict.values.some((entry) => entry.documentId === documentId),
        ) || [],
      ),
    ];
    const selectedEntry = input.selectedValue
      ? conflict.values.find((entry) => entry.value === input.selectedValue)
      : selectedDocumentIds.length === 1
        ? conflict.values.find(
            (entry) => entry.documentId === selectedDocumentIds[0],
          )
        : undefined;
    if (action === "prefer_source" && !selectedEntry)
      throw new HttpError(
        400,
        "prefer_source requires one selected source value or document.",
      );
    const resolvedDocuments =
      action === "prefer_source"
        ? [selectedEntry!.documentId]
        : action === "keep_both"
          ? conflict.values.map((entry) => entry.documentId)
          : [];
    const resolution: ConflictResolution = {
      resolutionId: makeId("resolution"),
      conflictId,
      projectId,
      sessionId: conflict.sessionId,
      sourceRevision: conflict.sourceRevision,
      planId:
        this.store.state.workSessions.find(
          (session) => session.sessionId === conflict.sessionId,
        )?.planId || null,
      action,
      selectedValue: action === "prefer_source" ? selectedEntry!.value : null,
      selectedDocumentIds: [...new Set(resolvedDocuments)],
      note: input.note?.slice(0, 1000) || "",
      resolvedBy: ownerId,
      createdAt: nowIso(),
    };
    await this.store.update((state) => {
      conflict.status = "resolved";
      conflict.resolution = {
        resolutionId: resolution.resolutionId,
        action: resolution.action,
        selectedValue: resolution.selectedValue,
        selectedDocumentIds: [...resolution.selectedDocumentIds],
        note: resolution.note,
      };
      conflict.resolvedAt = nowIso();
      state.conflictResolutions.push(resolution);
      const session = state.workSessions.find(
        (candidate) => candidate.sessionId === conflict.sessionId,
      );
      if (
        session &&
        !state.documentConflicts.some(
          (candidate) =>
            candidate.sessionId === session.sessionId &&
            candidate.status === "open" &&
            candidate.conflictId !== conflictId,
        )
      ) {
        session.status = "planned";
        session.updatedAt = nowIso();
      }
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId,
        action: "document.conflict.resolved",
        payload: {
          conflictId,
          resolutionId: resolution.resolutionId,
          resolutionAction: resolution.action,
          selectedDocumentIds: resolution.selectedDocumentIds,
          selectedValue: resolution.selectedValue,
        },
        createdAt: nowIso(),
      });
    });
    await this.events.publish({
      type: "document.conflict.resolved",
      projectId,
      sessionId: conflict.sessionId,
      payload: {
        sessionId: conflict.sessionId,
        conflictId,
        resolutionId: resolution.resolutionId,
        action: resolution.action,
      },
    });
    return conflict;
  }

  async createProject(
    ownerId: string,
    input: CreateProjectInput,
  ): Promise<Project> {
    const projectId = makeId("p");
    const goalId = makeId("goal");
    const sourceMarkdown = input.slidesMarkdown?.trim() || DEFAULT_MARKDOWN;
    const parsed = parseSlidesMarkdown(sourceMarkdown);
    const project: Project = {
      projectId,
      ownerId,
      name: input.name,
      themeId: input.themeId,
      themeVersion: input.themeVersion,
      currentDeckRevisionId: makeId("deckrev"),
      goalId,
      status: "ready",
      createdAt: nowIso(),
      updatedAt: nowIso(),
      sourceMarkdown,
      settings: { sensitiveMode: false },
      pages: [],
    };
    project.pages = parsed.map((slide, index) => {
      const pageId = makeId("page");
      const version = initialVersion(
        projectId,
        pageId,
        index + 1,
        project.name,
        slide,
      );
      return {
        pageId,
        projectId,
        currentVersionId: version.versionId,
        orderIndex: index,
        pageType: slide.pageType,
        locked: false,
        archived: false,
        factAnchorIds: slide.facts.map((fact) => fact.factId),
        factAnchors: slide.facts,
        editableLevel: version.editableLevel,
        status: "untouched",
        title: slide.title,
        body: slide.body,
        layout: slide.layout,
        sourceMarkdown: slide.sourceMarkdown,
        contract: slide.contract,
        versions: [version],
      } satisfies Page;
    });
    await this.store.update((state) => {
      state.projects.push(project);
      project.pages.forEach((page) => {
        const version = currentVersion(page);
        state.artifacts.push({
          artifactId: version.previewArtifactId!,
          projectId,
          pageId: page.pageId,
          versionId: version.versionId,
          kind: "svg_preview",
          promptSnapshotId: version.promptSnapshotId,
          model: "deterministic-local-preview",
          width: 960,
          height: 540,
          quality: "quick",
          source: "deterministic_svg",
          provenance: {
            stage: "initial_import",
            sourceRevision: version.sourceRevision,
          },
          createdAt: nowIso(),
        });
      });
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId,
        action: "project.created",
        payload: { pageCount: project.pages.length },
        createdAt: nowIso(),
      });
    });
    await this.events.publish({
      type: "project.created",
      projectId,
      payload: { name: project.name, pageCount: project.pages.length },
    });
    return project;
  }

  async duplicateProject(ownerId: string, projectId: string): Promise<Project> {
    const original = this.getProject(ownerId, projectId);
    return this.createProject(ownerId, {
      name: `${original.name} Copy`,
      themeId: original.themeId,
      themeVersion: original.themeVersion,
      slidesMarkdown: original.pages
        .map((page) => `# ${page.title}\n\n${page.body}`)
        .join("\n\n---\n\n"),
    });
  }

  async importSlides(
    ownerId: string,
    projectId: string,
    sourceMarkdown: string,
  ): Promise<Project> {
    const project = this.getProject(ownerId, projectId);
    const parsed = parseSlidesMarkdown(sourceMarkdown);
    const operationId = makeId("op");
    const orderedExisting = [...project.pages]
      .filter((page) => !page.archived)
      .sort((a, b) => a.orderIndex - b.orderIndex);
    const resolvedPageIds = parsed.map(
      (slide, index) => orderedExisting[index]?.pageId || makeId("page"),
    );
    const operation: EditOperation = {
      operationId,
      projectId,
      conversationId: makeId("conv"),
      mode: "multi",
      requestedPageIds: resolvedPageIds,
      resolvedPageIds,
      message: "Imported slides.md source",
      structuredPlan: {
        workflowMode: "import_document",
        targetScope: "multi",
        intent: "source_import",
        affectedPageIds: resolvedPageIds,
        pageDelta: { add: [], remove: [], split: [], merge: [] },
        changes: parsed.flatMap((slide) =>
          slide.facts.map((fact) => ({
            kind: "preserve_fact" as const,
            factId: fact.factId,
            target: "fact_anchor",
            value: fact.value,
          })),
        ),
        factImpact: { added: [], removed: [], changed: [] },
        sourceDocumentIds: [],
        conflictIds: [],
        unsupported: [],
        requiresConfirmation: false,
        confirmationReasons: [],
        estimatedCost: { imageUnits: 0, amount: 0, currency: "USD" },
        summary: `已导入 ${parsed.length} 页 slides.md，并保留位置匹配的 page_id。`,
      },
      factImpact: { added: [], removed: [], changed: [] },
      unsupportedItems: [],
      confirmationRequired: false,
      confirmedAt: nowIso(),
      resultVersionIds: [],
      status: "completed",
      createdAt: nowIso(),
      completedAt: nowIso(),
      failedPageIds: [],
      estimatedCost: 0,
    };
    await this.store.update((state) => {
      state.operations.push(operation);
      orderedExisting.slice(parsed.length).forEach((page) => {
        page.archived = true;
      });
      parsed.forEach((slide, index) => {
        const pageId = resolvedPageIds[index];
        const existing = orderedExisting[index];
        const version = initialVersion(
          projectId,
          pageId,
          index + 1,
          project.name,
          slide,
        );
        version.parentVersionId = existing?.currentVersionId || null;
        version.editOperationId = operationId;
        version.message = "Imported slides.md";
        const page: Page = existing || {
          pageId,
          projectId,
          currentVersionId: version.versionId,
          orderIndex: index,
          pageType: slide.pageType,
          locked: false,
          archived: false,
          factAnchorIds: [],
          factAnchors: [],
          editableLevel: version.editableLevel,
          status: "untouched",
          title: slide.title,
          body: slide.body,
          layout: slide.layout,
          sourceMarkdown: slide.sourceMarkdown,
          contract: slide.contract,
          versions: [],
        };
        page.currentVersionId = version.versionId;
        page.orderIndex = index;
        page.pageType = slide.pageType;
        page.factAnchorIds = slide.facts.map((fact) => fact.factId);
        page.factAnchors = slide.facts;
        page.title = slide.title;
        page.body = slide.body;
        page.layout = slide.layout;
        page.sourceMarkdown = slide.sourceMarkdown;
        page.contract = slide.contract;
        page.editableLevel = version.editableLevel;
        page.status = "untouched";
        page.versions.push(version);
        state.artifacts.push({
          artifactId: version.previewArtifactId!,
          projectId,
          pageId,
          versionId: version.versionId,
          kind: "svg_preview",
          promptSnapshotId: version.promptSnapshotId,
          model: "deterministic-local-preview",
          width: 960,
          height: 540,
          quality: "quick",
          source: "deterministic_svg",
          provenance: {
            stage: "source_import",
            sourceRevision: version.sourceRevision,
          },
          createdAt: nowIso(),
        });
        if (!existing) project.pages.push(page);
        operation.resultVersionIds.push(version.versionId);
      });
      project.sourceMarkdown = sourceMarkdown;
      project.currentDeckRevisionId = makeId("deckrev");
      project.updatedAt = nowIso();
    });
    await this.events.publish({
      type: "project.imported",
      projectId,
      operationId,
      payload: { pageIds: resolvedPageIds, pageCount: parsed.length },
    });
    return project;
  }

  async setArchived(
    ownerId: string,
    projectId: string,
    archived: boolean,
  ): Promise<Project> {
    const project = this.getProject(ownerId, projectId);
    await this.store.update(() => {
      project.status = archived ? "archived" : "ready";
      project.updatedAt = nowIso();
    });
    await this.events.publish({
      type: archived ? "project.archived" : "project.restored",
      projectId,
      payload: {},
    });
    return project;
  }

  listPages(ownerId: string, projectId: string): Page[] {
    return [...this.getProject(ownerId, projectId).pages]
      .filter((page) => !page.archived)
      .sort((a, b) => a.orderIndex - b.orderIndex);
  }

  getPage(ownerId: string, projectId: string, pageId: string): Page {
    const page = this.getProject(ownerId, projectId).pages.find(
      (candidate) => candidate.pageId === pageId,
    );
    if (!page) throw new HttpError(404, "Page not found.");
    return page;
  }

  async reorderPages(
    ownerId: string,
    projectId: string,
    pageIds: string[],
  ): Promise<Page[]> {
    const project = this.getProject(ownerId, projectId);
    const visible = project.pages.filter((page) => !page.archived);
    const expected = new Set(visible.map((page) => page.pageId));
    if (
      pageIds.length !== visible.length ||
      new Set(pageIds).size !== pageIds.length ||
      pageIds.some((pageId) => !expected.has(pageId))
    ) {
      throw new HttpError(
        400,
        "Reorder must include every visible page exactly once.",
      );
    }
    await this.store.update((state) => {
      pageIds.forEach((pageId, index) => {
        const page = project.pages.find(
          (candidate) => candidate.pageId === pageId,
        )!;
        page.orderIndex = index;
      });
      project.updatedAt = nowIso();
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId,
        action: "page.reordered",
        payload: { pageIds },
        createdAt: nowIso(),
      });
    });
    await this.events.publish({
      type: "page.reordered",
      projectId,
      payload: { pageIds },
    });
    return this.listPages(ownerId, projectId);
  }

  async archivePage(
    ownerId: string,
    projectId: string,
    pageId: string,
    archived: boolean,
  ): Promise<Page> {
    const project = this.getProject(ownerId, projectId);
    const page = this.getPage(ownerId, projectId, pageId);
    const visibleCount = project.pages.filter(
      (candidate) => !candidate.archived,
    ).length;
    if (archived && !page.archived && visibleCount <= 1)
      throw new HttpError(
        400,
        "A project must keep at least one visible page.",
      );
    await this.store.update((state) => {
      page.archived = archived;
      project.updatedAt = nowIso();
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId,
        action: archived ? "page.archived" : "page.restored",
        payload: { pageId },
        createdAt: nowIso(),
      });
    });
    await this.events.publish({
      type: archived ? "page.archived" : "page.restored",
      projectId,
      pageId,
      payload: { pageId },
    });
    return page;
  }

  async splitPage(
    ownerId: string,
    projectId: string,
    pageId: string,
  ): Promise<{ operation: EditOperation }> {
    const project = this.getProject(ownerId, projectId);
    const sourcePage = this.getPage(ownerId, projectId, pageId);
    if (sourcePage.archived)
      throw new HttpError(
        409,
        "Archived pages must be restored before splitting.",
      );
    if (sourcePage.body.trim().length < 20)
      throw new HttpError(400, "Page body is too short to split safely.");
    const operation: EditOperation = {
      operationId: makeId("op"),
      projectId,
      conversationId: makeId("conv"),
      mode: "single",
      requestedPageIds: [pageId],
      resolvedPageIds: [pageId, makeId("page")],
      message: "Split page",
      structuredPlan: {
        workflowMode: "page_entry",
        targetScope: "single",
        intent: "page_split",
        affectedPageIds: [pageId],
        pageDelta: { add: ["pending"], remove: [], split: [pageId], merge: [] },
        changes: [
          {
            kind: "layout_change",
            target: "page_count",
            value: "split_into_two_pages",
          },
        ],
        factImpact: { added: [], removed: [], changed: [] },
        sourceDocumentIds: [],
        conflictIds: [],
        unsupported: [],
        requiresConfirmation: true,
        confirmationReasons: ["页面数量变化必须确认。"],
        estimatedCost: { imageUnits: 0, amount: 0, currency: "USD" },
        summary: "Source page split into two stable page IDs.",
      },
      factImpact: { added: [], removed: [], changed: [] },
      unsupportedItems: [],
      confirmationRequired: true,
      confirmedAt: null,
      resultVersionIds: [],
      status: "planned",
      createdAt: nowIso(),
      completedAt: null,
      failedPageIds: [],
      estimatedCost: 0,
    };
    operation.structuredPlan.affectedPageIds = [...operation.resolvedPageIds];
    await this.store.update((state) => {
      state.operations.push(operation);
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId,
        action: "page.split.planned",
        payload: { pageId, operationId: operation.operationId },
        createdAt: nowIso(),
      });
    });
    await this.events.publish({
      type: "edit.confirmation.required",
      projectId,
      operationId: operation.operationId,
      pageId,
      payload: { pageId, plan: operation.structuredPlan },
    });
    return { operation };
  }

  private async applySplitPage(
    ownerId: string,
    projectId: string,
    pageId: string,
    operationIdOverride?: string,
  ): Promise<{ sourcePage: Page; newPage: Page; operation: EditOperation }> {
    const project = this.getProject(ownerId, projectId);
    const sourcePage = this.getPage(ownerId, projectId, pageId);
    if (sourcePage.archived)
      throw new HttpError(
        409,
        "Archived pages must be restored before splitting.",
      );
    const paragraphs = sourcePage.body
      .split(/\r?\n+/)
      .map((value) => value.trim())
      .filter(Boolean);
    let firstBody: string;
    let secondBody: string;
    if (paragraphs.length > 1) {
      const midpoint = Math.ceil(paragraphs.length / 2);
      firstBody = paragraphs.slice(0, midpoint).join("\n");
      secondBody = paragraphs.slice(midpoint).join("\n");
    } else {
      const body = sourcePage.body.trim();
      if (body.length < 20)
        throw new HttpError(400, "Page body is too short to split safely.");
      const midpoint = Math.ceil(body.length / 2);
      firstBody = body.slice(0, midpoint).trim();
      secondBody = body.slice(midpoint).trim();
    }
    const plannedOperation = operationIdOverride
      ? this.getOperation(ownerId, projectId, operationIdOverride)
      : null;
    if (plannedOperation && plannedOperation.status !== "confirmed")
      throw new HttpError(
        409,
        "Split operation must be confirmed before execution.",
      );
    const operationId = plannedOperation?.operationId || makeId("op");
    const sourceParent = sourcePage.currentVersionId;
    const sourceVersion: PageVersion = {
      ...currentVersion(sourcePage),
      versionId: makeId("ver"),
      parentVersionId: sourceParent,
      sourceRevision: makeId("deckrev"),
      slidesSourceHash: sourceHash(`# ${sourcePage.title}\n\n${firstBody}`),
      previewArtifactId: makeId("artifact"),
      svgArtifactId: makeId("svg"),
      pptxPageRenderId: null,
      promptSnapshotId: makeId("prompt"),
      editOperationId: operationId,
      qualityReportId: makeId("qa"),
      status: "ready",
      createdAt: nowIso(),
      message: "Split source page",
      body: firstBody,
      sourceMarkdown: `# ${sourcePage.title}\n\n${firstBody}`,
      contract: {
        ...sourcePage.contract,
        evidence: firstBody.split(/\r?\n/).filter(Boolean).slice(0, 5),
      },
      previewKind: "svg_fallback",
      previewSvg: makePreviewSvg(
        sourcePage.orderIndex + 1,
        sourcePage.title,
        firstBody,
        sourcePage.layout,
        project.name,
      ),
      renderNote:
        "Split page generated as SVG fallback; authoritative PowerPoint rendering is pending.",
      qaWarnings: ["Awaiting authoritative PowerPoint rendering after split."],
      factAnchors: sourcePage.factAnchors
        .filter(
          (fact) =>
            firstBody.includes(fact.value) || !secondBody.includes(fact.value),
        )
        .map((fact) => ({ ...fact })),
    };
    const continuationTitle = `${sourcePage.title}（续）`;
    const parsedContinuation = parseSlidesMarkdown(
      `# ${continuationTitle}\n\n${secondBody}`,
    )[0];
    const newPageId = plannedOperation?.resolvedPageIds[1] || makeId("page");
    const continuationVersion = initialVersion(
      projectId,
      newPageId,
      sourcePage.orderIndex + 2,
      project.name,
      parsedContinuation,
    );
    continuationVersion.editOperationId = operationId;
    continuationVersion.message = "Created by page split";
    const continuationPage: Page = {
      pageId: newPageId,
      projectId,
      currentVersionId: continuationVersion.versionId,
      orderIndex: sourcePage.orderIndex + 1,
      pageType: sourcePage.pageType,
      locked: false,
      archived: false,
      factAnchorIds: [],
      factAnchors: sourcePage.factAnchors
        .filter((fact) => secondBody.includes(fact.value))
        .map((fact) => ({
          ...fact,
          factId: makeId("fact"),
          source: `${fact.source}:split`,
        })),
      editableLevel: continuationVersion.editableLevel,
      status: "svg_fallback",
      title: continuationTitle,
      body: secondBody,
      layout: sourcePage.layout,
      sourceMarkdown: `# ${continuationTitle}\n\n${secondBody}`,
      contract: parsedContinuation.contract,
      versions: [continuationVersion],
    };
    continuationPage.factAnchorIds = continuationPage.factAnchors.map(
      (fact) => fact.factId,
    );
    continuationVersion.factAnchors = continuationPage.factAnchors.map(
      (fact) => ({ ...fact }),
    );
    const operation: EditOperation = plannedOperation || {
      operationId,
      projectId,
      conversationId: makeId("conv"),
      mode: "single",
      requestedPageIds: [pageId],
      resolvedPageIds: [pageId, newPageId],
      message: "Split page",
      structuredPlan: {
        workflowMode: "page_entry",
        targetScope: "single",
        intent: "page_split",
        affectedPageIds: [pageId, newPageId],
        pageDelta: { add: [newPageId], remove: [], split: [pageId], merge: [] },
        changes: [
          {
            kind: "layout_change",
            target: "page_count",
            value: "split_into_two_pages",
          },
        ],
        factImpact: { added: [], removed: [], changed: [] },
        sourceDocumentIds: [],
        conflictIds: [],
        unsupported: [],
        requiresConfirmation: true,
        confirmationReasons: ["页面数量变化必须确认。"],
        estimatedCost: { imageUnits: 0, amount: 0, currency: "USD" },
        summary: "Source page split into two stable page IDs.",
      },
      factImpact: { added: [], removed: [], changed: [] },
      unsupportedItems: [],
      confirmationRequired: true,
      confirmedAt: nowIso(),
      resultVersionIds: [
        sourceVersion.versionId,
        continuationVersion.versionId,
      ],
      status: "completed",
      createdAt: nowIso(),
      completedAt: nowIso(),
      failedPageIds: [],
      estimatedCost: 0,
    };
    await this.store.update((state) => {
      project.pages
        .filter(
          (page) => !page.archived && page.orderIndex > sourcePage.orderIndex,
        )
        .forEach((page) => {
          page.orderIndex += 1;
        });
      sourcePage.body = firstBody;
      sourcePage.sourceMarkdown = `# ${sourcePage.title}\n\n${firstBody}`;
      sourcePage.contract = {
        ...sourcePage.contract,
        evidence: firstBody.split(/\r?\n/).filter(Boolean).slice(0, 5),
      };
      sourcePage.currentVersionId = sourceVersion.versionId;
      sourcePage.status = "svg_fallback";
      sourcePage.versions.push(sourceVersion);
      sourcePage.factAnchors = sourcePage.factAnchors.filter(
        (fact) =>
          firstBody.includes(fact.value) || !secondBody.includes(fact.value),
      );
      sourcePage.factAnchorIds = sourcePage.factAnchors.map(
        (fact) => fact.factId,
      );
      project.pages.push(continuationPage);
      project.currentDeckRevisionId = sourceVersion.sourceRevision;
      project.updatedAt = nowIso();
      operation.resultVersionIds = [
        sourceVersion.versionId,
        continuationVersion.versionId,
      ];
      operation.confirmedAt = operation.confirmedAt || nowIso();
      operation.status = "completed";
      operation.completedAt = nowIso();
      if (!plannedOperation) state.operations.push(operation);
      [sourceVersion, continuationVersion].forEach((version, index) => {
        const artifactPageId = index === 0 ? pageId : newPageId;
        state.artifacts.push({
          artifactId: version.previewArtifactId!,
          projectId,
          pageId: artifactPageId,
          versionId: version.versionId,
          kind: "svg_preview",
          promptSnapshotId: version.promptSnapshotId,
          model: "deterministic-local-preview",
          width: 960,
          height: 540,
          quality: "quick",
          source: "deterministic_svg",
          provenance: { stage: "page_split", operationId },
          createdAt: nowIso(),
        });
      });
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId,
        action: "page.split",
        payload: { sourcePageId: pageId, newPageId, operationId },
        createdAt: nowIso(),
      });
    });
    await this.events.publish({
      type: "page.split",
      projectId,
      pageId,
      operationId,
      payload: {
        sourcePageId: pageId,
        newPageId,
        versionIds: operation.resultVersionIds,
      },
    });
    return { sourcePage, newPage: continuationPage, operation };
  }

  listMessages(
    ownerId: string,
    projectId: string,
    conversationId?: string,
  ): ConversationMessage[] {
    this.getProject(ownerId, projectId);
    return this.store.state.messages.filter(
      (message) =>
        message.projectId === projectId &&
        (!conversationId || message.conversationId === conversationId),
    );
  }

  projectHistory(ownerId: string, projectId: string): ProjectHistory {
    const project = this.getProject(ownerId, projectId);
    const messages = this.store.state.messages
      .filter((message) => message.projectId === projectId)
      .sort((a, b) => a.createdAt.localeCompare(b.createdAt));
    const sessions = this.store.state.workSessions
      .filter((session) => session.projectId === projectId)
      .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
    const sessionById = new Map(
      sessions.map((session) => [session.sessionId, session]),
    );
    const versionById = new Map(
      project.pages.flatMap((page) =>
        page.versions.map((version) => [
          version.versionId,
          { page, version },
        ] as const),
      ),
    );
    const pageById = new Map(project.pages.map((page) => [page.pageId, page]));

    const operations = this.store.state.operations
      .filter((operation) => operation.projectId === projectId)
      .sort((a, b) => b.createdAt.localeCompare(a.createdAt))
      .map((operation) => {
        const ledger = this.store.state.ledger.filter(
          (entry) => entry.operationId === operation.operationId,
        );
        const prompts = this.store.state.promptSnapshots.filter(
          (snapshot) => snapshot.operationId === operation.operationId,
        );
        const operationVersions = operation.resultVersionIds
          .map((versionId) => versionById.get(versionId))
          .filter(
            (item): item is NonNullable<typeof item> => item !== undefined,
          );
        const pageIds = [
          ...new Set([
            ...operation.resolvedPageIds,
            ...operationVersions.map(({ page }) => page.pageId),
          ]),
        ];
        const pages = pageIds.map((pageId) => {
          const page = pageById.get(pageId);
          const versionEntry = operationVersions.find(
            (item) => item.page.pageId === pageId,
          );
          const pageLedger = ledger.filter((entry) => entry.pageId === pageId);
          const pageWarnings = [
            ...(versionEntry?.version.qaWarnings || []),
            ...(operation.failedPageIds.includes(pageId)
              ? ["该页面处理失败，可单独重试。"]
              : []),
          ];
          const model =
            pageLedger.find((entry) => entry.model)?.model ||
            prompts.find((snapshot) => snapshot.pageId === pageId)?.model ||
            null;
          return {
            pageId,
            pageTitle: page?.title || pageId,
            versionId: versionEntry?.version.versionId || null,
            model,
            cost: pageLedger.reduce(
              (sum, entry) => sum + entry.settledAmount,
              0,
            ),
            qaStatus: operation.failedPageIds.includes(pageId)
              ? ("failed" as const)
              : !versionEntry
                ? ("pending" as const)
                : pageWarnings.length
                  ? ("warning" as const)
                  : ("passed" as const),
            warnings: [...new Set(pageWarnings)],
          };
        });
        const warnings = [
          ...operation.unsupportedItems,
          ...pages.flatMap((page) => page.warnings),
        ];
        const models = [
          ...new Set([
            ...ledger.map((entry) => entry.model),
            ...prompts.map((snapshot) => snapshot.model),
          ]),
        ].filter(Boolean);
        const actualCost = ledger.reduce(
          (sum, entry) => sum + entry.settledAmount,
          0,
        );
        const durationMs = operation.completedAt
          ? Math.max(
              0,
              new Date(operation.completedAt).getTime() -
                new Date(operation.createdAt).getTime(),
            )
          : null;
        const qaStatus = operation.status === "failed"
          ? ("failed" as const)
          : !["completed", "rolled_back"].includes(operation.status)
            ? ("pending" as const)
            : warnings.length || pages.some((page) => page.qaStatus === "warning")
              ? ("warning" as const)
              : ("passed" as const);
        return {
          operationId: operation.operationId,
          conversationId: operation.conversationId,
          sessionId: operation.sessionId || null,
          sessionStatus: operation.sessionId
            ? sessionById.get(operation.sessionId)?.status || null
            : null,
          status: operation.status,
          message: operation.message,
          createdAt: operation.createdAt,
          completedAt: operation.completedAt,
          durationMs,
          estimatedCost: operation.estimatedCost,
          actualCost,
          currency: operation.structuredPlan.estimatedCost.currency || "USD",
          pageIds,
          versionIds: operation.resultVersionIds,
          models,
          promptSnapshotIds: prompts.map(
            (snapshot) => snapshot.promptSnapshotId,
          ),
          qaStatus,
          warnings: [...new Set(warnings)],
          pages,
        };
      });

    return { messages, operations, sessions };
  }

  async createChatTurn(
    ownerId: string,
    input: ChatTurnInput,
    options: { deferExecution?: boolean } = {},
  ): Promise<{ operation: EditOperation; messages: ConversationMessage[] }> {
    const project = this.getProject(ownerId, input.projectId);
    const session = input.sessionId
      ? this.sessionForProject(ownerId, project.projectId, input.sessionId)
      : undefined;
    const blockingConflicts = this.store.state.documentConflicts.filter(
      (conflict) =>
        conflict.projectId === project.projectId &&
        (!session || conflict.sessionId === session.sessionId) &&
        conflict.status === "open" &&
        conflict.severity === "blocking",
    );
    if (blockingConflicts.length) {
      if (options.deferExecution)
        await this.events.publish({
          type: "plan.blocked",
          projectId: project.projectId,
          sessionId: session?.sessionId,
          payload: {
            sessionId: session?.sessionId,
            conflictIds: blockingConflicts.map(
              (conflict) => conflict.conflictId,
            ),
          },
        });
      throw new HttpError(
        409,
        `Resolve ${blockingConflicts.length} blocking document conflict(s) before editing pages.`,
      );
    }
    if (project.currentDeckRevisionId !== input.deckRevisionId)
      throw new HttpError(
        409,
        "Deck revision changed. Refresh project state before retrying.",
      );
    let pageIds = [...new Set(input.target.pageIds)];
    let candidateReasons: Record<string, string> | undefined;
    if (input.target.mode === "global") {
      const resolved = resolveSimilarPages(project, input.message);
      pageIds = resolved.pageIds;
      candidateReasons = resolved.reasons;
    }
    if (!pageIds.length)
      throw new HttpError(400, "At least one target page is required.");
    const pages = pageIds.map((pageId) =>
      this.getPage(ownerId, project.projectId, pageId),
    );
    const plan = buildEditPlan(
      project,
      pages,
      input.target.mode,
      input.message,
      {
        workflowMode:
          session?.workflowMode === "document_import"
            ? "import_document"
            : session?.workflowMode === "ppt_beautify"
              ? "pptx_beautify"
              : "page_entry",
        intent:
          session?.intent === "create"
            ? "create_deck"
            : session?.intent === "reference"
              ? "reference_only"
              : session
                ? "modify_project"
                : "edit_page",
        sourceDocumentIds: session
          ? this.store.state.documents
              .filter(
                (document) =>
                  document.projectId === project.projectId &&
                  document.sessionId === session.sessionId,
              )
              .map((document) => document.documentId)
          : [],
        conflictIds: session
          ? this.store.state.documentConflicts
              .filter(
                (conflict) =>
                  conflict.projectId === project.projectId &&
                  conflict.sessionId === session.sessionId &&
                  ["open", "resolved"].includes(conflict.status),
              )
              .map((conflict) => conflict.conflictId)
          : [],
      },
    );
    plan.candidateReasons = candidateReasons;
    const relayStatus = this.relay.status();
    if (relayStatus.configured) {
      try {
        const relay = await this.relay.structuredJson<{
          summary?: string;
          unsupported?: string[];
        }>({
          system:
            "Return JSON only. You may provide a short summary and unsupported limitations. Never change page IDs, facts, paths, or execute tools.",
          user: JSON.stringify({
            message: input.message,
            pageIds,
            deterministicPlan: plan,
          }),
        });
        if (relay.value.summary?.trim())
          plan.summary = relay.value.summary.trim();
        if (Array.isArray(relay.value.unsupported)) {
          const relayUnsupported = relay.value.unsupported
            .filter((item) => typeof item === "string")
            .map((item) => item.trim())
            .filter(Boolean)
            .slice(0, 5);
          plan.unsupported.push(...relayUnsupported);
          if (relayUnsupported.length)
            plan.confirmationReasons.push(
              "Relay 规划器标记了当前不能结构化执行的内容。",
            );
        }
        plan.requiresConfirmation =
          plan.requiresConfirmation || plan.unsupported.length > 0;
      } catch (error) {
        plan.unsupported.push(
          `Relay planner unavailable: ${(error as Error).message}`,
        );
        plan.confirmationReasons.push(
          "Relay 规划器不可用，计划需要人工确认后执行。",
        );
        plan.requiresConfirmation = true;
      }
    }
    const validatedPlan = editPlanSchema.safeParse(plan);
    if (!validatedPlan.success)
      throw new HttpError(
        422,
        `Structured edit plan validation failed: ${validatedPlan.error.issues.map((issue) => issue.path.join(".") + " " + issue.message).join("; ")}`,
      );
    Object.assign(plan, validatedPlan.data);
    const planErrors = validateEditPlan(
      project,
      pages,
      input.target.mode,
      plan,
    );
    if (planErrors.length)
      throw new HttpError(
        422,
        `Structured edit plan rejected: ${planErrors.join(" ")}`,
      );
    const conversationId = input.conversationId || makeId("conv");
    const operation: EditOperation = {
      operationId: makeId("op"),
      projectId: project.projectId,
      conversationId,
      sessionId: session?.sessionId || null,
      mode: input.target.mode,
      requestedPageIds: input.target.pageIds,
      resolvedPageIds: pageIds,
      message: input.message,
      structuredPlan: plan,
      factImpact: plan.factImpact,
      unsupportedItems: plan.unsupported,
      confirmationRequired: plan.requiresConfirmation,
      confirmedAt: null,
      resultVersionIds: [],
      status: "planned",
      createdAt: nowIso(),
      completedAt: null,
      failedPageIds: [],
      estimatedCost: plan.estimatedCost.amount,
      visualPreviews: [],
    };
    const userMessage: ConversationMessage = {
      messageId: makeId("msg"),
      projectId: project.projectId,
      conversationId,
      role: "user",
      text: input.message,
      createdAt: nowIso(),
      operationId: operation.operationId,
    };
    const planMessage: ConversationMessage = {
      messageId: makeId("msg"),
      projectId: project.projectId,
      conversationId,
      role: "assistant",
      text: plan.summary,
      createdAt: nowIso(),
      operationId: operation.operationId,
      meta: { plan, confirmationRequired: plan.requiresConfirmation },
    };
    await this.store.update((state) => {
      state.operations.push(operation);
      state.messages.push(userMessage, planMessage);
      if (session) {
        session.planId = operation.operationId;
        session.status =
          operation.confirmationRequired || options.deferExecution
            ? "planned"
            : "executing";
        session.updatedAt = nowIso();
      }
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId: project.projectId,
        action: "chat.plan.created",
        payload: {
          operationId: operation.operationId,
          mode: operation.mode,
          pageIds,
        },
        createdAt: nowIso(),
      });
    });
    await this.events.publish({
      type: options.deferExecution ? "plan.created" : "chat.plan.created",
      projectId: project.projectId,
      sessionId: session?.sessionId,
      operationId: operation.operationId,
      payload: { plan },
    });
    await this.prepareOperationVisualPreviews(ownerId, operation, pages);
    if (operation.confirmationRequired || options.deferExecution) {
      await this.events.publish({
        type: "edit.confirmation.required",
        projectId: project.projectId,
        sessionId: session?.sessionId,
        operationId: operation.operationId,
        payload: { pageIds, plan },
      });
      return { operation, messages: [userMessage, planMessage] };
    }
    await this.enqueueEdit(ownerId, operation.operationId);
    return { operation, messages: [userMessage, planMessage] };
  }

  getOperation(
    ownerId: string,
    projectId: string,
    operationId: string,
  ): EditOperation {
    this.getProject(ownerId, projectId);
    const operation = this.store.state.operations.find(
      (candidate) =>
        candidate.operationId === operationId &&
        candidate.projectId === projectId,
    );
    if (!operation) throw new HttpError(404, "Edit operation not found.");
    return operation;
  }

  private async applyDocumentStructurePlan(
    ownerId: string,
    operation: EditOperation,
  ): Promise<EditOperation> {
    const project = this.getProject(ownerId, operation.projectId);
    const session = operation.sessionId
      ? this.sessionForProject(ownerId, project.projectId, operation.sessionId)
      : undefined;
    const structurePlan = session?.structurePlan;
    if (!session || !structurePlan || structurePlan.operationId !== operation.operationId)
      throw new HttpError(409, "Document structure plan is no longer current.");
    await this.events.publish({
      type: "edit.started",
      projectId: project.projectId,
      sessionId: session.sessionId,
      operationId: operation.operationId,
      payload: { pageIds: operation.resolvedPageIds },
    });
    const creatingDeck = session.intent === "create";
    const appliedPages = structurePlan.pages.map((outlinePage, index) => {
      const parsed = parseSlidesMarkdown(
        `# ${outlinePage.title}\n\n${outlinePage.bodyPreview}`,
      )[0];
      const existing = creatingDeck
        ? undefined
        : project.pages.find(
            (page) => page.pageId === outlinePage.pageId && !page.archived,
          );
      if (!creatingDeck && !existing)
        throw new HttpError(
          409,
          `Structure plan page ${outlinePage.pageId} is no longer available.`,
        );
      const version = initialVersion(
        project.projectId,
        outlinePage.pageId,
        index + 1,
        project.name,
        parsed,
      );
      version.parentVersionId = existing?.currentVersionId || null;
      version.editOperationId = operation.operationId;
      version.sourceRevision = makeId("deckrev");
      const page: Page = existing
        ? existing
        : {
            pageId: outlinePage.pageId,
            projectId: project.projectId,
            currentVersionId: version.versionId,
            orderIndex: index,
            pageType: outlinePage.kind === "cover" ? "cover" : "content",
            locked: false,
            archived: false,
            factAnchorIds: parsed.facts.map((fact) => fact.factId),
            factAnchors: parsed.facts,
            editableLevel: "native_structure",
            status: "svg_fallback",
            title: parsed.title,
            body: parsed.body,
            layout: parsed.layout,
            sourceMarkdown: parsed.sourceMarkdown,
            contract: parsed.contract,
            versions: [],
          };
      if (existing) {
        page.orderIndex = index;
        page.pageType = outlinePage.kind === "cover" ? "cover" : "content";
        page.currentVersionId = version.versionId;
        page.factAnchorIds = parsed.facts.map((fact) => fact.factId);
        page.factAnchors = parsed.facts;
        page.editableLevel = version.editableLevel;
        page.status = "svg_fallback";
        page.title = parsed.title;
        page.body = parsed.body;
        page.layout = parsed.layout;
        page.sourceMarkdown = parsed.sourceMarkdown;
        page.contract = parsed.contract;
      }
      page.versions.push(version);
      return { page, version, existing };
    });
    await this.store.update((state) => {
      if (creatingDeck)
        project.pages
          .filter((page) => !page.archived)
          .forEach((page) => {
            page.archived = true;
          });
      for (const { page, version, existing } of appliedPages) {
        if (!existing) project.pages.push(page);
        operation.resultVersionIds.push(version.versionId);
        state.artifacts.push({
          artifactId: version.previewArtifactId!,
          projectId: project.projectId,
          pageId: page.pageId,
          versionId: version.versionId,
          kind: "svg_preview",
          promptSnapshotId: version.promptSnapshotId,
          model: "deterministic-local-preview",
          width: 960,
          height: 540,
          quality: "quick",
          source: "deterministic_svg",
          provenance: {
            stage: "document_structure_plan",
            operationId: operation.operationId,
            sourceDocumentIds:
              structurePlan.pages.find(
                (candidate) => candidate.pageId === page.pageId,
              )?.sourceDocumentIds || [],
          },
          createdAt: nowIso(),
        });
      }
      if (!structurePlan.inheritTheme) {
        project.themeId = "fastppt-editorial";
        project.themeVersion = "1.0.0";
      }
      project.currentDeckRevisionId = makeId("deckrev");
      project.updatedAt = nowIso();
      operation.status = "completed";
      operation.completedAt = nowIso();
      session.status = "completed";
      session.updatedAt = nowIso();
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId: project.projectId,
        action: "document.structure_plan.applied",
        payload: {
          sessionId: session.sessionId,
          operationId: operation.operationId,
          versionIds: operation.resultVersionIds,
        },
        createdAt: nowIso(),
      });
    });
    for (const { page, version } of appliedPages)
      await this.events.publish({
        type: "page.version.created",
        projectId: project.projectId,
        sessionId: session.sessionId,
        pageId: page.pageId,
        versionId: version.versionId,
        operationId: operation.operationId,
        payload: { parentVersionId: version.parentVersionId },
      });
    await this.events.publish({
      type: "edit.completed",
      projectId: project.projectId,
      sessionId: session.sessionId,
      operationId: operation.operationId,
      payload: { versionIds: operation.resultVersionIds, failedPageIds: [] },
    });
    return operation;
  }

  async confirmOperation(
    ownerId: string,
    projectId: string,
    operationId: string,
  ): Promise<EditOperation> {
    const operation = this.getOperation(ownerId, projectId, operationId);
    if (operation.status !== "planned")
      throw new HttpError(409, "Only planned operations can be confirmed.");
    if (this.operationNeedsVisualPreview(operation)) {
      const failed = (operation.visualPreviews || []).filter(
        (preview) => preview.status !== "ready",
      );
      if (failed.length)
        throw new HttpError(
          409,
          `视觉预览尚未就绪：${failed.map((preview) => preview.error || preview.pageId).join("、")}`,
        );
    }
    await this.store.update(() => {
      operation.status = "confirmed";
      operation.confirmedAt = nowIso();
    });
    if (
      operation.structuredPlan.intent === "create_deck_from_documents" ||
      operation.structuredPlan.intent === "document_structure_plan"
    )
      return this.applyDocumentStructurePlan(ownerId, operation);
    if (operation.structuredPlan.intent === "page_split") {
      const sourcePageId = operation.requestedPageIds[0];
      const result = await this.applySplitPage(
        ownerId,
        projectId,
        sourcePageId,
        operationId,
      );
      return result.operation;
    }
    return this.enqueueEdit(ownerId, operationId);
  }

  private async enqueueEdit(
    ownerId: string,
    operationId: string,
  ): Promise<EditOperation> {
    const operation = this.getOperation(
      ownerId,
      this.store.state.operations.find(
        (candidate) => candidate.operationId === operationId,
      )?.projectId || "",
      operationId,
    );
    const project = this.getProject(ownerId, operation.projectId);
    try {
      await this.jobs.enqueue(
        createDurableJob({
          jobId: operationId,
          projectId: project.projectId,
          operationId,
          kind: "edit",
          payload: { operationId },
        }),
        () => this.applyOperation(ownerId, operationId).then(() => undefined),
        true,
      );
    } catch (error) {
      if (operation.status !== "failed") throw error;
    }
    return operation;
  }

  async cancelOperation(
    ownerId: string,
    projectId: string,
    operationId: string,
  ): Promise<EditOperation> {
    const operation = this.getOperation(ownerId, projectId, operationId);
    if (!["planned", "confirmed", "applying"].includes(operation.status))
      throw new HttpError(
        409,
        "Operation cannot be cancelled in its current state.",
      );
    await this.store.update(() => {
      operation.status = "rolled_back";
      operation.completedAt = nowIso();
    });
    await this.events.publish({
      type: "edit.rolled_back",
      projectId,
      operationId,
      payload: { cancelled: true },
    });
    return operation;
  }

  async retryFailedPages(
    ownerId: string,
    projectId: string,
    operationId: string,
  ): Promise<EditOperation> {
    const project = this.getProject(ownerId, projectId);
    const source = this.getOperation(ownerId, projectId, operationId);
    if (!["completed", "failed"].includes(source.status))
      throw new HttpError(
        409,
        "Only completed operations with failed pages can be retried.",
      );
    const pageIds = [...new Set(source.failedPageIds)];
    if (!pageIds.length)
      throw new HttpError(409, "This operation has no failed pages to retry.");
    const pages = pageIds.map((pageId) =>
      this.getPage(ownerId, projectId, pageId),
    );
    const plan = buildEditPlan(project, pages, source.mode, source.message);
    if (source.structuredPlan.candidateReasons) {
      plan.candidateReasons = Object.fromEntries(
        pageIds.flatMap((pageId) => {
          const reason = source.structuredPlan.candidateReasons?.[pageId];
          return reason ? [[pageId, reason]] : [];
        }),
      );
    }
    const now = nowIso();
    const retry: EditOperation = {
      operationId: makeId("op"),
      projectId,
      conversationId: source.conversationId,
      mode: source.mode,
      requestedPageIds: pageIds,
      resolvedPageIds: pageIds,
      message: source.message,
      structuredPlan: {
        ...plan,
        summary: `仅重试上次失败的 ${pageIds.length} 个页面。`,
      },
      factImpact: plan.factImpact,
      unsupportedItems: plan.unsupported,
      confirmationRequired: plan.requiresConfirmation,
      confirmedAt: now,
      resultVersionIds: [],
      status: "confirmed",
      createdAt: now,
      completedAt: null,
      failedPageIds: [],
      estimatedCost: plan.estimatedCost.amount,
      parentOperationId: source.operationId,
    };
    await this.store.update((state) => {
      state.operations.push(retry);
      state.messages
        .filter(
          (message) =>
            message.operationId === source.operationId &&
            Array.isArray(message.meta?.failedPageIds),
        )
        .forEach((message) => {
          message.meta = {
            ...message.meta,
            retryOperationId: retry.operationId,
          };
        });
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId,
        action: "operation.retry_failed_pages",
        payload: {
          sourceOperationId: source.operationId,
          retryOperationId: retry.operationId,
          pageIds,
        },
        createdAt: now,
      });
    });
    await this.events.publish({
      type: "edit.retry.started",
      projectId,
      operationId: retry.operationId,
      payload: { sourceOperationId: source.operationId, pageIds },
    });
    return this.enqueueEdit(ownerId, retry.operationId);
  }

  private async applyOperation(
    ownerId: string,
    operationId: string,
  ): Promise<EditOperation> {
    const operation = this.store.state.operations.find(
      (candidate) => candidate.operationId === operationId,
    );
    if (!operation) throw new HttpError(404, "Edit operation not found.");
    const project = this.getProject(ownerId, operation.projectId);
    await this.store.update((state) => {
      operation.status = "applying";
      if (!state.jobs.some((job) => job.jobId === operationId))
        state.jobs.push(
          createDurableJob({
            jobId: operationId,
            projectId: project.projectId,
            operationId,
            kind: "edit",
            payload: { operationId },
          }),
        );
    });
    await this.events.publish({
      type: "edit.started",
      projectId: project.projectId,
      sessionId: operation.sessionId || undefined,
      operationId,
      payload: { pageIds: operation.resolvedPageIds },
    });
    for (let index = 0; index < operation.resolvedPageIds.length; index += 1) {
      if (operation.status === "rolled_back") break;
      const pageId = operation.resolvedPageIds[index];
      const page = this.getPage(ownerId, project.projectId, pageId);
      const priorVersion = currentVersion(page);
      if (
        page.versions.some((version) =>
          operation.resultVersionIds.includes(version.versionId),
        )
      )
        continue;
      operation.failedPageIds = operation.failedPageIds.filter(
        (failedPageId) => failedPageId !== pageId,
      );
      const imageChange = operation.structuredPlan.changes.some(
        (change) => change.kind === "image_replace",
      );
      const visualCandidate = operation.visualPreviews?.find(
        (preview) => preview.pageId === pageId && preview.status === "ready",
      );
      const relayStatus = this.relay.status();
      const reserve: UsageLedgerEntry | null = visualCandidate
        ? null
        : {
        ledgerId: makeId("ledger"),
        projectId: project.projectId,
        operationId,
        pageId,
        versionId: null,
        model: imageChange ? relayStatus.imageModel : relayStatus.model,
        unitPrice: imageChange
          ? relayStatus.priceSnapshot.imageUnit
          : relayStatus.priceSnapshot.tokenUnit,
        reservedAmount: imageChange ? relayStatus.priceSnapshot.imageUnit : 0,
        settledAmount: 0,
        refundedAmount: 0,
        status: "reserved",
        createdAt: nowIso(),
          };
      try {
        await this.store.update((state) => {
          if (reserve) state.ledger.push(reserve);
          page.status = "generating";
        });
        await this.events.publish({
          type: "edit.progress",
          projectId: project.projectId,
          pageId,
          sessionId: operation.sessionId || undefined,
          operationId,
          payload: {
            completed: index,
            total: operation.resolvedPageIds.length,
            stage: "plan_validated",
          },
        });
        const prompt = composePrompt(
          project,
          page,
          operation.structuredPlan,
          operation.message,
          this.promptContext(operation),
        );
        const edited = previewTextChange(page, operation.message);
        const layoutChange = operation.structuredPlan.changes.find(
          (change) => change.kind === "layout_change",
        )?.value;
        const nextLayout = layoutChange || page.layout;
        let imageAssetPath: string | null = null;
        let imageRequestId: string | null = null;
        let imageModel: string | null = null;
        let generatedVisualArtifactId: string | null =
          visualCandidate?.artifactId || null;
        if (imageChange && visualCandidate?.artifactId) {
          const candidateArtifact = this.store.state.artifacts.find(
            (artifact) => artifact.artifactId === visualCandidate.artifactId,
          );
          imageAssetPath = candidateArtifact?.assetPath || null;
          imageModel = candidateArtifact?.model || relayStatus.imageModel;
          imageRequestId = String(
            candidateArtifact?.provenance.relayRequestId || "",
          );
          if (!imageAssetPath)
            throw new Error("已确认的视觉候选缺少对象存储内容。");
        } else if (imageChange) {
          if (!relayStatus.configured)
            throw new Error(
              "视觉修改需要已配置的 Relay 图像模型；未生成占位图片。",
            );
          const generated = await this.relay.generateImage({
            prompt: `${prompt.prompt}\nGenerate a presentation visual for this page. Preserve every locked fact and do not render text inside the image.`,
            width: 1600,
            height: 900,
            quality: "high",
          });
          imageRequestId = generated.requestId;
          imageModel = generated.model;
          generatedVisualArtifactId = makeId("artifact");
          imageAssetPath = await this.persistGeneratedImage(
            ownerId,
            project.projectId,
            pageId,
            generatedVisualArtifactId,
            generated.bytes,
            generated.mimeType,
          );
        }
        const removedFacts = operation.factImpact.removed.map((value) =>
          value.includes(":") ? value.slice(value.indexOf(":") + 1) : value,
        );
        const changedFacts = factChanges(operation.message);
        if (
          (removedFacts.length > 0 ||
            operation.factImpact.changed.length > 0) &&
          !operation.confirmedAt
        )
          throw new Error("事实变更必须经过确认后才能执行。");
        const contentQa = runContentQa(
          project,
          page,
          edited,
          operation.structuredPlan,
          Boolean(operation.confirmedAt),
        );
        if (!contentQa.passed) throw new Error(contentQa.errors.join(" "));
        const nextFactAnchors = applyFactChanges(
          page.factAnchors,
          operation.message,
          Boolean(operation.confirmedAt),
        ).filter(
          (fact) =>
            !removedFacts.includes(fact.value) &&
            !changedFacts.some((change) => change.oldValue === fact.value),
        );
        let quickPreview = await this.quickPreview.render(
          project,
          page,
          edited.title,
          edited.body,
          nextLayout,
        );
        if (!imageChange && visualCandidate?.artifactId) {
          const candidateArtifact = this.store.state.artifacts.find(
            (artifact) => artifact.artifactId === visualCandidate.artifactId,
          );
          if (candidateArtifact?.assetPath) {
            quickPreview = {
              ...quickPreview,
              svg: (await this.objectStorage.getBytes(candidateArtifact.assetPath)).toString("utf8"),
            };
          }
        }
        let authoritativeRenderId: string | null = null;
        let authoritativeNote = "";
        let authoritativeArtifact: PreviewArtifact | null = null;
        const renderer =
          process.env.POWERPOINT_RENDERER === "powerpoint" ||
          process.env.POWERPOINT_RENDERER === "com";
        const previewArtifactId = makeId("artifact");
        const visualArtifactId = imageChange
          ? generatedVisualArtifactId
          : priorVersion.visualArtifactId || null;
        const nonEditableRegions = visualArtifactId
          ? [...new Set([...priorVersion.nonEditableRegions, "visual_anchor"])]
          : [...priorVersion.nonEditableRegions];
        const version: PageVersion = {
          versionId: makeId("ver"),
          pageId,
          parentVersionId: page.currentVersionId,
          sourceRevision: makeId("deckrev"),
          pageContractPath: `contracts/${pageId}.json`,
          slidesSourceHash: sourceHash(`# ${edited.title}\n\n${edited.body}`),
          previewArtifactId,
          visualArtifactId,
          svgArtifactId: makeId("svg"),
          pptxPageRenderId: null,
          promptSnapshotId:
            visualCandidate?.promptSnapshotId ||
            `prompt_${prompt.hash.slice(0, 20)}`,
          editOperationId: operationId,
          qualityReportId: makeId("qa"),
          status: "previewing",
          createdAt: nowIso(),
          message: operation.message,
          title: edited.title,
          body: edited.body,
          layout: nextLayout,
          sourceMarkdown: `# ${edited.title}\n\n${edited.body}`,
          contract: {
            ...page.contract,
            conclusion: edited.title,
            evidence: edited.body.split(/\r?\n/).filter(Boolean).slice(0, 5),
            mustKeep: nextFactAnchors.map((fact) => fact.value),
          },
          previewKind: "quick_preview",
          previewSvg: quickPreview.svg,
          editableLevel: visualArtifactId
            ? "native_partial"
            : "native_structure",
          nonEditableRegions,
          qaWarnings: [...operation.unsupportedItems, ...contentQa.warnings],
          renderNote: quickPreview.note,
          factAnchors: nextFactAnchors.map((fact) => ({ ...fact })),
        };
        await this.store.update((state) => {
          page.versions.push(version);
          page.currentVersionId = version.versionId;
          page.title = version.title;
          page.body = version.body;
          page.layout = version.layout;
          page.sourceMarkdown =
            version.sourceMarkdown || `# ${version.title}\n\n${version.body}`;
          if (version.contract)
            page.contract = {
              ...version.contract,
              evidence: [...version.contract.evidence],
              mustKeep: [...version.contract.mustKeep],
              canTrim: [...version.contract.canTrim],
              components: [...version.contract.components],
            };
          page.editableLevel = version.editableLevel;
          if (
            removedFacts.length > 0 ||
            operation.factImpact.changed.length > 0
          ) {
            page.factAnchors = nextFactAnchors;
            page.factAnchorIds = page.factAnchors.map((fact) => fact.factId);
          }
          page.status = "quick_preview";
          operation.resultVersionIds.push(version.versionId);
          if (reserve) reserve.versionId = version.versionId;
          const artifact: PreviewArtifact = {
            artifactId: version.previewArtifactId!,
            projectId: project.projectId,
            pageId,
            versionId: version.versionId,
            kind: "svg_preview",
            promptSnapshotId: version.promptSnapshotId,
            model: relayStatus.model,
            width: imageChange ? 1600 : 960,
            height: imageChange ? 900 : 540,
            quality: "quick",
            source: "deterministic_svg",
            provenance: {
              promptHash: prompt.hash,
              plan: operation.structuredPlan,
              engine: quickPreview.engine,
              relayRequestId: imageRequestId,
            },
            createdAt: nowIso(),
          };
          state.artifacts.push(artifact);
          if (imageChange && generatedVisualArtifactId && !visualCandidate) {
            state.artifacts.push({
              artifactId: generatedVisualArtifactId,
              projectId: project.projectId,
              pageId,
              versionId: version.versionId,
              kind: "visual_preview",
              promptSnapshotId: version.promptSnapshotId,
              model: imageModel || relayStatus.imageModel,
              width: 1600,
              height: 900,
              quality: "high",
              source: "relay_image",
              provenance: {
                promptHash: prompt.hash,
                plan: operation.structuredPlan,
                relayRequestId: imageRequestId,
              },
              createdAt: nowIso(),
              assetPath: imageAssetPath,
            });
          }
          if (visualCandidate?.artifactId) {
            const candidateArtifact = state.artifacts.find(
              (candidate) => candidate.artifactId === visualCandidate.artifactId,
            );
            if (candidateArtifact) candidateArtifact.versionId = version.versionId;
          }
          const snapshot: PromptSnapshot = {
            promptSnapshotId: version.promptSnapshotId,
            projectId: project.projectId,
            pageId,
            operationId,
            prompt: prompt.prompt,
            hash: prompt.hash,
            model: process.env.RELAY_MODEL || "deterministic-local-planner",
            createdAt: nowIso(),
          };
          if (
            !state.promptSnapshots.some(
              (candidate) =>
                candidate.promptSnapshotId === snapshot.promptSnapshotId,
            )
          )
            state.promptSnapshots.push(snapshot);
        });
        await this.events.publish({
          type: "page.version.created",
          projectId: project.projectId,
          sessionId: operation.sessionId || undefined,
          pageId,
          versionId: version.versionId,
          operationId,
          payload: { parentVersionId: version.parentVersionId },
        });
        await this.events.publish({
          type: "preview.quick.ready",
          projectId: project.projectId,
          sessionId: operation.sessionId || undefined,
          pageId,
          versionId: version.versionId,
          operationId,
          payload: { previewKind: "quick_preview" },
        });
        if (renderer) {
          try {
            const pageRenderDir = path.resolve(
              process.env.DATA_DIR || "./data",
              "page-renders",
              ownerId,
              project.projectId,
              pageId,
              version.versionId,
            );
            const renderProject = { ...project, pages: [page] };
            const exportResult = await runPptxExport(
              renderProject,
              pageRenderDir,
              `${pageId}-${version.versionId}.pptx`,
              await this.exportVisualAssets(renderProject),
            );
            const renderResult = await runPowerPointRender(
              exportResult.outputPath,
              path.join(pageRenderDir, "powerpoint"),
            );
            if (renderResult.renders.length !== 1)
              throw new Error(
                `Expected one authoritative page render, received ${renderResult.renders.length}.`,
              );
            const render = renderResult.renders[0];
            const renderBytes = await fs.readFile(render.path);
            if (
              renderBytes.length < 1_000 ||
              render.width <= 0 ||
              render.height <= 0
            )
              throw new Error(
                "PowerPoint returned an empty or invalid PNG render.",
              );
            authoritativeRenderId = makeId("artifact");
            const stored = await this.objectStorage.putFile(
              render.path,
              `projects/${ownerId}/${project.projectId}/pages/${pageId}/versions/${version.versionId}/${authoritativeRenderId}.png`,
              "image/png",
            );
            authoritativeArtifact = {
              artifactId: authoritativeRenderId,
              projectId: project.projectId,
              pageId,
              versionId: version.versionId,
              kind: "pptx_page_render",
              promptSnapshotId: version.promptSnapshotId,
              model: renderResult.renderer,
              width: render.width,
              height: render.height,
              quality: "authoritative",
              source: "powerpoint_com",
              provenance: {
                renderer: renderResult.renderer,
                sha256: sha256(renderBytes),
                sourceRevision: version.sourceRevision,
                slideIndex: render.slide_index,
              },
              createdAt: nowIso(),
              assetPath: stored.objectKey,
            };
            authoritativeNote = `PowerPoint COM rendered ${renderResult.renders.length} page(s).`;
          } catch (renderError) {
            authoritativeNote = `PowerPoint COM render unavailable: ${(renderError as Error).message}`;
          }
        }
        const authoritative = Boolean(authoritativeRenderId);
        await this.store.update((state) => {
          if (authoritativeArtifact)
            state.artifacts.push(authoritativeArtifact);
          version.status = "ready";
          version.previewKind = authoritative
            ? "pptx_authoritative"
            : "svg_fallback";
          version.pptxPageRenderId = authoritativeRenderId;
          version.renderNote = authoritative
            ? authoritativeNote
            : `PowerPoint worker is unavailable. This is an SVG fallback and may differ from PPTX.${authoritativeNote ? ` ${authoritativeNote}` : ""}`;
          if (!authoritative)
            version.qaWarnings.push(
              authoritativeNote ||
                "Awaiting authoritative PowerPoint rendering.",
            );
          page.status = authoritative ? "authoritative" : "svg_fallback";
          if (reserve) {
            reserve.status = "settled";
            reserve.settledAmount = reserve.reservedAmount;
          }
          project.currentDeckRevisionId = version.sourceRevision;
          project.updatedAt = nowIso();
        });
        if (authoritative) {
          await this.events.publish({
            type: "preview.pptx.ready",
            projectId: project.projectId,
            sessionId: operation.sessionId || undefined,
            pageId,
            versionId: version.versionId,
            operationId,
            payload: { renderId: version.pptxPageRenderId },
          });
        } else {
          await this.events.publish({
            type: "render.warning",
            projectId: project.projectId,
            sessionId: operation.sessionId || undefined,
            pageId,
            versionId: version.versionId,
            operationId,
            payload: {
              reason: "powerpoint_worker_unavailable",
              previewKind: "svg_fallback",
            },
          });
        }
      } catch (error) {
        await this.store.update(() => {
          page.status = "failed";
          operation.failedPageIds.push(pageId);
          if (reserve) {
            reserve.status = "refunded";
            reserve.refundedAmount = reserve.reservedAmount;
          }
        });
        await this.events.publish({
          type: "edit.failed",
          projectId: project.projectId,
          sessionId: operation.sessionId || undefined,
          pageId,
          operationId,
          payload: { message: (error as Error).message },
        });
      }
    }
    await this.store.update(() => {
      operation.status =
        operation.failedPageIds.length === operation.resolvedPageIds.length
          ? "failed"
          : "completed";
      operation.completedAt = nowIso();
    });
    if (operation.status === "failed")
      throw new Error(
        `Edit operation ${operationId} failed for every target page.`,
      );
    const completionMessage: ConversationMessage = {
      messageId: makeId("msg"),
      projectId: project.projectId,
      conversationId: operation.conversationId,
      role: "assistant",
      text: operation.failedPageIds.length
        ? `操作已完成，${operation.resultVersionIds.length} 页成功，${operation.failedPageIds.length} 页可单独重试。`
        : `操作已完成，生成 ${operation.resultVersionIds.length} 个不可变页面版本。`,
      createdAt: nowIso(),
      operationId,
      meta: {
        versionIds: operation.resultVersionIds,
        failedPageIds: operation.failedPageIds,
      },
    };
    await this.store.update((state) => state.messages.push(completionMessage));
    await this.events.publish({
      type: "edit.completed",
      projectId: project.projectId,
      sessionId: operation.sessionId || undefined,
      operationId,
      payload: {
        versionIds: operation.resultVersionIds,
        failedPageIds: operation.failedPageIds,
      },
    });
    return operation;
  }

  listVersions(
    ownerId: string,
    projectId: string,
    pageId: string,
  ): PageVersion[] {
    return [...this.getPage(ownerId, projectId, pageId).versions].sort((a, b) =>
      b.createdAt.localeCompare(a.createdAt),
    );
  }

  async rollbackVersion(
    ownerId: string,
    projectId: string,
    versionId: string,
  ): Promise<Page> {
    const project = this.getProject(ownerId, projectId);
    const page = project.pages.find((candidate) =>
      candidate.versions.some((version) => version.versionId === versionId),
    );
    const version = page?.versions.find(
      (candidate) => candidate.versionId === versionId,
    );
    if (!page || !version) throw new HttpError(404, "Version not found.");
    await this.store.update(() => {
      page.currentVersionId = version.versionId;
      page.title = version.title;
      page.body = version.body;
      page.layout = version.layout;
      page.sourceMarkdown =
        version.sourceMarkdown || `# ${version.title}\n\n${version.body}`;
      if (version.contract)
        page.contract = {
          ...version.contract,
          evidence: [...version.contract.evidence],
          mustKeep: [...version.contract.mustKeep],
          canTrim: [...version.contract.canTrim],
          components: [...version.contract.components],
        };
      page.editableLevel = version.editableLevel;
      if (version.factAnchors) {
        page.factAnchors = version.factAnchors.map((fact) => ({ ...fact }));
        page.factAnchorIds = page.factAnchors.map((fact) => fact.factId);
      }
      page.status = "rolled_back";
      project.currentDeckRevisionId = makeId("deckrev");
      project.updatedAt = nowIso();
    });
    await this.store.update((state) =>
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId,
        action: "version.rollback",
        payload: { pageId: page.pageId, versionId },
        createdAt: nowIso(),
      }),
    );
    await this.events.publish({
      type: "edit.rolled_back",
      projectId,
      pageId: page.pageId,
      versionId,
      payload: { restoredVersionId: versionId },
    });
    return page;
  }

  async rollbackOperation(
    ownerId: string,
    projectId: string,
    operationId: string,
  ): Promise<EditOperation> {
    const project = this.getProject(ownerId, projectId);
    const operation = this.getOperation(ownerId, projectId, operationId);
    if (!["completed", "failed"].includes(operation.status))
      throw new HttpError(
        409,
        "Only completed operations can be rolled back as a group.",
      );
    await this.store.update((state) => {
      for (const page of project.pages) {
        const applied = page.versions.find((version) =>
          operation.resultVersionIds.includes(version.versionId),
        );
        if (!applied) continue;
        const parent = applied.parentVersionId
          ? page.versions.find(
              (version) => version.versionId === applied.parentVersionId,
            )
          : null;
        if (!parent) continue;
        page.currentVersionId = parent.versionId;
        page.title = parent.title;
        page.body = parent.body;
        page.layout = parent.layout;
        page.sourceMarkdown =
          parent.sourceMarkdown || `# ${parent.title}\n\n${parent.body}`;
        if (parent.contract)
          page.contract = {
            ...parent.contract,
            evidence: [...parent.contract.evidence],
            mustKeep: [...parent.contract.mustKeep],
            canTrim: [...parent.contract.canTrim],
            components: [...parent.contract.components],
          };
        page.editableLevel = parent.editableLevel;
        if (parent.factAnchors) {
          page.factAnchors = parent.factAnchors.map((fact) => ({ ...fact }));
          page.factAnchorIds = page.factAnchors.map((fact) => fact.factId);
        }
        page.status = "rolled_back";
      }
      if (operation.structuredPlan.intent === "page_split") {
        const continuation = project.pages.find(
          (page) => page.pageId === operation.resolvedPageIds[1],
        );
        if (continuation) continuation.archived = true;
      }
      operation.status = "rolled_back";
      operation.completedAt = nowIso();
      project.currentDeckRevisionId = makeId("deckrev");
      project.updatedAt = nowIso();
      state.messages
        .filter(
          (message) =>
            message.operationId === operationId &&
            Array.isArray(message.meta?.versionIds),
        )
        .forEach((message) => {
          message.meta = { ...message.meta, rolledBack: true };
        });
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId,
        action: "operation.rollback",
        payload: { operationId },
        createdAt: nowIso(),
      });
    });
    await this.events.publish({
      type: "edit.rolled_back",
      projectId,
      operationId,
      payload: {
        operationRollback: true,
        versionIds: operation.resultVersionIds,
      },
    });
    return operation;
  }

  async createExport(ownerId: string, projectId: string): Promise<ExportJob> {
    const project = this.getProject(ownerId, projectId);
    const exportId = makeId("export");
    const rendererConfigured =
      process.env.POWERPOINT_RENDERER === "powerpoint" ||
      process.env.POWERPOINT_RENDERER === "com";
    const exportEngine =
      process.env.PPTX_EXPORT_ENGINE === "legacy"
        ? "legacy_python_pptx_development_fallback"
        : "ppt_master_svg_to_drawingml";
    const job: ExportJob = {
      exportId,
      projectId,
      status: "queued",
      artifactPath: null,
      renderMode: rendererConfigured ? "powerpoint" : "svg_fallback",
      qaWarnings: rendererConfigured
        ? []
        : [
            "PowerPoint COM render worker is unavailable; export is explicitly marked SVG fallback.",
          ],
      createdAt: nowIso(),
      completedAt: null,
      exportEngine,
      qaStatus: "pending",
    };
    await this.store.update((state) => {
      state.exports.push(job);
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId,
        action: "export.created",
        payload: { exportId, renderMode: job.renderMode, exportEngine },
        createdAt: nowIso(),
      });
    });
    void this.jobs.enqueue(
      createDurableJob({
        jobId: exportId,
        projectId,
        kind: "export",
        payload: { exportId },
      }),
      () => this.executeExport(project, job),
    );
    return job;
  }

  private async executeExport(project: Project, job: ExportJob): Promise<void> {
    try {
      await this.store.update((state) => {
        job.status = "running";
      });
      const outputDir = path.resolve(
        process.env.DATA_DIR || "./data",
        "exports",
        project.ownerId,
        project.projectId,
      );
      const exportResult = await runPptxExport(
        project,
        outputDir,
        `${safeFileName(project.name)}-${job.exportId}.pptx`,
        await this.exportVisualAssets(project),
      );
      job.artifactName = path.basename(exportResult.outputPath);
      if (!exportResult.qaPath)
        throw new Error("PPTX export worker did not return its QA receipt.");
      const qa = JSON.parse(await fs.readFile(exportResult.qaPath, "utf8")) as {
        export_engine?: ExportJob["exportEngine"];
        static_structure_passed?: boolean;
        svg_quality?: { summary?: { warnings?: number } };
        pptx_postflight?: { status?: string };
        delivery_check?: { status?: string };
      };
      job.exportEngine = qa.export_engine || job.exportEngine;
      const isDevelopmentFallback =
        job.exportEngine === "legacy_python_pptx_development_fallback";
      if (isDevelopmentFallback) {
        job.qaStatus = "not-run-development-fallback";
      } else {
        if (qa.static_structure_passed !== true)
          throw new Error("Static PPTX structure QA failed.");
        const postflightStatus = qa.pptx_postflight?.status;
        const deliveryStatus = qa.delivery_check?.status;
        if (
          ![
            "passed",
            "passed-with-warnings",
            "passed-with-advisories",
          ].includes(String(postflightStatus))
        ) {
          throw new Error(
            `PPTX postflight QA failed with status ${String(postflightStatus || "missing")}.`,
          );
        }
        if (
          !["passed", "passed-with-advisories"].includes(String(deliveryStatus))
        ) {
          throw new Error(
            `PPTX delivery QA failed with status ${String(deliveryStatus || "missing")}.`,
          );
        }
        job.qaStatus =
          Number(qa.svg_quality?.summary?.warnings || 0) > 0 ||
          postflightStatus !== "passed" ||
          deliveryStatus !== "passed"
            ? "passed-with-warnings"
            : "passed";
      }
      const storedArtifact = await this.objectStorage.putFile(
        exportResult.outputPath,
        `projects/${project.ownerId}/${project.projectId}/exports/${job.artifactName}`,
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      );
      job.artifactObjectKey = storedArtifact.objectKey;
      if (job.renderMode === "powerpoint") {
        try {
          const renderResult = await runPowerPointRender(
            exportResult.outputPath,
            path.join(outputDir, "powerpoint"),
          );
          job.qaWarnings.push(
            `PowerPoint COM rendered ${renderResult.renders.length} slide(s).`,
          );
        } catch (renderError) {
          job.renderMode = "svg_fallback";
          job.qaWarnings.push(
            `PowerPoint COM render failed; SVG fallback retained: ${(renderError as Error).message}`,
          );
        }
      }
      await this.store.update(() => {
        job.status = "completed";
        job.artifactPath = exportResult.outputPath;
        job.completedAt = nowIso();
      });
      await this.events.publish({
        type: "export.completed",
        projectId: project.projectId,
        exportId: job.exportId,
        payload: { renderMode: job.renderMode, qaWarnings: job.qaWarnings },
      });
    } catch (error) {
      await this.store.update(() => {
        job.status = "failed";
        job.qaStatus = "failed";
        job.qaWarnings.push((error as Error).message);
        job.completedAt = nowIso();
      });
      await this.events.publish({
        type: "export.failed",
        projectId: project.projectId,
        exportId: job.exportId,
        payload: { message: (error as Error).message },
      });
      throw error;
    }
  }

  async resumeJobs(): Promise<void> {
    const pending = this.store.state.jobs.filter((job) =>
      ["queued", "running"].includes(job.status),
    );
    await this.jobs.resume(pending, (durable) => {
      if (durable.kind === "edit") {
        const operation = this.store.state.operations.find(
          (candidate) => candidate.operationId === durable.operationId,
        );
        const project = this.store.state.projects.find(
          (candidate) => candidate.projectId === durable.projectId,
        );
        if (
          !operation ||
          !project ||
          !["confirmed", "applying"].includes(operation.status)
        )
          return null;
        return () =>
          this.applyOperation(project.ownerId, operation.operationId).then(
            () => undefined,
          );
      }
      const exportId = String(durable.payload.exportId || durable.jobId);
      const job = this.store.state.exports.find(
        (candidate) => candidate.exportId === exportId,
      );
      const project = this.store.state.projects.find(
        (candidate) => candidate.projectId === durable.projectId,
      );
      if (!job || !project) return null;
      job.status = "queued";
      return () => this.executeExport(project, job);
    });
  }

  getExport(ownerId: string, projectId: string, exportId: string): ExportJob {
    this.getProject(ownerId, projectId);
    const job = this.store.state.exports.find(
      (candidate) =>
        candidate.exportId === exportId && candidate.projectId === projectId,
    );
    if (!job) throw new HttpError(404, "Export not found.");
    return job;
  }

  async issueOfficePreviewGrant(
    ownerId: string,
    projectId: string,
    exportId: string,
  ): Promise<{
    grantId: string;
    token: string;
    expiresAt: string;
    publicPath: string;
  }> {
    const project = this.getProject(ownerId, projectId);
    if (project.settings?.sensitiveMode)
      throw new HttpError(403, "Office preview is disabled in sensitive mode.");
    const job = this.getExport(ownerId, projectId, exportId);
    if (job.status !== "completed")
      throw new HttpError(409, "Export is not ready.");
    const token = makeId("office");
    const expiresAt = new Date(Date.now() + 10 * 60 * 1000).toISOString();
    const grant = {
      grantId: makeId("grant"),
      exportId,
      projectId,
      issuedTo: ownerId,
      tokenHash: sha256(token),
      expiresAt,
      maxFetches: 6,
      fetchCount: 0,
      revokedAt: null,
      createdAt: nowIso(),
    };
    await this.store.update((state) => {
      state.officePreviewGrants.push(grant);
      state.auditLogs.push({
        auditId: makeId("audit"),
        ownerId,
        projectId,
        action: "office_preview.issued",
        payload: { grantId: grant.grantId, exportId },
        createdAt: nowIso(),
      });
    });
    await this.events.publish({
      type: "office_preview.issued",
      projectId,
      exportId,
      payload: { grantId: grant.grantId, expiresAt },
    });
    return {
      grantId: grant.grantId,
      token,
      expiresAt,
      publicPath: `/public/office-preview/${encodeURIComponent(token)}`,
    };
  }

  async readOfficePreviewGrant(
    token: string,
    method: "GET" | "HEAD",
    range?: string,
  ): Promise<{
    bytes: Buffer;
    fileName: string;
    contentRange?: string;
    fullLength: number;
  }> {
    const grant = this.store.state.officePreviewGrants.find(
      (candidate) => candidate.tokenHash === sha256(token),
    );
    if (
      !grant ||
      grant.revokedAt ||
      new Date(grant.expiresAt).getTime() <= Date.now() ||
      grant.fetchCount >= grant.maxFetches
    )
      throw new HttpError(404, "Preview not found.");
    const project = this.store.state.projects.find(
      (candidate) => candidate.projectId === grant.projectId,
    );
    if (!project || project.settings?.sensitiveMode)
      throw new HttpError(404, "Preview not found.");
    const artifact = await this.readExportArtifact(
      project.ownerId,
      grant.projectId,
      grant.exportId,
    );
    await this.store.update((state) => {
      const current = state.officePreviewGrants.find(
        (candidate) => candidate.grantId === grant.grantId,
      );
      if (current) current.fetchCount += 1;
    });
    if (method === "HEAD")
      return {
        bytes: Buffer.alloc(0),
        fileName: artifact.fileName,
        fullLength: artifact.bytes.length,
      };
    if (!range) return { ...artifact, fullLength: artifact.bytes.length };
    const match = /^bytes=(\d+)-(\d*)$/i.exec(range);
    if (!match) return { ...artifact, fullLength: artifact.bytes.length };
    const start = Number(match[1]);
    const end = match[2] ? Number(match[2]) : artifact.bytes.length - 1;
    if (start >= artifact.bytes.length || end < start)
      throw new HttpError(416, "Requested range is not satisfiable.");
    const boundedEnd = Math.min(end, artifact.bytes.length - 1);
    return {
      bytes: artifact.bytes.subarray(start, boundedEnd + 1),
      fileName: artifact.fileName,
      contentRange: `bytes ${start}-${boundedEnd}/${artifact.bytes.length}`,
      fullLength: artifact.bytes.length,
    };
  }

  async readExportArtifact(
    ownerId: string,
    projectId: string,
    exportId: string,
  ): Promise<{ bytes: Buffer; fileName: string }> {
    const job = this.getExport(ownerId, projectId, exportId);
    if (job.status !== "completed")
      throw new HttpError(409, "Export is not ready.");
    if (job.artifactObjectKey)
      return {
        bytes: await this.objectStorage.getBytes(job.artifactObjectKey),
        fileName: job.artifactName || `${exportId}.pptx`,
      };
    if (!job.artifactPath)
      throw new HttpError(409, "Export artifact is unavailable.");
    const root = path.resolve(process.env.DATA_DIR || "./data");
    const artifact = path.resolve(job.artifactPath);
    if (!artifact.startsWith(`${root}${path.sep}`))
      throw new HttpError(403, "Artifact path is outside the project store.");
    return {
      bytes: await fs.readFile(artifact),
      fileName: job.artifactName || path.basename(artifact),
    };
  }

  async readAuthoritativeRender(
    ownerId: string,
    projectId: string,
    artifactId: string,
  ): Promise<{ bytes: Buffer; width: number; height: number }> {
    const project = this.getProject(ownerId, projectId);
    const artifact = this.store.state.artifacts.find(
      (candidate) =>
        candidate.artifactId === artifactId &&
        candidate.projectId === projectId,
    );
    if (
      !artifact ||
      artifact.kind !== "pptx_page_render" ||
      artifact.source !== "powerpoint_com" ||
      !artifact.assetPath
    )
      throw new HttpError(404, "Authoritative page render not found.");
    const page = project.pages.find(
      (candidate) => candidate.pageId === artifact.pageId,
    );
    const version = page?.versions.find(
      (candidate) => candidate.versionId === artifact.versionId,
    );
    if (!version || version.pptxPageRenderId !== artifact.artifactId)
      throw new HttpError(
        404,
        "Authoritative page render is not bound to this project version.",
      );
    return {
      bytes: await this.objectStorage.getBytes(artifact.assetPath),
      width: artifact.width,
      height: artifact.height,
    };
  }

  async readVisualPreview(
    ownerId: string,
    projectId: string,
    artifactId: string,
  ): Promise<{ bytes: Buffer; width: number; height: number; contentType: string }> {
    this.getProject(ownerId, projectId);
    const artifact = this.store.state.artifacts.find(
      (candidate) =>
        candidate.artifactId === artifactId &&
        candidate.projectId === projectId &&
        candidate.kind === "visual_preview" &&
        candidate.assetPath,
    );
    if (!artifact?.assetPath)
      throw new HttpError(404, "Visual preview not found.");
    const extension = path.extname(artifact.assetPath).toLowerCase();
    const contentType =
      extension === ".svg"
        ? "image/svg+xml"
        : extension === ".jpg" || extension === ".jpeg"
          ? "image/jpeg"
          : extension === ".webp"
            ? "image/webp"
            : "image/png";
    return {
      bytes: await this.objectStorage.getBytes(artifact.assetPath),
      width: artifact.width,
      height: artifact.height,
      contentType,
    };
  }

  usage(ownerId: string, projectId: string): UsageLedgerEntry[] {
    this.getProject(ownerId, projectId);
    return this.store.state.ledger.filter(
      (entry) => entry.projectId === projectId,
    );
  }

  relayStatus() {
    return this.relay.status();
  }

  audit(ownerId: string, projectId: string) {
    this.getProject(ownerId, projectId);
    return this.store.state.auditLogs.filter(
      (entry) => entry.ownerId === ownerId && entry.projectId === projectId,
    );
  }

  artifacts(ownerId: string, projectId: string) {
    this.getProject(ownerId, projectId);
    return this.store.state.artifacts.filter(
      (artifact) => artifact.projectId === projectId,
    );
  }

  promptSnapshots(ownerId: string, projectId: string) {
    this.getProject(ownerId, projectId);
    return this.store.state.promptSnapshots.filter(
      (snapshot) => snapshot.projectId === projectId,
    );
  }
}

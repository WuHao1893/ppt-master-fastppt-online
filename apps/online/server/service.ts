import fs from 'node:fs/promises';
import path from 'node:path';
import type {
  ConversationMessage,
  EditOperation,
  ExportJob,
  Page,
  PageVersion,
  Project,
  PromptSnapshot,
  UsageLedgerEntry,
  PreviewArtifact,
} from '../shared/models.js';
import { editPlanSchema, type ChatTurnInput, type CreateProjectInput } from '../shared/protocol.js';
import { parseSlidesMarkdown, sourceHash } from './contracts.js';
import { EventBus } from './events.js';
import { HttpError } from './errors.js';
import { applyFactChanges, buildEditPlan, composePrompt, factChanges, previewTextChange, resolveSimilarPages, validateEditPlan } from './plans.js';
import { makePreviewSvg } from './preview.js';
import type { StateStore } from './store.js';
import { makeId, nowIso, safeFileName, sha256 } from './utils.js';
import { runPowerPointRender, runPptxExport, type PptxVisualAsset } from './workerBridge.js';
import { RelayModelAdapter } from './relay.js';
import { SlidevQuickPreviewWorker } from './quickPreview.js';
import { ObjectStorage, type ObjectStoreStatus } from './objectStore.js';
import { runContentQa } from './contentQa.js';
import { DurableJobQueue } from './jobQueue.js';

const DEFAULT_MARKDOWN = `# FastPPT Online

选择页面并通过聊天精修，所有版本都保留稳定 page_id。

---

# 从快速预览到权威渲染

快速预览用于即时反馈，最终真相来自同版本的 PPTX PowerPoint 渲染。

---

# 事实、成本与可编辑交付

事实锚点默认锁定。图片生成先预留成本，最终 PPTX 不使用整页截图冒充可编辑内容。`;

function initialVersion(projectId: string, pageId: string, pageNumber: number, projectName: string, slide: ReturnType<typeof parseSlidesMarkdown>[number]): PageVersion {
  const versionId = makeId('ver');
  return {
    versionId,
    pageId,
    parentVersionId: null,
    sourceRevision: `deckrev_${sha256(`${projectId}:${pageId}`).slice(0, 12)}`,
    pageContractPath: `contracts/${pageId}.json`,
    slidesSourceHash: sourceHash(slide.sourceMarkdown),
    previewArtifactId: makeId('artifact'),
    visualArtifactId: null,
    svgArtifactId: makeId('svg'),
    pptxPageRenderId: null,
    promptSnapshotId: makeId('prompt'),
    editOperationId: `import_${projectId}`,
    qualityReportId: makeId('qa'),
    status: 'ready',
    createdAt: nowIso(),
    message: 'Initial import',
    title: slide.title,
    body: slide.body,
    layout: slide.layout,
    sourceMarkdown: slide.sourceMarkdown,
    contract: { ...slide.contract, evidence: [...slide.contract.evidence], mustKeep: [...slide.contract.mustKeep], canTrim: [...slide.contract.canTrim], components: [...slide.contract.components] },
    previewKind: 'svg_fallback',
    previewSvg: makePreviewSvg(pageNumber, slide.title, slide.body, slide.layout, projectName),
    editableLevel: 'native_partial',
    nonEditableRegions: [],
    qaWarnings: ['PowerPoint authority has not rendered this initial version.'],
    renderNote: 'Quick SVG preview. The editable PPTX authority is pending.',
    factAnchors: slide.facts.map((fact) => ({ ...fact })),
  };
}

function currentVersion(page: Page): PageVersion {
  return page.versions.find((version) => version.versionId === page.currentVersionId) || page.versions[page.versions.length - 1];
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

  private async exportVisualAssets(project: Project): Promise<Record<string, PptxVisualAsset>> {
    const assets: Record<string, PptxVisualAsset> = {};
    for (const page of project.pages.filter((candidate) => !candidate.archived)) {
      const version = currentVersion(page);
      if (!version.visualArtifactId) continue;
      const artifact = this.store.state.artifacts.find((candidate) => candidate.artifactId === version.visualArtifactId);
      if (!artifact?.assetPath || artifact.source !== 'relay_image') {
        throw new Error(`The current visual asset for page ${page.pageId} is missing or invalid.`);
      }
      const extension = path.extname(artifact.assetPath).toLowerCase();
      if (!['.png', '.jpg', '.jpeg', '.webp'].includes(extension)) {
        throw new Error(`The current visual asset for page ${page.pageId} has an unsupported image format.`);
      }
      assets[page.pageId] = {
        bytes: await this.objectStorage.getBytes(artifact.assetPath),
        extension: extension as PptxVisualAsset['extension'],
      };
    }
    return assets;
  }

  private async persistGeneratedImage(ownerId: string, projectId: string, pageId: string, versionId: string, bytes: Buffer, mimeType: string): Promise<string> {
    if (bytes.length === 0 || bytes.length > 20 * 1024 * 1024) throw new Error('Relay image output exceeded the 20 MB artifact limit.');
    const extension = mimeType.includes('jpeg') || mimeType.includes('jpg') ? 'jpg' : mimeType.includes('webp') ? 'webp' : 'png';
    const directory = path.resolve(process.env.DATA_DIR || './data', 'artifacts', ownerId, projectId, pageId);
    await fs.mkdir(directory, { recursive: true });
    const outputPath = path.resolve(directory, `${versionId}.${extension}`);
    await fs.writeFile(outputPath, bytes, { flag: 'wx' });
    const stored = await this.objectStorage.putFile(outputPath, `projects/${ownerId}/${projectId}/pages/${pageId}/${versionId}.${extension}`, mimeType);
    return stored.objectKey;
  }

  listProjects(ownerId: string, includeArchived = false): Project[] {
    return this.store.state.projects
      .filter((project) => project.ownerId === ownerId && (includeArchived || project.status !== 'archived'))
      .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
  }

  getProject(ownerId: string, projectId: string): Project {
    const project = this.store.state.projects.find((candidate) => candidate.projectId === projectId);
    if (!project || project.ownerId !== ownerId) throw new HttpError(404, 'Project not found.');
    return project;
  }

  async createProject(ownerId: string, input: CreateProjectInput): Promise<Project> {
    const projectId = makeId('p');
    const goalId = makeId('goal');
    const sourceMarkdown = input.slidesMarkdown?.trim() || DEFAULT_MARKDOWN;
    const parsed = parseSlidesMarkdown(sourceMarkdown);
    const project: Project = {
      projectId,
      ownerId,
      name: input.name,
      themeId: input.themeId,
      themeVersion: input.themeVersion,
      currentDeckRevisionId: makeId('deckrev'),
      goalId,
      status: 'ready',
      createdAt: nowIso(),
      updatedAt: nowIso(),
      sourceMarkdown,
      pages: [],
    };
    project.pages = parsed.map((slide, index) => {
      const pageId = makeId('page');
      const version = initialVersion(projectId, pageId, index + 1, project.name, slide);
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
        status: 'untouched',
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
          artifactId: version.previewArtifactId!, projectId, pageId: page.pageId, versionId: version.versionId,
          kind: 'svg_preview', promptSnapshotId: version.promptSnapshotId, model: 'deterministic-local-preview',
          width: 960, height: 540, quality: 'quick', source: 'deterministic_svg',
          provenance: { stage: 'initial_import', sourceRevision: version.sourceRevision }, createdAt: nowIso(),
        });
      });
      state.auditLogs.push({ auditId: makeId('audit'), ownerId, projectId, action: 'project.created', payload: { pageCount: project.pages.length }, createdAt: nowIso() });
    });
    await this.events.publish({ type: 'project.created', projectId, payload: { name: project.name, pageCount: project.pages.length } });
    return project;
  }

  async duplicateProject(ownerId: string, projectId: string): Promise<Project> {
    const original = this.getProject(ownerId, projectId);
    return this.createProject(ownerId, {
      name: `${original.name} Copy`,
      themeId: original.themeId,
      themeVersion: original.themeVersion,
      slidesMarkdown: original.pages.map((page) => `# ${page.title}\n\n${page.body}`).join('\n\n---\n\n'),
    });
  }

  async importSlides(ownerId: string, projectId: string, sourceMarkdown: string): Promise<Project> {
    const project = this.getProject(ownerId, projectId);
    const parsed = parseSlidesMarkdown(sourceMarkdown);
    const operationId = makeId('op');
    const orderedExisting = [...project.pages].filter((page) => !page.archived).sort((a, b) => a.orderIndex - b.orderIndex);
    const resolvedPageIds = parsed.map((slide, index) => orderedExisting[index]?.pageId || makeId('page'));
    const operation: EditOperation = {
      operationId,
      projectId,
      conversationId: makeId('conv'),
      mode: 'multi',
      requestedPageIds: resolvedPageIds,
      resolvedPageIds,
      message: 'Imported slides.md source',
      structuredPlan: {
        intent: 'source_import',
        affectedPageIds: resolvedPageIds,
        changes: parsed.flatMap((slide) => slide.facts.map((fact) => ({ kind: 'preserve_fact' as const, factId: fact.factId, target: 'fact_anchor', value: fact.value }))),
        factImpact: { added: [], removed: [], changed: [] },
        unsupported: [],
        requiresConfirmation: false,
        estimatedCost: { imageUnits: 0, amount: 0, currency: 'USD' },
        summary: `已导入 ${parsed.length} 页 slides.md，并保留位置匹配的 page_id。`,
      },
      factImpact: { added: [], removed: [], changed: [] },
      unsupportedItems: [],
      confirmationRequired: false,
      confirmedAt: nowIso(),
      resultVersionIds: [],
      status: 'completed',
      createdAt: nowIso(),
      completedAt: nowIso(),
      failedPageIds: [],
      estimatedCost: 0,
    };
    await this.store.update((state) => {
      state.operations.push(operation);
      orderedExisting.slice(parsed.length).forEach((page) => { page.archived = true; });
      parsed.forEach((slide, index) => {
        const pageId = resolvedPageIds[index];
        const existing = orderedExisting[index];
        const version = initialVersion(projectId, pageId, index + 1, project.name, slide);
        version.parentVersionId = existing?.currentVersionId || null;
        version.editOperationId = operationId;
        version.message = 'Imported slides.md';
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
          status: 'untouched',
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
        page.status = 'untouched';
        page.versions.push(version);
        state.artifacts.push({
          artifactId: version.previewArtifactId!, projectId, pageId, versionId: version.versionId,
          kind: 'svg_preview', promptSnapshotId: version.promptSnapshotId, model: 'deterministic-local-preview',
          width: 960, height: 540, quality: 'quick', source: 'deterministic_svg',
          provenance: { stage: 'source_import', sourceRevision: version.sourceRevision }, createdAt: nowIso(),
        });
        if (!existing) project.pages.push(page);
        operation.resultVersionIds.push(version.versionId);
      });
      project.sourceMarkdown = sourceMarkdown;
      project.currentDeckRevisionId = makeId('deckrev');
      project.updatedAt = nowIso();
    });
    await this.events.publish({ type: 'project.imported', projectId, operationId, payload: { pageIds: resolvedPageIds, pageCount: parsed.length } });
    return project;
  }

  async setArchived(ownerId: string, projectId: string, archived: boolean): Promise<Project> {
    const project = this.getProject(ownerId, projectId);
    await this.store.update(() => {
      project.status = archived ? 'archived' : 'ready';
      project.updatedAt = nowIso();
    });
    await this.events.publish({ type: archived ? 'project.archived' : 'project.restored', projectId, payload: {} });
    return project;
  }

  listPages(ownerId: string, projectId: string): Page[] {
    return [...this.getProject(ownerId, projectId).pages].filter((page) => !page.archived).sort((a, b) => a.orderIndex - b.orderIndex);
  }

  getPage(ownerId: string, projectId: string, pageId: string): Page {
    const page = this.getProject(ownerId, projectId).pages.find((candidate) => candidate.pageId === pageId);
    if (!page) throw new HttpError(404, 'Page not found.');
    return page;
  }

  async reorderPages(ownerId: string, projectId: string, pageIds: string[]): Promise<Page[]> {
    const project = this.getProject(ownerId, projectId);
    const visible = project.pages.filter((page) => !page.archived);
    const expected = new Set(visible.map((page) => page.pageId));
    if (pageIds.length !== visible.length || new Set(pageIds).size !== pageIds.length || pageIds.some((pageId) => !expected.has(pageId))) {
      throw new HttpError(400, 'Reorder must include every visible page exactly once.');
    }
    await this.store.update((state) => {
      pageIds.forEach((pageId, index) => {
        const page = project.pages.find((candidate) => candidate.pageId === pageId)!;
        page.orderIndex = index;
      });
      project.updatedAt = nowIso();
      state.auditLogs.push({ auditId: makeId('audit'), ownerId, projectId, action: 'page.reordered', payload: { pageIds }, createdAt: nowIso() });
    });
    await this.events.publish({ type: 'page.reordered', projectId, payload: { pageIds } });
    return this.listPages(ownerId, projectId);
  }

  async archivePage(ownerId: string, projectId: string, pageId: string, archived: boolean): Promise<Page> {
    const project = this.getProject(ownerId, projectId);
    const page = this.getPage(ownerId, projectId, pageId);
    const visibleCount = project.pages.filter((candidate) => !candidate.archived).length;
    if (archived && !page.archived && visibleCount <= 1) throw new HttpError(400, 'A project must keep at least one visible page.');
    await this.store.update((state) => {
      page.archived = archived;
      project.updatedAt = nowIso();
      state.auditLogs.push({ auditId: makeId('audit'), ownerId, projectId, action: archived ? 'page.archived' : 'page.restored', payload: { pageId }, createdAt: nowIso() });
    });
    await this.events.publish({ type: archived ? 'page.archived' : 'page.restored', projectId, pageId, payload: { pageId } });
    return page;
  }

  async splitPage(ownerId: string, projectId: string, pageId: string): Promise<{ operation: EditOperation }> {
    const project = this.getProject(ownerId, projectId);
    const sourcePage = this.getPage(ownerId, projectId, pageId);
    if (sourcePage.archived) throw new HttpError(409, 'Archived pages must be restored before splitting.');
    if (sourcePage.body.trim().length < 20) throw new HttpError(400, 'Page body is too short to split safely.');
    const operation: EditOperation = {
      operationId: makeId('op'), projectId, conversationId: makeId('conv'), mode: 'single', requestedPageIds: [pageId],
      resolvedPageIds: [pageId, makeId('page')], message: 'Split page',
      structuredPlan: {
        intent: 'page_split', affectedPageIds: [pageId],
        changes: [{ kind: 'layout_change', target: 'page_count', value: 'split_into_two_pages' }],
        factImpact: { added: [], removed: [], changed: [] }, unsupported: [], requiresConfirmation: true,
        estimatedCost: { imageUnits: 0, amount: 0, currency: 'USD' }, summary: 'Source page split into two stable page IDs.',
      },
      factImpact: { added: [], removed: [], changed: [] }, unsupportedItems: [], confirmationRequired: true,
      confirmedAt: null, resultVersionIds: [], status: 'planned', createdAt: nowIso(), completedAt: null,
      failedPageIds: [], estimatedCost: 0,
    };
    operation.structuredPlan.affectedPageIds = [...operation.resolvedPageIds];
    await this.store.update((state) => {
      state.operations.push(operation);
      state.auditLogs.push({ auditId: makeId('audit'), ownerId, projectId, action: 'page.split.planned', payload: { pageId, operationId: operation.operationId }, createdAt: nowIso() });
    });
    await this.events.publish({ type: 'edit.confirmation.required', projectId, operationId: operation.operationId, pageId, payload: { pageId, plan: operation.structuredPlan } });
    return { operation };
  }

  private async applySplitPage(ownerId: string, projectId: string, pageId: string, operationIdOverride?: string): Promise<{ sourcePage: Page; newPage: Page; operation: EditOperation }> {
    const project = this.getProject(ownerId, projectId);
    const sourcePage = this.getPage(ownerId, projectId, pageId);
    if (sourcePage.archived) throw new HttpError(409, 'Archived pages must be restored before splitting.');
    const paragraphs = sourcePage.body.split(/\r?\n+/).map((value) => value.trim()).filter(Boolean);
    let firstBody: string;
    let secondBody: string;
    if (paragraphs.length > 1) {
      const midpoint = Math.ceil(paragraphs.length / 2);
      firstBody = paragraphs.slice(0, midpoint).join('\n');
      secondBody = paragraphs.slice(midpoint).join('\n');
    } else {
      const body = sourcePage.body.trim();
      if (body.length < 20) throw new HttpError(400, 'Page body is too short to split safely.');
      const midpoint = Math.ceil(body.length / 2);
      firstBody = body.slice(0, midpoint).trim();
      secondBody = body.slice(midpoint).trim();
    }
    const plannedOperation = operationIdOverride ? this.getOperation(ownerId, projectId, operationIdOverride) : null;
    if (plannedOperation && plannedOperation.status !== 'confirmed') throw new HttpError(409, 'Split operation must be confirmed before execution.');
    const operationId = plannedOperation?.operationId || makeId('op');
    const sourceParent = sourcePage.currentVersionId;
    const sourceVersion: PageVersion = {
      ...currentVersion(sourcePage),
      versionId: makeId('ver'),
      parentVersionId: sourceParent,
      sourceRevision: makeId('deckrev'),
      slidesSourceHash: sourceHash(`# ${sourcePage.title}\n\n${firstBody}`),
      previewArtifactId: makeId('artifact'),
      svgArtifactId: makeId('svg'),
      pptxPageRenderId: null,
      promptSnapshotId: makeId('prompt'),
      editOperationId: operationId,
      qualityReportId: makeId('qa'),
      status: 'ready',
      createdAt: nowIso(),
      message: 'Split source page',
      body: firstBody,
      sourceMarkdown: `# ${sourcePage.title}\n\n${firstBody}`,
      contract: { ...sourcePage.contract, evidence: firstBody.split(/\r?\n/).filter(Boolean).slice(0, 5) },
      previewKind: 'svg_fallback',
      previewSvg: makePreviewSvg(sourcePage.orderIndex + 1, sourcePage.title, firstBody, sourcePage.layout, project.name),
      renderNote: 'Split page generated as SVG fallback; authoritative PowerPoint rendering is pending.',
      qaWarnings: ['Awaiting authoritative PowerPoint rendering after split.'],
      factAnchors: sourcePage.factAnchors.filter((fact) => firstBody.includes(fact.value) || !secondBody.includes(fact.value)).map((fact) => ({ ...fact })),
    };
    const continuationTitle = `${sourcePage.title}（续）`;
    const parsedContinuation = parseSlidesMarkdown(`# ${continuationTitle}\n\n${secondBody}`)[0];
    const newPageId = plannedOperation?.resolvedPageIds[1] || makeId('page');
    const continuationVersion = initialVersion(projectId, newPageId, sourcePage.orderIndex + 2, project.name, parsedContinuation);
    continuationVersion.editOperationId = operationId;
    continuationVersion.message = 'Created by page split';
    const continuationPage: Page = {
      pageId: newPageId,
      projectId,
      currentVersionId: continuationVersion.versionId,
      orderIndex: sourcePage.orderIndex + 1,
      pageType: sourcePage.pageType,
      locked: false,
      archived: false,
      factAnchorIds: [],
      factAnchors: sourcePage.factAnchors.filter((fact) => secondBody.includes(fact.value)).map((fact) => ({ ...fact, factId: makeId('fact'), source: `${fact.source}:split` })),
      editableLevel: continuationVersion.editableLevel,
      status: 'svg_fallback',
      title: continuationTitle,
      body: secondBody,
      layout: sourcePage.layout,
      sourceMarkdown: `# ${continuationTitle}\n\n${secondBody}`,
      contract: parsedContinuation.contract,
      versions: [continuationVersion],
    };
    continuationPage.factAnchorIds = continuationPage.factAnchors.map((fact) => fact.factId);
    continuationVersion.factAnchors = continuationPage.factAnchors.map((fact) => ({ ...fact }));
    const operation: EditOperation = plannedOperation || {
      operationId, projectId, conversationId: makeId('conv'), mode: 'single', requestedPageIds: [pageId],
      resolvedPageIds: [pageId, newPageId], message: 'Split page',
      structuredPlan: {
        intent: 'page_split', affectedPageIds: [pageId, newPageId],
        changes: [{ kind: 'layout_change', target: 'page_count', value: 'split_into_two_pages' }],
        factImpact: { added: [], removed: [], changed: [] }, unsupported: [], requiresConfirmation: true,
        estimatedCost: { imageUnits: 0, amount: 0, currency: 'USD' }, summary: 'Source page split into two stable page IDs.',
      },
      factImpact: { added: [], removed: [], changed: [] }, unsupportedItems: [], confirmationRequired: true,
      confirmedAt: nowIso(), resultVersionIds: [sourceVersion.versionId, continuationVersion.versionId], status: 'completed',
      createdAt: nowIso(), completedAt: nowIso(), failedPageIds: [], estimatedCost: 0,
    };
    await this.store.update((state) => {
      project.pages.filter((page) => !page.archived && page.orderIndex > sourcePage.orderIndex).forEach((page) => { page.orderIndex += 1; });
      sourcePage.body = firstBody;
      sourcePage.sourceMarkdown = `# ${sourcePage.title}\n\n${firstBody}`;
      sourcePage.contract = { ...sourcePage.contract, evidence: firstBody.split(/\r?\n/).filter(Boolean).slice(0, 5) };
      sourcePage.currentVersionId = sourceVersion.versionId;
      sourcePage.status = 'svg_fallback';
      sourcePage.versions.push(sourceVersion);
      sourcePage.factAnchors = sourcePage.factAnchors.filter((fact) => firstBody.includes(fact.value) || !secondBody.includes(fact.value));
      sourcePage.factAnchorIds = sourcePage.factAnchors.map((fact) => fact.factId);
      project.pages.push(continuationPage);
      project.currentDeckRevisionId = sourceVersion.sourceRevision;
      project.updatedAt = nowIso();
      operation.resultVersionIds = [sourceVersion.versionId, continuationVersion.versionId];
      operation.confirmedAt = operation.confirmedAt || nowIso();
      operation.status = 'completed';
      operation.completedAt = nowIso();
      if (!plannedOperation) state.operations.push(operation);
      [sourceVersion, continuationVersion].forEach((version, index) => {
        const artifactPageId = index === 0 ? pageId : newPageId;
        state.artifacts.push({
          artifactId: version.previewArtifactId!, projectId, pageId: artifactPageId, versionId: version.versionId,
          kind: 'svg_preview', promptSnapshotId: version.promptSnapshotId, model: 'deterministic-local-preview',
          width: 960, height: 540, quality: 'quick', source: 'deterministic_svg',
          provenance: { stage: 'page_split', operationId }, createdAt: nowIso(),
        });
      });
      state.auditLogs.push({ auditId: makeId('audit'), ownerId, projectId, action: 'page.split', payload: { sourcePageId: pageId, newPageId, operationId }, createdAt: nowIso() });
    });
    await this.events.publish({ type: 'page.split', projectId, pageId, operationId, payload: { sourcePageId: pageId, newPageId, versionIds: operation.resultVersionIds } });
    return { sourcePage, newPage: continuationPage, operation };
  }

  listMessages(ownerId: string, projectId: string, conversationId?: string): ConversationMessage[] {
    this.getProject(ownerId, projectId);
    return this.store.state.messages.filter((message) => message.projectId === projectId && (!conversationId || message.conversationId === conversationId));
  }

  async createChatTurn(ownerId: string, input: ChatTurnInput): Promise<{ operation: EditOperation; messages: ConversationMessage[] }> {
    const project = this.getProject(ownerId, input.projectId);
    if (project.currentDeckRevisionId !== input.deckRevisionId) throw new HttpError(409, 'Deck revision changed. Refresh project state before retrying.');
    let pageIds = [...new Set(input.target.pageIds)];
    let candidateReasons: Record<string, string> | undefined;
    if (input.target.mode === 'global') {
      const resolved = resolveSimilarPages(project, input.message);
      pageIds = resolved.pageIds;
      candidateReasons = resolved.reasons;
    }
    if (!pageIds.length) throw new HttpError(400, 'At least one target page is required.');
    const pages = pageIds.map((pageId) => this.getPage(ownerId, project.projectId, pageId));
    const plan = buildEditPlan(project, pages, input.target.mode, input.message);
    plan.candidateReasons = candidateReasons;
    const relayStatus = this.relay.status();
    if (relayStatus.configured) {
      try {
        const relay = await this.relay.structuredJson<{ summary?: string; unsupported?: string[] }>({
          system: 'Return JSON only. You may provide a short summary and unsupported limitations. Never change page IDs, facts, paths, or execute tools.',
          user: JSON.stringify({ message: input.message, pageIds, deterministicPlan: plan }),
        });
        if (relay.value.summary?.trim()) plan.summary = relay.value.summary.trim();
        if (Array.isArray(relay.value.unsupported)) plan.unsupported.push(...relay.value.unsupported.filter((item) => typeof item === 'string').slice(0, 5));
        plan.requiresConfirmation = plan.requiresConfirmation || plan.unsupported.length > 0;
      } catch (error) {
        plan.unsupported.push(`Relay planner unavailable: ${(error as Error).message}`);
        plan.requiresConfirmation = true;
      }
    }
    const validatedPlan = editPlanSchema.safeParse(plan);
    if (!validatedPlan.success) throw new HttpError(422, `Structured edit plan validation failed: ${validatedPlan.error.issues.map((issue) => issue.path.join('.') + ' ' + issue.message).join('; ')}`);
    Object.assign(plan, validatedPlan.data);
    const planErrors = validateEditPlan(project, pages, input.target.mode, plan);
    if (planErrors.length) throw new HttpError(422, `Structured edit plan rejected: ${planErrors.join(' ')}`);
    const conversationId = input.conversationId || makeId('conv');
    const operation: EditOperation = {
      operationId: makeId('op'),
      projectId: project.projectId,
      conversationId,
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
      status: 'planned',
      createdAt: nowIso(),
      completedAt: null,
      failedPageIds: [],
      estimatedCost: plan.estimatedCost.amount,
    };
    const userMessage: ConversationMessage = {
      messageId: makeId('msg'), projectId: project.projectId, conversationId, role: 'user', text: input.message, createdAt: nowIso(), operationId: operation.operationId,
    };
    const planMessage: ConversationMessage = {
      messageId: makeId('msg'), projectId: project.projectId, conversationId, role: 'assistant', text: plan.summary, createdAt: nowIso(), operationId: operation.operationId,
      meta: { plan, confirmationRequired: plan.requiresConfirmation },
    };
    await this.store.update((state) => {
      state.operations.push(operation);
      state.messages.push(userMessage, planMessage);
      state.auditLogs.push({ auditId: makeId('audit'), ownerId, projectId: project.projectId, action: 'chat.plan.created', payload: { operationId: operation.operationId, mode: operation.mode, pageIds }, createdAt: nowIso() });
    });
    await this.events.publish({ type: 'chat.plan.created', projectId: project.projectId, operationId: operation.operationId, payload: { plan } });
    if (operation.confirmationRequired) {
      await this.events.publish({ type: 'edit.confirmation.required', projectId: project.projectId, operationId: operation.operationId, payload: { pageIds, plan } });
      return { operation, messages: [userMessage, planMessage] };
    }
    await this.enqueueEdit(ownerId, operation.operationId);
    return { operation, messages: [userMessage, planMessage] };
  }

  getOperation(ownerId: string, projectId: string, operationId: string): EditOperation {
    this.getProject(ownerId, projectId);
    const operation = this.store.state.operations.find((candidate) => candidate.operationId === operationId && candidate.projectId === projectId);
    if (!operation) throw new HttpError(404, 'Edit operation not found.');
    return operation;
  }

  async confirmOperation(ownerId: string, projectId: string, operationId: string): Promise<EditOperation> {
    const operation = this.getOperation(ownerId, projectId, operationId);
    if (operation.status !== 'planned') throw new HttpError(409, 'Only planned operations can be confirmed.');
    await this.store.update(() => {
      operation.status = 'confirmed';
      operation.confirmedAt = nowIso();
    });
    if (operation.structuredPlan.intent === 'page_split') {
      const sourcePageId = operation.requestedPageIds[0];
      const result = await this.applySplitPage(ownerId, projectId, sourcePageId, operationId);
      return result.operation;
    }
    return this.enqueueEdit(ownerId, operationId);
  }

  private async enqueueEdit(ownerId: string, operationId: string): Promise<EditOperation> {
    const operation = this.getOperation(ownerId, this.store.state.operations.find((candidate) => candidate.operationId === operationId)?.projectId || '', operationId);
    const project = this.getProject(ownerId, operation.projectId);
    await this.jobs.enqueue({
      jobId: operationId,
      projectId: project.projectId,
      operationId,
      kind: 'edit',
      status: 'queued',
      payload: { operationId },
      createdAt: nowIso(),
      completedAt: null,
    }, () => this.applyOperation(ownerId, operationId).then(() => undefined), true);
    return operation;
  }

  async cancelOperation(ownerId: string, projectId: string, operationId: string): Promise<EditOperation> {
    const operation = this.getOperation(ownerId, projectId, operationId);
    if (!['planned', 'confirmed', 'applying'].includes(operation.status)) throw new HttpError(409, 'Operation cannot be cancelled in its current state.');
    await this.store.update(() => {
      operation.status = 'rolled_back';
      operation.completedAt = nowIso();
    });
    await this.events.publish({ type: 'edit.rolled_back', projectId, operationId, payload: { cancelled: true } });
    return operation;
  }

  async retryFailedPages(ownerId: string, projectId: string, operationId: string): Promise<EditOperation> {
    const project = this.getProject(ownerId, projectId);
    const source = this.getOperation(ownerId, projectId, operationId);
    if (!['completed', 'failed'].includes(source.status)) throw new HttpError(409, 'Only completed operations with failed pages can be retried.');
    const pageIds = [...new Set(source.failedPageIds)];
    if (!pageIds.length) throw new HttpError(409, 'This operation has no failed pages to retry.');
    const pages = pageIds.map((pageId) => this.getPage(ownerId, projectId, pageId));
    const plan = buildEditPlan(project, pages, source.mode, source.message);
    if (source.structuredPlan.candidateReasons) {
      plan.candidateReasons = Object.fromEntries(pageIds.flatMap((pageId) => {
        const reason = source.structuredPlan.candidateReasons?.[pageId];
        return reason ? [[pageId, reason]] : [];
      }));
    }
    const now = nowIso();
    const retry: EditOperation = {
      operationId: makeId('op'),
      projectId,
      conversationId: source.conversationId,
      mode: source.mode,
      requestedPageIds: pageIds,
      resolvedPageIds: pageIds,
      message: source.message,
      structuredPlan: { ...plan, summary: `仅重试上次失败的 ${pageIds.length} 个页面。` },
      factImpact: plan.factImpact,
      unsupportedItems: plan.unsupported,
      confirmationRequired: plan.requiresConfirmation,
      confirmedAt: now,
      resultVersionIds: [],
      status: 'confirmed',
      createdAt: now,
      completedAt: null,
      failedPageIds: [],
      estimatedCost: plan.estimatedCost.amount,
      parentOperationId: source.operationId,
    };
    await this.store.update((state) => {
      state.operations.push(retry);
      state.messages.filter((message) => message.operationId === source.operationId && Array.isArray(message.meta?.failedPageIds)).forEach((message) => {
        message.meta = { ...message.meta, retryOperationId: retry.operationId };
      });
      state.auditLogs.push({ auditId: makeId('audit'), ownerId, projectId, action: 'operation.retry_failed_pages', payload: { sourceOperationId: source.operationId, retryOperationId: retry.operationId, pageIds }, createdAt: now });
    });
    await this.events.publish({ type: 'edit.retry.started', projectId, operationId: retry.operationId, payload: { sourceOperationId: source.operationId, pageIds } });
    return this.enqueueEdit(ownerId, retry.operationId);
  }

  private async applyOperation(ownerId: string, operationId: string): Promise<EditOperation> {
    const operation = this.store.state.operations.find((candidate) => candidate.operationId === operationId);
    if (!operation) throw new HttpError(404, 'Edit operation not found.');
    const project = this.getProject(ownerId, operation.projectId);
    await this.store.update((state) => {
      operation.status = 'applying';
      if (!state.jobs.some((job) => job.jobId === operationId)) state.jobs.push({ jobId: operationId, projectId: project.projectId, operationId, kind: 'edit', status: 'running', payload: { operationId }, createdAt: nowIso(), completedAt: null });
    });
    await this.events.publish({ type: 'edit.started', projectId: project.projectId, operationId, payload: { pageIds: operation.resolvedPageIds } });
    for (let index = 0; index < operation.resolvedPageIds.length; index += 1) {
      if (operation.status === 'rolled_back') break;
      const pageId = operation.resolvedPageIds[index];
      const page = this.getPage(ownerId, project.projectId, pageId);
      const priorVersion = currentVersion(page);
      if (page.versions.some((version) => operation.resultVersionIds.includes(version.versionId))) continue;
      operation.failedPageIds = operation.failedPageIds.filter((failedPageId) => failedPageId !== pageId);
      const imageChange = operation.structuredPlan.changes.some((change) => change.kind === 'image_replace');
      const relayStatus = this.relay.status();
      const reserve: UsageLedgerEntry = {
        ledgerId: makeId('ledger'), projectId: project.projectId, operationId, pageId, versionId: null,
        model: imageChange ? relayStatus.imageModel : relayStatus.model,
        unitPrice: imageChange ? relayStatus.priceSnapshot.imageUnit : relayStatus.priceSnapshot.tokenUnit,
        reservedAmount: imageChange ? relayStatus.priceSnapshot.imageUnit : 0,
        settledAmount: 0, refundedAmount: 0, status: 'reserved', createdAt: nowIso(),
      };
      try {
        await this.store.update((state) => {
          state.ledger.push(reserve);
          page.status = 'generating';
        });
        await this.events.publish({
          type: 'edit.progress', projectId: project.projectId, pageId, operationId,
          payload: { completed: index, total: operation.resolvedPageIds.length, stage: 'plan_validated' },
        });
        const prompt = composePrompt(project, page, operation.structuredPlan, operation.message);
        const edited = previewTextChange(page, operation.message);
        const layoutChange = operation.structuredPlan.changes.find((change) => change.kind === 'layout_change')?.value;
        const nextLayout = layoutChange || page.layout;
        let imageAssetPath: string | null = null;
        let imageRequestId: string | null = null;
        let imageModel: string | null = null;
        if (imageChange) {
          if (!relayStatus.configured) throw new Error('视觉修改需要已配置的 Relay 图像模型；未生成占位图片。');
          const generated = await this.relay.generateImage({ prompt: `${prompt.prompt}\nGenerate a presentation visual for this page. Preserve every locked fact and do not render text inside the image.`, width: 1600, height: 900, quality: 'high' });
          imageRequestId = generated.requestId;
          imageModel = generated.model;
          imageAssetPath = await this.persistGeneratedImage(ownerId, project.projectId, pageId, makeId('asset'), generated.bytes, generated.mimeType);
        }
        const removedFacts = operation.factImpact.removed.map((value) => value.includes(':') ? value.slice(value.indexOf(':') + 1) : value);
        const changedFacts = factChanges(operation.message);
        if ((removedFacts.length > 0 || operation.factImpact.changed.length > 0) && !operation.confirmedAt) throw new Error('事实变更必须经过确认后才能执行。');
        const contentQa = runContentQa(project, page, edited, operation.structuredPlan, Boolean(operation.confirmedAt));
        if (!contentQa.passed) throw new Error(contentQa.errors.join(' '));
        const nextFactAnchors = applyFactChanges(page.factAnchors, operation.message, Boolean(operation.confirmedAt)).filter((fact) => !removedFacts.includes(fact.value) && !changedFacts.some((change) => change.oldValue === fact.value));
        const quickPreview = await this.quickPreview.render(project, page, edited.title, edited.body, nextLayout);
        let authoritativeRenderId: string | null = null;
        let authoritativeNote = '';
        let authoritativeArtifact: PreviewArtifact | null = null;
        const renderer = process.env.POWERPOINT_RENDERER === 'powerpoint' || process.env.POWERPOINT_RENDERER === 'com';
        const previewArtifactId = makeId('artifact');
        const visualArtifactId = imageChange ? previewArtifactId : priorVersion.visualArtifactId || null;
        const nonEditableRegions = visualArtifactId
          ? [...new Set([...priorVersion.nonEditableRegions, 'visual_anchor'])]
          : [...priorVersion.nonEditableRegions];
        const version: PageVersion = {
          versionId: makeId('ver'),
          pageId,
          parentVersionId: page.currentVersionId,
          sourceRevision: makeId('deckrev'),
          pageContractPath: `contracts/${pageId}.json`,
          slidesSourceHash: sourceHash(`# ${edited.title}\n\n${edited.body}`),
          previewArtifactId,
          visualArtifactId,
          svgArtifactId: makeId('svg'),
          pptxPageRenderId: null,
          promptSnapshotId: `prompt_${prompt.hash.slice(0, 20)}`,
          editOperationId: operationId,
          qualityReportId: makeId('qa'),
          status: 'previewing',
          createdAt: nowIso(),
          message: operation.message,
           title: edited.title,
           body: edited.body,
           layout: nextLayout,
           sourceMarkdown: `# ${edited.title}\n\n${edited.body}`,
            contract: { ...page.contract, conclusion: edited.title, evidence: edited.body.split(/\r?\n/).filter(Boolean).slice(0, 5), mustKeep: nextFactAnchors.map((fact) => fact.value) },
           previewKind: 'quick_preview',
          previewSvg: quickPreview.svg,
          editableLevel: visualArtifactId ? 'native_partial' : 'native_structure',
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
           page.sourceMarkdown = version.sourceMarkdown || `# ${version.title}\n\n${version.body}`;
           if (version.contract) page.contract = { ...version.contract, evidence: [...version.contract.evidence], mustKeep: [...version.contract.mustKeep], canTrim: [...version.contract.canTrim], components: [...version.contract.components] };
          page.editableLevel = version.editableLevel;
          if (removedFacts.length > 0 || operation.factImpact.changed.length > 0) {
            page.factAnchors = nextFactAnchors;
            page.factAnchorIds = page.factAnchors.map((fact) => fact.factId);
          }
          page.status = 'quick_preview';
          operation.resultVersionIds.push(version.versionId);
          reserve.versionId = version.versionId;
          const artifact: PreviewArtifact = {
            artifactId: version.previewArtifactId!, projectId: project.projectId, pageId, versionId: version.versionId,
            kind: imageChange ? 'visual_preview' : 'svg_preview',
            promptSnapshotId: version.promptSnapshotId, model: imageModel || relayStatus.model,
            width: imageChange ? 1600 : 960, height: imageChange ? 900 : 540, quality: imageChange ? 'high' : 'quick', source: imageChange ? 'relay_image' : 'deterministic_svg',
            provenance: { promptHash: prompt.hash, plan: operation.structuredPlan, engine: quickPreview.engine, relayRequestId: imageRequestId }, createdAt: nowIso(), assetPath: imageAssetPath,
          };
          state.artifacts.push(artifact);
          const snapshot: PromptSnapshot = {
            promptSnapshotId: version.promptSnapshotId, projectId: project.projectId, pageId, operationId,
            prompt: prompt.prompt, hash: prompt.hash, model: process.env.RELAY_MODEL || 'deterministic-local-planner', createdAt: nowIso(),
          };
          state.promptSnapshots.push(snapshot);
        });
        await this.events.publish({ type: 'page.version.created', projectId: project.projectId, pageId, versionId: version.versionId, operationId, payload: { parentVersionId: version.parentVersionId } });
        await this.events.publish({ type: 'preview.quick.ready', projectId: project.projectId, pageId, versionId: version.versionId, operationId, payload: { previewKind: 'quick_preview' } });
        if (renderer) {
          try {
            const pageRenderDir = path.resolve(process.env.DATA_DIR || './data', 'page-renders', ownerId, project.projectId, pageId, version.versionId);
            const renderProject = { ...project, pages: [page] };
            const exportResult = await runPptxExport(
              renderProject,
              pageRenderDir,
              `${pageId}-${version.versionId}.pptx`,
              await this.exportVisualAssets(renderProject),
            );
            const renderResult = await runPowerPointRender(exportResult.outputPath, path.join(pageRenderDir, 'powerpoint'));
            if (renderResult.renders.length !== 1) throw new Error(`Expected one authoritative page render, received ${renderResult.renders.length}.`);
            const render = renderResult.renders[0];
            const renderBytes = await fs.readFile(render.path);
            if (renderBytes.length < 1_000 || render.width <= 0 || render.height <= 0) throw new Error('PowerPoint returned an empty or invalid PNG render.');
            authoritativeRenderId = makeId('artifact');
            const stored = await this.objectStorage.putFile(
              render.path,
              `projects/${ownerId}/${project.projectId}/pages/${pageId}/versions/${version.versionId}/${authoritativeRenderId}.png`,
              'image/png',
            );
            authoritativeArtifact = {
              artifactId: authoritativeRenderId,
              projectId: project.projectId,
              pageId,
              versionId: version.versionId,
              kind: 'pptx_page_render',
              promptSnapshotId: version.promptSnapshotId,
              model: renderResult.renderer,
              width: render.width,
              height: render.height,
              quality: 'authoritative',
              source: 'powerpoint_com',
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
          if (authoritativeArtifact) state.artifacts.push(authoritativeArtifact);
          version.status = 'ready';
          version.previewKind = authoritative ? 'pptx_authoritative' : 'svg_fallback';
          version.pptxPageRenderId = authoritativeRenderId;
          version.renderNote = authoritative
            ? authoritativeNote
            : `PowerPoint worker is unavailable. This is an SVG fallback and may differ from PPTX.${authoritativeNote ? ` ${authoritativeNote}` : ''}`;
          if (!authoritative) version.qaWarnings.push(authoritativeNote || 'Awaiting authoritative PowerPoint rendering.');
          page.status = authoritative ? 'authoritative' : 'svg_fallback';
          reserve.status = 'settled';
          reserve.settledAmount = reserve.reservedAmount;
          project.currentDeckRevisionId = version.sourceRevision;
          project.updatedAt = nowIso();
        });
        if (authoritative) {
          await this.events.publish({ type: 'preview.pptx.ready', projectId: project.projectId, pageId, versionId: version.versionId, operationId, payload: { renderId: version.pptxPageRenderId } });
        } else {
          await this.events.publish({ type: 'render.warning', projectId: project.projectId, pageId, versionId: version.versionId, operationId, payload: { reason: 'powerpoint_worker_unavailable', previewKind: 'svg_fallback' } });
        }
      } catch (error) {
        await this.store.update(() => {
          page.status = 'failed';
          operation.failedPageIds.push(pageId);
          reserve.status = 'refunded';
          reserve.refundedAmount = reserve.reservedAmount;
        });
        await this.events.publish({ type: 'edit.failed', projectId: project.projectId, pageId, operationId, payload: { message: (error as Error).message } });
      }
    }
    await this.store.update(() => {
      operation.status = operation.failedPageIds.length === operation.resolvedPageIds.length ? 'failed' : 'completed';
      operation.completedAt = nowIso();
    });
    await this.store.update((state) => {
      const durable = state.jobs.find((job) => job.jobId === operationId);
      if (durable) { durable.status = operation.status === 'failed' ? 'failed' : 'completed'; durable.completedAt = operation.completedAt; }
    });
    const completionMessage: ConversationMessage = {
      messageId: makeId('msg'), projectId: project.projectId, conversationId: operation.conversationId, role: 'assistant',
      text: operation.failedPageIds.length
        ? `操作已完成，${operation.resultVersionIds.length} 页成功，${operation.failedPageIds.length} 页可单独重试。`
        : `操作已完成，生成 ${operation.resultVersionIds.length} 个不可变页面版本。`,
      createdAt: nowIso(), operationId, meta: { versionIds: operation.resultVersionIds, failedPageIds: operation.failedPageIds },
    };
    await this.store.update((state) => state.messages.push(completionMessage));
    await this.events.publish({ type: 'edit.completed', projectId: project.projectId, operationId, payload: { versionIds: operation.resultVersionIds, failedPageIds: operation.failedPageIds } });
    return operation;
  }

  listVersions(ownerId: string, projectId: string, pageId: string): PageVersion[] {
    return [...this.getPage(ownerId, projectId, pageId).versions].sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  }

  async rollbackVersion(ownerId: string, projectId: string, versionId: string): Promise<Page> {
    const project = this.getProject(ownerId, projectId);
    const page = project.pages.find((candidate) => candidate.versions.some((version) => version.versionId === versionId));
    const version = page?.versions.find((candidate) => candidate.versionId === versionId);
    if (!page || !version) throw new HttpError(404, 'Version not found.');
    await this.store.update(() => {
      page.currentVersionId = version.versionId;
       page.title = version.title;
       page.body = version.body;
       page.layout = version.layout;
       page.sourceMarkdown = version.sourceMarkdown || `# ${version.title}\n\n${version.body}`;
       if (version.contract) page.contract = { ...version.contract, evidence: [...version.contract.evidence], mustKeep: [...version.contract.mustKeep], canTrim: [...version.contract.canTrim], components: [...version.contract.components] };
      page.editableLevel = version.editableLevel;
      if (version.factAnchors) {
        page.factAnchors = version.factAnchors.map((fact) => ({ ...fact }));
        page.factAnchorIds = page.factAnchors.map((fact) => fact.factId);
      }
      page.status = 'rolled_back';
      project.currentDeckRevisionId = makeId('deckrev');
      project.updatedAt = nowIso();
    });
    await this.store.update((state) => state.auditLogs.push({ auditId: makeId('audit'), ownerId, projectId, action: 'version.rollback', payload: { pageId: page.pageId, versionId }, createdAt: nowIso() }));
    await this.events.publish({ type: 'edit.rolled_back', projectId, pageId: page.pageId, versionId, payload: { restoredVersionId: versionId } });
    return page;
  }

  async rollbackOperation(ownerId: string, projectId: string, operationId: string): Promise<EditOperation> {
    const project = this.getProject(ownerId, projectId);
    const operation = this.getOperation(ownerId, projectId, operationId);
    if (!['completed', 'failed'].includes(operation.status)) throw new HttpError(409, 'Only completed operations can be rolled back as a group.');
    await this.store.update((state) => {
      for (const page of project.pages) {
        const applied = page.versions.find((version) => operation.resultVersionIds.includes(version.versionId));
        if (!applied) continue;
        const parent = applied.parentVersionId ? page.versions.find((version) => version.versionId === applied.parentVersionId) : null;
        if (!parent) continue;
        page.currentVersionId = parent.versionId;
         page.title = parent.title;
         page.body = parent.body;
         page.layout = parent.layout;
         page.sourceMarkdown = parent.sourceMarkdown || `# ${parent.title}\n\n${parent.body}`;
         if (parent.contract) page.contract = { ...parent.contract, evidence: [...parent.contract.evidence], mustKeep: [...parent.contract.mustKeep], canTrim: [...parent.contract.canTrim], components: [...parent.contract.components] };
        page.editableLevel = parent.editableLevel;
        if (parent.factAnchors) {
          page.factAnchors = parent.factAnchors.map((fact) => ({ ...fact }));
          page.factAnchorIds = page.factAnchors.map((fact) => fact.factId);
        }
        page.status = 'rolled_back';
      }
      if (operation.structuredPlan.intent === 'page_split') {
        const continuation = project.pages.find((page) => page.pageId === operation.resolvedPageIds[1]);
        if (continuation) continuation.archived = true;
      }
      operation.status = 'rolled_back';
      operation.completedAt = nowIso();
      project.currentDeckRevisionId = makeId('deckrev');
      project.updatedAt = nowIso();
      state.messages.filter((message) => message.operationId === operationId && Array.isArray(message.meta?.versionIds)).forEach((message) => {
        message.meta = { ...message.meta, rolledBack: true };
      });
      state.auditLogs.push({ auditId: makeId('audit'), ownerId, projectId, action: 'operation.rollback', payload: { operationId }, createdAt: nowIso() });
    });
    await this.events.publish({ type: 'edit.rolled_back', projectId, operationId, payload: { operationRollback: true, versionIds: operation.resultVersionIds } });
    return operation;
  }

  async createExport(ownerId: string, projectId: string): Promise<ExportJob> {
    const project = this.getProject(ownerId, projectId);
    const exportId = makeId('export');
    const rendererConfigured = process.env.POWERPOINT_RENDERER === 'powerpoint' || process.env.POWERPOINT_RENDERER === 'com';
    const exportEngine = process.env.PPTX_EXPORT_ENGINE === 'legacy'
      ? 'legacy_python_pptx_development_fallback'
      : 'ppt_master_svg_to_drawingml';
    const job: ExportJob = {
      exportId,
      projectId,
      status: 'queued',
      artifactPath: null,
      renderMode: rendererConfigured ? 'powerpoint' : 'svg_fallback',
      qaWarnings: rendererConfigured ? [] : ['PowerPoint COM render worker is unavailable; export is explicitly marked SVG fallback.'],
      createdAt: nowIso(),
      completedAt: null,
      exportEngine,
      qaStatus: 'pending',
    };
    await this.store.update((state) => {
      state.exports.push(job);
      state.jobs.push({ jobId: exportId, projectId, kind: 'export', status: 'queued', payload: { exportId }, createdAt: nowIso(), completedAt: null });
      state.auditLogs.push({ auditId: makeId('audit'), ownerId, projectId, action: 'export.created', payload: { exportId, renderMode: job.renderMode, exportEngine }, createdAt: nowIso() });
    });
    void this.jobs.enqueue({
      jobId: exportId,
      projectId,
      kind: 'export',
      status: 'queued',
      payload: { exportId },
      createdAt: job.createdAt,
      completedAt: null,
    }, () => this.executeExport(project, job));
    return job;
  }

  private async executeExport(project: Project, job: ExportJob): Promise<void> {
    try {
      await this.store.update((state) => {
        job.status = 'running';
        const durable = state.jobs.find((candidate) => candidate.jobId === job.exportId);
        if (durable) durable.status = 'running';
      });
      const outputDir = path.resolve(process.env.DATA_DIR || './data', 'exports', project.ownerId, project.projectId);
      const exportResult = await runPptxExport(
        project,
        outputDir,
        `${safeFileName(project.name)}-${job.exportId}.pptx`,
        await this.exportVisualAssets(project),
      );
      job.artifactName = path.basename(exportResult.outputPath);
      if (!exportResult.qaPath) throw new Error('PPTX export worker did not return its QA receipt.');
      const qa = JSON.parse(await fs.readFile(exportResult.qaPath, 'utf8')) as {
        export_engine?: ExportJob['exportEngine'];
        static_structure_passed?: boolean;
        svg_quality?: { summary?: { warnings?: number } };
        pptx_postflight?: { status?: string };
        delivery_check?: { status?: string };
      };
      job.exportEngine = qa.export_engine || job.exportEngine;
      const isDevelopmentFallback = job.exportEngine === 'legacy_python_pptx_development_fallback';
      if (isDevelopmentFallback) {
        job.qaStatus = 'not-run-development-fallback';
      } else {
        if (qa.static_structure_passed !== true) throw new Error('Static PPTX structure QA failed.');
        const postflightStatus = qa.pptx_postflight?.status;
        const deliveryStatus = qa.delivery_check?.status;
        if (!['passed', 'passed-with-warnings', 'passed-with-advisories'].includes(String(postflightStatus))) {
          throw new Error(`PPTX postflight QA failed with status ${String(postflightStatus || 'missing')}.`);
        }
        if (!['passed', 'passed-with-advisories'].includes(String(deliveryStatus))) {
          throw new Error(`PPTX delivery QA failed with status ${String(deliveryStatus || 'missing')}.`);
        }
        job.qaStatus = Number(qa.svg_quality?.summary?.warnings || 0) > 0 || postflightStatus !== 'passed' || deliveryStatus !== 'passed'
          ? 'passed-with-warnings'
          : 'passed';
      }
      const storedArtifact = await this.objectStorage.putFile(exportResult.outputPath, `projects/${project.ownerId}/${project.projectId}/exports/${job.artifactName}`, 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
      job.artifactObjectKey = storedArtifact.objectKey;
      if (job.renderMode === 'powerpoint') {
        try {
          const renderResult = await runPowerPointRender(exportResult.outputPath, path.join(outputDir, 'powerpoint'));
          job.qaWarnings.push(`PowerPoint COM rendered ${renderResult.renders.length} slide(s).`);
        } catch (renderError) {
          job.renderMode = 'svg_fallback';
          job.qaWarnings.push(`PowerPoint COM render failed; SVG fallback retained: ${(renderError as Error).message}`);
        }
      }
      await this.store.update(() => {
        job.status = 'completed';
        job.artifactPath = exportResult.outputPath;
        job.completedAt = nowIso();
      });
      await this.store.update((state) => {
        const durable = state.jobs.find((candidate) => candidate.jobId === job.exportId);
        if (durable) { durable.status = 'completed'; durable.completedAt = job.completedAt; }
      });
      await this.events.publish({ type: 'export.completed', projectId: project.projectId, exportId: job.exportId, payload: { renderMode: job.renderMode, qaWarnings: job.qaWarnings } });
    } catch (error) {
      await this.store.update(() => {
        job.status = 'failed';
        job.qaStatus = 'failed';
        job.qaWarnings.push((error as Error).message);
        job.completedAt = nowIso();
      });
      await this.store.update((state) => {
        const durable = state.jobs.find((candidate) => candidate.jobId === job.exportId);
        if (durable) { durable.status = 'failed'; durable.completedAt = job.completedAt; }
      });
      await this.events.publish({ type: 'export.failed', projectId: project.projectId, exportId: job.exportId, payload: { message: (error as Error).message } });
    }
  }

  async resumeJobs(): Promise<void> {
    const pending = this.store.state.jobs.filter((job) => ['queued', 'running'].includes(job.status));
    await this.jobs.resume(pending, (durable) => {
      if (durable.kind === 'edit') {
        const operation = this.store.state.operations.find((candidate) => candidate.operationId === durable.operationId);
        const project = this.store.state.projects.find((candidate) => candidate.projectId === durable.projectId);
        if (!operation || !project || !['confirmed', 'applying'].includes(operation.status)) return null;
        return () => this.applyOperation(project.ownerId, operation.operationId).then(() => undefined);
      }
      const exportId = String(durable.payload.exportId || durable.jobId);
      const job = this.store.state.exports.find((candidate) => candidate.exportId === exportId);
      const project = this.store.state.projects.find((candidate) => candidate.projectId === durable.projectId);
      if (!job || !project) return null;
      job.status = 'queued';
      return () => this.executeExport(project, job);
    });
  }

  getExport(ownerId: string, projectId: string, exportId: string): ExportJob {
    this.getProject(ownerId, projectId);
    const job = this.store.state.exports.find((candidate) => candidate.exportId === exportId && candidate.projectId === projectId);
    if (!job) throw new HttpError(404, 'Export not found.');
    return job;
  }

  async readExportArtifact(ownerId: string, projectId: string, exportId: string): Promise<{ bytes: Buffer; fileName: string }> {
    const job = this.getExport(ownerId, projectId, exportId);
    if (job.status !== 'completed') throw new HttpError(409, 'Export is not ready.');
    if (job.artifactObjectKey) return { bytes: await this.objectStorage.getBytes(job.artifactObjectKey), fileName: job.artifactName || `${exportId}.pptx` };
    if (!job.artifactPath) throw new HttpError(409, 'Export artifact is unavailable.');
    const root = path.resolve(process.env.DATA_DIR || './data');
    const artifact = path.resolve(job.artifactPath);
    if (!artifact.startsWith(`${root}${path.sep}`)) throw new HttpError(403, 'Artifact path is outside the project store.');
    return { bytes: await fs.readFile(artifact), fileName: job.artifactName || path.basename(artifact) };
  }

  async readAuthoritativeRender(ownerId: string, projectId: string, artifactId: string): Promise<{ bytes: Buffer; width: number; height: number }> {
    const project = this.getProject(ownerId, projectId);
    const artifact = this.store.state.artifacts.find((candidate) => candidate.artifactId === artifactId && candidate.projectId === projectId);
    if (!artifact || artifact.kind !== 'pptx_page_render' || artifact.source !== 'powerpoint_com' || !artifact.assetPath) throw new HttpError(404, 'Authoritative page render not found.');
    const page = project.pages.find((candidate) => candidate.pageId === artifact.pageId);
    const version = page?.versions.find((candidate) => candidate.versionId === artifact.versionId);
    if (!version || version.pptxPageRenderId !== artifact.artifactId) throw new HttpError(404, 'Authoritative page render is not bound to this project version.');
    return { bytes: await this.objectStorage.getBytes(artifact.assetPath), width: artifact.width, height: artifact.height };
  }

  usage(ownerId: string, projectId: string): UsageLedgerEntry[] {
    this.getProject(ownerId, projectId);
    return this.store.state.ledger.filter((entry) => entry.projectId === projectId);
  }

  relayStatus() {
    return this.relay.status();
  }

  audit(ownerId: string, projectId: string) {
    this.getProject(ownerId, projectId);
    return this.store.state.auditLogs.filter((entry) => entry.ownerId === ownerId && entry.projectId === projectId);
  }

  artifacts(ownerId: string, projectId: string) {
    this.getProject(ownerId, projectId);
    return this.store.state.artifacts.filter((artifact) => artifact.projectId === projectId);
  }

  promptSnapshots(ownerId: string, projectId: string) {
    this.getProject(ownerId, projectId);
    return this.store.state.promptSnapshots.filter((snapshot) => snapshot.projectId === projectId);
  }
}

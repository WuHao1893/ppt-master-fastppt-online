export type ProjectStatus = 'draft' | 'processing' | 'ready' | 'failed' | 'archived';
export type PageType = 'cover' | 'toc' | 'content' | 'section' | 'ending' | 'other';
export type EditableLevel = 'visual' | 'text_layer' | 'native_partial' | 'native_structure';
export type VersionStatus = 'draft' | 'previewing' | 'rendering' | 'ready' | 'rejected' | 'failed';
export type PageStatus =
  | 'untouched'
  | 'quick_preview'
  | 'generating'
  | 'authoritative'
  | 'svg_fallback'
  | 'failed'
  | 'rolled_back';
export type EditMode = 'single' | 'multi' | 'global';
export type OperationStatus = 'planned' | 'confirmed' | 'applying' | 'completed' | 'failed' | 'rolled_back';
export type PreviewKind = 'quick_preview' | 'pptx_authoritative' | 'svg_fallback';

export interface User {
  userId: string;
  email: string;
  name: string;
  createdAt: string;
}

export interface AuthSession {
  sessionId: string;
  userId: string;
  expiresAt: string;
  createdAt: string;
  revokedAt: string | null;
}

export interface FactAnchor {
  factId: string;
  value: string;
  source: string;
  locked: boolean;
}

export interface PageContract {
  conclusion: string;
  evidence: string[];
  mustKeep: string[];
  canTrim: string[];
  layout: string;
  density: 'airy' | 'balanced' | 'dense';
  components: string[];
}

export interface PageVersion {
  versionId: string;
  pageId: string;
  parentVersionId: string | null;
  sourceRevision: string;
  pageContractPath: string;
  slidesSourceHash: string;
  previewArtifactId: string | null;
  /** Non-sensitive artifact id for a page-local visual image carried into export. */
  visualArtifactId?: string | null;
  svgArtifactId: string | null;
  pptxPageRenderId: string | null;
  promptSnapshotId: string;
  editOperationId: string;
  qualityReportId: string | null;
  status: VersionStatus;
  createdAt: string;
  message: string;
  title: string;
  body: string;
  layout: string;
  sourceMarkdown?: string;
  contract?: PageContract;
  previewKind: PreviewKind;
  previewSvg: string;
  editableLevel: EditableLevel;
  nonEditableRegions: string[];
  qaWarnings: string[];
  renderNote: string;
  factAnchors?: FactAnchor[];
}

export interface Page {
  pageId: string;
  projectId: string;
  currentVersionId: string;
  orderIndex: number;
  pageType: PageType;
  locked: boolean;
  archived?: boolean;
  factAnchorIds: string[];
  factAnchors: FactAnchor[];
  editableLevel: EditableLevel;
  status: PageStatus;
  title: string;
  body: string;
  layout: string;
  sourceMarkdown: string;
  contract: PageContract;
  versions: PageVersion[];
}

export interface Project {
  projectId: string;
  ownerId: string;
  name: string;
  themeId: string;
  themeVersion: string;
  currentDeckRevisionId: string;
  goalId: string;
  status: ProjectStatus;
  createdAt: string;
  updatedAt: string;
  pages: Page[];
  sourceMarkdown: string;
}

export interface EditChange {
  kind: 'preserve_fact' | 'rewrite_text' | 'layout_change' | 'style_change' | 'image_replace' | 'unsupported';
  target: string;
  value?: string;
  constraint?: string;
  factId?: string;
}

export interface EditPlan {
  intent: string;
  affectedPageIds: string[];
  changes: EditChange[];
  factImpact: { added: string[]; removed: string[]; changed: string[] };
  unsupported: string[];
  requiresConfirmation: boolean;
  estimatedCost: { imageUnits: number; amount: number; currency: string };
  summary: string;
  candidateReasons?: Record<string, string>;
}

export interface EditOperation {
  operationId: string;
  projectId: string;
  conversationId: string;
  mode: EditMode;
  requestedPageIds: string[];
  resolvedPageIds: string[];
  message: string;
  structuredPlan: EditPlan;
  factImpact: EditPlan['factImpact'];
  unsupportedItems: string[];
  confirmationRequired: boolean;
  confirmedAt: string | null;
  resultVersionIds: string[];
  status: OperationStatus;
  createdAt: string;
  completedAt: string | null;
  failedPageIds: string[];
  estimatedCost: number;
  parentOperationId?: string | null;
}

export interface ConversationMessage {
  messageId: string;
  projectId: string;
  conversationId: string;
  role: 'user' | 'assistant' | 'system';
  text: string;
  createdAt: string;
  operationId?: string;
  meta?: Record<string, unknown>;
}

export interface UsageLedgerEntry {
  ledgerId: string;
  projectId: string;
  operationId: string;
  pageId: string | null;
  versionId: string | null;
  model: string;
  unitPrice: number;
  reservedAmount: number;
  settledAmount: number;
  refundedAmount: number;
  status: 'reserved' | 'settled' | 'refunded' | 'unknown';
  createdAt: string;
}

export interface ExportJob {
  exportId: string;
  projectId: string;
  status: 'queued' | 'running' | 'completed' | 'failed';
  artifactPath: string | null;
  renderMode: 'powerpoint' | 'svg_fallback';
  qaWarnings: string[];
  createdAt: string;
  completedAt: string | null;
  artifactName?: string | null;
  artifactObjectKey?: string | null;
  exportEngine?: 'ppt_master_svg_to_drawingml' | 'legacy_python_pptx_development_fallback';
  qaStatus?: 'pending' | 'passed' | 'passed-with-warnings' | 'failed' | 'not-run-development-fallback';
}

export interface DurableJob {
  jobId: string;
  projectId: string;
  operationId?: string | null;
  kind: 'export' | 'edit';
  status: 'queued' | 'running' | 'completed' | 'failed';
  payload: Record<string, unknown>;
  createdAt: string;
  completedAt: string | null;
}

export interface PreviewArtifact {
  artifactId: string;
  projectId: string;
  pageId: string;
  versionId: string;
  kind: 'visual_preview' | 'svg_preview' | 'pptx_page_render';
  promptSnapshotId: string;
  model: string;
  width: number;
  height: number;
  quality: string;
  source: 'deterministic_svg' | 'relay_image' | 'powerpoint_com';
  provenance: Record<string, unknown>;
  createdAt: string;
  assetPath?: string | null;
}

export interface PromptSnapshot {
  promptSnapshotId: string;
  projectId: string;
  pageId: string;
  operationId: string;
  prompt: string;
  hash: string;
  model: string;
  createdAt: string;
}

export interface AuditLog {
  auditId: string;
  ownerId: string;
  projectId?: string;
  action: string;
  requestId?: string;
  payload: Record<string, unknown>;
  createdAt: string;
}

export interface EventEnvelope {
  seq: number;
  type: string;
  projectId: string;
  pageId?: string;
  versionId?: string;
  operationId?: string;
  exportId?: string;
  payload: Record<string, unknown>;
  createdAt: string;
}

export interface StoreState {
  users: User[];
  authSessions: AuthSession[];
  projects: Project[];
  operations: EditOperation[];
  messages: ConversationMessage[];
  ledger: UsageLedgerEntry[];
  exports: ExportJob[];
  artifacts: PreviewArtifact[];
  promptSnapshots: PromptSnapshot[];
  auditLogs: AuditLog[];
  events: EventEnvelope[];
  seq: number;
  jobs: DurableJob[];
}

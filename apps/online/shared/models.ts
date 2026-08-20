export type ProjectStatus =
  | "draft"
  | "processing"
  | "ready"
  | "failed"
  | "archived";
export type PageType =
  | "cover"
  | "toc"
  | "content"
  | "section"
  | "ending"
  | "other";
export type EditableLevel =
  | "visual"
  | "text_layer"
  | "native_partial"
  | "native_structure";
export type VersionStatus =
  | "draft"
  | "previewing"
  | "rendering"
  | "ready"
  | "rejected"
  | "failed";
export type PageStatus =
  | "untouched"
  | "quick_preview"
  | "generating"
  | "authoritative"
  | "svg_fallback"
  | "failed"
  | "rolled_back";
export type EditMode = "single" | "multi" | "global";
export type OperationStatus =
  | "planned"
  | "confirmed"
  | "applying"
  | "completed"
  | "failed"
  | "rolled_back";
export type PreviewKind =
  | "quick_preview"
  | "pptx_authoritative"
  | "svg_fallback";

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
  density: "airy" | "balanced" | "dense";
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
  settings?: ProjectSettings;
}

export interface ProjectSettings {
  sensitiveMode: boolean;
}

export type WorkflowMode = "document_import" | "page_entry" | "ppt_beautify";
export type WorkSessionStatus =
  | "draft"
  | "collecting"
  | "parsing"
  | "planned"
  | "blocked"
  | "executing"
  | "completed"
  | "failed"
  | "cancelled";
export interface WorkSession {
  sessionId: string;
  projectId: string;
  workflowMode: WorkflowMode;
  intent: "create" | "modify" | "reference";
  sourceDocumentIds: string[];
  sourceRevision: string;
  planId: string | null;
  pageBudget: number;
  inheritTheme: boolean;
  structurePlan: DocumentStructurePlan | null;
  createdBy: string;
  status: WorkSessionStatus;
  createdAt: string;
  updatedAt: string;
}

export interface DocumentStructurePlan {
  operationId: string;
  intent: "create" | "modify" | "reference";
  pageBudget: number;
  inheritTheme: boolean;
  preservesPageCount: boolean;
  totalPages: number;
  pages: Array<{
    pageId: string;
    kind: "cover" | "transition" | "content";
    title: string;
    bodyPreview: string;
    sourceDocumentIds: string[];
  }>;
  createdAt: string;
}

export type DocumentParseStatus =
  | "queued"
  | "parsing"
  | "ready"
  | "warning"
  | "blocked"
  | "failed"
  | "uploaded"
  | "partial"
  | "rejected";
export interface DocumentFact {
  factId: string;
  key: string;
  value: string;
  normalizedValue: string;
  kind:
    | "number"
    | "date"
    | "person"
    | "organization"
    | "metric"
    | "claim"
    | "term"
    | "source";
  sourceDocumentId: string;
  sourceLocator: string;
  confidence: number;
  locked: boolean;
  location: string;
  context: string;
}
export interface DocumentSource {
  documentId: string;
  projectId: string;
  sessionId: string;
  fileName: string;
  fileType: "md" | "docx" | "pdf" | "pptx";
  sizeBytes: number;
  sha256: string;
  sourceVersionId: string;
  uploadedBy: string;
  objectKey: string;
  parseStatus: DocumentParseStatus;
  includedInContext: boolean;
  parsedText: string;
  structure: DocumentStructure | null;
  facts: DocumentFact[];
  warnings: string[];
  error: string | null;
  createdAt: string;
  updatedAt: string;
}

export interface DocumentVersion {
  sourceVersionId: string;
  documentId: string;
  projectId: string;
  sha256: string;
  sizeBytes: number;
  mediaType: string;
  objectKey: string;
  createdBy: string;
  createdAt: string;
}

export interface DocumentSlideStructure {
  index: number;
  title: string;
  body: string;
  layoutName: string;
  shapeCount: number;
  textShapeCount: number;
  imageCount: number;
  tableCount: number;
  editableLevel: EditableLevel;
  nonEditableRegions: string[];
}

export interface DocumentStructure {
  headings: string[];
  lineCount: number;
  characterCount: number;
  pageCount?: number;
  paragraphCount?: number;
  tableCount?: number;
  hyperlinkCount?: number;
  imageCount?: number;
  properties?: Record<string, string>;
  slideWidth?: number;
  slideHeight?: number;
  slides?: DocumentSlideStructure[];
}

export interface DocumentConflict {
  conflictId: string;
  projectId: string;
  sessionId: string;
  sourceRevision: string;
  key: string;
  values: Array<{
    value: string;
    documentId: string;
    fileName: string;
    location: string;
    sourceLocator: string;
    originalText: string;
    confidence: number;
    affectedPageIds: string[];
  }>;
  severity: "warning" | "blocking";
  status: "open" | "resolved";
  resolution: {
    resolutionId: string;
    action: "prefer_source" | "keep_both" | "ignore";
    selectedValue: string | null;
    selectedDocumentIds: string[];
    note: string;
  } | null;
  createdAt: string;
  resolvedAt: string | null;
}

export interface ConflictResolution {
  resolutionId: string;
  conflictId: string;
  projectId: string;
  sessionId: string;
  sourceRevision: string;
  planId: string | null;
  action: "prefer_source" | "keep_both" | "ignore";
  selectedValue: string | null;
  selectedDocumentIds: string[];
  note: string;
  resolvedBy: string;
  createdAt: string;
}

export interface EditChange {
  kind:
    | "preserve_fact"
    | "rewrite_text"
    | "layout_change"
    | "style_change"
    | "image_replace"
    | "unsupported";
  target: string;
  value?: string;
  constraint?: string;
  factId?: string;
}

export type PlanWorkflowMode =
  | "import_document"
  | "page_entry"
  | "pptx_beautify";
export interface EditPlan {
  workflowMode: PlanWorkflowMode;
  targetScope: EditMode;
  intent: string;
  affectedPageIds: string[];
  pageDelta: {
    add: string[];
    remove: string[];
    split: string[];
    merge: string[];
  };
  changes: EditChange[];
  factImpact: { added: string[]; removed: string[]; changed: string[] };
  sourceDocumentIds: string[];
  conflictIds: string[];
  unsupported: string[];
  requiresConfirmation: boolean;
  confirmationReasons: string[];
  estimatedCost: { imageUnits: number; amount: number; currency: string };
  summary: string;
  candidateReasons?: Record<string, string>;
}

export interface EditOperation {
  operationId: string;
  projectId: string;
  conversationId: string;
  sessionId?: string | null;
  mode: EditMode;
  requestedPageIds: string[];
  resolvedPageIds: string[];
  message: string;
  structuredPlan: EditPlan;
  factImpact: EditPlan["factImpact"];
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
  visualPreviews?: OperationVisualPreview[];
}

export interface OperationVisualPreview {
  pageId: string;
  artifactId: string | null;
  promptSnapshotId: string;
  source: "relay_image" | "slidev_svg";
  status: "generating" | "ready" | "failed";
  ledgerId: string | null;
  error: string | null;
}

export interface ConversationMessage {
  messageId: string;
  projectId: string;
  conversationId: string;
  role: "user" | "assistant" | "system";
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
  status: "reserved" | "settled" | "refunded" | "unknown";
  createdAt: string;
}

export interface ExportJob {
  exportId: string;
  projectId: string;
  status: "queued" | "running" | "completed" | "failed";
  artifactPath: string | null;
  renderMode: "powerpoint" | "svg_fallback";
  qaWarnings: string[];
  createdAt: string;
  completedAt: string | null;
  artifactName?: string | null;
  artifactObjectKey?: string | null;
  exportEngine?:
    | "ppt_master_svg_to_drawingml"
    | "legacy_python_pptx_development_fallback";
  qaStatus?:
    | "pending"
    | "passed"
    | "passed-with-warnings"
    | "failed"
    | "not-run-development-fallback";
}

export interface DurableJob {
  jobId: string;
  idempotencyKey: string;
  projectId: string;
  operationId?: string | null;
  kind: "export" | "edit";
  status: "queued" | "running" | "completed" | "failed";
  payload: Record<string, unknown>;
  attemptCount: number;
  maxAttempts: number;
  availableAt: string;
  leaseOwner: string | null;
  leaseExpiresAt: string | null;
  lastError: string | null;
  startedAt: string | null;
  createdAt: string;
  updatedAt: string;
  completedAt: string | null;
}

export interface PreviewArtifact {
  artifactId: string;
  projectId: string;
  pageId: string;
  versionId: string | null;
  kind: "visual_preview" | "svg_preview" | "pptx_page_render";
  promptSnapshotId: string;
  model: string;
  width: number;
  height: number;
  quality: string;
  source: "deterministic_svg" | "relay_image" | "powerpoint_com";
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

export interface OperationHistoryPage {
  pageId: string;
  pageTitle: string;
  versionId: string | null;
  model: string | null;
  cost: number;
  qaStatus: "passed" | "warning" | "failed" | "pending";
  warnings: string[];
}

export interface OperationHistoryEntry {
  operationId: string;
  conversationId: string;
  sessionId: string | null;
  sessionStatus: WorkSessionStatus | null;
  status: OperationStatus;
  message: string;
  createdAt: string;
  completedAt: string | null;
  durationMs: number | null;
  estimatedCost: number;
  actualCost: number;
  currency: string;
  pageIds: string[];
  versionIds: string[];
  models: string[];
  promptSnapshotIds: string[];
  qaStatus: "passed" | "warning" | "failed" | "pending";
  warnings: string[];
  pages: OperationHistoryPage[];
}

export interface ProjectHistory {
  messages: ConversationMessage[];
  operations: OperationHistoryEntry[];
  sessions: WorkSession[];
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

export interface OfficePreviewGrant {
  grantId: string;
  exportId: string;
  projectId: string;
  issuedTo: string;
  tokenHash: string;
  expiresAt: string;
  maxFetches: number;
  fetchCount: number;
  revokedAt: string | null;
  createdAt: string;
}

export interface EventEnvelope {
  seq: number;
  type: string;
  projectId: string;
  sessionId?: string;
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
  officePreviewGrants: OfficePreviewGrant[];
  workSessions: WorkSession[];
  documents: DocumentSource[];
  documentVersions: DocumentVersion[];
  documentConflicts: DocumentConflict[];
  conflictResolutions: ConflictResolution[];
}

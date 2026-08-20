import fs from "node:fs/promises";
import path from "node:path";
import { Pool, type PoolClient } from "pg";
import type { DurableJob, StoreState } from "../shared/models.js";

export interface StateStore {
  readonly kind: "file" | "postgres";
  state: StoreState;
  init(): Promise<void>;
  refresh(): Promise<void>;
  update<T>(mutator: (state: StoreState) => T | Promise<T>): Promise<T>;
  enqueueJob(job: DurableJob): Promise<DurableJob>;
  claimJob(workerId: string, leaseMs: number): Promise<DurableJob | null>;
  heartbeatJob(jobId: string, workerId: string, leaseMs: number): Promise<boolean>;
  completeJob(jobId: string, workerId: string): Promise<DurableJob | null>;
  failJob(
    jobId: string,
    workerId: string,
    error: string,
    retryDelayMs: number,
  ): Promise<DurableJob | null>;
  releaseExpiredJobLeases(): Promise<number>;
}

export function emptyState(): StoreState {
  return {
    users: [],
    authSessions: [],
    projects: [],
    operations: [],
    messages: [],
    ledger: [],
    exports: [],
    artifacts: [],
    promptSnapshots: [],
    auditLogs: [],
    events: [],
    seq: 0,
    jobs: [],
    officePreviewGrants: [],
    workSessions: [],
    documents: [],
    documentVersions: [],
    documentConflicts: [],
    conflictResolutions: [],
  };
}

function normalizeDurableJob(job: Partial<DurableJob> & {
  jobId: string;
  projectId: string;
  kind: DurableJob["kind"];
}): DurableJob {
  const createdAt = job.createdAt || new Date().toISOString();
  return {
    jobId: job.jobId,
    idempotencyKey: job.idempotencyKey || `${job.kind}:${job.jobId}`,
    projectId: job.projectId,
    operationId: job.operationId || null,
    kind: job.kind,
    status: job.status || "queued",
    payload: job.payload || {},
    attemptCount: Number(job.attemptCount || 0),
    maxAttempts: Number(job.maxAttempts || 3),
    availableAt: job.availableAt || createdAt,
    leaseOwner: job.leaseOwner || null,
    leaseExpiresAt: job.leaseExpiresAt || null,
    lastError: job.lastError || null,
    startedAt: job.startedAt || null,
    createdAt,
    updatedAt: job.updatedAt || createdAt,
    completedAt: job.completedAt || null,
  };
}

function reconcileStoreState(target: StoreState, fresh: StoreState): void {
  const replaceObject = <T extends object>(targetItem: T, freshItem: T): void => {
    for (const key of Object.keys(targetItem)) delete (targetItem as any)[key];
    Object.assign(targetItem, structuredClone(freshItem));
  };
  const syncArray = <T extends object>(
    targetItems: T[],
    freshItems: T[],
    key: (item: T) => string | number,
    reconcile: (targetItem: T, freshItem: T) => void = replaceObject,
  ): void => {
    const existing = new Map(targetItems.map((item) => [key(item), item]));
    targetItems.splice(
      0,
      targetItems.length,
      ...freshItems.map((freshItem) => {
        const targetItem = existing.get(key(freshItem));
        if (!targetItem) return structuredClone(freshItem);
        reconcile(targetItem, freshItem);
        return targetItem;
      }),
    );
  };
  const syncPage = (
    targetPage: StoreState["projects"][number]["pages"][number],
    freshPage: StoreState["projects"][number]["pages"][number],
  ): void => {
    const versions = targetPage.versions;
    const { versions: freshVersions, ...freshFields } = freshPage;
    for (const key of Object.keys(targetPage)) delete (targetPage as any)[key];
    Object.assign(targetPage, structuredClone(freshFields), { versions });
    syncArray(
      versions,
      freshVersions,
      (version) => version.versionId,
    );
  };
  const syncProject = (
    targetProject: StoreState["projects"][number],
    freshProject: StoreState["projects"][number],
  ): void => {
    const pages = targetProject.pages;
    const { pages: freshPages, ...freshFields } = freshProject;
    for (const key of Object.keys(targetProject))
      delete (targetProject as any)[key];
    Object.assign(targetProject, structuredClone(freshFields), { pages });
    syncArray(pages, freshPages, (page) => page.pageId, syncPage);
  };

  syncArray(target.users, fresh.users, (item) => item.userId);
  syncArray(target.authSessions, fresh.authSessions, (item) => item.sessionId);
  syncArray(target.projects, fresh.projects, (item) => item.projectId, syncProject);
  syncArray(target.operations, fresh.operations, (item) => item.operationId);
  syncArray(target.messages, fresh.messages, (item) => item.messageId);
  syncArray(target.ledger, fresh.ledger, (item) => item.ledgerId);
  syncArray(target.exports, fresh.exports, (item) => item.exportId);
  syncArray(target.artifacts, fresh.artifacts, (item) => item.artifactId);
  syncArray(
    target.promptSnapshots,
    fresh.promptSnapshots,
    (item) => item.promptSnapshotId,
  );
  syncArray(target.auditLogs, fresh.auditLogs, (item) => item.auditId);
  syncArray(target.events, fresh.events, (item) => item.seq);
  syncArray(target.jobs, fresh.jobs, (item) => item.jobId);
  syncArray(
    target.officePreviewGrants,
    fresh.officePreviewGrants,
    (item) => item.grantId,
  );
  syncArray(target.workSessions, fresh.workSessions, (item) => item.sessionId);
  syncArray(target.documents, fresh.documents, (item) => item.documentId);
  syncArray(
    target.documentVersions,
    fresh.documentVersions,
    (item) => item.sourceVersionId,
  );
  syncArray(
    target.documentConflicts,
    fresh.documentConflicts,
    (item) => item.conflictId,
  );
  syncArray(
    target.conflictResolutions,
    fresh.conflictResolutions,
    (item) => item.resolutionId,
  );
  target.seq = fresh.seq;
}

abstract class BaseStore implements StateStore {
  abstract readonly kind: "file" | "postgres";
  state: StoreState = emptyState();
  protected mutationQueue = Promise.resolve();

  abstract init(): Promise<void>;
  abstract refresh(): Promise<void>;
  protected abstract persist(previousState?: StoreState): Promise<void>;

  async update<T>(mutator: (state: StoreState) => T | Promise<T>): Promise<T> {
    let result: T;
    const task = this.mutationQueue.then(async () => {
      const previousState = structuredClone(this.state);
      result = await mutator(this.state);
      await this.persist(previousState);
    });
    this.mutationQueue = task.then(
      () => undefined,
      () => undefined,
    );
    await task;
    return result!;
  }

  async enqueueJob(job: DurableJob): Promise<DurableJob> {
    return this.update((state) => {
      const existing = state.jobs.find(
        (candidate) =>
          candidate.jobId === job.jobId ||
          candidate.idempotencyKey === job.idempotencyKey,
      );
      if (existing) return existing;
      state.jobs.push(job);
      return job;
    });
  }

  async claimJob(workerId: string, leaseMs: number): Promise<DurableJob | null> {
    return this.update((state) => {
      const now = Date.now();
      const job = state.jobs
        .filter(
          (candidate) =>
            (candidate.status === "queued" &&
              new Date(candidate.availableAt).getTime() <= now) ||
            (candidate.status === "running" &&
              (!candidate.leaseExpiresAt ||
                new Date(candidate.leaseExpiresAt).getTime() <= now)),
        )
        .sort((a, b) => a.availableAt.localeCompare(b.availableAt))[0];
      if (!job) return null;
      const timestamp = new Date(now).toISOString();
      job.status = "running";
      job.attemptCount += 1;
      job.leaseOwner = workerId;
      job.leaseExpiresAt = new Date(now + leaseMs).toISOString();
      job.startedAt ||= timestamp;
      job.updatedAt = timestamp;
      return job;
    });
  }

  async heartbeatJob(
    jobId: string,
    workerId: string,
    leaseMs: number,
  ): Promise<boolean> {
    return this.update((state) => {
      const job = state.jobs.find(
        (candidate) =>
          candidate.jobId === jobId &&
          candidate.status === "running" &&
          candidate.leaseOwner === workerId,
      );
      if (!job) return false;
      job.leaseExpiresAt = new Date(Date.now() + leaseMs).toISOString();
      job.updatedAt = new Date().toISOString();
      return true;
    });
  }

  async completeJob(
    jobId: string,
    workerId: string,
  ): Promise<DurableJob | null> {
    return this.update((state) => {
      const job = state.jobs.find(
        (candidate) =>
          candidate.jobId === jobId &&
          candidate.status === "running" &&
          candidate.leaseOwner === workerId,
      );
      if (!job) return null;
      const timestamp = new Date().toISOString();
      job.status = "completed";
      job.completedAt = timestamp;
      job.updatedAt = timestamp;
      job.leaseOwner = null;
      job.leaseExpiresAt = null;
      job.lastError = null;
      return job;
    });
  }

  async failJob(
    jobId: string,
    workerId: string,
    error: string,
    retryDelayMs: number,
  ): Promise<DurableJob | null> {
    return this.update((state) => {
      const job = state.jobs.find(
        (candidate) =>
          candidate.jobId === jobId &&
          candidate.status === "running" &&
          candidate.leaseOwner === workerId,
      );
      if (!job) return null;
      const now = Date.now();
      const retry = job.attemptCount < job.maxAttempts;
      job.status = retry ? "queued" : "failed";
      job.availableAt = new Date(now + (retry ? retryDelayMs : 0)).toISOString();
      job.completedAt = retry ? null : new Date(now).toISOString();
      job.updatedAt = new Date(now).toISOString();
      job.leaseOwner = null;
      job.leaseExpiresAt = null;
      job.lastError = error.slice(0, 4_000);
      return job;
    });
  }

  async releaseExpiredJobLeases(): Promise<number> {
    return this.update((state) => {
      const now = Date.now();
      let released = 0;
      for (const job of state.jobs) {
        if (
          job.status === "running" &&
          (!job.leaseExpiresAt || new Date(job.leaseExpiresAt).getTime() <= now)
        ) {
          job.status = "queued";
          job.availableAt = new Date(now).toISOString();
          job.leaseOwner = null;
          job.leaseExpiresAt = null;
          job.updatedAt = new Date(now).toISOString();
          released += 1;
        }
      }
      return released;
    });
  }
}

export class FileStore extends BaseStore {
  readonly kind = "file" as const;
  private readonly filePath: string;

  constructor(dataDir: string) {
    super();
    this.filePath = path.resolve(dataDir, "store.json");
  }

  async init(): Promise<void> {
    await fs.mkdir(path.dirname(this.filePath), { recursive: true });
    try {
      const raw = await fs.readFile(this.filePath, "utf8");
      this.state = { ...emptyState(), ...JSON.parse(raw) } as StoreState;
      this.state.jobs = this.state.jobs.map((job) => normalizeDurableJob(job));
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
      await this.persist();
    }
  }

  async refresh(): Promise<void> {
    // File mode is single-process development persistence.
  }

  protected async persist(): Promise<void> {
    const tempPath = `${this.filePath}.tmp`;
    await fs.writeFile(tempPath, JSON.stringify(this.state, null, 2), "utf8");
    await fs.rename(tempPath, this.filePath);
  }
}

export class PostgresStore extends BaseStore {
  readonly kind = "postgres" as const;
  private readonly pool: Pool;

  constructor(databaseUrl: string) {
    super();
    this.pool = new Pool({ connectionString: databaseUrl, max: 5 });
  }

  async init(): Promise<void> {
    await this.pool.query(`
      CREATE TABLE IF NOT EXISTS users (user_id TEXT PRIMARY KEY, email TEXT NOT NULL UNIQUE, name TEXT NOT NULL, created_at TIMESTAMPTZ NOT NULL);
      CREATE TABLE IF NOT EXISTS auth_sessions (session_id TEXT PRIMARY KEY, user_id TEXT NOT NULL REFERENCES users(user_id) ON DELETE CASCADE, expires_at TIMESTAMPTZ NOT NULL, created_at TIMESTAMPTZ NOT NULL, revoked_at TIMESTAMPTZ);
      CREATE TABLE IF NOT EXISTS projects (project_id TEXT PRIMARY KEY, owner_id TEXT NOT NULL, name TEXT NOT NULL, theme_id TEXT NOT NULL, theme_version TEXT NOT NULL, current_deck_revision_id TEXT NOT NULL, goal_id TEXT NOT NULL, status TEXT NOT NULL, source_markdown TEXT NOT NULL DEFAULT '', sensitive_mode BOOLEAN NOT NULL DEFAULT FALSE, created_at TIMESTAMPTZ NOT NULL, updated_at TIMESTAMPTZ NOT NULL);
      ALTER TABLE projects ADD COLUMN IF NOT EXISTS sensitive_mode BOOLEAN NOT NULL DEFAULT FALSE;
      CREATE TABLE IF NOT EXISTS pages (page_id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, current_version_id TEXT, order_index INTEGER NOT NULL, page_type TEXT NOT NULL, locked BOOLEAN NOT NULL DEFAULT FALSE, archived BOOLEAN NOT NULL DEFAULT FALSE, fact_anchor_ids JSONB NOT NULL DEFAULT '[]'::jsonb, editable_level TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'untouched', payload JSONB NOT NULL DEFAULT '{}'::jsonb);
      ALTER TABLE pages ADD COLUMN IF NOT EXISTS payload JSONB NOT NULL DEFAULT '{}'::jsonb;
      CREATE TABLE IF NOT EXISTS page_versions (version_id TEXT PRIMARY KEY, page_id TEXT NOT NULL REFERENCES pages(page_id) ON DELETE CASCADE, parent_version_id TEXT, source_revision TEXT NOT NULL, page_contract_path TEXT NOT NULL, slides_source_hash TEXT NOT NULL, preview_artifact_id TEXT, svg_artifact_id TEXT, pptx_page_render_id TEXT, prompt_snapshot_id TEXT NOT NULL, edit_operation_id TEXT NOT NULL, quality_report_id TEXT, status TEXT NOT NULL, created_at TIMESTAMPTZ NOT NULL, payload JSONB NOT NULL DEFAULT '{}'::jsonb);
      CREATE TABLE IF NOT EXISTS edit_operations (operation_id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, conversation_id TEXT NOT NULL, mode TEXT NOT NULL, requested_page_ids JSONB NOT NULL DEFAULT '[]'::jsonb, resolved_page_ids JSONB NOT NULL DEFAULT '[]'::jsonb, message TEXT NOT NULL, structured_plan JSONB NOT NULL, fact_impact JSONB NOT NULL, unsupported_items JSONB NOT NULL DEFAULT '[]'::jsonb, confirmation_required BOOLEAN NOT NULL DEFAULT FALSE, confirmed_at TIMESTAMPTZ, result_version_ids JSONB NOT NULL DEFAULT '[]'::jsonb, status TEXT NOT NULL, created_at TIMESTAMPTZ NOT NULL, completed_at TIMESTAMPTZ, payload JSONB NOT NULL DEFAULT '{}'::jsonb);
      ALTER TABLE edit_operations ADD COLUMN IF NOT EXISTS payload JSONB NOT NULL DEFAULT '{}'::jsonb;
      CREATE TABLE IF NOT EXISTS usage_ledger (ledger_id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, operation_id TEXT NOT NULL, page_id TEXT, version_id TEXT, model TEXT NOT NULL, unit_price NUMERIC(12, 6) NOT NULL DEFAULT 0, reserved_amount NUMERIC(12, 6) NOT NULL DEFAULT 0, settled_amount NUMERIC(12, 6) NOT NULL DEFAULT 0, refunded_amount NUMERIC(12, 6) NOT NULL DEFAULT 0, status TEXT NOT NULL, created_at TIMESTAMPTZ NOT NULL);
      CREATE TABLE IF NOT EXISTS preview_artifacts (artifact_id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, page_id TEXT NOT NULL REFERENCES pages(page_id) ON DELETE CASCADE, version_id TEXT NOT NULL, kind TEXT NOT NULL, prompt_snapshot_id TEXT NOT NULL, model TEXT NOT NULL, width INTEGER NOT NULL, height INTEGER NOT NULL, quality TEXT NOT NULL, source TEXT NOT NULL, provenance JSONB NOT NULL DEFAULT '{}'::jsonb, created_at TIMESTAMPTZ NOT NULL, asset_path TEXT);
      ALTER TABLE preview_artifacts ADD COLUMN IF NOT EXISTS asset_path TEXT;
      ALTER TABLE preview_artifacts ALTER COLUMN version_id DROP NOT NULL;
      CREATE TABLE IF NOT EXISTS prompt_snapshots (prompt_snapshot_id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, page_id TEXT NOT NULL REFERENCES pages(page_id) ON DELETE CASCADE, operation_id TEXT NOT NULL, prompt TEXT NOT NULL, hash TEXT NOT NULL, model TEXT NOT NULL, created_at TIMESTAMPTZ NOT NULL);
      CREATE TABLE IF NOT EXISTS export_jobs (export_id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, status TEXT NOT NULL, artifact_path TEXT, artifact_object_key TEXT, render_mode TEXT NOT NULL, qa_warnings JSONB NOT NULL DEFAULT '[]'::jsonb, artifact_name TEXT, created_at TIMESTAMPTZ NOT NULL, completed_at TIMESTAMPTZ);
      ALTER TABLE export_jobs ADD COLUMN IF NOT EXISTS artifact_object_key TEXT;
      CREATE TABLE IF NOT EXISTS conversation_messages (message_id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, conversation_id TEXT NOT NULL, role TEXT NOT NULL, text TEXT NOT NULL, payload JSONB NOT NULL DEFAULT '{}'::jsonb, created_at TIMESTAMPTZ NOT NULL);
      CREATE TABLE IF NOT EXISTS event_log (seq BIGINT PRIMARY KEY, project_id TEXT NOT NULL, payload JSONB NOT NULL, created_at TIMESTAMPTZ NOT NULL);
      CREATE TABLE IF NOT EXISTS jobs (job_id TEXT PRIMARY KEY, idempotency_key TEXT NOT NULL, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, operation_id TEXT, kind TEXT NOT NULL, status TEXT NOT NULL, payload JSONB NOT NULL DEFAULT '{}'::jsonb, attempt_count INTEGER NOT NULL DEFAULT 0, max_attempts INTEGER NOT NULL DEFAULT 3, available_at TIMESTAMPTZ NOT NULL DEFAULT NOW(), lease_owner TEXT, lease_expires_at TIMESTAMPTZ, last_error TEXT, started_at TIMESTAMPTZ, created_at TIMESTAMPTZ NOT NULL, updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(), completed_at TIMESTAMPTZ);
      ALTER TABLE jobs ADD COLUMN IF NOT EXISTS idempotency_key TEXT;
      ALTER TABLE jobs ADD COLUMN IF NOT EXISTS attempt_count INTEGER NOT NULL DEFAULT 0;
      ALTER TABLE jobs ADD COLUMN IF NOT EXISTS max_attempts INTEGER NOT NULL DEFAULT 3;
      ALTER TABLE jobs ADD COLUMN IF NOT EXISTS available_at TIMESTAMPTZ NOT NULL DEFAULT NOW();
      ALTER TABLE jobs ADD COLUMN IF NOT EXISTS lease_owner TEXT;
      ALTER TABLE jobs ADD COLUMN IF NOT EXISTS lease_expires_at TIMESTAMPTZ;
      ALTER TABLE jobs ADD COLUMN IF NOT EXISTS last_error TEXT;
      ALTER TABLE jobs ADD COLUMN IF NOT EXISTS started_at TIMESTAMPTZ;
      ALTER TABLE jobs ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW();
      UPDATE jobs SET idempotency_key = kind || ':' || job_id WHERE idempotency_key IS NULL OR idempotency_key = '';
      ALTER TABLE jobs ALTER COLUMN idempotency_key SET NOT NULL;
      CREATE UNIQUE INDEX IF NOT EXISTS jobs_idempotency_key_idx ON jobs(idempotency_key);
      CREATE INDEX IF NOT EXISTS jobs_claim_idx ON jobs(status, available_at, lease_expires_at);
      CREATE TABLE IF NOT EXISTS audit_logs (audit_id TEXT PRIMARY KEY, owner_id TEXT NOT NULL, project_id TEXT, action TEXT NOT NULL, request_id TEXT, payload JSONB NOT NULL DEFAULT '{}'::jsonb, created_at TIMESTAMPTZ NOT NULL DEFAULT NOW());
      CREATE TABLE IF NOT EXISTS office_preview_grants (grant_id TEXT PRIMARY KEY, export_id TEXT NOT NULL, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, issued_to TEXT NOT NULL, token_hash TEXT NOT NULL UNIQUE, expires_at TIMESTAMPTZ NOT NULL, max_fetches INTEGER NOT NULL, fetch_count INTEGER NOT NULL DEFAULT 0, revoked_at TIMESTAMPTZ, created_at TIMESTAMPTZ NOT NULL);
      CREATE TABLE IF NOT EXISTS work_sessions (session_id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, workflow_mode TEXT NOT NULL, intent TEXT NOT NULL, source_document_ids JSONB NOT NULL DEFAULT '[]'::jsonb, source_revision TEXT NOT NULL DEFAULT '', plan_id TEXT, page_budget INTEGER NOT NULL DEFAULT 8, inherit_theme BOOLEAN NOT NULL DEFAULT TRUE, structure_plan JSONB, created_by TEXT NOT NULL DEFAULT '', status TEXT NOT NULL, created_at TIMESTAMPTZ NOT NULL, updated_at TIMESTAMPTZ NOT NULL);
      ALTER TABLE work_sessions ADD COLUMN IF NOT EXISTS source_document_ids JSONB NOT NULL DEFAULT '[]'::jsonb;
      ALTER TABLE work_sessions ADD COLUMN IF NOT EXISTS source_revision TEXT NOT NULL DEFAULT '';
      ALTER TABLE work_sessions ADD COLUMN IF NOT EXISTS plan_id TEXT;
      ALTER TABLE work_sessions ADD COLUMN IF NOT EXISTS page_budget INTEGER NOT NULL DEFAULT 8;
      ALTER TABLE work_sessions ADD COLUMN IF NOT EXISTS inherit_theme BOOLEAN NOT NULL DEFAULT TRUE;
      ALTER TABLE work_sessions ADD COLUMN IF NOT EXISTS structure_plan JSONB;
      ALTER TABLE work_sessions ADD COLUMN IF NOT EXISTS created_by TEXT NOT NULL DEFAULT '';
      CREATE TABLE IF NOT EXISTS document_sources (document_id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, session_id TEXT NOT NULL REFERENCES work_sessions(session_id) ON DELETE CASCADE, file_name TEXT NOT NULL, file_type TEXT NOT NULL, size_bytes BIGINT NOT NULL, sha256 TEXT NOT NULL, source_version_id TEXT NOT NULL DEFAULT '', uploaded_by TEXT NOT NULL DEFAULT '', object_key TEXT NOT NULL, parse_status TEXT NOT NULL, included_in_context BOOLEAN NOT NULL DEFAULT TRUE, payload JSONB NOT NULL DEFAULT '{}'::jsonb, created_at TIMESTAMPTZ NOT NULL, updated_at TIMESTAMPTZ NOT NULL, UNIQUE(project_id, sha256));
      ALTER TABLE document_sources ADD COLUMN IF NOT EXISTS source_version_id TEXT NOT NULL DEFAULT '';
      ALTER TABLE document_sources ADD COLUMN IF NOT EXISTS uploaded_by TEXT NOT NULL DEFAULT '';
      CREATE TABLE IF NOT EXISTS document_versions (source_version_id TEXT PRIMARY KEY, document_id TEXT NOT NULL REFERENCES document_sources(document_id) ON DELETE CASCADE, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, sha256 TEXT NOT NULL, size_bytes BIGINT NOT NULL, media_type TEXT NOT NULL, object_key TEXT NOT NULL, created_by TEXT NOT NULL, created_at TIMESTAMPTZ NOT NULL);
      CREATE TABLE IF NOT EXISTS document_conflicts (conflict_id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, session_id TEXT NOT NULL REFERENCES work_sessions(session_id) ON DELETE CASCADE, conflict_key TEXT NOT NULL, severity TEXT NOT NULL, status TEXT NOT NULL, payload JSONB NOT NULL DEFAULT '{}'::jsonb, created_at TIMESTAMPTZ NOT NULL, resolved_at TIMESTAMPTZ);
      CREATE TABLE IF NOT EXISTS conflict_resolutions (resolution_id TEXT PRIMARY KEY, conflict_id TEXT NOT NULL REFERENCES document_conflicts(conflict_id) ON DELETE CASCADE, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, session_id TEXT NOT NULL REFERENCES work_sessions(session_id) ON DELETE CASCADE, source_revision TEXT NOT NULL DEFAULT '', plan_id TEXT, action TEXT NOT NULL, selected_value TEXT, selected_document_ids JSONB NOT NULL DEFAULT '[]'::jsonb, note TEXT NOT NULL DEFAULT '', resolved_by TEXT NOT NULL, created_at TIMESTAMPTZ NOT NULL);
      ALTER TABLE conflict_resolutions ADD COLUMN IF NOT EXISTS source_revision TEXT NOT NULL DEFAULT '';
      ALTER TABLE conflict_resolutions ADD COLUMN IF NOT EXISTS plan_id TEXT;
    `);
    await this.refresh();
  }

  async refresh(): Promise<void> {
    const [
      users,
      sessions,
      projects,
      pages,
      versions,
      operations,
      messages,
      ledger,
      exports,
      artifacts,
      prompts,
      audits,
      events,
      jobs,
      officeGrants,
      workSessions,
      documents,
      documentVersions,
      conflicts,
      conflictResolutions,
    ] = await Promise.all([
      this.pool.query("SELECT * FROM users"),
      this.pool.query("SELECT * FROM auth_sessions"),
      this.pool.query("SELECT * FROM projects"),
      this.pool.query("SELECT * FROM pages"),
      this.pool.query("SELECT * FROM page_versions"),
      this.pool.query("SELECT * FROM edit_operations"),
      this.pool.query("SELECT * FROM conversation_messages"),
      this.pool.query("SELECT * FROM usage_ledger"),
      this.pool.query("SELECT * FROM export_jobs"),
      this.pool.query("SELECT * FROM preview_artifacts"),
      this.pool.query("SELECT * FROM prompt_snapshots"),
      this.pool.query("SELECT * FROM audit_logs"),
      this.pool.query("SELECT * FROM event_log ORDER BY seq"),
      this.pool.query("SELECT * FROM jobs"),
      this.pool.query("SELECT * FROM office_preview_grants"),
      this.pool.query("SELECT * FROM work_sessions"),
      this.pool.query("SELECT * FROM document_sources"),
      this.pool.query("SELECT * FROM document_versions"),
      this.pool.query("SELECT * FROM document_conflicts"),
      this.pool.query("SELECT * FROM conflict_resolutions"),
    ]);
    const json = <T>(value: unknown, fallback: T): T =>
      typeof value === "string"
        ? (JSON.parse(value) as T)
        : ((value as T) ?? fallback);
    const state = emptyState();
    state.users = users.rows.map((row: any) => ({
      userId: row.user_id,
      email: row.email,
      name: row.name,
      createdAt: new Date(row.created_at).toISOString(),
    }));
    state.authSessions = sessions.rows.map((row: any) => ({
      sessionId: row.session_id,
      userId: row.user_id,
      expiresAt: new Date(row.expires_at).toISOString(),
      createdAt: new Date(row.created_at).toISOString(),
      revokedAt: row.revoked_at ? new Date(row.revoked_at).toISOString() : null,
    }));
    state.projects = projects.rows.map((row: any) => ({
      projectId: row.project_id,
      ownerId: row.owner_id,
      name: row.name,
      themeId: row.theme_id,
      themeVersion: row.theme_version,
      currentDeckRevisionId: row.current_deck_revision_id,
      goalId: row.goal_id,
      status: row.status,
      sourceMarkdown: row.source_markdown,
      settings: { sensitiveMode: Boolean(row.sensitive_mode) },
      createdAt: new Date(row.created_at).toISOString(),
      updatedAt: new Date(row.updated_at).toISOString(),
      pages: [],
    }));
    const pageMap = new Map<string, any>();
    for (const row of pages.rows as any[]) {
      const payload = json<Record<string, unknown>>(row.payload, {});
      const page: any = {
        ...payload,
        pageId: row.page_id,
        projectId: row.project_id,
        currentVersionId: row.current_version_id,
        orderIndex: row.order_index,
        pageType: row.page_type,
        locked: row.locked,
        archived: row.archived,
        factAnchorIds: json<string[]>(row.fact_anchor_ids, []),
        editableLevel: row.editable_level,
        status: row.status,
        versions: [],
      };
      pageMap.set(page.pageId, page);
      state.projects
        .find((project) => project.projectId === page.projectId)
        ?.pages.push(page);
    }
    for (const row of versions.rows as any[]) {
      const payload = json<Record<string, unknown>>(row.payload, {});
      const version = {
        ...payload,
        versionId: row.version_id,
        pageId: row.page_id,
        parentVersionId: row.parent_version_id,
        sourceRevision: row.source_revision,
        pageContractPath: row.page_contract_path,
        slidesSourceHash: row.slides_source_hash,
        previewArtifactId: row.preview_artifact_id,
        svgArtifactId: row.svg_artifact_id,
        pptxPageRenderId: row.pptx_page_render_id,
        promptSnapshotId: row.prompt_snapshot_id,
        editOperationId: row.edit_operation_id,
        qualityReportId: row.quality_report_id,
        status: row.status,
        createdAt: new Date(row.created_at).toISOString(),
      };
      pageMap.get(row.page_id)?.versions.push(version);
    }
    state.operations = operations.rows.map(
      (row: any) =>
        ({
          ...json<Record<string, unknown>>(row.payload, {}),
          operationId: row.operation_id,
          projectId: row.project_id,
          conversationId: row.conversation_id,
          mode: row.mode,
          requestedPageIds: json<string[]>(row.requested_page_ids, []),
          resolvedPageIds: json<string[]>(row.resolved_page_ids, []),
          message: row.message,
          structuredPlan: json(row.structured_plan, {}),
          factImpact: json(row.fact_impact, {}),
          unsupportedItems: json<string[]>(row.unsupported_items, []),
          confirmationRequired: row.confirmation_required,
          confirmedAt: row.confirmed_at
            ? new Date(row.confirmed_at).toISOString()
            : null,
          resultVersionIds: json<string[]>(row.result_version_ids, []),
          status: row.status,
          createdAt: new Date(row.created_at).toISOString(),
          completedAt: row.completed_at
            ? new Date(row.completed_at).toISOString()
            : null,
        }) as any,
    );
    state.messages = messages.rows.map(
      (row: any) =>
        ({
          ...json<Record<string, unknown>>(row.payload, {}),
          messageId: row.message_id,
          projectId: row.project_id,
          conversationId: row.conversation_id,
          role: row.role,
          text: row.text,
          createdAt: new Date(row.created_at).toISOString(),
        }) as any,
    );
    state.ledger = ledger.rows.map((row: any) => ({
      ledgerId: row.ledger_id,
      projectId: row.project_id,
      operationId: row.operation_id,
      pageId: row.page_id,
      versionId: row.version_id,
      model: row.model,
      unitPrice: Number(row.unit_price),
      reservedAmount: Number(row.reserved_amount),
      settledAmount: Number(row.settled_amount),
      refundedAmount: Number(row.refunded_amount),
      status: row.status,
      createdAt: new Date(row.created_at).toISOString(),
    }));
    state.exports = exports.rows.map((row: any) => ({
      exportId: row.export_id,
      projectId: row.project_id,
      status: row.status,
      artifactPath: row.artifact_path,
      artifactObjectKey: row.artifact_object_key,
      renderMode: row.render_mode,
      qaWarnings: json<string[]>(row.qa_warnings, []),
      artifactName: row.artifact_name,
      createdAt: new Date(row.created_at).toISOString(),
      completedAt: row.completed_at
        ? new Date(row.completed_at).toISOString()
        : null,
    }));
    state.artifacts = artifacts.rows.map((row: any) => ({
      artifactId: row.artifact_id,
      projectId: row.project_id,
      pageId: row.page_id,
      versionId: row.version_id,
      kind: row.kind,
      promptSnapshotId: row.prompt_snapshot_id,
      model: row.model,
      width: row.width,
      height: row.height,
      quality: row.quality,
      source: row.source,
      provenance: json(row.provenance, {}),
      createdAt: new Date(row.created_at).toISOString(),
      assetPath: row.asset_path,
    }));
    state.promptSnapshots = prompts.rows.map((row: any) => ({
      promptSnapshotId: row.prompt_snapshot_id,
      projectId: row.project_id,
      pageId: row.page_id,
      operationId: row.operation_id,
      prompt: row.prompt,
      hash: row.hash,
      model: row.model,
      createdAt: new Date(row.created_at).toISOString(),
    }));
    state.auditLogs = audits.rows.map((row: any) => ({
      auditId: String(row.audit_id),
      ownerId: row.owner_id,
      projectId: row.project_id || undefined,
      action: row.action,
      requestId: row.request_id || undefined,
      payload: json(row.payload, {}),
      createdAt: new Date(row.created_at).toISOString(),
    }));
    state.events = events.rows.map((row: any) =>
      json(row.payload, {}),
    ) as any[];
    state.seq = state.events.reduce(
      (max, event: any) => Math.max(max, Number(event.seq || 0)),
      0,
    );
    state.jobs = jobs.rows.map((row: any) => ({
      jobId: row.job_id,
      idempotencyKey: row.idempotency_key || `${row.kind}:${row.job_id}`,
      projectId: row.project_id,
      operationId: row.operation_id,
      kind: row.kind,
      status: row.status,
      payload: json(row.payload, {}),
      attemptCount: Number(row.attempt_count || 0),
      maxAttempts: Number(row.max_attempts || 3),
      availableAt: new Date(row.available_at || row.created_at).toISOString(),
      leaseOwner: row.lease_owner || null,
      leaseExpiresAt: row.lease_expires_at
        ? new Date(row.lease_expires_at).toISOString()
        : null,
      lastError: row.last_error || null,
      startedAt: row.started_at
        ? new Date(row.started_at).toISOString()
        : null,
      createdAt: new Date(row.created_at).toISOString(),
      updatedAt: new Date(row.updated_at || row.created_at).toISOString(),
      completedAt: row.completed_at
        ? new Date(row.completed_at).toISOString()
        : null,
    }));
    state.officePreviewGrants = officeGrants.rows.map((row: any) => ({
      grantId: row.grant_id,
      exportId: row.export_id,
      projectId: row.project_id,
      issuedTo: row.issued_to,
      tokenHash: row.token_hash,
      expiresAt: new Date(row.expires_at).toISOString(),
      maxFetches: row.max_fetches,
      fetchCount: row.fetch_count,
      revokedAt: row.revoked_at ? new Date(row.revoked_at).toISOString() : null,
      createdAt: new Date(row.created_at).toISOString(),
    }));
    state.workSessions = workSessions.rows.map((row: any) => ({
      sessionId: row.session_id,
      projectId: row.project_id,
      workflowMode: row.workflow_mode,
      intent: row.intent,
      sourceDocumentIds: json<string[]>(row.source_document_ids, []),
      sourceRevision: row.source_revision || "",
      planId: row.plan_id || null,
      pageBudget: Number(row.page_budget || 8),
      inheritTheme: row.inherit_theme !== false,
      structurePlan: json(row.structure_plan, null),
      createdBy: row.created_by || "",
      status: row.status,
      createdAt: new Date(row.created_at).toISOString(),
      updatedAt: new Date(row.updated_at).toISOString(),
    }));
    state.documents = documents.rows.map(
      (row: any) =>
        ({
          ...json<Record<string, unknown>>(row.payload, {}),
          documentId: row.document_id,
          projectId: row.project_id,
          sessionId: row.session_id,
          fileName: row.file_name,
          fileType: row.file_type,
          sizeBytes: Number(row.size_bytes),
          sha256: row.sha256,
          sourceVersionId: row.source_version_id || "",
          uploadedBy: row.uploaded_by || "",
          objectKey: row.object_key,
          parseStatus: row.parse_status,
          includedInContext: row.included_in_context,
          createdAt: new Date(row.created_at).toISOString(),
          updatedAt: new Date(row.updated_at).toISOString(),
        }) as any,
    );
    state.documentVersions = documentVersions.rows.map((row: any) => ({
      sourceVersionId: row.source_version_id,
      documentId: row.document_id,
      projectId: row.project_id,
      sha256: row.sha256,
      sizeBytes: Number(row.size_bytes),
      mediaType: row.media_type,
      objectKey: row.object_key,
      createdBy: row.created_by,
      createdAt: new Date(row.created_at).toISOString(),
    }));
    state.documentConflicts = conflicts.rows.map(
      (row: any) =>
        ({
          ...json<Record<string, unknown>>(row.payload, {}),
          conflictId: row.conflict_id,
          projectId: row.project_id,
          sessionId: row.session_id,
          key: row.conflict_key,
          severity: row.severity,
          status: row.status,
          createdAt: new Date(row.created_at).toISOString(),
          resolvedAt: row.resolved_at
            ? new Date(row.resolved_at).toISOString()
            : null,
        }) as any,
    );
    state.conflictResolutions = conflictResolutions.rows.map((row: any) => ({
      resolutionId: row.resolution_id,
      conflictId: row.conflict_id,
      projectId: row.project_id,
      sessionId: row.session_id,
      sourceRevision: row.source_revision || "",
      planId: row.plan_id || null,
      action: row.action,
      selectedValue: row.selected_value,
      selectedDocumentIds: json<string[]>(row.selected_document_ids, []),
      note: row.note,
      resolvedBy: row.resolved_by,
      createdAt: new Date(row.created_at).toISOString(),
    }));
    this.state = state;
  }

  override async update<T>(
    mutator: (state: StoreState) => T | Promise<T>,
  ): Promise<T> {
    let result: T;
    const task = this.mutationQueue.then(async () => {
      const client = await this.pool.connect();
      let previousState: StoreState | null = null;
      try {
        await client.query("BEGIN");
        await client.query(
          "SELECT pg_advisory_xact_lock(hashtext('fastppt-online-store'))",
        );
        const retainedState = this.state;
        await this.refresh();
        const freshState = this.state;
        this.state = retainedState;
        reconcileStoreState(this.state, freshState);
        previousState = structuredClone(this.state);
        result = await mutator(this.state);
        await this.persistTransaction(previousState, client);
        await client.query("COMMIT");
      } catch (error) {
        await client.query("ROLLBACK").catch(() => undefined);
        if (previousState) reconcileStoreState(this.state, previousState);
        throw error;
      } finally {
        client.release();
      }
    });
    this.mutationQueue = task.then(
      () => undefined,
      () => undefined,
    );
    await task;
    return result!;
  }

  protected async persist(previousState: StoreState = emptyState()): Promise<void> {
    await this.persistTransaction(previousState);
  }

  private async persistTransaction(
    previousState: StoreState,
    transactionClient?: PoolClient,
  ): Promise<void> {
    const client = transactionClient || (await this.pool.connect());
    const ownsTransaction = !transactionClient;
    const json = (value: unknown): string => JSON.stringify(value ?? null);
    const changed = <T>(
      current: T[],
      previous: T[],
      key: (item: T) => string | number,
    ): T[] => {
      const previousById = new Map(
        previous.map((item) => [key(item), JSON.stringify(item)]),
      );
      return current.filter(
        (item) => previousById.get(key(item)) !== JSON.stringify(item),
      );
    };
    const removed = <T>(
      current: T[],
      previous: T[],
      key: (item: T) => string | number,
    ): Array<string | number> => {
      const currentIds = new Set(current.map(key));
      return previous.map(key).filter((id) => !currentIds.has(id));
    };
    const currentPages = this.state.projects.flatMap((project) => project.pages);
    const previousPages = previousState.projects.flatMap(
      (project) => project.pages,
    );
    const currentVersions = currentPages.flatMap((page) => page.versions);
    const previousVersions = previousPages.flatMap((page) => page.versions);
    try {
      if (ownsTransaction) {
        await client.query("BEGIN");
        await client.query(
          "SELECT pg_advisory_xact_lock(hashtext('fastppt-online-store'))",
        );
      }
      const removedConflictIds = removed(
        this.state.documentConflicts,
        previousState.documentConflicts,
        (item) => item.conflictId,
      );
      if (removedConflictIds.length)
        await client.query(
          "DELETE FROM document_conflicts WHERE conflict_id = ANY($1::text[])",
          [removedConflictIds],
        );
      const removedEventSeqs = removed(
        this.state.events,
        previousState.events,
        (item) => item.seq,
      );
      if (removedEventSeqs.length)
        await client.query("DELETE FROM event_log WHERE seq = ANY($1::bigint[])", [
          removedEventSeqs,
        ]);
      for (const user of changed(
        this.state.users,
        previousState.users,
        (item) => item.userId,
      ))
        await client.query(
          "INSERT INTO users (user_id,email,name,created_at) VALUES ($1,$2,$3,$4) ON CONFLICT (user_id) DO UPDATE SET email=EXCLUDED.email,name=EXCLUDED.name",
          [user.userId, user.email, user.name, user.createdAt],
        );
      for (const session of changed(
        this.state.authSessions,
        previousState.authSessions,
        (item) => item.sessionId,
      ))
        await client.query(
          "INSERT INTO auth_sessions (session_id,user_id,expires_at,created_at,revoked_at) VALUES ($1,$2,$3,$4,$5) ON CONFLICT (session_id) DO UPDATE SET expires_at=EXCLUDED.expires_at,revoked_at=EXCLUDED.revoked_at",
          [
            session.sessionId,
            session.userId,
            session.expiresAt,
            session.createdAt,
            session.revokedAt,
          ],
        );
      for (const project of changed(
        this.state.projects,
        previousState.projects,
        (item) => item.projectId,
      ))
        await client.query(
          "INSERT INTO projects (project_id,owner_id,name,theme_id,theme_version,current_deck_revision_id,goal_id,status,source_markdown,sensitive_mode,created_at,updated_at) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12) ON CONFLICT (project_id) DO UPDATE SET owner_id=EXCLUDED.owner_id,name=EXCLUDED.name,theme_id=EXCLUDED.theme_id,theme_version=EXCLUDED.theme_version,current_deck_revision_id=EXCLUDED.current_deck_revision_id,goal_id=EXCLUDED.goal_id,status=EXCLUDED.status,source_markdown=EXCLUDED.source_markdown,sensitive_mode=EXCLUDED.sensitive_mode,updated_at=EXCLUDED.updated_at",
          [
            project.projectId,
            project.ownerId,
            project.name,
            project.themeId,
            project.themeVersion,
            project.currentDeckRevisionId,
            project.goalId,
            project.status,
            project.sourceMarkdown,
            project.settings?.sensitiveMode || false,
            project.createdAt,
            project.updatedAt,
          ],
        );
      for (const page of changed(
        currentPages,
        previousPages,
        (item) => item.pageId,
      )) {
          const { versions, ...pagePayload } = page;
          await client.query(
            "INSERT INTO pages (page_id,project_id,current_version_id,order_index,page_type,locked,archived,fact_anchor_ids,editable_level,status,payload) VALUES ($1,$2,$3,$4,$5,$6,$7,$8::jsonb,$9,$10,$11::jsonb) ON CONFLICT (page_id) DO UPDATE SET current_version_id=EXCLUDED.current_version_id,order_index=EXCLUDED.order_index,page_type=EXCLUDED.page_type,locked=EXCLUDED.locked,archived=EXCLUDED.archived,fact_anchor_ids=EXCLUDED.fact_anchor_ids,editable_level=EXCLUDED.editable_level,status=EXCLUDED.status,payload=EXCLUDED.payload",
            [
              page.pageId,
              page.projectId,
              page.currentVersionId,
              page.orderIndex,
              page.pageType,
              page.locked,
              page.archived || false,
              json(page.factAnchorIds),
              page.editableLevel,
              page.status,
              json(pagePayload),
            ],
          );
      }
      for (const version of changed(
        currentVersions,
        previousVersions,
        (item) => item.versionId,
      ))
        await client.query(
              "INSERT INTO page_versions (version_id,page_id,parent_version_id,source_revision,page_contract_path,slides_source_hash,preview_artifact_id,svg_artifact_id,pptx_page_render_id,prompt_snapshot_id,edit_operation_id,quality_report_id,status,created_at,payload) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15::jsonb) ON CONFLICT (version_id) DO UPDATE SET pptx_page_render_id=EXCLUDED.pptx_page_render_id,status=EXCLUDED.status,payload=EXCLUDED.payload",
              [
                version.versionId,
                version.pageId,
                version.parentVersionId,
                version.sourceRevision,
                version.pageContractPath,
                version.slidesSourceHash,
                version.previewArtifactId,
                version.svgArtifactId,
                version.pptxPageRenderId,
                version.promptSnapshotId,
                version.editOperationId,
                version.qualityReportId,
                version.status,
                version.createdAt,
                json(version),
              ],
            );
      for (const operation of changed(
        this.state.operations,
        previousState.operations,
        (item) => item.operationId,
      ))
        await client.query(
          "INSERT INTO edit_operations (operation_id,project_id,conversation_id,mode,requested_page_ids,resolved_page_ids,message,structured_plan,fact_impact,unsupported_items,confirmation_required,confirmed_at,result_version_ids,status,created_at,completed_at,payload) VALUES ($1,$2,$3,$4,$5::jsonb,$6::jsonb,$7,$8::jsonb,$9::jsonb,$10::jsonb,$11,$12,$13::jsonb,$14,$15,$16,$17::jsonb) ON CONFLICT (operation_id) DO UPDATE SET structured_plan=EXCLUDED.structured_plan,fact_impact=EXCLUDED.fact_impact,unsupported_items=EXCLUDED.unsupported_items,confirmation_required=EXCLUDED.confirmation_required,confirmed_at=EXCLUDED.confirmed_at,result_version_ids=EXCLUDED.result_version_ids,status=EXCLUDED.status,completed_at=EXCLUDED.completed_at,payload=EXCLUDED.payload",
          [
            operation.operationId,
            operation.projectId,
            operation.conversationId,
            operation.mode,
            json(operation.requestedPageIds),
            json(operation.resolvedPageIds),
            operation.message,
            json(operation.structuredPlan),
            json(operation.factImpact),
            json(operation.unsupportedItems),
            operation.confirmationRequired,
            operation.confirmedAt,
            json(operation.resultVersionIds),
            operation.status,
            operation.createdAt,
            operation.completedAt,
            json(operation),
          ],
        );
      for (const message of changed(
        this.state.messages,
        previousState.messages,
        (item) => item.messageId,
      ))
        await client.query(
          "INSERT INTO conversation_messages (message_id,project_id,conversation_id,role,text,payload,created_at) VALUES ($1,$2,$3,$4,$5,$6::jsonb,$7) ON CONFLICT (message_id) DO UPDATE SET text=EXCLUDED.text,payload=EXCLUDED.payload",
          [
            message.messageId,
            message.projectId,
            message.conversationId,
            message.role,
            message.text,
            json(message),
            message.createdAt,
          ],
        );
      for (const entry of changed(
        this.state.ledger,
        previousState.ledger,
        (item) => item.ledgerId,
      ))
        await client.query(
          "INSERT INTO usage_ledger (ledger_id,project_id,operation_id,page_id,version_id,model,unit_price,reserved_amount,settled_amount,refunded_amount,status,created_at) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12) ON CONFLICT (ledger_id) DO UPDATE SET version_id=EXCLUDED.version_id,settled_amount=EXCLUDED.settled_amount,refunded_amount=EXCLUDED.refunded_amount,status=EXCLUDED.status",
          [
            entry.ledgerId,
            entry.projectId,
            entry.operationId,
            entry.pageId,
            entry.versionId,
            entry.model,
            entry.unitPrice,
            entry.reservedAmount,
            entry.settledAmount,
            entry.refundedAmount,
            entry.status,
            entry.createdAt,
          ],
        );
      for (const job of changed(
        this.state.exports,
        previousState.exports,
        (item) => item.exportId,
      ))
        await client.query(
          "INSERT INTO export_jobs (export_id,project_id,status,artifact_path,artifact_object_key,render_mode,qa_warnings,artifact_name,created_at,completed_at) VALUES ($1,$2,$3,$4,$5,$6,$7::jsonb,$8,$9,$10) ON CONFLICT (export_id) DO UPDATE SET status=EXCLUDED.status,artifact_path=EXCLUDED.artifact_path,artifact_object_key=EXCLUDED.artifact_object_key,render_mode=EXCLUDED.render_mode,qa_warnings=EXCLUDED.qa_warnings,artifact_name=EXCLUDED.artifact_name,completed_at=EXCLUDED.completed_at",
          [
            job.exportId,
            job.projectId,
            job.status,
            job.artifactPath,
            job.artifactObjectKey || null,
            job.renderMode,
            json(job.qaWarnings),
            job.artifactName || null,
            job.createdAt,
            job.completedAt,
          ],
        );
      for (const artifact of changed(
        this.state.artifacts,
        previousState.artifacts,
        (item) => item.artifactId,
      ))
        await client.query(
          "INSERT INTO preview_artifacts (artifact_id,project_id,page_id,version_id,kind,prompt_snapshot_id,model,width,height,quality,source,provenance,created_at,asset_path) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12::jsonb,$13,$14) ON CONFLICT (artifact_id) DO UPDATE SET version_id=EXCLUDED.version_id,provenance=EXCLUDED.provenance,asset_path=EXCLUDED.asset_path",
          [
            artifact.artifactId,
            artifact.projectId,
            artifact.pageId,
            artifact.versionId,
            artifact.kind,
            artifact.promptSnapshotId,
            artifact.model,
            artifact.width,
            artifact.height,
            artifact.quality,
            artifact.source,
            json(artifact.provenance),
            artifact.createdAt,
            artifact.assetPath || null,
          ],
        );
      for (const snapshot of changed(
        this.state.promptSnapshots,
        previousState.promptSnapshots,
        (item) => item.promptSnapshotId,
      ))
        await client.query(
          "INSERT INTO prompt_snapshots (prompt_snapshot_id,project_id,page_id,operation_id,prompt,hash,model,created_at) VALUES ($1,$2,$3,$4,$5,$6,$7,$8) ON CONFLICT (prompt_snapshot_id) DO NOTHING",
          [
            snapshot.promptSnapshotId,
            snapshot.projectId,
            snapshot.pageId,
            snapshot.operationId,
            snapshot.prompt,
            snapshot.hash,
            snapshot.model,
            snapshot.createdAt,
          ],
        );
      for (const audit of changed(
        this.state.auditLogs,
        previousState.auditLogs,
        (item) => item.auditId,
      ))
        await client.query(
          "INSERT INTO audit_logs (audit_id,owner_id,project_id,action,request_id,payload,created_at) VALUES ($1,$2,$3,$4,$5,$6::jsonb,$7) ON CONFLICT (audit_id) DO NOTHING",
          [
            audit.auditId,
            audit.ownerId,
            audit.projectId || null,
            audit.action,
            audit.requestId || null,
            json(audit.payload),
            audit.createdAt,
          ],
        );
      for (const event of changed(
        this.state.events,
        previousState.events,
        (item) => item.seq,
      ))
        await client.query(
          "INSERT INTO event_log (seq,project_id,payload,created_at) VALUES ($1,$2,$3::jsonb,$4) ON CONFLICT (seq) DO UPDATE SET project_id=EXCLUDED.project_id,payload=EXCLUDED.payload,created_at=EXCLUDED.created_at",
          [event.seq, event.projectId, json(event), event.createdAt],
        );
      for (const job of changed(
        this.state.jobs,
        previousState.jobs,
        (item) => item.jobId,
      ))
        await client.query(
          "INSERT INTO jobs (job_id,idempotency_key,project_id,operation_id,kind,status,payload,attempt_count,max_attempts,available_at,lease_owner,lease_expires_at,last_error,started_at,created_at,updated_at,completed_at) VALUES ($1,$2,$3,$4,$5,$6,$7::jsonb,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17) ON CONFLICT (job_id) DO UPDATE SET status=EXCLUDED.status,payload=EXCLUDED.payload,attempt_count=EXCLUDED.attempt_count,max_attempts=EXCLUDED.max_attempts,available_at=EXCLUDED.available_at,lease_owner=EXCLUDED.lease_owner,lease_expires_at=EXCLUDED.lease_expires_at,last_error=EXCLUDED.last_error,started_at=EXCLUDED.started_at,updated_at=EXCLUDED.updated_at,completed_at=EXCLUDED.completed_at",
          [
            job.jobId,
            job.idempotencyKey,
            job.projectId,
            job.operationId || null,
            job.kind,
            job.status,
            json(job.payload),
            job.attemptCount,
            job.maxAttempts,
            job.availableAt,
            job.leaseOwner,
            job.leaseExpiresAt,
            job.lastError,
            job.startedAt,
            job.createdAt,
            job.updatedAt,
            job.completedAt,
          ],
        );
      for (const grant of changed(
        this.state.officePreviewGrants,
        previousState.officePreviewGrants,
        (item) => item.grantId,
      ))
        await client.query(
          "INSERT INTO office_preview_grants (grant_id,export_id,project_id,issued_to,token_hash,expires_at,max_fetches,fetch_count,revoked_at,created_at) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10) ON CONFLICT (grant_id) DO UPDATE SET fetch_count=EXCLUDED.fetch_count,revoked_at=EXCLUDED.revoked_at",
          [
            grant.grantId,
            grant.exportId,
            grant.projectId,
            grant.issuedTo,
            grant.tokenHash,
            grant.expiresAt,
            grant.maxFetches,
            grant.fetchCount,
            grant.revokedAt,
            grant.createdAt,
          ],
        );
      for (const session of changed(
        this.state.workSessions,
        previousState.workSessions,
        (item) => item.sessionId,
      ))
        await client.query(
          "INSERT INTO work_sessions (session_id,project_id,workflow_mode,intent,source_document_ids,source_revision,plan_id,page_budget,inherit_theme,structure_plan,created_by,status,created_at,updated_at) VALUES ($1,$2,$3,$4,$5::jsonb,$6,$7,$8,$9,$10::jsonb,$11,$12,$13,$14) ON CONFLICT (session_id) DO UPDATE SET workflow_mode=EXCLUDED.workflow_mode,intent=EXCLUDED.intent,source_document_ids=EXCLUDED.source_document_ids,source_revision=EXCLUDED.source_revision,plan_id=EXCLUDED.plan_id,page_budget=EXCLUDED.page_budget,inherit_theme=EXCLUDED.inherit_theme,structure_plan=EXCLUDED.structure_plan,created_by=EXCLUDED.created_by,status=EXCLUDED.status,updated_at=EXCLUDED.updated_at",
          [
            session.sessionId,
            session.projectId,
            session.workflowMode,
            session.intent,
            json(session.sourceDocumentIds),
            session.sourceRevision,
            session.planId,
            session.pageBudget,
            session.inheritTheme,
            json(session.structurePlan),
            session.createdBy,
            session.status,
            session.createdAt,
            session.updatedAt,
          ],
        );
      for (const document of changed(
        this.state.documents,
        previousState.documents,
        (item) => item.documentId,
      ))
        await client.query(
          "INSERT INTO document_sources (document_id,project_id,session_id,file_name,file_type,size_bytes,sha256,source_version_id,uploaded_by,object_key,parse_status,included_in_context,payload,created_at,updated_at) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13::jsonb,$14,$15) ON CONFLICT (document_id) DO UPDATE SET parse_status=EXCLUDED.parse_status,included_in_context=EXCLUDED.included_in_context,source_version_id=EXCLUDED.source_version_id,uploaded_by=EXCLUDED.uploaded_by,payload=EXCLUDED.payload,updated_at=EXCLUDED.updated_at",
          [
            document.documentId,
            document.projectId,
            document.sessionId,
            document.fileName,
            document.fileType,
            document.sizeBytes,
            document.sha256,
            document.sourceVersionId,
            document.uploadedBy,
            document.objectKey,
            document.parseStatus,
            document.includedInContext,
            json(document),
            document.createdAt,
            document.updatedAt,
          ],
        );
      for (const version of changed(
        this.state.documentVersions,
        previousState.documentVersions,
        (item) => item.sourceVersionId,
      ))
        await client.query(
          "INSERT INTO document_versions (source_version_id,document_id,project_id,sha256,size_bytes,media_type,object_key,created_by,created_at) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9) ON CONFLICT (source_version_id) DO NOTHING",
          [
            version.sourceVersionId,
            version.documentId,
            version.projectId,
            version.sha256,
            version.sizeBytes,
            version.mediaType,
            version.objectKey,
            version.createdBy,
            version.createdAt,
          ],
        );
      for (const conflict of changed(
        this.state.documentConflicts,
        previousState.documentConflicts,
        (item) => item.conflictId,
      ))
        await client.query(
          "INSERT INTO document_conflicts (conflict_id,project_id,session_id,conflict_key,severity,status,payload,created_at,resolved_at) VALUES ($1,$2,$3,$4,$5,$6,$7::jsonb,$8,$9) ON CONFLICT (conflict_id) DO UPDATE SET severity=EXCLUDED.severity,status=EXCLUDED.status,payload=EXCLUDED.payload,resolved_at=EXCLUDED.resolved_at",
          [
            conflict.conflictId,
            conflict.projectId,
            conflict.sessionId,
            conflict.key,
            conflict.severity,
            conflict.status,
            json(conflict),
            conflict.createdAt,
            conflict.resolvedAt,
          ],
        );
      for (const resolution of changed(
        this.state.conflictResolutions,
        previousState.conflictResolutions,
        (item) => item.resolutionId,
      ))
        await client.query(
          "INSERT INTO conflict_resolutions (resolution_id,conflict_id,project_id,session_id,source_revision,plan_id,action,selected_value,selected_document_ids,note,resolved_by,created_at) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9::jsonb,$10,$11,$12) ON CONFLICT (resolution_id) DO NOTHING",
          [
            resolution.resolutionId,
            resolution.conflictId,
            resolution.projectId,
            resolution.sessionId,
            resolution.sourceRevision,
            resolution.planId,
            resolution.action,
            resolution.selectedValue,
            json(resolution.selectedDocumentIds),
            resolution.note,
            resolution.resolvedBy,
            resolution.createdAt,
          ],
        );
      if (ownsTransaction) await client.query("COMMIT");
    } catch (error) {
      if (ownsTransaction) await client.query("ROLLBACK");
      throw error;
    } finally {
      if (ownsTransaction) client.release();
    }
  }

  private durableJobFromRow(row: any): DurableJob {
    const json = <T>(value: unknown, fallback: T): T =>
      typeof value === "string"
        ? (JSON.parse(value) as T)
        : ((value as T) ?? fallback);
    return {
      jobId: row.job_id,
      idempotencyKey: row.idempotency_key,
      projectId: row.project_id,
      operationId: row.operation_id,
      kind: row.kind,
      status: row.status,
      payload: json(row.payload, {}),
      attemptCount: Number(row.attempt_count || 0),
      maxAttempts: Number(row.max_attempts || 3),
      availableAt: new Date(row.available_at).toISOString(),
      leaseOwner: row.lease_owner || null,
      leaseExpiresAt: row.lease_expires_at
        ? new Date(row.lease_expires_at).toISOString()
        : null,
      lastError: row.last_error || null,
      startedAt: row.started_at
        ? new Date(row.started_at).toISOString()
        : null,
      createdAt: new Date(row.created_at).toISOString(),
      updatedAt: new Date(row.updated_at).toISOString(),
      completedAt: row.completed_at
        ? new Date(row.completed_at).toISOString()
        : null,
    };
  }

  private syncLocalJob(job: DurableJob): DurableJob {
    const existing = this.state.jobs.find(
      (candidate) => candidate.jobId === job.jobId,
    );
    if (existing) {
      Object.assign(existing, job);
      return existing;
    }
    this.state.jobs.push(job);
    return job;
  }

  private async serialized<T>(run: () => Promise<T>): Promise<T> {
    let result: T;
    const task = this.mutationQueue.then(async () => {
      result = await run();
    });
    this.mutationQueue = task.then(
      () => undefined,
      () => undefined,
    );
    await task;
    return result!;
  }

  override async enqueueJob(job: DurableJob): Promise<DurableJob> {
    return this.serialized(async () => {
      const result = await this.pool.query(
        `INSERT INTO jobs (job_id,idempotency_key,project_id,operation_id,kind,status,payload,attempt_count,max_attempts,available_at,lease_owner,lease_expires_at,last_error,started_at,created_at,updated_at,completed_at)
         VALUES ($1,$2,$3,$4,$5,'queued',$6::jsonb,$7,$8,$9,NULL,NULL,NULL,NULL,$10,$11,NULL)
         ON CONFLICT (idempotency_key) DO UPDATE SET idempotency_key=EXCLUDED.idempotency_key
         RETURNING *`,
        [
          job.jobId,
          job.idempotencyKey,
          job.projectId,
          job.operationId || null,
          job.kind,
          JSON.stringify(job.payload),
          job.attemptCount,
          job.maxAttempts,
          job.availableAt,
          job.createdAt,
          job.updatedAt,
        ],
      );
      return this.syncLocalJob(this.durableJobFromRow(result.rows[0]));
    });
  }

  override async claimJob(
    workerId: string,
    leaseMs: number,
  ): Promise<DurableJob | null> {
    return this.serialized(async () => {
      const result = await this.pool.query(
        `WITH candidate AS (
           SELECT job_id FROM jobs
           WHERE (status='queued' AND available_at <= NOW())
              OR (status='running' AND (lease_expires_at IS NULL OR lease_expires_at <= NOW()))
           ORDER BY available_at, created_at
           FOR UPDATE SKIP LOCKED
           LIMIT 1
         )
         UPDATE jobs AS job
         SET status='running', attempt_count=job.attempt_count+1,
             lease_owner=$1, lease_expires_at=NOW()+($2 * INTERVAL '1 millisecond'),
             started_at=COALESCE(job.started_at,NOW()), updated_at=NOW(), completed_at=NULL
         FROM candidate
         WHERE job.job_id=candidate.job_id
         RETURNING job.*`,
        [workerId, leaseMs],
      );
      if (!result.rows[0]) return null;
      return this.syncLocalJob(this.durableJobFromRow(result.rows[0]));
    });
  }

  override async heartbeatJob(
    jobId: string,
    workerId: string,
    leaseMs: number,
  ): Promise<boolean> {
    return this.serialized(async () => {
      const result = await this.pool.query(
        `UPDATE jobs SET lease_expires_at=NOW()+($3 * INTERVAL '1 millisecond'),updated_at=NOW()
         WHERE job_id=$1 AND status='running' AND lease_owner=$2 RETURNING *`,
        [jobId, workerId, leaseMs],
      );
      if (!result.rows[0]) return false;
      this.syncLocalJob(this.durableJobFromRow(result.rows[0]));
      return true;
    });
  }

  override async completeJob(
    jobId: string,
    workerId: string,
  ): Promise<DurableJob | null> {
    return this.serialized(async () => {
      const result = await this.pool.query(
        `UPDATE jobs SET status='completed',completed_at=NOW(),updated_at=NOW(),
           lease_owner=NULL,lease_expires_at=NULL,last_error=NULL
         WHERE job_id=$1 AND status='running' AND lease_owner=$2 RETURNING *`,
        [jobId, workerId],
      );
      if (!result.rows[0]) return null;
      return this.syncLocalJob(this.durableJobFromRow(result.rows[0]));
    });
  }

  override async failJob(
    jobId: string,
    workerId: string,
    error: string,
    retryDelayMs: number,
  ): Promise<DurableJob | null> {
    return this.serialized(async () => {
      const result = await this.pool.query(
        `UPDATE jobs
         SET status=CASE WHEN attempt_count < max_attempts THEN 'queued' ELSE 'failed' END,
             available_at=CASE WHEN attempt_count < max_attempts THEN NOW()+($4 * INTERVAL '1 millisecond') ELSE available_at END,
             completed_at=CASE WHEN attempt_count < max_attempts THEN NULL ELSE NOW() END,
             lease_owner=NULL,lease_expires_at=NULL,last_error=$3,updated_at=NOW()
         WHERE job_id=$1 AND status='running' AND lease_owner=$2 RETURNING *`,
        [jobId, workerId, error.slice(0, 4_000), retryDelayMs],
      );
      if (!result.rows[0]) return null;
      return this.syncLocalJob(this.durableJobFromRow(result.rows[0]));
    });
  }

  override async releaseExpiredJobLeases(): Promise<number> {
    return this.serialized(async () => {
      const result = await this.pool.query(
        `UPDATE jobs SET status='queued',available_at=NOW(),lease_owner=NULL,
           lease_expires_at=NULL,updated_at=NOW()
         WHERE status='running' AND (lease_expires_at IS NULL OR lease_expires_at <= NOW())
         RETURNING *`,
      );
      for (const row of result.rows)
        this.syncLocalJob(this.durableJobFromRow(row));
      return result.rows.length;
    });
  }

  async close(): Promise<void> {
    await this.pool.end();
  }
}

export async function createStore(): Promise<StateStore> {
  const databaseUrl = process.env.DATABASE_URL?.trim();
  if (process.env.NODE_ENV === "production" && !databaseUrl)
    throw new Error(
      "DATABASE_URL is required in production; file persistence is development-only.",
    );
  if (databaseUrl) {
    const postgres = new PostgresStore(databaseUrl);
    try {
      await postgres.init();
      return postgres;
    } catch (error) {
      await postgres.close().catch(() => undefined);
      if (process.env.NODE_ENV === "production") throw error;
      console.warn(
        `DATABASE_URL unavailable; falling back to file persistence: ${(error as Error).message}`,
      );
    }
  }
  const file = new FileStore(process.env.DATA_DIR || "./data");
  await file.init();
  return file;
}

import fs from 'node:fs/promises';
import path from 'node:path';
import { Pool } from 'pg';
import type { StoreState } from '../shared/models.js';

export interface StateStore {
  readonly kind: 'file' | 'postgres';
  state: StoreState;
  init(): Promise<void>;
  update<T>(mutator: (state: StoreState) => T | Promise<T>): Promise<T>;
}

export function emptyState(): StoreState {
  return { users: [], authSessions: [], projects: [], operations: [], messages: [], ledger: [], exports: [], artifacts: [], promptSnapshots: [], auditLogs: [], events: [], seq: 0, jobs: [] };
}

abstract class BaseStore implements StateStore {
  abstract readonly kind: 'file' | 'postgres';
  state: StoreState = emptyState();
  private mutationQueue = Promise.resolve();

  abstract init(): Promise<void>;
  protected abstract persist(): Promise<void>;

  async update<T>(mutator: (state: StoreState) => T | Promise<T>): Promise<T> {
    let result: T;
    const task = this.mutationQueue.then(async () => {
      result = await mutator(this.state);
      await this.persist();
    });
    this.mutationQueue = task.then(() => undefined, () => undefined);
    await task;
    return result!;
  }
}

export class FileStore extends BaseStore {
  readonly kind = 'file' as const;
  private readonly filePath: string;

  constructor(dataDir: string) {
    super();
    this.filePath = path.resolve(dataDir, 'store.json');
  }

  async init(): Promise<void> {
    await fs.mkdir(path.dirname(this.filePath), { recursive: true });
    try {
      const raw = await fs.readFile(this.filePath, 'utf8');
      this.state = { ...emptyState(), ...JSON.parse(raw) } as StoreState;
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error;
      await this.persist();
    }
  }

  protected async persist(): Promise<void> {
    const tempPath = `${this.filePath}.tmp`;
    await fs.writeFile(tempPath, JSON.stringify(this.state, null, 2), 'utf8');
    await fs.rename(tempPath, this.filePath);
  }
}

export class PostgresStore extends BaseStore {
  readonly kind = 'postgres' as const;
  private readonly pool: Pool;

  constructor(databaseUrl: string) {
    super();
    this.pool = new Pool({ connectionString: databaseUrl, max: 5 });
  }

  async init(): Promise<void> {
    await this.pool.query(`
      CREATE TABLE IF NOT EXISTS users (user_id TEXT PRIMARY KEY, email TEXT NOT NULL UNIQUE, name TEXT NOT NULL, created_at TIMESTAMPTZ NOT NULL);
      CREATE TABLE IF NOT EXISTS auth_sessions (session_id TEXT PRIMARY KEY, user_id TEXT NOT NULL REFERENCES users(user_id) ON DELETE CASCADE, expires_at TIMESTAMPTZ NOT NULL, created_at TIMESTAMPTZ NOT NULL, revoked_at TIMESTAMPTZ);
      CREATE TABLE IF NOT EXISTS projects (project_id TEXT PRIMARY KEY, owner_id TEXT NOT NULL, name TEXT NOT NULL, theme_id TEXT NOT NULL, theme_version TEXT NOT NULL, current_deck_revision_id TEXT NOT NULL, goal_id TEXT NOT NULL, status TEXT NOT NULL, source_markdown TEXT NOT NULL DEFAULT '', created_at TIMESTAMPTZ NOT NULL, updated_at TIMESTAMPTZ NOT NULL);
      CREATE TABLE IF NOT EXISTS pages (page_id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, current_version_id TEXT, order_index INTEGER NOT NULL, page_type TEXT NOT NULL, locked BOOLEAN NOT NULL DEFAULT FALSE, archived BOOLEAN NOT NULL DEFAULT FALSE, fact_anchor_ids JSONB NOT NULL DEFAULT '[]'::jsonb, editable_level TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'untouched', payload JSONB NOT NULL DEFAULT '{}'::jsonb);
      ALTER TABLE pages ADD COLUMN IF NOT EXISTS payload JSONB NOT NULL DEFAULT '{}'::jsonb;
      CREATE TABLE IF NOT EXISTS page_versions (version_id TEXT PRIMARY KEY, page_id TEXT NOT NULL REFERENCES pages(page_id) ON DELETE CASCADE, parent_version_id TEXT, source_revision TEXT NOT NULL, page_contract_path TEXT NOT NULL, slides_source_hash TEXT NOT NULL, preview_artifact_id TEXT, svg_artifact_id TEXT, pptx_page_render_id TEXT, prompt_snapshot_id TEXT NOT NULL, edit_operation_id TEXT NOT NULL, quality_report_id TEXT, status TEXT NOT NULL, created_at TIMESTAMPTZ NOT NULL, payload JSONB NOT NULL DEFAULT '{}'::jsonb);
      CREATE TABLE IF NOT EXISTS edit_operations (operation_id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, conversation_id TEXT NOT NULL, mode TEXT NOT NULL, requested_page_ids JSONB NOT NULL DEFAULT '[]'::jsonb, resolved_page_ids JSONB NOT NULL DEFAULT '[]'::jsonb, message TEXT NOT NULL, structured_plan JSONB NOT NULL, fact_impact JSONB NOT NULL, unsupported_items JSONB NOT NULL DEFAULT '[]'::jsonb, confirmation_required BOOLEAN NOT NULL DEFAULT FALSE, confirmed_at TIMESTAMPTZ, result_version_ids JSONB NOT NULL DEFAULT '[]'::jsonb, status TEXT NOT NULL, created_at TIMESTAMPTZ NOT NULL, completed_at TIMESTAMPTZ, payload JSONB NOT NULL DEFAULT '{}'::jsonb);
      ALTER TABLE edit_operations ADD COLUMN IF NOT EXISTS payload JSONB NOT NULL DEFAULT '{}'::jsonb;
      CREATE TABLE IF NOT EXISTS usage_ledger (ledger_id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, operation_id TEXT NOT NULL, page_id TEXT, version_id TEXT, model TEXT NOT NULL, unit_price NUMERIC(12, 6) NOT NULL DEFAULT 0, reserved_amount NUMERIC(12, 6) NOT NULL DEFAULT 0, settled_amount NUMERIC(12, 6) NOT NULL DEFAULT 0, refunded_amount NUMERIC(12, 6) NOT NULL DEFAULT 0, status TEXT NOT NULL, created_at TIMESTAMPTZ NOT NULL);
      CREATE TABLE IF NOT EXISTS preview_artifacts (artifact_id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, page_id TEXT NOT NULL REFERENCES pages(page_id) ON DELETE CASCADE, version_id TEXT NOT NULL, kind TEXT NOT NULL, prompt_snapshot_id TEXT NOT NULL, model TEXT NOT NULL, width INTEGER NOT NULL, height INTEGER NOT NULL, quality TEXT NOT NULL, source TEXT NOT NULL, provenance JSONB NOT NULL DEFAULT '{}'::jsonb, created_at TIMESTAMPTZ NOT NULL, asset_path TEXT);
      ALTER TABLE preview_artifacts ADD COLUMN IF NOT EXISTS asset_path TEXT;
      CREATE TABLE IF NOT EXISTS prompt_snapshots (prompt_snapshot_id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, page_id TEXT NOT NULL REFERENCES pages(page_id) ON DELETE CASCADE, operation_id TEXT NOT NULL, prompt TEXT NOT NULL, hash TEXT NOT NULL, model TEXT NOT NULL, created_at TIMESTAMPTZ NOT NULL);
      CREATE TABLE IF NOT EXISTS export_jobs (export_id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, status TEXT NOT NULL, artifact_path TEXT, artifact_object_key TEXT, render_mode TEXT NOT NULL, qa_warnings JSONB NOT NULL DEFAULT '[]'::jsonb, artifact_name TEXT, created_at TIMESTAMPTZ NOT NULL, completed_at TIMESTAMPTZ);
      ALTER TABLE export_jobs ADD COLUMN IF NOT EXISTS artifact_object_key TEXT;
      CREATE TABLE IF NOT EXISTS conversation_messages (message_id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, conversation_id TEXT NOT NULL, role TEXT NOT NULL, text TEXT NOT NULL, payload JSONB NOT NULL DEFAULT '{}'::jsonb, created_at TIMESTAMPTZ NOT NULL);
      CREATE TABLE IF NOT EXISTS event_log (seq BIGINT PRIMARY KEY, project_id TEXT NOT NULL, payload JSONB NOT NULL, created_at TIMESTAMPTZ NOT NULL);
      CREATE TABLE IF NOT EXISTS jobs (job_id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE, operation_id TEXT, kind TEXT NOT NULL, status TEXT NOT NULL, payload JSONB NOT NULL DEFAULT '{}'::jsonb, created_at TIMESTAMPTZ NOT NULL, completed_at TIMESTAMPTZ);
      CREATE TABLE IF NOT EXISTS audit_logs (audit_id TEXT PRIMARY KEY, owner_id TEXT NOT NULL, project_id TEXT, action TEXT NOT NULL, request_id TEXT, payload JSONB NOT NULL DEFAULT '{}'::jsonb, created_at TIMESTAMPTZ NOT NULL DEFAULT NOW());
    `);
    const [users, sessions, projects, pages, versions, operations, messages, ledger, exports, artifacts, prompts, audits, events, jobs] = await Promise.all([
      this.pool.query('SELECT * FROM users'), this.pool.query('SELECT * FROM auth_sessions'), this.pool.query('SELECT * FROM projects'), this.pool.query('SELECT * FROM pages'), this.pool.query('SELECT * FROM page_versions'),
      this.pool.query('SELECT * FROM edit_operations'), this.pool.query('SELECT * FROM conversation_messages'), this.pool.query('SELECT * FROM usage_ledger'), this.pool.query('SELECT * FROM export_jobs'),
      this.pool.query('SELECT * FROM preview_artifacts'), this.pool.query('SELECT * FROM prompt_snapshots'), this.pool.query('SELECT * FROM audit_logs'), this.pool.query('SELECT * FROM event_log ORDER BY seq'), this.pool.query('SELECT * FROM jobs'),
    ]);
    const json = <T>(value: unknown, fallback: T): T => typeof value === 'string' ? JSON.parse(value) as T : (value as T ?? fallback);
    const state = emptyState();
    state.users = users.rows.map((row: any) => ({ userId: row.user_id, email: row.email, name: row.name, createdAt: new Date(row.created_at).toISOString() }));
    state.authSessions = sessions.rows.map((row: any) => ({ sessionId: row.session_id, userId: row.user_id, expiresAt: new Date(row.expires_at).toISOString(), createdAt: new Date(row.created_at).toISOString(), revokedAt: row.revoked_at ? new Date(row.revoked_at).toISOString() : null }));
    state.projects = projects.rows.map((row: any) => ({ projectId: row.project_id, ownerId: row.owner_id, name: row.name, themeId: row.theme_id, themeVersion: row.theme_version, currentDeckRevisionId: row.current_deck_revision_id, goalId: row.goal_id, status: row.status, sourceMarkdown: row.source_markdown, createdAt: new Date(row.created_at).toISOString(), updatedAt: new Date(row.updated_at).toISOString(), pages: [] }));
    const pageMap = new Map<string, any>();
    for (const row of pages.rows as any[]) {
      const payload = json<Record<string, unknown>>(row.payload, {});
      const page: any = { ...payload, pageId: row.page_id, projectId: row.project_id, currentVersionId: row.current_version_id, orderIndex: row.order_index, pageType: row.page_type, locked: row.locked, archived: row.archived, factAnchorIds: json<string[]>(row.fact_anchor_ids, []), editableLevel: row.editable_level, status: row.status, versions: [] };
      pageMap.set(page.pageId, page);
      state.projects.find((project) => project.projectId === page.projectId)?.pages.push(page);
    }
    for (const row of versions.rows as any[]) {
      const payload = json<Record<string, unknown>>(row.payload, {});
      const version = { ...payload, versionId: row.version_id, pageId: row.page_id, parentVersionId: row.parent_version_id, sourceRevision: row.source_revision, pageContractPath: row.page_contract_path, slidesSourceHash: row.slides_source_hash, previewArtifactId: row.preview_artifact_id, svgArtifactId: row.svg_artifact_id, pptxPageRenderId: row.pptx_page_render_id, promptSnapshotId: row.prompt_snapshot_id, editOperationId: row.edit_operation_id, qualityReportId: row.quality_report_id, status: row.status, createdAt: new Date(row.created_at).toISOString() };
      pageMap.get(row.page_id)?.versions.push(version);
    }
    state.operations = operations.rows.map((row: any) => ({ ...json<Record<string, unknown>>(row.payload, {}), operationId: row.operation_id, projectId: row.project_id, conversationId: row.conversation_id, mode: row.mode, requestedPageIds: json<string[]>(row.requested_page_ids, []), resolvedPageIds: json<string[]>(row.resolved_page_ids, []), message: row.message, structuredPlan: json(row.structured_plan, {}), factImpact: json(row.fact_impact, {}), unsupportedItems: json<string[]>(row.unsupported_items, []), confirmationRequired: row.confirmation_required, confirmedAt: row.confirmed_at ? new Date(row.confirmed_at).toISOString() : null, resultVersionIds: json<string[]>(row.result_version_ids, []), status: row.status, createdAt: new Date(row.created_at).toISOString(), completedAt: row.completed_at ? new Date(row.completed_at).toISOString() : null } as any));
    state.messages = messages.rows.map((row: any) => ({ ...json<Record<string, unknown>>(row.payload, {}), messageId: row.message_id, projectId: row.project_id, conversationId: row.conversation_id, role: row.role, text: row.text, createdAt: new Date(row.created_at).toISOString() } as any));
    state.ledger = ledger.rows.map((row: any) => ({ ledgerId: row.ledger_id, projectId: row.project_id, operationId: row.operation_id, pageId: row.page_id, versionId: row.version_id, model: row.model, unitPrice: Number(row.unit_price), reservedAmount: Number(row.reserved_amount), settledAmount: Number(row.settled_amount), refundedAmount: Number(row.refunded_amount), status: row.status, createdAt: new Date(row.created_at).toISOString() }));
    state.exports = exports.rows.map((row: any) => ({ exportId: row.export_id, projectId: row.project_id, status: row.status, artifactPath: row.artifact_path, artifactObjectKey: row.artifact_object_key, renderMode: row.render_mode, qaWarnings: json<string[]>(row.qa_warnings, []), artifactName: row.artifact_name, createdAt: new Date(row.created_at).toISOString(), completedAt: row.completed_at ? new Date(row.completed_at).toISOString() : null }));
    state.artifacts = artifacts.rows.map((row: any) => ({ artifactId: row.artifact_id, projectId: row.project_id, pageId: row.page_id, versionId: row.version_id, kind: row.kind, promptSnapshotId: row.prompt_snapshot_id, model: row.model, width: row.width, height: row.height, quality: row.quality, source: row.source, provenance: json(row.provenance, {}), createdAt: new Date(row.created_at).toISOString(), assetPath: row.asset_path }));
    state.promptSnapshots = prompts.rows.map((row: any) => ({ promptSnapshotId: row.prompt_snapshot_id, projectId: row.project_id, pageId: row.page_id, operationId: row.operation_id, prompt: row.prompt, hash: row.hash, model: row.model, createdAt: new Date(row.created_at).toISOString() }));
    state.auditLogs = audits.rows.map((row: any) => ({ auditId: String(row.audit_id), ownerId: row.owner_id, projectId: row.project_id || undefined, action: row.action, requestId: row.request_id || undefined, payload: json(row.payload, {}), createdAt: new Date(row.created_at).toISOString() }));
    state.events = events.rows.map((row: any) => json(row.payload, {})) as any[];
    state.seq = state.events.reduce((max, event: any) => Math.max(max, Number(event.seq || 0)), 0);
    state.jobs = jobs.rows.map((row: any) => ({ jobId: row.job_id, projectId: row.project_id, operationId: row.operation_id, kind: row.kind, status: row.status, payload: json(row.payload, {}), createdAt: new Date(row.created_at).toISOString(), completedAt: row.completed_at ? new Date(row.completed_at).toISOString() : null }));
    this.state = state;
    if (!projects.rows.length) await this.persist();
  }

  protected async persist(): Promise<void> {
    const client = await this.pool.connect();
    const json = (value: unknown): string => JSON.stringify(value ?? null);
    try {
      await client.query('BEGIN');
      await client.query("SELECT pg_advisory_xact_lock(hashtext('fastppt-online-store'))");
      for (const user of this.state.users) await client.query('INSERT INTO users (user_id,email,name,created_at) VALUES ($1,$2,$3,$4) ON CONFLICT (user_id) DO UPDATE SET email=EXCLUDED.email,name=EXCLUDED.name', [user.userId, user.email, user.name, user.createdAt]);
      for (const session of this.state.authSessions) await client.query('INSERT INTO auth_sessions (session_id,user_id,expires_at,created_at,revoked_at) VALUES ($1,$2,$3,$4,$5) ON CONFLICT (session_id) DO UPDATE SET expires_at=EXCLUDED.expires_at,revoked_at=EXCLUDED.revoked_at', [session.sessionId, session.userId, session.expiresAt, session.createdAt, session.revokedAt]);
      for (const project of this.state.projects) await client.query('INSERT INTO projects (project_id,owner_id,name,theme_id,theme_version,current_deck_revision_id,goal_id,status,source_markdown,created_at,updated_at) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11) ON CONFLICT (project_id) DO UPDATE SET owner_id=EXCLUDED.owner_id,name=EXCLUDED.name,theme_id=EXCLUDED.theme_id,theme_version=EXCLUDED.theme_version,current_deck_revision_id=EXCLUDED.current_deck_revision_id,goal_id=EXCLUDED.goal_id,status=EXCLUDED.status,source_markdown=EXCLUDED.source_markdown,updated_at=EXCLUDED.updated_at', [project.projectId, project.ownerId, project.name, project.themeId, project.themeVersion, project.currentDeckRevisionId, project.goalId, project.status, project.sourceMarkdown, project.createdAt, project.updatedAt]);
      for (const project of this.state.projects) for (const page of project.pages) {
        const { versions, ...pagePayload } = page;
        await client.query('INSERT INTO pages (page_id,project_id,current_version_id,order_index,page_type,locked,archived,fact_anchor_ids,editable_level,status,payload) VALUES ($1,$2,$3,$4,$5,$6,$7,$8::jsonb,$9,$10,$11::jsonb) ON CONFLICT (page_id) DO UPDATE SET current_version_id=EXCLUDED.current_version_id,order_index=EXCLUDED.order_index,page_type=EXCLUDED.page_type,locked=EXCLUDED.locked,archived=EXCLUDED.archived,fact_anchor_ids=EXCLUDED.fact_anchor_ids,editable_level=EXCLUDED.editable_level,status=EXCLUDED.status,payload=EXCLUDED.payload', [page.pageId, page.projectId, page.currentVersionId, page.orderIndex, page.pageType, page.locked, page.archived || false, json(page.factAnchorIds), page.editableLevel, page.status, json(pagePayload)]);
        for (const version of versions) await client.query('INSERT INTO page_versions (version_id,page_id,parent_version_id,source_revision,page_contract_path,slides_source_hash,preview_artifact_id,svg_artifact_id,pptx_page_render_id,prompt_snapshot_id,edit_operation_id,quality_report_id,status,created_at,payload) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15::jsonb) ON CONFLICT (version_id) DO UPDATE SET pptx_page_render_id=EXCLUDED.pptx_page_render_id,status=EXCLUDED.status,payload=EXCLUDED.payload', [version.versionId, version.pageId, version.parentVersionId, version.sourceRevision, version.pageContractPath, version.slidesSourceHash, version.previewArtifactId, version.svgArtifactId, version.pptxPageRenderId, version.promptSnapshotId, version.editOperationId, version.qualityReportId, version.status, version.createdAt, json(version)]);
      }
      for (const operation of this.state.operations) await client.query('INSERT INTO edit_operations (operation_id,project_id,conversation_id,mode,requested_page_ids,resolved_page_ids,message,structured_plan,fact_impact,unsupported_items,confirmation_required,confirmed_at,result_version_ids,status,created_at,completed_at,payload) VALUES ($1,$2,$3,$4,$5::jsonb,$6::jsonb,$7,$8::jsonb,$9::jsonb,$10::jsonb,$11,$12,$13::jsonb,$14,$15,$16,$17::jsonb) ON CONFLICT (operation_id) DO UPDATE SET structured_plan=EXCLUDED.structured_plan,fact_impact=EXCLUDED.fact_impact,unsupported_items=EXCLUDED.unsupported_items,confirmation_required=EXCLUDED.confirmation_required,confirmed_at=EXCLUDED.confirmed_at,result_version_ids=EXCLUDED.result_version_ids,status=EXCLUDED.status,completed_at=EXCLUDED.completed_at,payload=EXCLUDED.payload', [operation.operationId, operation.projectId, operation.conversationId, operation.mode, json(operation.requestedPageIds), json(operation.resolvedPageIds), operation.message, json(operation.structuredPlan), json(operation.factImpact), json(operation.unsupportedItems), operation.confirmationRequired, operation.confirmedAt, json(operation.resultVersionIds), operation.status, operation.createdAt, operation.completedAt, json(operation)]);
      for (const message of this.state.messages) await client.query('INSERT INTO conversation_messages (message_id,project_id,conversation_id,role,text,payload,created_at) VALUES ($1,$2,$3,$4,$5,$6::jsonb,$7) ON CONFLICT (message_id) DO UPDATE SET text=EXCLUDED.text,payload=EXCLUDED.payload', [message.messageId, message.projectId, message.conversationId, message.role, message.text, json(message), message.createdAt]);
      for (const entry of this.state.ledger) await client.query('INSERT INTO usage_ledger (ledger_id,project_id,operation_id,page_id,version_id,model,unit_price,reserved_amount,settled_amount,refunded_amount,status,created_at) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12) ON CONFLICT (ledger_id) DO UPDATE SET version_id=EXCLUDED.version_id,settled_amount=EXCLUDED.settled_amount,refunded_amount=EXCLUDED.refunded_amount,status=EXCLUDED.status', [entry.ledgerId, entry.projectId, entry.operationId, entry.pageId, entry.versionId, entry.model, entry.unitPrice, entry.reservedAmount, entry.settledAmount, entry.refundedAmount, entry.status, entry.createdAt]);
      for (const job of this.state.exports) await client.query('INSERT INTO export_jobs (export_id,project_id,status,artifact_path,artifact_object_key,render_mode,qa_warnings,artifact_name,created_at,completed_at) VALUES ($1,$2,$3,$4,$5,$6,$7::jsonb,$8,$9,$10) ON CONFLICT (export_id) DO UPDATE SET status=EXCLUDED.status,artifact_path=EXCLUDED.artifact_path,artifact_object_key=EXCLUDED.artifact_object_key,render_mode=EXCLUDED.render_mode,qa_warnings=EXCLUDED.qa_warnings,artifact_name=EXCLUDED.artifact_name,completed_at=EXCLUDED.completed_at', [job.exportId, job.projectId, job.status, job.artifactPath, job.artifactObjectKey || null, job.renderMode, json(job.qaWarnings), job.artifactName || null, job.createdAt, job.completedAt]);
      for (const artifact of this.state.artifacts) await client.query('INSERT INTO preview_artifacts (artifact_id,project_id,page_id,version_id,kind,prompt_snapshot_id,model,width,height,quality,source,provenance,created_at,asset_path) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12::jsonb,$13,$14) ON CONFLICT (artifact_id) DO UPDATE SET provenance=EXCLUDED.provenance,asset_path=EXCLUDED.asset_path', [artifact.artifactId, artifact.projectId, artifact.pageId, artifact.versionId, artifact.kind, artifact.promptSnapshotId, artifact.model, artifact.width, artifact.height, artifact.quality, artifact.source, json(artifact.provenance), artifact.createdAt, artifact.assetPath || null]);
      for (const snapshot of this.state.promptSnapshots) await client.query('INSERT INTO prompt_snapshots (prompt_snapshot_id,project_id,page_id,operation_id,prompt,hash,model,created_at) VALUES ($1,$2,$3,$4,$5,$6,$7,$8) ON CONFLICT (prompt_snapshot_id) DO NOTHING', [snapshot.promptSnapshotId, snapshot.projectId, snapshot.pageId, snapshot.operationId, snapshot.prompt, snapshot.hash, snapshot.model, snapshot.createdAt]);
      for (const audit of this.state.auditLogs) await client.query('INSERT INTO audit_logs (audit_id,owner_id,project_id,action,request_id,payload,created_at) VALUES ($1,$2,$3,$4,$5,$6::jsonb,$7) ON CONFLICT (audit_id) DO NOTHING', [audit.auditId, audit.ownerId, audit.projectId || null, audit.action, audit.requestId || null, json(audit.payload), audit.createdAt]);
      for (const event of this.state.events) await client.query('INSERT INTO event_log (seq,project_id,payload,created_at) VALUES ($1,$2,$3::jsonb,$4) ON CONFLICT (seq) DO UPDATE SET project_id=EXCLUDED.project_id,payload=EXCLUDED.payload,created_at=EXCLUDED.created_at', [event.seq, event.projectId, json(event), event.createdAt]);
      for (const job of this.state.jobs) await client.query('INSERT INTO jobs (job_id,project_id,operation_id,kind,status,payload,created_at,completed_at) VALUES ($1,$2,$3,$4,$5,$6::jsonb,$7,$8) ON CONFLICT (job_id) DO UPDATE SET status=EXCLUDED.status,payload=EXCLUDED.payload,completed_at=EXCLUDED.completed_at', [job.jobId, job.projectId, job.operationId || null, job.kind, job.status, json(job.payload), job.createdAt, job.completedAt]);
      await client.query('COMMIT');
    } catch (error) {
      await client.query('ROLLBACK');
      throw error;
    } finally {
      client.release();
    }
  }

  async close(): Promise<void> {
    await this.pool.end();
  }
}

export async function createStore(): Promise<StateStore> {
  const databaseUrl = process.env.DATABASE_URL?.trim();
  if (process.env.NODE_ENV === 'production' && !databaseUrl) throw new Error('DATABASE_URL is required in production; file persistence is development-only.');
  if (databaseUrl) {
    const postgres = new PostgresStore(databaseUrl);
    try {
      await postgres.init();
      return postgres;
    } catch (error) {
      await postgres.close().catch(() => undefined);
      if (process.env.NODE_ENV === 'production') throw error;
      console.warn(`DATABASE_URL unavailable; falling back to file persistence: ${(error as Error).message}`);
    }
  }
  const file = new FileStore(process.env.DATA_DIR || './data');
  await file.init();
  return file;
}

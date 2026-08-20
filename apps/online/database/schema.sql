-- FastPPT Online persistence contract.
-- The local file store mirrors these entities for zero-dependency development;
-- PostgreSQL deployments use this schema before enabling DATABASE_URL.

CREATE TABLE IF NOT EXISTS projects (
  project_id TEXT PRIMARY KEY,
  owner_id TEXT NOT NULL,
  name TEXT NOT NULL,
  theme_id TEXT NOT NULL,
  theme_version TEXT NOT NULL,
  current_deck_revision_id TEXT NOT NULL,
  goal_id TEXT NOT NULL,
  status TEXT NOT NULL CHECK (status IN ('draft', 'processing', 'ready', 'failed', 'archived')),
  source_markdown TEXT NOT NULL DEFAULT '',
  sensitive_mode BOOLEAN NOT NULL DEFAULT FALSE,
  created_at TIMESTAMPTZ NOT NULL,
  updated_at TIMESTAMPTZ NOT NULL
);

ALTER TABLE projects ADD COLUMN IF NOT EXISTS sensitive_mode BOOLEAN NOT NULL DEFAULT FALSE;

CREATE TABLE IF NOT EXISTS users (
  user_id TEXT PRIMARY KEY,
  email TEXT NOT NULL UNIQUE,
  name TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL
);

CREATE TABLE IF NOT EXISTS auth_sessions (
  session_id TEXT PRIMARY KEY,
  user_id TEXT NOT NULL REFERENCES users(user_id) ON DELETE CASCADE,
  expires_at TIMESTAMPTZ NOT NULL,
  created_at TIMESTAMPTZ NOT NULL,
  revoked_at TIMESTAMPTZ
);

CREATE TABLE IF NOT EXISTS pages (
  page_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
  current_version_id TEXT,
  order_index INTEGER NOT NULL,
  page_type TEXT NOT NULL,
  locked BOOLEAN NOT NULL DEFAULT FALSE,
  archived BOOLEAN NOT NULL DEFAULT FALSE,
  fact_anchor_ids JSONB NOT NULL DEFAULT '[]'::jsonb,
  editable_level TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'untouched',
  payload JSONB NOT NULL DEFAULT '{}'::jsonb
);

CREATE TABLE IF NOT EXISTS page_versions (
  version_id TEXT PRIMARY KEY,
  page_id TEXT NOT NULL REFERENCES pages(page_id) ON DELETE CASCADE,
  parent_version_id TEXT,
  source_revision TEXT NOT NULL,
  page_contract_path TEXT NOT NULL,
  slides_source_hash TEXT NOT NULL,
  preview_artifact_id TEXT,
  svg_artifact_id TEXT,
  pptx_page_render_id TEXT,
  prompt_snapshot_id TEXT NOT NULL,
  edit_operation_id TEXT NOT NULL,
  quality_report_id TEXT,
  status TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL,
  payload JSONB NOT NULL DEFAULT '{}'::jsonb
);

CREATE TABLE IF NOT EXISTS edit_operations (
  operation_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
  conversation_id TEXT NOT NULL,
  mode TEXT NOT NULL,
  requested_page_ids JSONB NOT NULL DEFAULT '[]'::jsonb,
  resolved_page_ids JSONB NOT NULL DEFAULT '[]'::jsonb,
  message TEXT NOT NULL,
  structured_plan JSONB NOT NULL,
  fact_impact JSONB NOT NULL,
  unsupported_items JSONB NOT NULL DEFAULT '[]'::jsonb,
  confirmation_required BOOLEAN NOT NULL DEFAULT FALSE,
  confirmed_at TIMESTAMPTZ,
  result_version_ids JSONB NOT NULL DEFAULT '[]'::jsonb,
  status TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL,
  completed_at TIMESTAMPTZ,
  payload JSONB NOT NULL DEFAULT '{}'::jsonb
);

CREATE TABLE IF NOT EXISTS usage_ledger (
  ledger_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
  operation_id TEXT NOT NULL,
  page_id TEXT,
  version_id TEXT,
  model TEXT NOT NULL,
  unit_price NUMERIC(12, 6) NOT NULL DEFAULT 0,
  reserved_amount NUMERIC(12, 6) NOT NULL DEFAULT 0,
  settled_amount NUMERIC(12, 6) NOT NULL DEFAULT 0,
  refunded_amount NUMERIC(12, 6) NOT NULL DEFAULT 0,
  status TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL
);

CREATE TABLE IF NOT EXISTS assets (
  artifact_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
  page_id TEXT,
  kind TEXT NOT NULL,
  object_key TEXT NOT NULL,
  provenance JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at TIMESTAMPTZ NOT NULL
);

CREATE TABLE IF NOT EXISTS preview_artifacts (
  artifact_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
  page_id TEXT NOT NULL REFERENCES pages(page_id) ON DELETE CASCADE,
  version_id TEXT,
  kind TEXT NOT NULL,
  prompt_snapshot_id TEXT NOT NULL,
  model TEXT NOT NULL,
  width INTEGER NOT NULL,
  height INTEGER NOT NULL,
  quality TEXT NOT NULL,
  source TEXT NOT NULL,
  provenance JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at TIMESTAMPTZ NOT NULL
);

ALTER TABLE preview_artifacts ALTER COLUMN version_id DROP NOT NULL;

CREATE TABLE IF NOT EXISTS export_jobs (
  export_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
  status TEXT NOT NULL,
  artifact_path TEXT,
  artifact_object_key TEXT,
  render_mode TEXT NOT NULL,
  qa_warnings JSONB NOT NULL DEFAULT '[]'::jsonb,
  artifact_name TEXT,
  created_at TIMESTAMPTZ NOT NULL,
  completed_at TIMESTAMPTZ
);

CREATE TABLE IF NOT EXISTS conversation_messages (
  message_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
  conversation_id TEXT NOT NULL,
  role TEXT NOT NULL,
  text TEXT NOT NULL,
  payload JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at TIMESTAMPTZ NOT NULL
);

CREATE TABLE IF NOT EXISTS event_log (
  seq BIGINT PRIMARY KEY,
  project_id TEXT NOT NULL,
  payload JSONB NOT NULL,
  created_at TIMESTAMPTZ NOT NULL
);

CREATE TABLE IF NOT EXISTS prompt_snapshots (
  prompt_snapshot_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
  page_id TEXT NOT NULL REFERENCES pages(page_id) ON DELETE CASCADE,
  operation_id TEXT NOT NULL,
  prompt TEXT NOT NULL,
  hash TEXT NOT NULL,
  model TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL
);

CREATE TABLE IF NOT EXISTS jobs (
  job_id TEXT PRIMARY KEY,
  idempotency_key TEXT NOT NULL,
  project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
  operation_id TEXT,
  kind TEXT NOT NULL,
  status TEXT NOT NULL,
  payload JSONB NOT NULL DEFAULT '{}'::jsonb,
  attempt_count INTEGER NOT NULL DEFAULT 0,
  max_attempts INTEGER NOT NULL DEFAULT 3,
  available_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  lease_owner TEXT,
  lease_expires_at TIMESTAMPTZ,
  last_error TEXT,
  started_at TIMESTAMPTZ,
  created_at TIMESTAMPTZ NOT NULL,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  completed_at TIMESTAMPTZ
);

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

CREATE TABLE IF NOT EXISTS audit_logs (
  audit_id TEXT PRIMARY KEY,
  owner_id TEXT NOT NULL,
  project_id TEXT,
  action TEXT NOT NULL,
  request_id TEXT,
  payload JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS office_preview_grants (
  grant_id TEXT PRIMARY KEY,
  export_id TEXT NOT NULL,
  project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
  issued_to TEXT NOT NULL,
  token_hash TEXT NOT NULL UNIQUE,
  expires_at TIMESTAMPTZ NOT NULL,
  max_fetches INTEGER NOT NULL,
  fetch_count INTEGER NOT NULL DEFAULT 0,
  revoked_at TIMESTAMPTZ,
  created_at TIMESTAMPTZ NOT NULL
);

CREATE TABLE IF NOT EXISTS work_sessions (
  session_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
  workflow_mode TEXT NOT NULL,
  intent TEXT NOT NULL,
  source_document_ids JSONB NOT NULL DEFAULT '[]'::jsonb,
  source_revision TEXT NOT NULL DEFAULT '',
  plan_id TEXT,
  page_budget INTEGER NOT NULL DEFAULT 8,
  inherit_theme BOOLEAN NOT NULL DEFAULT TRUE,
  structure_plan JSONB,
  created_by TEXT NOT NULL DEFAULT '',
  status TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL,
  updated_at TIMESTAMPTZ NOT NULL
);

ALTER TABLE work_sessions ADD COLUMN IF NOT EXISTS source_document_ids JSONB NOT NULL DEFAULT '[]'::jsonb;
ALTER TABLE work_sessions ADD COLUMN IF NOT EXISTS source_revision TEXT NOT NULL DEFAULT '';
ALTER TABLE work_sessions ADD COLUMN IF NOT EXISTS plan_id TEXT;
ALTER TABLE work_sessions ADD COLUMN IF NOT EXISTS page_budget INTEGER NOT NULL DEFAULT 8;
ALTER TABLE work_sessions ADD COLUMN IF NOT EXISTS inherit_theme BOOLEAN NOT NULL DEFAULT TRUE;
ALTER TABLE work_sessions ADD COLUMN IF NOT EXISTS structure_plan JSONB;
ALTER TABLE work_sessions ADD COLUMN IF NOT EXISTS created_by TEXT NOT NULL DEFAULT '';

CREATE TABLE IF NOT EXISTS document_sources (
  document_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
  session_id TEXT NOT NULL REFERENCES work_sessions(session_id) ON DELETE CASCADE,
  file_name TEXT NOT NULL,
  file_type TEXT NOT NULL,
  size_bytes BIGINT NOT NULL,
  sha256 TEXT NOT NULL,
  source_version_id TEXT NOT NULL,
  uploaded_by TEXT NOT NULL,
  object_key TEXT NOT NULL,
  parse_status TEXT NOT NULL,
  included_in_context BOOLEAN NOT NULL DEFAULT TRUE,
  payload JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at TIMESTAMPTZ NOT NULL,
  updated_at TIMESTAMPTZ NOT NULL,
  UNIQUE(project_id, sha256)
);

CREATE TABLE IF NOT EXISTS document_conflicts (
  conflict_id TEXT PRIMARY KEY,
  project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
  session_id TEXT NOT NULL REFERENCES work_sessions(session_id) ON DELETE CASCADE,
  conflict_key TEXT NOT NULL,
  severity TEXT NOT NULL,
  status TEXT NOT NULL,
  payload JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at TIMESTAMPTZ NOT NULL,
  resolved_at TIMESTAMPTZ
);

CREATE TABLE IF NOT EXISTS document_versions (
  source_version_id TEXT PRIMARY KEY,
  document_id TEXT NOT NULL REFERENCES document_sources(document_id) ON DELETE CASCADE,
  project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
  sha256 TEXT NOT NULL,
  size_bytes BIGINT NOT NULL,
  media_type TEXT NOT NULL,
  object_key TEXT NOT NULL,
  created_by TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL
);

CREATE TABLE IF NOT EXISTS conflict_resolutions (
  resolution_id TEXT PRIMARY KEY,
  conflict_id TEXT NOT NULL REFERENCES document_conflicts(conflict_id) ON DELETE CASCADE,
  project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
  session_id TEXT NOT NULL REFERENCES work_sessions(session_id) ON DELETE CASCADE,
  source_revision TEXT NOT NULL DEFAULT '',
  plan_id TEXT,
  action TEXT NOT NULL,
  selected_value TEXT,
  selected_document_ids JSONB NOT NULL DEFAULT '[]'::jsonb,
  note TEXT NOT NULL DEFAULT '',
  resolved_by TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL
);

ALTER TABLE conflict_resolutions ADD COLUMN IF NOT EXISTS source_revision TEXT NOT NULL DEFAULT '';
ALTER TABLE conflict_resolutions ADD COLUMN IF NOT EXISTS plan_id TEXT;

ALTER TABLE document_sources ADD COLUMN IF NOT EXISTS source_version_id TEXT NOT NULL DEFAULT '';
ALTER TABLE document_sources ADD COLUMN IF NOT EXISTS uploaded_by TEXT NOT NULL DEFAULT '';

CREATE INDEX IF NOT EXISTS pages_project_order_idx ON pages(project_id, order_index);
CREATE INDEX IF NOT EXISTS versions_page_created_idx ON page_versions(page_id, created_at DESC);
CREATE INDEX IF NOT EXISTS ledger_project_idx ON usage_ledger(project_id, created_at DESC);
CREATE INDEX IF NOT EXISTS artifacts_page_created_idx ON preview_artifacts(page_id, created_at DESC);
CREATE INDEX IF NOT EXISTS prompts_project_created_idx ON prompt_snapshots(project_id, created_at DESC);
CREATE INDEX IF NOT EXISTS office_grants_project_created_idx ON office_preview_grants(project_id, created_at DESC);
CREATE INDEX IF NOT EXISTS work_sessions_project_updated_idx ON work_sessions(project_id, updated_at DESC);
CREATE INDEX IF NOT EXISTS documents_project_session_idx ON document_sources(project_id, session_id, created_at DESC);
CREATE INDEX IF NOT EXISTS document_versions_document_created_idx ON document_versions(document_id, created_at DESC);
CREATE INDEX IF NOT EXISTS conflicts_project_session_idx ON document_conflicts(project_id, session_id, status);
CREATE INDEX IF NOT EXISTS conflict_resolutions_session_created_idx ON conflict_resolutions(session_id, created_at DESC);
CREATE UNIQUE INDEX IF NOT EXISTS jobs_idempotency_key_idx ON jobs(idempotency_key);
CREATE INDEX IF NOT EXISTS jobs_claim_idx ON jobs(status, available_at, lease_expires_at);

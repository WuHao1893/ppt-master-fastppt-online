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
  created_at TIMESTAMPTZ NOT NULL,
  updated_at TIMESTAMPTZ NOT NULL
);

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
  version_id TEXT NOT NULL,
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
  project_id TEXT NOT NULL REFERENCES projects(project_id) ON DELETE CASCADE,
  operation_id TEXT,
  kind TEXT NOT NULL,
  status TEXT NOT NULL,
  payload JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at TIMESTAMPTZ NOT NULL,
  completed_at TIMESTAMPTZ
);

CREATE TABLE IF NOT EXISTS audit_logs (
  audit_id TEXT PRIMARY KEY,
  owner_id TEXT NOT NULL,
  project_id TEXT,
  action TEXT NOT NULL,
  request_id TEXT,
  payload JSONB NOT NULL DEFAULT '{}'::jsonb,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS pages_project_order_idx ON pages(project_id, order_index);
CREATE INDEX IF NOT EXISTS versions_page_created_idx ON page_versions(page_id, created_at DESC);
CREATE INDEX IF NOT EXISTS ledger_project_idx ON usage_ledger(project_id, created_at DESC);
CREATE INDEX IF NOT EXISTS artifacts_page_created_idx ON preview_artifacts(page_id, created_at DESC);
CREATE INDEX IF NOT EXISTS prompts_project_created_idx ON prompt_snapshots(project_id, created_at DESC);

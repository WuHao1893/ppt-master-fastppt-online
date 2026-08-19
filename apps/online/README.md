# FastPPT Online

This directory is the browser-based online workbench described in
`FastPPT_在线聊天精修_spec.md`. The browser sends business requests to the
server, while the server owns contracts, facts, prompts, cost reservations,
render status, and export artifacts. The export worker consumes the retained
official `skills/ppt-master` distribution to rebuild each page from validated
SVG into native DrawingML.

## Run locally

From this directory:

```powershell
npm install
$env:PYTHON_BIN = (Resolve-Path "..\..\..\.venv\Scripts\python.exe").Path
& $env:PYTHON_BIN -m pip install -r worker\requirements.txt
$env:PPTX_EXPORT_ENGINE = "ppt-master"
npm run dev
```

The API listens on `http://127.0.0.1:8787`, the WebSocket event stream on
`ws://127.0.0.1:8788`, and the Vite workbench on `http://localhost:5173`.
`DATA_DIR` defaults to `./data`; set `DATABASE_URL` to use PostgreSQL. When
PostgreSQL is unavailable the server uses the file store and exposes the
chosen persistence mode through `/api/v1/health`.

Install `worker/requirements.txt` into `PYTHON_BIN`. The default exporter creates
a temporary lockless Quick Generate workspace, runs the ppt-master final SVG
checker, converts `svg_output/` to native DrawingML, and requires both package
postflight and delivery checks to pass. `PPTX_EXPORT_ENGINE=legacy` is rejected
unless `FASTPPT_ALLOW_LEGACY_EXPORT=true`; it is a development fallback and not
the release path.

The default `POWERPOINT_RENDERER=unavailable` is deliberately visible as
`SVG 回退` in the workbench. On Windows with desktop PowerPoint and `pywin32`, set
`POWERPOINT_RENDERER=com` to run `worker/render_powerpoint.py`. This renders the
same exported PPTX through Microsoft PowerPoint and records authoritative
1920x1080 page images. The PNG is registered in object storage and served to
the browser through an owner-scoped artifact endpoint; the workbench replaces
the quick SVG only after that exact-version PNG is available. Configure a short,
writable `POWERPOINT_STAGING_DIR` (for example `C:\fppt-com`) so COM never sees
deep project or `DATA_DIR` paths.

For local service dependencies, `docker compose up -d` starts PostgreSQL,
Redis, and MinIO with the schema mounted from `database/schema.sql`; the init
container creates the `fastppt-online` bucket. The API can still use its file
store and local object store for development. Production mode requires
PostgreSQL and S3-compatible object storage and fails startup instead of
silently using local fallbacks. Durable edit/export jobs are persisted before
execution, dispatched with the bounded `JOB_CONCURRENCY` worker, and re-enqueued
after an interrupted process restart. `REDIS_URL` remains reserved for a future
independently scaled transport.

This release is deliberately a single-writer deployment. Set
`API_INSTANCE_COUNT=1`; production startup rejects a larger value because the
PostgreSQL store still keeps a process-local working snapshot. The health
response reports `deploymentMode=single_api_writer`. Horizontal API writers
require transactional row-level commands and database-assigned event sequences
in a future release.

## Behaviour covered

- Persisted, revocable login sessions and owner-scoped project access. Production
  login requires `AUTH_ALLOWED_EMAILS` and the server-side `AUTH_LOGIN_CODE`.
  An allowlisted address with the correct code is a controlled invitation and
  can create the first user in an empty production database; `ALLOW_DEV_LOGIN`
  only controls arbitrary development self-registration.
- `slides.md` parsing into stable `project_id`/`page_id`, page contracts and
  locked fact anchors.
- Current-page chat executes after plan validation without an Apply button.
- Multi-page and global modes return candidate pages, reasons, cost, fact
  impact, and a confirmation operation before any page changes.
- Every page edit appends an immutable `version_id`; history can compare and
  restore versions. A failed page refunds its reserved cost independently.
- Page reorder and soft archive/restore keep version and ledger history bound
  to the original stable `page_id`; export and the workbench omit archived pages.
- Page split creates a new stable continuation `page_id`, appends an immutable
  source-page version, carries only matching fact anchors to each resulting
  page, and emits a `page.split` event for the affected project.
- WebSocket events carry monotonic `seq`, project/page/version/operation IDs.
- Quick SVG preview is retained until the edit settles. The UI distinguishes
  `快速预览`, `PPTX 权威渲染`, and `SVG 回退`.
- The Python export worker authors complete 1280x720 SVG pages, validates them,
  creates native text/shapes/pictures through ppt-master, and writes checker,
  postflight, delivery, and computed static-structure receipts to a QA sidecar.
- Relay-generated visuals are downloaded from local or S3-compatible object
  storage into an isolated worker staging directory and embedded only as marked
  local picture regions. Text and structural shapes remain native and editable.
- `RelayModelAdapter` keeps model credentials server-side and supports
  structured JSON, request IDs, image URL inputs, timeout, retry, model, and
  price/configuration snapshots. Without relay configuration, the deterministic
  planner remains the explicit local fallback.

## API outline

The REST surface follows the Spec under `/api/v1`: auth, projects, import,
pages, page split, chat turns, operation confirmation/cancellation, version
rollback, group rollback, failed-page retry, render/export, and usage ledger. The WebSocket server accepts a
one-time `fastppt-ticket.<ticket>` subprotocol plus
`?projectId=<id>&afterSeq=<n>`, then replays missed events before subscribing to
live updates. Session tokens are carried by an HttpOnly cookie or REST bearer
header, not a WebSocket URL.

For the normalized PostgreSQL entity contract, see `database/schema.sql`.
The local JSON store is a development fallback, not a production replacement
for PostgreSQL/object storage/queue services.

## Verification

```powershell
npm run format:check
npm run build
$env:PYTHON_BIN = (Resolve-Path "..\..\..\.venv\Scripts\python.exe").Path
npm test
npm run test:pptmaster
npm run test:golden
npm run test:powerpoint
npm run smoke
npm run test:e2e
```

`test:pptmaster` requires the retained official `skills/ppt-master` files and
verifies native text/shapes, a local image region, zero full-slide rasters, and
all QA receipts. `test:golden` runs the seven-page Golden Deck covering cover,
two-column, timeline, dense data, chart, image, and complex-flow contracts;
the complex flow is checked for explicit partial-editability marking.
`test:powerpoint` additionally requires Windows PowerPoint and renders a PPTX
whose original path exceeds 255 characters through the short COM staging area.
The API smoke flow
exercises login, stable page IDs, confirmation, event replay, group rollback,
failed-page retry, page split, renderer status, authenticated PowerPoint PNG
bytes, and downloadable PPTX export.
`test:e2e` requires the local API/Web server and drives a real browser through
login, fact confirmation, execution, group rollback, multi-page confirmation,
decoded authoritative PNG dimensions when COM rendering is enabled, and an
offline WebSocket interval with missed-event replay after reconnection.
Browser screenshots used for visual QA are written under `output/playwright/`.

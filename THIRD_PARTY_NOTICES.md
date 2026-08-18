# Third-Party Notices

This file records upstream projects used as the source or design reference for
the FastPPT Online workbench.

## ppt-master

- Source: https://github.com/hugohe3/ppt-master
- Fork: https://github.com/WuHao1893/ppt-master-fastppt-online
- Base branch: `main`
- Base commit: `d45820250dfedee6ab25e1783dafb4ad9f44935a`
- License: MIT
- Copyright: Copyright (c) 2025-2026 Hugo He
- Planned changes in this fork: online page-scoped chat editing, multi-page and
  global edit orchestration, versioning, FastPPT integration, and authoritative
  PPTX rendering. The upstream `LICENSE` file is retained unchanged.

## FastPPT reference project

- Source: https://github.com/FastR-D/FastPPT/tree/slidev
- Local reference: `D:/Users/吴昊/Desktop/PPT图片/FastPPT-slidev`
- Reference commit: `a30cd8b9b020368193740a79171b47aa0de18db5`
- Status: reference-only at this stage; no FastPPT source code is copied into
  this repository by this commit.
- Intended references: image-first generation, high-quality prompt composition,
  fact anchors, editable PPTX conversion, and QA/rollback safeguards.

Before copying any FastPPT code or assets, verify the upstream license and add
the required attribution and license text here.

## FastPPT Online runtime dependencies

The online application under `apps/online` is new code in this fork. Its
runtime dependencies are consumed as unmodified packages; their license files
remain in `apps/online/node_modules` and are not copied into the upstream
skill package.

| Package family | Source | SPDX | Introduced |
| --- | --- | --- | --- |
| Fastify / `@fastify/*` | https://github.com/fastify/fastify | MIT | 2026-08-18 |
| React / React DOM | https://github.com/facebook/react | MIT | 2026-08-18 |
| Vite / TypeScript | https://github.com/vitejs/vite; https://github.com/microsoft/TypeScript | MIT; Apache-2.0 | 2026-08-18 |
| `ws` / `pg` / `zod` | https://github.com/websockets/ws; https://github.com/brianc/node-postgres; https://github.com/colinhacks/zod | MIT | 2026-08-18 |
| AWS SDK for JavaScript (`@aws-sdk/client-s3`) | https://github.com/aws/aws-sdk-js-v3 | Apache-2.0 | 2026-08-18 |
| Lucide React | https://github.com/lucide-icons/lucide | ISC | 2026-08-18 |
| `python-pptx` worker dependency | https://github.com/scanny/python-pptx | MIT | 2026-08-18 |
| `pywin32` PowerPoint COM worker dependency | https://github.com/mhammond/pywin32 | PSF-2.0 | 2026-08-18 |
| `skia-pathops` / `uharfbuzz` | https://github.com/fonttools/skia-pathops; https://github.com/harfbuzz/uharfbuzz | BSD-3-Clause; Apache-2.0 | 2026-08-18 |
| Pillow / XlsxWriter | https://github.com/python-pillow/Pillow; https://github.com/jmcnamara/XlsxWriter | HPND; BSD-2-Clause | 2026-08-18 |

The FastPPT reference project remains reference-only. No student-project
Gateway, MCP, Codex/Claude Harness, or source asset is copied into the online
runtime.

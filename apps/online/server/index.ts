import fs from "node:fs/promises";
import { isIP } from "node:net";
import path from "node:path";
import { fileURLToPath } from "node:url";
import cors from "@fastify/cors";
import multipart from "@fastify/multipart";
import fastifyStatic from "@fastify/static";
import Fastify, {
  type FastifyInstance,
  type FastifyReply,
  type FastifyRequest,
} from "fastify";
import ipaddr from "ipaddr.js";
import { WebSocketServer, type WebSocket } from "ws";
import {
  chatTurnSchema,
  confirmSchema,
  createProjectSchema,
  loginSchema,
} from "../shared/protocol.js";
import { AuthService } from "./auth.js";
import { EventBus } from "./events.js";
import { HttpError } from "./errors.js";
import { OnlineService } from "./service.js";
import { createStore, type StateStore } from "./store.js";

const serverDir = path.dirname(fileURLToPath(import.meta.url));

interface AppContext {
  app: FastifyInstance;
  store: StateStore;
  auth: AuthService;
  service: OnlineService;
  events: EventBus;
}

function bearerToken(request: FastifyRequest): string | undefined {
  const header = request.headers.authorization || "";
  return header.startsWith("Bearer ")
    ? header.slice("Bearer ".length).trim()
    : undefined;
}

function cookieToken(request: FastifyRequest): string | undefined {
  const cookie = request.headers.cookie || "";
  const value = cookie
    .split(";")
    .map((part) => part.trim())
    .find((part) => part.startsWith("fastppt_session="));
  return value
    ? decodeURIComponent(value.slice("fastppt_session=".length))
    : undefined;
}

function authToken(request: FastifyRequest): string | undefined {
  return bearerToken(request) || cookieToken(request);
}

function sessionCookie(token: string, maxAgeSeconds: number): string {
  const secure =
    process.env.COOKIE_SECURE === "true" ||
    process.env.NODE_ENV === "production"
      ? "; Secure"
      : "";
  return `fastppt_session=${encodeURIComponent(token)}; Path=/; Max-Age=${maxAgeSeconds}; HttpOnly; SameSite=Lax${secure}`;
}

function publicExport(
  job: Awaited<ReturnType<OnlineService["getExport"]>>,
): Awaited<ReturnType<OnlineService["getExport"]>> {
  return { ...job, artifactPath: null, artifactObjectKey: null };
}

function publicDocument(
  document: ReturnType<OnlineService["getDocument"]>,
  includeText = false,
) {
  const { objectKey: _objectKey, parsedText, ...safe } = document;
  return { ...safe, parsedText: includeText ? parsedText : undefined };
}

function userId(request: FastifyRequest): string {
  const user = (request as FastifyRequest & { user?: { userId: string } }).user;
  if (!user) throw new HttpError(401, "Login required.");
  return user.userId;
}

function bodyAsRecord(request: FastifyRequest): Record<string, unknown> {
  return (
    request.body && typeof request.body === "object" ? request.body : {}
  ) as Record<string, unknown>;
}

function officePreviewBaseUrl(request: FastifyRequest): string {
  const configured = process.env.PUBLIC_PREVIEW_BASE_URL?.trim().replace(
    /\/$/,
    "",
  );
  if (!configured) {
    if (process.env.NODE_ENV === "production")
      throw new HttpError(
        503,
        "PUBLIC_PREVIEW_BASE_URL is required for Office Viewer.",
      );
    return `${request.protocol}://${request.host}`;
  }
  let parsed: URL;
  try {
    parsed = new URL(configured);
  } catch {
    throw new HttpError(
      503,
      "PUBLIC_PREVIEW_BASE_URL must be an absolute URL.",
    );
  }
  if (process.env.NODE_ENV === "production") {
    const hostname = parsed.hostname.toLowerCase().replace(/^\[|\]$/g, "");
    const literalRange = isIP(hostname) ? ipaddr.parse(hostname).range() : null;
    if (
      parsed.protocol !== "https:" ||
      parsed.username ||
      parsed.password ||
      hostname === "localhost" ||
      hostname.endsWith(".local") ||
      (literalRange && literalRange !== "unicast")
    ) {
      throw new HttpError(
        503,
        "PUBLIC_PREVIEW_BASE_URL must use public HTTPS without credentials.",
      );
    }
  }
  return configured;
}

function replyError(reply: FastifyReply, error: unknown): void {
  if (error instanceof HttpError) {
    void reply
      .code(error.statusCode)
      .send({ error: { message: error.message } });
    return;
  }
  if (error && typeof error === "object" && "issues" in error) {
    void reply
      .code(400)
      .send({
        error: {
          message: "Request validation failed.",
          details: (error as { issues: unknown }).issues,
        },
      });
    return;
  }
  if (
    error &&
    typeof error === "object" &&
    "statusCode" in error &&
    typeof (error as { statusCode: unknown }).statusCode === "number"
  ) {
    const statusCode = (error as { statusCode: number }).statusCode;
    void reply
      .code(statusCode)
      .send({ error: { message: (error as unknown as Error).message } });
    return;
  }
  console.error(error);
  void reply.code(500).send({ error: { message: "Unexpected server error." } });
}

export async function createApp(): Promise<AppContext> {
  const corsOrigin =
    process.env.CORS_ORIGIN?.trim() ||
    (process.env.NODE_ENV === "production" ? "" : "http://localhost:5173");
  if (!corsOrigin || corsOrigin === "*")
    throw new Error("CORS_ORIGIN must contain explicit trusted origins.");
  const apiInstanceCount = Number(process.env.API_INSTANCE_COUNT || 1);
  if (!Number.isInteger(apiInstanceCount) || apiInstanceCount < 1)
    throw new Error("API_INSTANCE_COUNT must be a positive integer.");
  const store = await createStore();
  const events = new EventBus(store);
  const auth = new AuthService(store);
  const service = new OnlineService(store, events);
  const app = Fastify({ logger: process.env.NODE_ENV !== "test" });
  await app.register(cors, {
    origin: corsOrigin
      .split(",")
      .map((value) => value.trim())
      .filter(Boolean),
    credentials: true,
  });
  await app.register(multipart, {
    limits: {
      fileSize: Number(
        process.env.MAX_DOCUMENT_BYTES ||
          process.env.MAX_UPLOAD_BYTES ||
          26_214_400,
      ),
      files: 1,
    },
  });
  const staticCandidates = [
    path.resolve(serverDir, "..", "web", "dist"),
    path.resolve(serverDir, "..", "..", "web", "dist"),
  ];
  const staticRoots = await Promise.all(
    staticCandidates.map(async (candidate) => {
      try {
        await fs.access(candidate);
        return candidate;
      } catch {
        return null;
      }
    }),
  );
  const staticRoot = staticRoots.find((candidate): candidate is string =>
    Boolean(candidate),
  );
  if (staticRoot)
    await app.register(fastifyStatic, { root: staticRoot, prefix: "/" });

  app.decorateRequest("user", null);
  app.addHook("preHandler", async (request) => {
    if (store.kind === "postgres" && request.url.startsWith("/api/"))
      await store.update(() => undefined);
    const publicPath =
      !request.url.startsWith("/api/") ||
      request.url.startsWith("/api/v1/auth/login") ||
      request.url.startsWith("/api/v1/auth/me") ||
      request.url.startsWith("/api/v1/health") ||
      request.url.startsWith("/api/v2/auth/login") ||
      request.url.startsWith("/api/v2/auth/me") ||
      request.url.startsWith("/api/v2/health");
    if (publicPath) return;
    const origin = request.headers.origin;
    const configuredOrigins = corsOrigin
      .split(",")
      .map((value) => value.trim())
      .filter(Boolean);
    if (
      origin &&
      ["POST", "PUT", "PATCH", "DELETE"].includes(request.method) &&
      !configuredOrigins.includes(origin)
    )
      throw new HttpError(403, "Origin is not allowed for this write request.");
    const user = auth.verify(authToken(request));
    if (!user) throw new HttpError(401, "Login required.");
    (request as FastifyRequest & { user?: unknown }).user = user;
  });
  app.addHook("onSend", async (request, reply, payload) => {
    if (
      request.url.startsWith("/api/") &&
      !reply.hasHeader("Cache-Control")
    )
      reply.header("Cache-Control", "no-store");
    return payload;
  });

  app.setErrorHandler((error, _request, reply) => replyError(reply, error));

  app.get("/api/v1/health", async () => ({
    ok: true,
    service: "fastppt-online-api",
    persistence: store.kind,
    renderer:
      process.env.POWERPOINT_RENDERER === "powerpoint" ||
      process.env.POWERPOINT_RENDERER === "com"
        ? "powerpoint_com"
        : "unavailable",
    exporter:
      process.env.PPTX_EXPORT_ENGINE === "legacy"
        ? "legacy_development_fallback"
        : "ppt_master_svg_to_drawingml",
    queue: process.env.QUEUE_MODE || "durable_store",
    deploymentMode:
      store.kind === "postgres"
        ? "transactional_postgres_multi_writer"
        : "development_single_process",
    apiInstanceCount,
    objectStorage: service.objectStorageStatus(),
    relay: service.relayStatus(),
  }));

  app.post("/api/v1/auth/login", async (request, reply) => {
    const input = loginSchema.parse(bodyAsRecord(request));
    const result = await auth.login(input.email, input.name, input.accessCode);
    return reply
      .header("Set-Cookie", sessionCookie(result.token, 60 * 60 * 24 * 7))
      .send(result);
  });

  app.post("/api/v1/auth/logout", async (request, reply) => {
    await auth.revoke(authToken(request));
    return reply.header("Set-Cookie", sessionCookie("", 0)).send({ ok: true });
  });

  app.post("/api/v1/auth/ws-ticket", async (request) => {
    return { ticket: auth.issueWebSocketTicket(userId(request)) };
  });

  app.get("/api/v1/auth/me", async (request) => {
    const user = auth.verify(authToken(request));
    return { user };
  });

  app.get("/api/v1/projects", async (request) => {
    const query = request.query as { includeArchived?: string };
    return {
      projects: service.listProjects(
        userId(request),
        query.includeArchived === "true",
      ),
    };
  });

  app.post("/api/v1/projects", async (request, reply) => {
    const input = createProjectSchema.parse(bodyAsRecord(request));
    const project = await service.createProject(userId(request), input);
    return reply.code(201).send({ project });
  });

  app.get("/api/v1/projects/:projectId", async (request) => {
    const params = request.params as { projectId: string };
    return { project: service.getProject(userId(request), params.projectId) };
  });

  app.post("/api/v1/projects/:projectId/copy", async (request, reply) => {
    const params = request.params as { projectId: string };
    return reply
      .code(201)
      .send({
        project: await service.duplicateProject(
          userId(request),
          params.projectId,
        ),
      });
  });

  app.post("/api/v1/projects/:projectId/archive", async (request) => {
    const params = request.params as { projectId: string };
    return {
      project: await service.setArchived(
        userId(request),
        params.projectId,
        true,
      ),
    };
  });

  app.post("/api/v1/projects/:projectId/restore", async (request) => {
    const params = request.params as { projectId: string };
    return {
      project: await service.setArchived(
        userId(request),
        params.projectId,
        false,
      ),
    };
  });

  app.post("/api/v1/projects/:projectId/import", async (request, reply) => {
    const params = request.params as { projectId: string };
    let markdown = String(bodyAsRecord(request).slidesMarkdown || "");
    if (request.isMultipart()) {
      const part = await request.file();
      if (!part) throw new HttpError(400, "slides.md file is required.");
      if (!part.filename.toLowerCase().endsWith(".md"))
        throw new HttpError(400, "Only UTF-8 .md files can be imported.");
      markdown = (await part.toBuffer()).toString("utf8");
    }
    if (!markdown.trim())
      throw new HttpError(
        400,
        "slidesMarkdown or a slides.md upload is required.",
      );
    const imported = await service.importSlides(
      userId(request),
      params.projectId,
      markdown,
    );
    return reply
      .code(200)
      .send({ project: imported, importedInto: imported.projectId });
  });

  app.get("/api/v1/projects/:projectId/pages", async (request) => {
    const params = request.params as { projectId: string };
    return { pages: service.listPages(userId(request), params.projectId) };
  });

  app.get("/api/v1/projects/:projectId/pages/:pageId", async (request) => {
    const params = request.params as { projectId: string; pageId: string };
    return {
      page: service.getPage(userId(request), params.projectId, params.pageId),
    };
  });

  app.get(
    "/api/v1/projects/:projectId/pages/:pageId/versions",
    async (request) => {
      const params = request.params as { projectId: string; pageId: string };
      return {
        versions: service.listVersions(
          userId(request),
          params.projectId,
          params.pageId,
        ),
      };
    },
  );

  app.post("/api/v1/projects/:projectId/pages/reorder", async (request) => {
    const params = request.params as { projectId: string };
    const body = bodyAsRecord(request);
    const pageIds = Array.isArray(body.pageIds) ? body.pageIds.map(String) : [];
    return {
      pages: await service.reorderPages(
        userId(request),
        params.projectId,
        pageIds,
      ),
    };
  });

  app.post(
    "/api/v1/projects/:projectId/pages/:pageId/archive",
    async (request) => {
      const params = request.params as { projectId: string; pageId: string };
      return {
        page: await service.archivePage(
          userId(request),
          params.projectId,
          params.pageId,
          true,
        ),
      };
    },
  );

  app.post(
    "/api/v1/projects/:projectId/pages/:pageId/restore",
    async (request) => {
      const params = request.params as { projectId: string; pageId: string };
      return {
        page: await service.archivePage(
          userId(request),
          params.projectId,
          params.pageId,
          false,
        ),
      };
    },
  );

  app.post(
    "/api/v1/projects/:projectId/pages/:pageId/split",
    async (request) => {
      const params = request.params as { projectId: string; pageId: string };
      return service.splitPage(
        userId(request),
        params.projectId,
        params.pageId,
      );
    },
  );

  app.get("/api/v1/projects/:projectId/messages", async (request) => {
    const params = request.params as { projectId: string };
    const query = request.query as { conversationId?: string };
    return {
      messages: service.listMessages(
        userId(request),
        params.projectId,
        query.conversationId,
      ),
    };
  });

  app.get("/api/v1/projects/:projectId/history", async (request) => {
    const params = request.params as { projectId: string };
    return service.projectHistory(userId(request), params.projectId);
  });

  app.post("/api/v1/projects/:projectId/chat/turns", async (request, reply) => {
    const params = request.params as { projectId: string };
    const input = chatTurnSchema.parse({
      ...bodyAsRecord(request),
      projectId: params.projectId,
    });
    const result = await service.createChatTurn(userId(request), input);
    return reply
      .code(result.operation.confirmationRequired ? 202 : 200)
      .send(result);
  });

  app.post(
    "/api/v1/projects/:projectId/edit-operations/:operationId/confirm",
    async (request) => {
      const params = request.params as {
        projectId: string;
        operationId: string;
      };
      confirmSchema.parse({ operationId: params.operationId });
      return {
        operation: await service.confirmOperation(
          userId(request),
          params.projectId,
          params.operationId,
        ),
      };
    },
  );

  app.post(
    "/api/v1/projects/:projectId/edit-operations/:operationId/cancel",
    async (request) => {
      const params = request.params as {
        projectId: string;
        operationId: string;
      };
      return {
        operation: await service.cancelOperation(
          userId(request),
          params.projectId,
          params.operationId,
        ),
      };
    },
  );

  app.post(
    "/api/v1/projects/:projectId/edit-operations/:operationId/rollback",
    async (request) => {
      const params = request.params as {
        projectId: string;
        operationId: string;
      };
      return {
        operation: await service.rollbackOperation(
          userId(request),
          params.projectId,
          params.operationId,
        ),
      };
    },
  );

  app.post(
    "/api/v1/projects/:projectId/edit-operations/:operationId/retry-failed",
    async (request) => {
      const params = request.params as {
        projectId: string;
        operationId: string;
      };
      return {
        operation: await service.retryFailedPages(
          userId(request),
          params.projectId,
          params.operationId,
        ),
      };
    },
  );

  app.get(
    "/api/v1/projects/:projectId/edit-operations/:operationId",
    async (request) => {
      const params = request.params as {
        projectId: string;
        operationId: string;
      };
      return {
        operation: service.getOperation(
          userId(request),
          params.projectId,
          params.operationId,
        ),
      };
    },
  );

  app.post(
    "/api/v1/projects/:projectId/versions/:versionId/rollback",
    async (request) => {
      const params = request.params as { projectId: string; versionId: string };
      return {
        page: await service.rollbackVersion(
          userId(request),
          params.projectId,
          params.versionId,
        ),
      };
    },
  );

  app.post("/api/v1/projects/:projectId/render", async (request, reply) => {
    const params = request.params as { projectId: string };
    return reply
      .code(202)
      .send({
        export: await service.createExport(userId(request), params.projectId),
      });
  });

  app.get("/api/v1/projects/:projectId/exports/:exportId", async (request) => {
    const params = request.params as { projectId: string; exportId: string };
    return {
      export: publicExport(
        service.getExport(userId(request), params.projectId, params.exportId),
      ),
    };
  });

  app.get(
    "/api/v1/projects/:projectId/exports/:exportId/download",
    async (request, reply) => {
      const params = request.params as { projectId: string; exportId: string };
      const artifact = await service.readExportArtifact(
        userId(request),
        params.projectId,
        params.exportId,
      );
      return reply
        .type(
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        .header(
          "Content-Disposition",
          `attachment; filename="${artifact.fileName}"`,
        )
        .send(artifact.bytes);
    },
  );

  app.get("/api/v1/projects/:projectId/usage", async (request) => {
    const params = request.params as { projectId: string };
    return { ledger: service.usage(userId(request), params.projectId) };
  });

  app.get("/api/v1/projects/:projectId/audit", async (request) => {
    const params = request.params as { projectId: string };
    return { audit: service.audit(userId(request), params.projectId) };
  });

  app.get("/api/v1/projects/:projectId/artifacts", async (request) => {
    const params = request.params as { projectId: string };
    return {
      artifacts: service
        .artifacts(userId(request), params.projectId)
        .map((artifact) => ({ ...artifact, assetPath: null })),
    };
  });

  app.get(
    "/api/v1/projects/:projectId/artifacts/:artifactId/content",
    async (request, reply) => {
      const params = request.params as {
        projectId: string;
        artifactId: string;
      };
      const artifact = await service.readAuthoritativeRender(
        userId(request),
        params.projectId,
        params.artifactId,
      );
      return reply
        .type("image/png")
        .header("Cache-Control", "private, max-age=31536000, immutable")
        .header("Content-Disposition", "inline")
        .header("X-Content-Type-Options", "nosniff")
        .header("Content-Length", String(artifact.bytes.length))
        .send(artifact.bytes);
    },
  );

  app.get("/api/v1/projects/:projectId/prompt-snapshots", async (request) => {
    const params = request.params as { projectId: string };
    return {
      promptSnapshots: service.promptSnapshots(
        userId(request),
        params.projectId,
      ),
    };
  });

  app.get("/api/v2/health", async () => ({
    ok: true,
    service: "fastppt-online-api",
    apiVersion: "v2",
    persistence: store.kind,
  }));

  app.post("/api/v2/auth/login", async (request, reply) => {
    const input = loginSchema.parse(bodyAsRecord(request));
    const result = await auth.login(input.email, input.name, input.accessCode);
    return reply
      .header("Set-Cookie", sessionCookie(result.token, 60 * 60 * 24 * 7))
      .send(result);
  });

  app.get("/api/v2/auth/me", async (request) => ({
    user: auth.verify(authToken(request)),
  }));

  app.get("/api/v2/projects", async (request) => {
    const query = request.query as { includeArchived?: string };
    return {
      projects: service.listProjects(
        userId(request),
        query.includeArchived === "true",
      ),
    };
  });

  app.post("/api/v2/projects", async (request, reply) => {
    const input = createProjectSchema.parse(bodyAsRecord(request));
    return reply
      .code(201)
      .send({ project: await service.createProject(userId(request), input) });
  });

  app.get("/api/v2/projects/:projectId", async (request) => {
    const params = request.params as { projectId: string };
    return { project: service.getProject(userId(request), params.projectId) };
  });

  app.post(
    "/api/v2/projects/:projectId/work-sessions",
    async (request, reply) => {
      const params = request.params as { projectId: string };
      const body = bodyAsRecord(request);
      const workflowMode = String(body.workflowMode || "");
      if (
        !["document_import", "page_entry", "ppt_beautify"].includes(
          workflowMode,
        )
      )
        throw new HttpError(400, "Invalid workflowMode.");
      const intent =
        body.intent === "create" || body.intent === "reference"
          ? body.intent
          : "modify";
      const pageBudget =
        body.pageBudget === undefined ? undefined : Number(body.pageBudget);
      if (
        pageBudget !== undefined &&
        (!Number.isInteger(pageBudget) || pageBudget < 1 || pageBudget > 100)
      )
        throw new HttpError(400, "pageBudget must be an integer from 1 to 100.");
      if (
        body.inheritTheme !== undefined &&
        typeof body.inheritTheme !== "boolean"
      )
        throw new HttpError(400, "inheritTheme must be a boolean.");
      const session = await service.createWorkSession(
        userId(request),
        params.projectId,
        {
          workflowMode: workflowMode as
            | "document_import"
            | "page_entry"
            | "ppt_beautify",
          intent,
          pageBudget,
          inheritTheme: body.inheritTheme as boolean | undefined,
        },
      );
      return reply.code(201).send({ session });
    },
  );

  app.get("/api/v2/projects/:projectId/work-sessions", async (request) => {
    const params = request.params as { projectId: string };
    return {
      sessions: service.listWorkSessions(userId(request), params.projectId),
    };
  });

  app.patch(
    "/api/v2/projects/:projectId/work-sessions/:sessionId",
    async (request) => {
      const params = request.params as { projectId: string; sessionId: string };
      const body = bodyAsRecord(request);
      if (
        body.intent !== undefined &&
        !["create", "modify", "reference"].includes(String(body.intent))
      )
        throw new HttpError(400, "Invalid intent.");
      if (
        body.pageBudget !== undefined &&
        (!Number.isInteger(Number(body.pageBudget)) ||
          Number(body.pageBudget) < 1 ||
          Number(body.pageBudget) > 100)
      )
        throw new HttpError(400, "pageBudget must be an integer from 1 to 100.");
      if (
        body.inheritTheme !== undefined &&
        typeof body.inheritTheme !== "boolean"
      )
        throw new HttpError(400, "inheritTheme must be a boolean.");
      return {
        session: await service.updateWorkSession(
          userId(request),
          params.projectId,
          params.sessionId,
          {
            intent: body.intent as
              | "create"
              | "modify"
              | "reference"
              | undefined,
            pageBudget:
              body.pageBudget === undefined
                ? undefined
                : Number(body.pageBudget),
            inheritTheme: body.inheritTheme as boolean | undefined,
          },
        ),
      };
    },
  );

  app.post(
    "/api/v2/projects/:projectId/work-sessions/:sessionId/structure-plan",
    async (request, reply) => {
      const params = request.params as { projectId: string; sessionId: string };
      const body = bodyAsRecord(request);
      const pageBudget =
        body.pageBudget === undefined ? undefined : Number(body.pageBudget);
      if (
        pageBudget !== undefined &&
        (!Number.isInteger(pageBudget) || pageBudget < 1 || pageBudget > 100)
      )
        throw new HttpError(400, "pageBudget must be an integer from 1 to 100.");
      if (
        body.inheritTheme !== undefined &&
        typeof body.inheritTheme !== "boolean"
      )
        throw new HttpError(400, "inheritTheme must be a boolean.");
      return reply.code(201).send(
        await service.createDocumentStructurePlan(
          userId(request),
          params.projectId,
          params.sessionId,
          {
            pageBudget,
            inheritTheme: body.inheritTheme as boolean | undefined,
          },
        ),
      );
    },
  );

  app.post(
    "/api/v2/projects/:projectId/documents/upload",
    async (request, reply) => {
      const params = request.params as { projectId: string };
      const query = request.query as { sessionId?: string };
      if (!query.sessionId) throw new HttpError(400, "sessionId is required.");
      const part = await request.file();
      if (!part) throw new HttpError(400, "A document file is required.");
      const result = await service.uploadDocument(
        userId(request),
        params.projectId,
        query.sessionId,
        {
          fileName: part.filename,
          contentType: part.mimetype,
          bytes: await part.toBuffer(),
        },
      );
      return reply
        .code(result.duplicate ? 200 : 201)
        .send({
          document: publicDocument(result.document),
          duplicate: result.duplicate,
        });
    },
  );

  app.get("/api/v2/projects/:projectId/documents", async (request) => {
    const params = request.params as { projectId: string };
    const query = request.query as { sessionId?: string };
    return {
      documents: service
        .listDocuments(userId(request), params.projectId, query.sessionId)
        .map((document) => publicDocument(document)),
    };
  });

  app.get(
    "/api/v2/projects/:projectId/documents/:documentId",
    async (request) => {
      const params = request.params as {
        projectId: string;
        documentId: string;
      };
      return {
        document: publicDocument(
          service.getDocument(
            userId(request),
            params.projectId,
            params.documentId,
          ),
          true,
        ),
      };
    },
  );

  app.post(
    "/api/v2/projects/:projectId/documents/:documentId/parse",
    async (request) => {
      const params = request.params as {
        projectId: string;
        documentId: string;
      };
      return {
        document: publicDocument(
          await service.parseDocument(
            userId(request),
            params.projectId,
            params.documentId,
          ),
        ),
      };
    },
  );

  app.patch(
    "/api/v2/projects/:projectId/documents/:documentId/context",
    async (request) => {
      const params = request.params as {
        projectId: string;
        documentId: string;
      };
      const body = bodyAsRecord(request);
      if (typeof body.includedInContext !== "boolean")
        throw new HttpError(400, "includedInContext must be a boolean.");
      return {
        document: publicDocument(
          await service.setDocumentContext(
            userId(request),
            params.projectId,
            params.documentId,
            body.includedInContext,
          ),
        ),
      };
    },
  );

  app.get("/api/v2/projects/:projectId/conflicts", async (request) => {
    const params = request.params as { projectId: string };
    const query = request.query as { sessionId?: string };
    return {
      conflicts: service.listDocumentConflicts(
        userId(request),
        params.projectId,
        query.sessionId,
      ),
    };
  });

  app.post(
    "/api/v2/projects/:projectId/conflicts/:conflictId/resolve",
    async (request) => {
      const params = request.params as {
        projectId: string;
        conflictId: string;
      };
      const body = bodyAsRecord(request);
      const action = String(body.action || "prefer_source");
      if (!["prefer_source", "keep_both", "ignore"].includes(action))
        throw new HttpError(400, "Invalid conflict resolution action.");
      if (
        action === "prefer_source" &&
        !String(body.selectedValue || "") &&
        !Array.isArray(body.selectedDocumentIds)
      )
        throw new HttpError(
          400,
          "prefer_source requires selectedValue or selectedDocumentIds.",
        );
      return {
        conflict: await service.resolveDocumentConflict(
          userId(request),
          params.projectId,
          params.conflictId,
          {
            action: action as "prefer_source" | "keep_both" | "ignore",
            selectedValue: body.selectedValue
              ? String(body.selectedValue)
              : undefined,
            selectedDocumentIds: Array.isArray(body.selectedDocumentIds)
              ? body.selectedDocumentIds.map(String)
              : undefined,
            note: String(body.note || ""),
          },
        ),
      };
    },
  );

  app.patch("/api/v2/projects/:projectId/settings", async (request) => {
    const params = request.params as { projectId: string };
    const body = bodyAsRecord(request);
    if (typeof body.sensitiveMode !== "boolean")
      throw new HttpError(400, "sensitiveMode must be a boolean.");
    return {
      project: await service.updateProjectSettings(
        userId(request),
        params.projectId,
        { sensitiveMode: body.sensitiveMode },
      ),
    };
  });

  app.post("/api/v2/projects/:projectId/chat/turns", async (request, reply) => {
    const params = request.params as { projectId: string };
    const input = chatTurnSchema.parse({
      ...bodyAsRecord(request),
      projectId: params.projectId,
    });
    const result = await service.createChatTurn(userId(request), input);
    return reply
      .code(result.operation.confirmationRequired ? 202 : 200)
      .send(result);
  });

  app.post("/api/v2/projects/:projectId/plans", async (request, reply) => {
    const params = request.params as { projectId: string };
    const input = chatTurnSchema.parse({
      ...bodyAsRecord(request),
      projectId: params.projectId,
    });
    const result = await service.createChatTurn(userId(request), input, {
      deferExecution: true,
    });
    return reply
      .code(201)
      .send({
        planId: result.operation.operationId,
        plan: result.operation.structuredPlan,
        operation: result.operation,
        messages: result.messages,
      });
  });

  app.get("/api/v2/projects/:projectId/plans/:planId", async (request) => {
    const params = request.params as { projectId: string; planId: string };
    const operation = service.getOperation(
      userId(request),
      params.projectId,
      params.planId,
    );
    return {
      planId: operation.operationId,
      plan: operation.structuredPlan,
      operation,
    };
  });

  app.post(
    "/api/v2/projects/:projectId/plans/:planId/confirm",
    async (request) => {
      const params = request.params as { projectId: string; planId: string };
      return {
        operation: await service.confirmOperation(
          userId(request),
          params.projectId,
          params.planId,
        ),
      };
    },
  );

  app.post(
    "/api/v2/projects/:projectId/plans/:planId/cancel",
    async (request) => {
      const params = request.params as { projectId: string; planId: string };
      return {
        operation: await service.cancelOperation(
          userId(request),
          params.projectId,
          params.planId,
        ),
      };
    },
  );

  app.get("/api/v2/projects/:projectId/conversations", async (request) => {
    const params = request.params as { projectId: string };
    return {
      conversations: service.listMessages(userId(request), params.projectId),
    };
  });

  app.get("/api/v2/projects/:projectId/versions", async (request) => {
    const params = request.params as { projectId: string };
    const project = service.getProject(userId(request), params.projectId);
    return { versions: project.pages.flatMap((page) => page.versions) };
  });

  app.get(
    "/api/v2/projects/:projectId/visual-previews/:artifactId/content",
    async (request, reply) => {
      const params = request.params as {
        projectId: string;
        artifactId: string;
      };
      const preview = await service.readVisualPreview(
        userId(request),
        params.projectId,
        params.artifactId,
      );
      return reply
        .type(preview.contentType)
        .header("Cache-Control", "private, max-age=300")
        .header("X-Content-Type-Options", "nosniff")
        .send(preview.bytes);
    },
  );

  app.post(
    "/api/v2/projects/:projectId/versions/:versionId/rollback",
    async (request) => {
      const params = request.params as { projectId: string; versionId: string };
      return {
        page: await service.rollbackVersion(
          userId(request),
          params.projectId,
          params.versionId,
        ),
      };
    },
  );

  app.post(
    "/api/v2/projects/:projectId/operations/:operationId/rollback",
    async (request) => {
      const params = request.params as {
        projectId: string;
        operationId: string;
      };
      return {
        operation: await service.rollbackOperation(
          userId(request),
          params.projectId,
          params.operationId,
        ),
      };
    },
  );

  app.post("/api/v2/projects/:projectId/exports", async (request, reply) => {
    const params = request.params as { projectId: string };
    return reply
      .code(202)
      .send({
        export: await service.createExport(userId(request), params.projectId),
      });
  });

  app.get("/api/v2/projects/:projectId/exports/:exportId", async (request) => {
    const params = request.params as { projectId: string; exportId: string };
    return {
      export: publicExport(
        service.getExport(userId(request), params.projectId, params.exportId),
      ),
    };
  });

  app.get(
    "/api/v2/projects/:projectId/exports/:exportId/download",
    async (request, reply) => {
      const params = request.params as { projectId: string; exportId: string };
      const artifact = await service.readExportArtifact(
        userId(request),
        params.projectId,
        params.exportId,
      );
      return reply
        .type(
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        .header(
          "Content-Disposition",
          `attachment; filename="${artifact.fileName}"`,
        )
        .send(artifact.bytes);
    },
  );

  app.post(
    "/api/v2/projects/:projectId/exports/:exportId/office-preview",
    async (request) => {
      const params = request.params as { projectId: string; exportId: string };
      const previewBase = officePreviewBaseUrl(request);
      const grant = await service.issueOfficePreviewGrant(
        userId(request),
        params.projectId,
        params.exportId,
      );
      const publicUrl = `${previewBase}${grant.publicPath}`;
      return {
        grantId: grant.grantId,
        expiresAt: grant.expiresAt,
        publicUrl,
        officeViewerUrl: `https://view.officeapps.live.com/op/view.aspx?src=${encodeURIComponent(publicUrl)}`,
      };
    },
  );

  app.route({
    method: ["GET", "HEAD"],
    url: "/public/office-preview/:opaqueToken",
    handler: async (request, reply) => {
      const params = request.params as { opaqueToken: string };
      const method = request.method === "HEAD" ? "HEAD" : "GET";
      const artifact = await service.readOfficePreviewGrant(
        params.opaqueToken,
        method,
        request.headers.range,
      );
      reply
        .type(
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        .header(
          "Content-Disposition",
          `inline; filename*=UTF-8''${encodeURIComponent(artifact.fileName)}`,
        )
        .header("X-Content-Type-Options", "nosniff")
        .header("Accept-Ranges", "bytes")
        .header(
          "Content-Length",
          String(
            method === "HEAD" ? artifact.fullLength : artifact.bytes.length,
          ),
        );
      if (artifact.contentRange)
        reply.code(206).header("Content-Range", artifact.contentRange);
      return method === "HEAD" ? reply.send() : reply.send(artifact.bytes);
    },
  });

  return { app, store, auth, service, events };
}

export async function startServer(): Promise<void> {
  const context = await createApp();
  await context.service.resumeJobs();
  const port = Number(process.env.PORT || 8787);
  const wsPort = Number(process.env.WS_PORT || port + 1);
  const host = process.env.HOST || "127.0.0.1";
  const wss = new WebSocketServer({
    port: wsPort,
    host: "0.0.0.0",
    handleProtocols: (protocols) =>
      [...protocols].find((protocol) =>
        protocol.startsWith("fastppt-ticket."),
      ) || false,
  });
  wss.on("connection", (socket: WebSocket, request) => {
    try {
      const url = new URL(
        request.url || "/",
        `http://${request.headers.host || "localhost"}`,
      );
      const projectId = url.searchParams.get("projectId") || "";
      const protocol = request.headers["sec-websocket-protocol"];
      const protocolValue = Array.isArray(protocol)
        ? protocol.join(",")
        : protocol || "";
      const ticket = protocolValue
        .split(",")
        .map((value: string) => value.trim())
        .find((value: string) => value.startsWith("fastppt-ticket."))
        ?.slice("fastppt-ticket.".length);
      const user = context.auth.consumeWebSocketTicket(ticket);
      if (
        !user ||
        !context.service
          .listProjects(user.userId, true)
          .some((project) => project.projectId === projectId)
      ) {
        socket.close(1008, "Unauthorized");
        return;
      }
      const afterSeq = Number(url.searchParams.get("afterSeq") || 0);
      context.events
        .history(projectId, afterSeq)
        .forEach((event) => socket.send(JSON.stringify(event)));
      const unsubscribe = context.events.subscribe(projectId, (event) => {
        if (socket.readyState === socket.OPEN)
          socket.send(JSON.stringify(event));
      });
      socket.on("close", unsubscribe);
    } catch {
      socket.close(1008, "Invalid WebSocket request");
    }
  });
  await context.app.listen({ port, host });
  console.log(
    `FastPPT Online API listening at http://${host}:${port} (WebSocket :${wsPort}, ${context.store.kind} persistence)`,
  );
}

if (
  process.argv[1] &&
  path.resolve(process.argv[1]) === path.resolve(fileURLToPath(import.meta.url))
) {
  startServer().catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
}

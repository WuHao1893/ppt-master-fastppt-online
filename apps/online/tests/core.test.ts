import assert from "node:assert/strict";
import fs from "node:fs/promises";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { AuthService } from "../server/auth.js";
import { EventBus } from "../server/events.js";
import { buildEditPlan } from "../server/plans.js";
import { SlidevQuickPreviewWorker } from "../server/quickPreview.js";
import { downloadRemoteImage, RelayModelAdapter } from "../server/relay.js";
import { OnlineService } from "../server/service.js";
import { FileStore } from "../server/store.js";
import { DurableJobQueue } from "../server/jobQueue.js";
import { parseControlledDocument } from "../server/documentParser.js";
import { editPlanSchema } from "../shared/protocol.js";

async function tempStore(): Promise<{ directory: string; store: FileStore }> {
  const directory = await fs.mkdtemp(
    path.join(os.tmpdir(), "fastppt-online-test-"),
  );
  const store = new FileStore(directory);
  await store.init();
  return { directory, store };
}

test("structured plan reports locked fact removal and requires confirmation", async () => {
  const { directory, store } = await tempStore();
  try {
    const service = new OnlineService(store, new EventBus(store));
    const project = await service.createProject("owner_plan", {
      name: "Fact deck",
      themeId: "test",
      themeVersion: "1",
      slidesMarkdown: "# Metric\n\n42% baseline remains locked.",
    });
    const page = project.pages[0];
    const plan = buildEditPlan(
      project,
      [page],
      "single",
      "正文改为：No metric remains in this body.",
    );
    assert.equal(editPlanSchema.safeParse(plan).success, true);
    assert.equal(plan.requiresConfirmation, true);
    assert.ok(plan.factImpact.removed.some((value) => value.includes("42%")));
  } finally {
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test("confirmed operation removes a declared fact and group rollback restores its version snapshot", async () => {
  const { directory, store } = await tempStore();
  const previousDataDir = process.env.DATA_DIR;
  process.env.DATA_DIR = directory;
  try {
    const service = new OnlineService(store, new EventBus(store));
    const project = await service.createProject("owner_rollback", {
      name: "Rollback deck",
      themeId: "test",
      themeVersion: "1",
      slidesMarkdown: "# Metric\n\n42% baseline remains locked.",
    });
    const pageId = project.pages[0].pageId;
    const turn = await service.createChatTurn("owner_rollback", {
      projectId: project.projectId,
      deckRevisionId: project.currentDeckRevisionId,
      target: { mode: "single", pageIds: [pageId] },
      message: "正文改为：No metric remains in this body.",
      clientRevision: 0,
    });
    assert.equal(turn.operation.status, "planned");
    const applied = await service.confirmOperation(
      "owner_rollback",
      project.projectId,
      turn.operation.operationId,
    );
    assert.equal(applied.status, "completed");
    assert.equal(
      service
        .getPage("owner_rollback", project.projectId, pageId)
        .factAnchors.some((fact) => fact.value === "42%"),
      false,
    );
    await service.rollbackOperation(
      "owner_rollback",
      project.projectId,
      applied.operationId,
    );
    const restored = service.getPage(
      "owner_rollback",
      project.projectId,
      pageId,
    );
    assert.equal(restored.body.includes("42%"), true);
    assert.equal(
      restored.factAnchors.some((fact) => fact.value === "42%"),
      true,
    );
  } finally {
    if (previousDataDir === undefined) delete process.env.DATA_DIR;
    else process.env.DATA_DIR = previousDataDir;
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test("fact replacement is explicit, confirmed, and carried into the next anchor snapshot", async () => {
  const { directory, store } = await tempStore();
  try {
    const service = new OnlineService(store, new EventBus(store));
    const project = await service.createProject("owner_change", {
      name: "Fact change deck",
      themeId: "test",
      themeVersion: "1",
      slidesMarkdown: "# Metric\n\n42% baseline remains locked.",
    });
    const pageId = project.pages[0].pageId;
    const turn = await service.createChatTurn("owner_change", {
      projectId: project.projectId,
      deckRevisionId: project.currentDeckRevisionId,
      target: { mode: "single", pageIds: [pageId] },
      message: "把 42% 改为 43%",
      clientRevision: 0,
    });
    assert.equal(turn.operation.status, "planned");
    assert.ok(
      turn.operation.factImpact.changed.some((value) =>
        value.includes("42%->43%"),
      ),
    );
    const applied = await service.confirmOperation(
      "owner_change",
      project.projectId,
      turn.operation.operationId,
    );
    assert.equal(applied.status, "completed");
    const page = service.getPage("owner_change", project.projectId, pageId);
    assert.equal(page.body.includes("43%"), true);
    assert.equal(page.body.includes("42%"), false);
    assert.equal(
      page.factAnchors.some((fact) => fact.value === "43%"),
      true,
    );
    assert.equal(
      page.factAnchors.some((fact) => fact.value === "42%"),
      false,
    );
  } finally {
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test("signed session survives AuthService recreation and WebSocket tickets are one time", async () => {
  const { directory, store } = await tempStore();
  const previousSecret = process.env.AUTH_SECRET;
  process.env.AUTH_SECRET = "test-secret-at-least-local";
  try {
    const first = new AuthService(store);
    const login = await first.login("auth@example.test", "Auth Test");
    const recreated = new AuthService(store);
    assert.equal(recreated.verify(login.token)?.email, "auth@example.test");
    const ticket = recreated.issueWebSocketTicket(login.user.userId);
    assert.equal(
      recreated.consumeWebSocketTicket(ticket)?.userId,
      login.user.userId,
    );
    assert.equal(recreated.consumeWebSocketTicket(ticket), null);
    await recreated.revoke(login.token);
    assert.equal(new AuthService(store).verify(login.token), null);
  } finally {
    if (previousSecret === undefined) delete process.env.AUTH_SECRET;
    else process.env.AUTH_SECRET = previousSecret;
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test("production allowlist and access code can bootstrap the first user", async () => {
  const { directory, store } = await tempStore();
  const previous = {
    nodeEnv: process.env.NODE_ENV,
    secret: process.env.AUTH_SECRET,
    code: process.env.AUTH_LOGIN_CODE,
    allowed: process.env.AUTH_ALLOWED_EMAILS,
    devLogin: process.env.ALLOW_DEV_LOGIN,
  };
  process.env.NODE_ENV = "production";
  process.env.AUTH_SECRET = "production-test-secret-at-least-32-characters";
  process.env.AUTH_LOGIN_CODE = "invite-code";
  process.env.AUTH_ALLOWED_EMAILS = "first@example.test";
  process.env.ALLOW_DEV_LOGIN = "false";
  try {
    const auth = new AuthService(store);
    const login = await auth.login(
      "first@example.test",
      "First User",
      "invite-code",
    );
    assert.equal(login.user.email, "first@example.test");
    assert.equal(store.state.users.length, 1);
    assert.equal(auth.verify(login.token)?.userId, login.user.userId);
    await assert.rejects(
      () => auth.login("blocked@example.test", "Blocked", "invite-code"),
      /not allowed/i,
    );
    await assert.rejects(
      () => auth.login("first@example.test", "First User", "wrong-code"),
      /invalid access code/i,
    );
  } finally {
    if (previous.nodeEnv === undefined) delete process.env.NODE_ENV;
    else process.env.NODE_ENV = previous.nodeEnv;
    if (previous.secret === undefined) delete process.env.AUTH_SECRET;
    else process.env.AUTH_SECRET = previous.secret;
    if (previous.code === undefined) delete process.env.AUTH_LOGIN_CODE;
    else process.env.AUTH_LOGIN_CODE = previous.code;
    if (previous.allowed === undefined) delete process.env.AUTH_ALLOWED_EMAILS;
    else process.env.AUTH_ALLOWED_EMAILS = previous.allowed;
    if (previous.devLogin === undefined) delete process.env.ALLOW_DEV_LOGIN;
    else process.env.ALLOW_DEV_LOGIN = previous.devLogin;
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test("remote image download rejects unsafe targets, redirects, media types, and oversized bodies", async () => {
  const publicLookup = async (): Promise<Array<{ address: string }>> => [
    { address: "93.184.216.34" },
  ];
  let fetchCalls = 0;
  const unusedFetch = (async (): Promise<Response> => {
    fetchCalls += 1;
    throw new Error("fetch should not be reached");
  }) as typeof fetch;
  await assert.rejects(
    () =>
      downloadRemoteImage("http://images.example.test/a.png", undefined, {
        fetchImpl: unusedFetch,
        lookup: publicLookup,
      }),
    /must use HTTPS/i,
  );
  const blockedAddresses = [
    "127.0.0.1",
    "169.254.169.254",
    "198.18.0.1",
    "192.0.2.1",
    "240.0.0.1",
    "255.255.255.255",
    "[fec0::1]",
    "[64:ff9b::7f00:1]",
    "[64:ff9b:1::a00:1]",
    "[::ffff:127.0.0.1]",
    "[::ffff:10.0.0.1]",
    "[2002:7f00:1::1]",
    "[2001::1]",
    "[2001:db8::1]",
  ];
  for (const address of blockedAddresses) {
    await assert.rejects(
      () =>
        downloadRemoteImage(`https://${address}/a.png`, undefined, {
          fetchImpl: unusedFetch,
          lookup: publicLookup,
        }),
      /private or non-routable/i,
    );
  }
  await assert.rejects(
    () =>
      downloadRemoteImage("https://mixed.example.test/a.png", undefined, {
        fetchImpl: unusedFetch,
        lookup: async () => [{ address: "8.8.8.8" }, { address: "10.0.0.1" }],
      }),
    /private or non-routable/i,
  );
  assert.equal(fetchCalls, 0);

  const redirectFetch = (async (): Promise<Response> =>
    new Response(null, {
      status: 302,
      headers: { Location: "https://10.0.0.1/secret.png" },
    })) as typeof fetch;
  await assert.rejects(
    () =>
      downloadRemoteImage("https://images.example.test/a.png", undefined, {
        fetchImpl: redirectFetch,
        lookup: publicLookup,
      }),
    /private or non-routable/i,
  );

  const textFetch = (async (): Promise<Response> =>
    new Response("not an image", {
      status: 200,
      headers: { "Content-Type": "text/plain" },
    })) as typeof fetch;
  await assert.rejects(
    () =>
      downloadRemoteImage("https://images.example.test/a.png", undefined, {
        fetchImpl: textFetch,
        lookup: publicLookup,
      }),
    /unsupported media type/i,
  );

  const disguisedFetch = (async (): Promise<Response> =>
    new Response("not a png", {
      status: 200,
      headers: { "Content-Type": "image/png" },
    })) as typeof fetch;
  await assert.rejects(
    () =>
      downloadRemoteImage("https://images.example.test/a.png", undefined, {
        fetchImpl: disguisedFetch,
        lookup: publicLookup,
      }),
    /do not match/i,
  );

  const declaredLargeFetch = (async (): Promise<Response> =>
    new Response("x", {
      status: 200,
      headers: { "Content-Type": "image/png", "Content-Length": "100" },
    })) as typeof fetch;
  await assert.rejects(
    () =>
      downloadRemoteImage("https://images.example.test/a.png", undefined, {
        fetchImpl: declaredLargeFetch,
        lookup: publicLookup,
        maxBytes: 8,
      }),
    /exceeded the 8 byte limit/i,
  );

  const streamedLargeFetch = (async (): Promise<Response> =>
    new Response(
      new ReadableStream({
        start(controller) {
          controller.enqueue(new Uint8Array([1, 2, 3, 4]));
          controller.enqueue(new Uint8Array([5, 6, 7, 8]));
          controller.close();
        },
      }),
      { status: 200, headers: { "Content-Type": "image/png" } },
    )) as typeof fetch;
  await assert.rejects(
    () =>
      downloadRemoteImage("https://images.example.test/a.png", undefined, {
        fetchImpl: streamedLargeFetch,
        lookup: publicLookup,
        maxBytes: 6,
      }),
    /while streaming/i,
  );

  const png = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  const imageFetch = (async (): Promise<Response> =>
    new Response(png, {
      status: 200,
      headers: {
        "Content-Type": "image/png",
        "Content-Length": String(png.length),
      },
    })) as typeof fetch;
  const downloaded = await downloadRemoteImage(
    "https://images.example.test/a.png",
    undefined,
    { fetchImpl: imageFetch, lookup: publicLookup, maxBytes: 16 },
  );
  assert.equal(downloaded.bytes.equals(png), true);
  assert.equal(downloaded.mimeType, "image/png");
  await downloadRemoteImage("https://8.8.8.8/a.png", undefined, {
    fetchImpl: imageFetch,
    lookup: publicLookup,
    maxBytes: 16,
  });
  await downloadRemoteImage("https://[2606:4700:4700::1111]/a.png", undefined, {
    fetchImpl: imageFetch,
    lookup: publicLookup,
    maxBytes: 16,
  });
});

test("durable job queue persists terminal state and enforces its concurrency limit", async () => {
  const { directory, store } = await tempStore();
  try {
    const queue = new DurableJobQueue(store, 1);
    let active = 0;
    let peak = 0;
    const makeJob = (jobId: string) => {
      const timestamp = new Date().toISOString();
      return {
        jobId,
        idempotencyKey: `edit:${jobId}`,
        projectId: "queue-project",
        kind: "edit" as const,
        status: "queued" as const,
        payload: {},
        attemptCount: 0,
        maxAttempts: 3,
        availableAt: timestamp,
        leaseOwner: null,
        leaseExpiresAt: null,
        lastError: null,
        startedAt: null,
        createdAt: timestamp,
        updatedAt: timestamp,
        completedAt: null,
      };
    };
    const runner = async (): Promise<void> => {
      active += 1;
      peak = Math.max(peak, active);
      await new Promise((resolve) => setTimeout(resolve, 12));
      active -= 1;
    };
    await Promise.all([
      queue.enqueue(makeJob("job_a"), runner, true),
      queue.enqueue(makeJob("job_b"), runner, true),
    ]);
    assert.equal(peak, 1);
    assert.equal(
      store.state.jobs.every((job) => job.status === "completed"),
      true,
    );
    let retryAttempts = 0;
    const retryJob = { ...makeJob("job_retry"), maxAttempts: 2 };
    await queue.enqueue(
      retryJob,
      async () => {
        retryAttempts += 1;
        if (retryAttempts === 1) throw new Error("transient queue failure");
      },
      true,
    );
    assert.equal(retryAttempts, 2);
    assert.equal(
      store.state.jobs.find((job) => job.jobId === retryJob.jobId)?.attemptCount,
      2,
    );
    let duplicateRuns = 0;
    await queue.enqueue(
      { ...retryJob, jobId: "different_job_id" },
      async () => {
        duplicateRuns += 1;
      },
      true,
    );
    assert.equal(duplicateRuns, 0);
  } finally {
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test("retry operation targets only failed pages and keeps a parent audit link", async () => {
  const { directory, store } = await tempStore();
  const previous = {
    base: process.env.RELAY_BASE_URL,
    key: process.env.RELAY_API_KEY,
    data: process.env.DATA_DIR,
  };
  delete process.env.RELAY_BASE_URL;
  delete process.env.RELAY_API_KEY;
  process.env.DATA_DIR = directory;
  try {
    const service = new OnlineService(store, new EventBus(store));
    const project = await service.createProject("owner_retry", {
      name: "Retry deck",
      themeId: "test",
      themeVersion: "1",
      slidesMarkdown: "# Visual\n\nA page that requests a generated image.",
    });
    const pageId = project.pages[0].pageId;
    const first = await service.createChatTurn("owner_retry", {
      projectId: project.projectId,
      deckRevisionId: project.currentDeckRevisionId,
      target: { mode: "single", pageIds: [pageId] },
      message: "替换当前图片",
      clientRevision: 0,
    });
    assert.equal(first.operation.status, "planned");
    await store.update((state) => {
      const operation = state.operations.find(
        (candidate) => candidate.operationId === first.operation.operationId,
      )!;
      operation.status = "failed";
      operation.confirmedAt = new Date().toISOString();
      operation.failedPageIds = [pageId];
    });
    assert.equal(first.operation.status, "failed");
    assert.deepEqual(first.operation.failedPageIds, [pageId]);
    const retry = await service.retryFailedPages(
      "owner_retry",
      project.projectId,
      first.operation.operationId,
    );
    assert.equal(retry.parentOperationId, first.operation.operationId);
    assert.deepEqual(retry.resolvedPageIds, [pageId]);
    assert.equal(retry.status, "failed");
    const retryEntries = service
      .usage("owner_retry", project.projectId)
      .filter((entry) => entry.operationId === retry.operationId);
    assert.equal(retryEntries.length, 3);
    assert.equal(
      retryEntries.every((entry) => entry.status === "refunded"),
      true,
    );
    assert.equal(
      store.state.jobs.find((job) => job.jobId === retry.operationId)
        ?.attemptCount,
      3,
    );
  } finally {
    if (previous.base === undefined) delete process.env.RELAY_BASE_URL;
    else process.env.RELAY_BASE_URL = previous.base;
    if (previous.key === undefined) delete process.env.RELAY_API_KEY;
    else process.env.RELAY_API_KEY = previous.key;
    if (previous.data === undefined) delete process.env.DATA_DIR;
    else process.env.DATA_DIR = previous.data;
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test("Relay image and Slidev adapters perform real configured HTTP calls", async () => {
  const png = Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
    "base64",
  );
  let imageRequestId = "";
  let imageCalls = 0;
  let hmrCalls = 0;
  const server = http.createServer((request, response) => {
    if (request.url === "/v1/images") {
      imageCalls += 1;
      imageRequestId = String(request.headers["x-request-id"] || "");
      response.writeHead(200, { "Content-Type": "application/json" });
      response.end(
        JSON.stringify({
          data: [{ b64_json: png.toString("base64"), mime_type: "image/png" }],
        }),
      );
      return;
    }
    if (request.url === "/api/hmr") {
      hmrCalls += 1;
      response.writeHead(204);
      response.end();
      return;
    }
    if (request.url === "/v1/chat/completions") {
      response.writeHead(200, { "Content-Type": "application/json" });
      response.end(
        JSON.stringify({ choices: [{ message: { content: "{}" } }] }),
      );
      return;
    }
    response.writeHead(404);
    response.end();
  });
  await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
  const address = server.address();
  assert.ok(address && typeof address !== "string");
  const baseUrl = `http://127.0.0.1:${address.port}`;
  const previous = {
    base: process.env.RELAY_BASE_URL,
    key: process.env.RELAY_API_KEY,
    hmr: process.env.SLIDEV_HMR_URL,
    data: process.env.DATA_DIR,
  };
  process.env.RELAY_BASE_URL = baseUrl;
  process.env.RELAY_API_KEY = "relay-test-key";
  process.env.SLIDEV_HMR_URL = `${baseUrl}/api/hmr`;
  try {
    const image = await new RelayModelAdapter().generateImage({
      prompt: "A test presentation visual",
      width: 1600,
      height: 900,
    });
    assert.equal(image.bytes.equals(png), true);
    assert.ok(imageRequestId.startsWith("relay_"));
    const worker = new SlidevQuickPreviewWorker();
    const project: any = {
      projectId: "p_test",
      name: "Test",
      currentDeckRevisionId: "deckrev_1",
    };
    const page: any = { pageId: "page_test", orderIndex: 0 };
    const preview = await worker.render(
      project,
      page,
      "Title",
      "Body",
      "editorial",
    );
    assert.equal(preview.engine, "slidev_hmr");
    assert.equal(hmrCalls, 1);
    const { directory, store } = await tempStore();
    process.env.DATA_DIR = directory;
    try {
      const service = new OnlineService(store, new EventBus(store));
      const imageProject = await service.createProject("owner_image", {
        name: "Image deck",
        themeId: "test",
        themeVersion: "1",
        slidesMarkdown: "# Visual\n\nA visual page without locked metrics.",
      });
      const session = await service.createWorkSession(
        "owner_image",
        imageProject.projectId,
        { workflowMode: "ppt_beautify", intent: "modify" },
      );
      const originalVersionId = imageProject.pages[0].currentVersionId;
      const turn = await service.createChatTurn("owner_image", {
        projectId: imageProject.projectId,
        deckRevisionId: imageProject.currentDeckRevisionId,
        target: { mode: "single", pageIds: [imageProject.pages[0].pageId] },
        message: "把图片改为抽象几何视觉",
        clientRevision: 0,
        sessionId: session.sessionId,
      });
      assert.equal(turn.operation.status, "planned");
      assert.equal(turn.operation.visualPreviews?.[0]?.status, "ready");
      assert.equal(imageProject.pages[0].currentVersionId, originalVersionId);
      assert.deepEqual(turn.operation.resultVersionIds, []);
      assert.equal(imageCalls, 2);
      const candidateArtifactId = turn.operation.visualPreviews?.[0]?.artifactId;
      assert.ok(candidateArtifactId);
      const candidateArtifact = service
        .artifacts("owner_image", imageProject.projectId)
        .find((artifact) => artifact.artifactId === candidateArtifactId);
      assert.equal(candidateArtifact?.versionId, null);
      const candidateBytes = await service.readVisualPreview(
        "owner_image",
        imageProject.projectId,
        candidateArtifactId,
      );
      assert.equal(candidateBytes.bytes.equals(png), true);
      const preconfirmationEvents = store.state.events.filter(
        (event) => event.operationId === turn.operation.operationId,
      );
      assert.deepEqual(
        preconfirmationEvents.map((event) => event.type),
        [
          "chat.plan.created",
          "preview.visual.ready",
          "edit.confirmation.required",
        ],
      );
      assert.equal(
        preconfirmationEvents.every(
          (event) => event.sessionId === session.sessionId,
        ),
        true,
      );
      const confirmed = await service.confirmOperation(
        "owner_image",
        imageProject.projectId,
        turn.operation.operationId,
      );
      assert.equal(confirmed.status, "completed");
      assert.equal(imageCalls, 2);
      const relayArtifact = service
        .artifacts("owner_image", imageProject.projectId)
        .find((artifact) => artifact.source === "relay_image");
      assert.ok(
        relayArtifact?.assetPath?.startsWith(
          `projects/owner_image/${imageProject.projectId}/pages/`,
        ),
      );
      const currentImageVersion = imageProject.pages[0].versions.find(
        (version) =>
          version.versionId === imageProject.pages[0].currentVersionId,
      );
      assert.equal(
        currentImageVersion?.visualArtifactId,
        candidateArtifactId,
      );
      assert.equal(relayArtifact?.versionId, currentImageVersion?.versionId);
      assert.equal(
        service
          .usage("owner_image", imageProject.projectId)
          .filter(
            (entry) =>
              entry.operationId === turn.operation.operationId &&
              entry.status === "settled",
          ).length,
        1,
      );
    } finally {
      await fs.rm(directory, { recursive: true, force: true });
    }
  } finally {
    if (previous.base === undefined) delete process.env.RELAY_BASE_URL;
    else process.env.RELAY_BASE_URL = previous.base;
    if (previous.key === undefined) delete process.env.RELAY_API_KEY;
    else process.env.RELAY_API_KEY = previous.key;
    if (previous.hmr === undefined) delete process.env.SLIDEV_HMR_URL;
    else process.env.SLIDEV_HMR_URL = previous.hmr;
    if (previous.data === undefined) delete process.env.DATA_DIR;
    else process.env.DATA_DIR = previous.data;
    await new Promise<void>((resolve, reject) =>
      server.close((error) => (error ? reject(error) : resolve())),
    );
  }
});

test("Office preview grants are hashed, range-aware, and revoked by sensitive mode", async () => {
  const { directory, store } = await tempStore();
  const previousDataDir = process.env.DATA_DIR;
  process.env.DATA_DIR = directory;
  try {
    const service = new OnlineService(store, new EventBus(store));
    const project = await service.createProject("owner_office", {
      name: "Office deck",
      themeId: "test",
      themeVersion: "1",
      slidesMarkdown: "# Office\n\nPreview grant test.",
    });
    const exportDirectory = path.join(directory, "exports");
    const artifactPath = path.join(exportDirectory, "office-test.pptx");
    const artifactBytes = Buffer.from("PK-office-preview-test");
    await fs.mkdir(exportDirectory, { recursive: true });
    await fs.writeFile(artifactPath, artifactBytes);
    await store.update((state) =>
      state.exports.push({
        exportId: "export_office_test",
        projectId: project.projectId,
        status: "completed",
        artifactPath,
        renderMode: "svg_fallback",
        qaWarnings: [],
        createdAt: new Date().toISOString(),
        completedAt: new Date().toISOString(),
        artifactName: "office-test.pptx",
      }),
    );
    const issued = await service.issueOfficePreviewGrant(
      "owner_office",
      project.projectId,
      "export_office_test",
    );
    assert.equal(JSON.stringify(store.state).includes(issued.token), false);
    const head = await service.readOfficePreviewGrant(issued.token, "HEAD");
    assert.equal(head.fullLength, artifactBytes.length);
    const range = await service.readOfficePreviewGrant(
      issued.token,
      "GET",
      "bytes=3-8",
    );
    assert.equal(range.bytes.equals(artifactBytes.subarray(3, 9)), true);
    assert.equal(range.contentRange, `bytes 3-8/${artifactBytes.length}`);
    await service.updateProjectSettings("owner_office", project.projectId, {
      sensitiveMode: true,
    });
    await assert.rejects(
      () => service.readOfficePreviewGrant(issued.token, "GET"),
      /Preview not found/i,
    );
    await assert.rejects(
      () =>
        service.issueOfficePreviewGrant(
          "owner_office",
          project.projectId,
          "export_office_test",
        ),
      /sensitive mode/i,
    );
  } finally {
    if (previousDataDir === undefined) delete process.env.DATA_DIR;
    else process.env.DATA_DIR = previousDataDir;
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test("document sessions persist parsed facts and block chat until a numeric conflict is resolved", async () => {
  const { directory, store } = await tempStore();
  const previous = {
    data: process.env.DATA_DIR,
    python: process.env.PYTHON_BIN,
  };
  process.env.DATA_DIR = directory;
  const workspacePython = path.resolve(
    process.cwd(),
    "..",
    "..",
    "..",
    ".venv",
    "Scripts",
    "python.exe",
  );
  try {
    await fs.access(workspacePython);
    process.env.PYTHON_BIN = workspacePython;
  } catch {
    /* Use configured Python in non-workspace CI. */
  }
  try {
    const service = new OnlineService(store, new EventBus(store));
    const project = await service.createProject("owner_documents", {
      name: "Document deck",
      themeId: "test",
      themeVersion: "1",
      slidesMarkdown: "# Metric\n\nOriginal page.",
    });
    const session = await service.createWorkSession(
      "owner_documents",
      project.projectId,
      { workflowMode: "document_import", intent: "reference" },
    );
    const first = await service.uploadDocument(
      "owner_documents",
      project.projectId,
      session.sessionId,
      {
        fileName: "first.md",
        contentType: "text/markdown",
        bytes: Buffer.from("# Report\n\nRetention rate is 42%.", "utf8"),
      },
    );
    const duplicate = await service.uploadDocument(
      "owner_documents",
      project.projectId,
      session.sessionId,
      {
        fileName: "duplicate.md",
        contentType: "text/markdown",
        bytes: Buffer.from("# Report\n\nRetention rate is 42%.", "utf8"),
      },
    );
    assert.equal(duplicate.duplicate, true);
    const second = await service.uploadDocument(
      "owner_documents",
      project.projectId,
      session.sessionId,
      {
        fileName: "second.md",
        contentType: "text/markdown",
        bytes: Buffer.from("# Report\n\nRetention rate is 43%.", "utf8"),
      },
    );
    await service.parseDocument(
      "owner_documents",
      project.projectId,
      first.document.documentId,
    );
    await service.parseDocument(
      "owner_documents",
      project.projectId,
      second.document.documentId,
    );
    const conflicts = service.listDocumentConflicts(
      "owner_documents",
      project.projectId,
      session.sessionId,
    );
    assert.equal(conflicts.length, 1);
    assert.equal(conflicts[0].severity, "blocking");
    await assert.rejects(
      () =>
        service.createChatTurn("owner_documents", {
          projectId: project.projectId,
          deckRevisionId: project.currentDeckRevisionId,
          target: { mode: "single", pageIds: [project.pages[0].pageId] },
          message: "标题改短",
          clientRevision: 0,
        }),
      /blocking document conflict/i,
    );
    await service.resolveDocumentConflict(
      "owner_documents",
      project.projectId,
      conflicts[0].conflictId,
      { selectedValue: "42%" },
    );
    assert.equal(store.state.documentVersions.length, 2);
    assert.equal(store.state.conflictResolutions.length, 1);
    assert.equal(
      store.state.conflictResolutions[0].action,
      "prefer_source",
    );
    assert.equal(
      store.state.conflictResolutions[0].selectedDocumentIds[0],
      first.document.documentId,
    );

    const third = await service.uploadDocument(
      "owner_documents",
      project.projectId,
      session.sessionId,
      {
        fileName: "third.md",
        contentType: "text/markdown",
        bytes: Buffer.from("# Report\n\nRetention rate is 44%.", "utf8"),
      },
    );
    await service.parseDocument(
      "owner_documents",
      project.projectId,
      third.document.documentId,
    );
    const keepBothConflict = service
      .listDocumentConflicts(
        "owner_documents",
        project.projectId,
        session.sessionId,
      )
      .find((conflict) => conflict.status === "open");
    assert.ok(keepBothConflict);
    await service.resolveDocumentConflict(
      "owner_documents",
      project.projectId,
      keepBothConflict.conflictId,
      { action: "keep_both", note: "Both reporting periods are valid." },
    );
    assert.equal(store.state.conflictResolutions.at(-1)?.action, "keep_both");
    assert.equal(
      store.state.conflictResolutions.at(-1)?.selectedDocumentIds.length,
      3,
    );

    const fourth = await service.uploadDocument(
      "owner_documents",
      project.projectId,
      session.sessionId,
      {
        fileName: "fourth.md",
        contentType: "text/markdown",
        bytes: Buffer.from("# Report\n\nRetention rate is 45%.", "utf8"),
      },
    );
    await service.parseDocument(
      "owner_documents",
      project.projectId,
      fourth.document.documentId,
    );
    const ignoreConflict = service
      .listDocumentConflicts(
        "owner_documents",
        project.projectId,
        session.sessionId,
      )
      .find((conflict) => conflict.status === "open");
    assert.ok(ignoreConflict);
    await service.resolveDocumentConflict(
      "owner_documents",
      project.projectId,
      ignoreConflict.conflictId,
      { action: "ignore", note: "Exclude this metric from the deck." },
    );
    assert.equal(store.state.conflictResolutions.at(-1)?.action, "ignore");
    assert.equal(
      store.state.conflictResolutions.at(-1)?.selectedDocumentIds.length,
      0,
    );
    const turn = await service.createChatTurn("owner_documents", {
      projectId: project.projectId,
      deckRevisionId: project.currentDeckRevisionId,
      target: { mode: "single", pageIds: [project.pages[0].pageId] },
      message: "标题改短",
      clientRevision: 0,
    });
    assert.equal(turn.operation.status, "completed");
    assert.equal(
      store.state.documents.every(
        (document) => !document.objectKey.includes("\\"),
      ),
      true,
    );
  } finally {
    if (previous.data === undefined) delete process.env.DATA_DIR;
    else process.env.DATA_DIR = previous.data;
    if (previous.python === undefined) delete process.env.PYTHON_BIN;
    else process.env.PYTHON_BIN = previous.python;
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test("document parser rejects archive and extracted text limit violations", async () => {
  const fixturePath = path.resolve(
    process.cwd(),
    "tests",
    "fixtures",
    "pptx-structure.pptx",
  );
  const { directory } = await tempStore();
  const markdownPath = path.join(directory, "oversized.md");
  const previous = {
    entries: process.env.MAX_DOCUMENT_ARCHIVE_ENTRIES,
    text: process.env.MAX_DOCUMENT_TEXT_CHARS,
    python: process.env.PYTHON_BIN,
  };
  const workspacePython = path.resolve(
    process.cwd(),
    "..",
    "..",
    "..",
    ".venv",
    "Scripts",
    "python.exe",
  );
  try {
    await fs.access(workspacePython);
    process.env.PYTHON_BIN = workspacePython;
  } catch {
    /* Keep the caller's parser configuration in other environments. */
  }
  try {
    process.env.MAX_DOCUMENT_ARCHIVE_ENTRIES = "1";
    await assert.rejects(
      () => parseControlledDocument(fixturePath),
      /archive contains .* entries/i,
    );
    delete process.env.MAX_DOCUMENT_ARCHIVE_ENTRIES;
    process.env.MAX_DOCUMENT_TEXT_CHARS = "10";
    await fs.writeFile(markdownPath, "This text exceeds ten characters.", "utf8");
    await assert.rejects(
      () => parseControlledDocument(markdownPath),
      /parsed document contains .* characters/i,
    );
  } finally {
    if (previous.entries === undefined)
      delete process.env.MAX_DOCUMENT_ARCHIVE_ENTRIES;
    else process.env.MAX_DOCUMENT_ARCHIVE_ENTRIES = previous.entries;
    if (previous.text === undefined) delete process.env.MAX_DOCUMENT_TEXT_CHARS;
    else process.env.MAX_DOCUMENT_TEXT_CHARS = previous.text;
    if (previous.python === undefined) delete process.env.PYTHON_BIN;
    else process.env.PYTHON_BIN = previous.python;
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test("document import structure plans preserve existing pages or create a confirmed new deck", async () => {
  const { directory, store } = await tempStore();
  const previous = {
    data: process.env.DATA_DIR,
    python: process.env.PYTHON_BIN,
  };
  process.env.DATA_DIR = directory;
  const workspacePython = path.resolve(
    process.cwd(),
    "..",
    "..",
    "..",
    ".venv",
    "Scripts",
    "python.exe",
  );
  try {
    await fs.access(workspacePython);
    process.env.PYTHON_BIN = workspacePython;
  } catch {
    /* Keep the caller's parser configuration in other environments. */
  }
  try {
    const service = new OnlineService(store, new EventBus(store));
    const project = await service.createProject("owner_structure", {
      name: "Structure deck",
      themeId: "test-theme",
      themeVersion: "2",
      slidesMarkdown:
        "# Existing cover\n\nIntro.\n\n---\n\n# Existing content\n\nDetails.",
    });
    const existingIds = project.pages.map((page) => page.pageId);
    const modifySession = await service.createWorkSession(
      "owner_structure",
      project.projectId,
      { workflowMode: "document_import", intent: "modify" },
    );
    const source = await service.uploadDocument(
      "owner_structure",
      project.projectId,
      modifySession.sessionId,
      {
        fileName: "outline.md",
        contentType: "text/markdown",
        bytes: Buffer.from(
          "# Revised outline\n\nRetention is 42%.\n\n# Detail\n\nKeep this section.",
          "utf8",
        ),
      },
    );
    await service.parseDocument(
      "owner_structure",
      project.projectId,
      source.document.documentId,
    );
    const modifyPlan = await service.createDocumentStructurePlan(
      "owner_structure",
      project.projectId,
      modifySession.sessionId,
    );
    assert.equal(modifyPlan.operation.status, "planned");
    assert.equal(modifyPlan.structurePlan.preservesPageCount, true);
    assert.deepEqual(
      modifyPlan.structurePlan.pages.map((page) => page.pageId),
      existingIds,
    );
    assert.deepEqual(
      service
        .getProject("owner_structure", project.projectId)
        .pages.filter((page) => !page.archived)
        .map((page) => page.pageId),
      existingIds,
    );
    const modified = await service.confirmOperation(
      "owner_structure",
      project.projectId,
      modifyPlan.operation.operationId,
    );
    assert.equal(modified.status, "completed");
    const modifiedPages = service
      .getProject("owner_structure", project.projectId)
      .pages.filter((page) => !page.archived)
      .sort((a, b) => a.orderIndex - b.orderIndex);
    assert.deepEqual(
      modifiedPages.map((page) => page.pageId),
      existingIds,
    );
    assert.equal(modifiedPages.every((page) => page.versions.length === 2), true);
    assert.equal(
      modifiedPages.every(
        (page) =>
          page.versions.at(-1)?.parentVersionId === page.versions[0].versionId,
      ),
      true,
    );

    const createSession = await service.createWorkSession(
      "owner_structure",
      project.projectId,
      { workflowMode: "document_import", intent: "create" },
    );
    const createSource = await service.uploadDocument(
      "owner_structure",
      project.projectId,
      createSession.sessionId,
      {
        fileName: "new-deck.md",
        contentType: "text/markdown",
        bytes: Buffer.from("# New deck\n\nA new narrative.", "utf8"),
      },
    );
    await service.parseDocument(
      "owner_structure",
      project.projectId,
      createSource.document.documentId,
    );
    const createPlan = await service.createDocumentStructurePlan(
      "owner_structure",
      project.projectId,
      createSession.sessionId,
      { pageBudget: 4, inheritTheme: false },
    );
    assert.equal(createPlan.structurePlan.totalPages, 4);
    assert.equal(createPlan.structurePlan.pages[0].kind, "cover");
    assert.equal(createPlan.structurePlan.pages[1].kind, "transition");
    assert.equal(createPlan.operation.structuredPlan.pageDelta.add.length, 4);
    const beforeConfirmActive = service
      .getProject("owner_structure", project.projectId)
      .pages.filter((page) => !page.archived);
    assert.equal(beforeConfirmActive.length, 2);
    await service.confirmOperation(
      "owner_structure",
      project.projectId,
      createPlan.operation.operationId,
    );
    const afterCreate = service.getProject(
      "owner_structure",
      project.projectId,
    );
    assert.equal(afterCreate.pages.filter((page) => !page.archived).length, 4);
    assert.equal(afterCreate.pages.filter((page) => page.archived).length, 2);
    assert.equal(afterCreate.themeId, "fastppt-editorial");
  } finally {
    if (previous.data === undefined) delete process.env.DATA_DIR;
    else process.env.DATA_DIR = previous.data;
    if (previous.python === undefined) delete process.env.PYTHON_BIN;
    else process.env.PYTHON_BIN = previous.python;
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test("PPTX parsing preserves slide order, stable page IDs, and image editability warnings", async () => {
  const fixturePath = path.resolve(
    process.cwd(),
    "tests",
    "fixtures",
    "pptx-structure.pptx",
  );
  const fixtureBytes = await fs.readFile(fixturePath);
  const { directory, store } = await tempStore();
  const previous = {
    data: process.env.DATA_DIR,
    python: process.env.PYTHON_BIN,
  };
  process.env.DATA_DIR = directory;
  const workspacePython = path.resolve(
    process.cwd(),
    "..",
    "..",
    "..",
    ".venv",
    "Scripts",
    "python.exe",
  );
  try {
    await fs.access(workspacePython);
    process.env.PYTHON_BIN = workspacePython;
  } catch {
    /* Keep the caller's parser configuration in other environments. */
  }
  try {
    const parsed = await parseControlledDocument(fixturePath);
    assert.equal(parsed.structure.pageCount, 2);
    assert.deepEqual(
      parsed.structure.slides?.map((slide) => slide.title),
      ["Structure overview", "Visual reference"],
    );
    assert.equal(
      parsed.structure.slides?.[0]?.editableLevel,
      "native_structure",
    );
    assert.deepEqual(parsed.structure.slides?.[1]?.nonEditableRegions, [
      "slide:2:image-region",
    ]);

    const service = new OnlineService(store, new EventBus(store));
    const project = await service.createProject("owner_pptx_structure", {
      name: "PPTX structure deck",
      themeId: "test",
      themeVersion: "1",
      slidesMarkdown: "# Existing\n\nThis page will retain its stable ID.",
    });
    const session = await service.createWorkSession(
      "owner_pptx_structure",
      project.projectId,
      { workflowMode: "ppt_beautify", intent: "modify" },
    );
    const uploaded = await service.uploadDocument(
      "owner_pptx_structure",
      project.projectId,
      session.sessionId,
      {
        fileName: "structure.pptx",
        contentType:
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        bytes: fixtureBytes,
      },
    );
    const first = await service.parseDocument(
      "owner_pptx_structure",
      project.projectId,
      uploaded.document.documentId,
    );
    assert.equal(first.parseStatus, "ready");
    const firstIds = service
      .getProject("owner_pptx_structure", project.projectId)
      .pages.filter((page) => !page.archived)
      .sort((a, b) => a.orderIndex - b.orderIndex)
      .map((page) => page.pageId);
    assert.equal(firstIds.length, 2);
    assert.equal(firstIds[0], project.pages[0].pageId);
    const importedProject = service.getProject(
      "owner_pptx_structure",
      project.projectId,
    );
    assert.equal(
      importedProject.pages.find((page) => page.pageId === firstIds[0])
        ?.editableLevel,
      "native_structure",
    );
    const imagePage = importedProject.pages.find(
      (page) => page.pageId === firstIds[1],
    );
    assert.equal(imagePage?.editableLevel, "native_partial");
    assert.deepEqual(currentVersionForTest(imagePage)?.nonEditableRegions, [
      "slide:2:image-region",
    ]);

    await service.parseDocument(
      "owner_pptx_structure",
      project.projectId,
      uploaded.document.documentId,
    );
    const secondIds = service
      .getProject("owner_pptx_structure", project.projectId)
      .pages.filter((page) => !page.archived)
      .sort((a, b) => a.orderIndex - b.orderIndex)
      .map((page) => page.pageId);
    assert.deepEqual(secondIds, firstIds);
  } finally {
    if (previous.data === undefined) delete process.env.DATA_DIR;
    else process.env.DATA_DIR = previous.data;
    if (previous.python === undefined) delete process.env.PYTHON_BIN;
    else process.env.PYTHON_BIN = previous.python;
    await fs.rm(directory, { recursive: true, force: true });
  }
});

function currentVersionForTest(page: { currentVersionId: string; versions: Array<{ versionId: string; nonEditableRegions: string[] }> } | undefined) {
  return page?.versions.find((version) => version.versionId === page.currentVersionId);
}

import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { AuthService } from '../server/auth.js';
import { EventBus } from '../server/events.js';
import { buildEditPlan } from '../server/plans.js';
import { SlidevQuickPreviewWorker } from '../server/quickPreview.js';
import { downloadRemoteImage, RelayModelAdapter } from '../server/relay.js';
import { OnlineService } from '../server/service.js';
import { FileStore } from '../server/store.js';
import { DurableJobQueue } from '../server/jobQueue.js';
import { editPlanSchema } from '../shared/protocol.js';

async function tempStore(): Promise<{ directory: string; store: FileStore }> {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'fastppt-online-test-'));
  const store = new FileStore(directory);
  await store.init();
  return { directory, store };
}

test('structured plan reports locked fact removal and requires confirmation', async () => {
  const { directory, store } = await tempStore();
  try {
    const service = new OnlineService(store, new EventBus(store));
    const project = await service.createProject('owner_plan', { name: 'Fact deck', themeId: 'test', themeVersion: '1', slidesMarkdown: '# Metric\n\n42% baseline remains locked.' });
    const page = project.pages[0];
    const plan = buildEditPlan(project, [page], 'single', '正文改为：No metric remains in this body.');
    assert.equal(editPlanSchema.safeParse(plan).success, true);
    assert.equal(plan.requiresConfirmation, true);
    assert.ok(plan.factImpact.removed.some((value) => value.includes('42%')));
  } finally {
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test('confirmed operation removes a declared fact and group rollback restores its version snapshot', async () => {
  const { directory, store } = await tempStore();
  const previousDataDir = process.env.DATA_DIR;
  process.env.DATA_DIR = directory;
  try {
    const service = new OnlineService(store, new EventBus(store));
    const project = await service.createProject('owner_rollback', { name: 'Rollback deck', themeId: 'test', themeVersion: '1', slidesMarkdown: '# Metric\n\n42% baseline remains locked.' });
    const pageId = project.pages[0].pageId;
    const turn = await service.createChatTurn('owner_rollback', { projectId: project.projectId, deckRevisionId: project.currentDeckRevisionId, target: { mode: 'single', pageIds: [pageId] }, message: '正文改为：No metric remains in this body.', clientRevision: 0 });
    assert.equal(turn.operation.status, 'planned');
    const applied = await service.confirmOperation('owner_rollback', project.projectId, turn.operation.operationId);
    assert.equal(applied.status, 'completed');
    assert.equal(service.getPage('owner_rollback', project.projectId, pageId).factAnchors.some((fact) => fact.value === '42%'), false);
    await service.rollbackOperation('owner_rollback', project.projectId, applied.operationId);
    const restored = service.getPage('owner_rollback', project.projectId, pageId);
    assert.equal(restored.body.includes('42%'), true);
    assert.equal(restored.factAnchors.some((fact) => fact.value === '42%'), true);
  } finally {
    if (previousDataDir === undefined) delete process.env.DATA_DIR; else process.env.DATA_DIR = previousDataDir;
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test('fact replacement is explicit, confirmed, and carried into the next anchor snapshot', async () => {
  const { directory, store } = await tempStore();
  try {
    const service = new OnlineService(store, new EventBus(store));
    const project = await service.createProject('owner_change', { name: 'Fact change deck', themeId: 'test', themeVersion: '1', slidesMarkdown: '# Metric\n\n42% baseline remains locked.' });
    const pageId = project.pages[0].pageId;
    const turn = await service.createChatTurn('owner_change', { projectId: project.projectId, deckRevisionId: project.currentDeckRevisionId, target: { mode: 'single', pageIds: [pageId] }, message: '把 42% 改为 43%', clientRevision: 0 });
    assert.equal(turn.operation.status, 'planned');
    assert.ok(turn.operation.factImpact.changed.some((value) => value.includes('42%->43%')));
    const applied = await service.confirmOperation('owner_change', project.projectId, turn.operation.operationId);
    assert.equal(applied.status, 'completed');
    const page = service.getPage('owner_change', project.projectId, pageId);
    assert.equal(page.body.includes('43%'), true);
    assert.equal(page.body.includes('42%'), false);
    assert.equal(page.factAnchors.some((fact) => fact.value === '43%'), true);
    assert.equal(page.factAnchors.some((fact) => fact.value === '42%'), false);
  } finally {
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test('signed session survives AuthService recreation and WebSocket tickets are one time', async () => {
  const { directory, store } = await tempStore();
  const previousSecret = process.env.AUTH_SECRET;
  process.env.AUTH_SECRET = 'test-secret-at-least-local';
  try {
    const first = new AuthService(store);
    const login = await first.login('auth@example.test', 'Auth Test');
    const recreated = new AuthService(store);
    assert.equal(recreated.verify(login.token)?.email, 'auth@example.test');
    const ticket = recreated.issueWebSocketTicket(login.user.userId);
    assert.equal(recreated.consumeWebSocketTicket(ticket)?.userId, login.user.userId);
    assert.equal(recreated.consumeWebSocketTicket(ticket), null);
    await recreated.revoke(login.token);
    assert.equal(new AuthService(store).verify(login.token), null);
  } finally {
    if (previousSecret === undefined) delete process.env.AUTH_SECRET; else process.env.AUTH_SECRET = previousSecret;
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test('production allowlist and access code can bootstrap the first user', async () => {
  const { directory, store } = await tempStore();
  const previous = {
    nodeEnv: process.env.NODE_ENV,
    secret: process.env.AUTH_SECRET,
    code: process.env.AUTH_LOGIN_CODE,
    allowed: process.env.AUTH_ALLOWED_EMAILS,
    devLogin: process.env.ALLOW_DEV_LOGIN,
  };
  process.env.NODE_ENV = 'production';
  process.env.AUTH_SECRET = 'production-test-secret-at-least-32-characters';
  process.env.AUTH_LOGIN_CODE = 'invite-code';
  process.env.AUTH_ALLOWED_EMAILS = 'first@example.test';
  process.env.ALLOW_DEV_LOGIN = 'false';
  try {
    const auth = new AuthService(store);
    const login = await auth.login('first@example.test', 'First User', 'invite-code');
    assert.equal(login.user.email, 'first@example.test');
    assert.equal(store.state.users.length, 1);
    assert.equal(auth.verify(login.token)?.userId, login.user.userId);
    await assert.rejects(() => auth.login('blocked@example.test', 'Blocked', 'invite-code'), /not allowed/i);
    await assert.rejects(() => auth.login('first@example.test', 'First User', 'wrong-code'), /invalid access code/i);
  } finally {
    if (previous.nodeEnv === undefined) delete process.env.NODE_ENV; else process.env.NODE_ENV = previous.nodeEnv;
    if (previous.secret === undefined) delete process.env.AUTH_SECRET; else process.env.AUTH_SECRET = previous.secret;
    if (previous.code === undefined) delete process.env.AUTH_LOGIN_CODE; else process.env.AUTH_LOGIN_CODE = previous.code;
    if (previous.allowed === undefined) delete process.env.AUTH_ALLOWED_EMAILS; else process.env.AUTH_ALLOWED_EMAILS = previous.allowed;
    if (previous.devLogin === undefined) delete process.env.ALLOW_DEV_LOGIN; else process.env.ALLOW_DEV_LOGIN = previous.devLogin;
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test('remote image download rejects unsafe targets, redirects, media types, and oversized bodies', async () => {
  const publicLookup = async (): Promise<Array<{ address: string }>> => [{ address: '93.184.216.34' }];
  let fetchCalls = 0;
  const unusedFetch = (async (): Promise<Response> => {
    fetchCalls += 1;
    throw new Error('fetch should not be reached');
  }) as typeof fetch;
  await assert.rejects(() => downloadRemoteImage('http://images.example.test/a.png', undefined, { fetchImpl: unusedFetch, lookup: publicLookup }), /must use HTTPS/i);
  await assert.rejects(() => downloadRemoteImage('https://127.0.0.1/a.png', undefined, { fetchImpl: unusedFetch, lookup: publicLookup }), /private or non-routable/i);
  await assert.rejects(() => downloadRemoteImage('https://169.254.169.254/latest/meta-data', undefined, { fetchImpl: unusedFetch, lookup: publicLookup }), /private or non-routable/i);
  assert.equal(fetchCalls, 0);

  const redirectFetch = (async (): Promise<Response> => new Response(null, { status: 302, headers: { Location: 'https://10.0.0.1/secret.png' } })) as typeof fetch;
  await assert.rejects(() => downloadRemoteImage('https://images.example.test/a.png', undefined, { fetchImpl: redirectFetch, lookup: publicLookup }), /private or non-routable/i);

  const textFetch = (async (): Promise<Response> => new Response('not an image', { status: 200, headers: { 'Content-Type': 'text/plain' } })) as typeof fetch;
  await assert.rejects(() => downloadRemoteImage('https://images.example.test/a.png', undefined, { fetchImpl: textFetch, lookup: publicLookup }), /unsupported media type/i);

  const disguisedFetch = (async (): Promise<Response> => new Response('not a png', { status: 200, headers: { 'Content-Type': 'image/png' } })) as typeof fetch;
  await assert.rejects(() => downloadRemoteImage('https://images.example.test/a.png', undefined, { fetchImpl: disguisedFetch, lookup: publicLookup }), /do not match/i);

  const declaredLargeFetch = (async (): Promise<Response> => new Response('x', { status: 200, headers: { 'Content-Type': 'image/png', 'Content-Length': '100' } })) as typeof fetch;
  await assert.rejects(() => downloadRemoteImage('https://images.example.test/a.png', undefined, { fetchImpl: declaredLargeFetch, lookup: publicLookup, maxBytes: 8 }), /exceeded the 8 byte limit/i);

  const streamedLargeFetch = (async (): Promise<Response> => new Response(new ReadableStream({
    start(controller) {
      controller.enqueue(new Uint8Array([1, 2, 3, 4]));
      controller.enqueue(new Uint8Array([5, 6, 7, 8]));
      controller.close();
    },
  }), { status: 200, headers: { 'Content-Type': 'image/png' } })) as typeof fetch;
  await assert.rejects(() => downloadRemoteImage('https://images.example.test/a.png', undefined, { fetchImpl: streamedLargeFetch, lookup: publicLookup, maxBytes: 6 }), /while streaming/i);

  const png = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  const imageFetch = (async (): Promise<Response> => new Response(png, { status: 200, headers: { 'Content-Type': 'image/png', 'Content-Length': String(png.length) } })) as typeof fetch;
  const downloaded = await downloadRemoteImage('https://images.example.test/a.png', undefined, { fetchImpl: imageFetch, lookup: publicLookup, maxBytes: 16 });
  assert.equal(downloaded.bytes.equals(png), true);
  assert.equal(downloaded.mimeType, 'image/png');
});

test('durable job queue persists terminal state and enforces its concurrency limit', async () => {
  const { directory, store } = await tempStore();
  try {
    const queue = new DurableJobQueue(store, 1);
    let active = 0;
    let peak = 0;
    const makeJob = (jobId: string) => ({ jobId, projectId: 'queue-project', kind: 'edit' as const, status: 'queued' as const, payload: {}, createdAt: new Date().toISOString(), completedAt: null });
    const runner = async (): Promise<void> => {
      active += 1;
      peak = Math.max(peak, active);
      await new Promise((resolve) => setTimeout(resolve, 12));
      active -= 1;
    };
    await Promise.all([queue.enqueue(makeJob('job_a'), runner, true), queue.enqueue(makeJob('job_b'), runner, true)]);
    assert.equal(peak, 1);
    assert.equal(store.state.jobs.every((job) => job.status === 'completed'), true);
  } finally {
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test('retry operation targets only failed pages and keeps a parent audit link', async () => {
  const { directory, store } = await tempStore();
  const previous = { base: process.env.RELAY_BASE_URL, key: process.env.RELAY_API_KEY, data: process.env.DATA_DIR };
  delete process.env.RELAY_BASE_URL;
  delete process.env.RELAY_API_KEY;
  process.env.DATA_DIR = directory;
  try {
    const service = new OnlineService(store, new EventBus(store));
    const project = await service.createProject('owner_retry', { name: 'Retry deck', themeId: 'test', themeVersion: '1', slidesMarkdown: '# Visual\n\nA page that requests a generated image.' });
    const pageId = project.pages[0].pageId;
    const first = await service.createChatTurn('owner_retry', { projectId: project.projectId, deckRevisionId: project.currentDeckRevisionId, target: { mode: 'single', pageIds: [pageId] }, message: '替换当前图片', clientRevision: 0 });
    assert.equal(first.operation.status, 'failed');
    assert.deepEqual(first.operation.failedPageIds, [pageId]);
    const retry = await service.retryFailedPages('owner_retry', project.projectId, first.operation.operationId);
    assert.equal(retry.parentOperationId, first.operation.operationId);
    assert.deepEqual(retry.resolvedPageIds, [pageId]);
    assert.equal(retry.status, 'failed');
    const retryEntries = service.usage('owner_retry', project.projectId).filter((entry) => entry.operationId === retry.operationId);
    assert.equal(retryEntries.length, 1);
    assert.equal(retryEntries[0].status, 'refunded');
  } finally {
    if (previous.base === undefined) delete process.env.RELAY_BASE_URL; else process.env.RELAY_BASE_URL = previous.base;
    if (previous.key === undefined) delete process.env.RELAY_API_KEY; else process.env.RELAY_API_KEY = previous.key;
    if (previous.data === undefined) delete process.env.DATA_DIR; else process.env.DATA_DIR = previous.data;
    await fs.rm(directory, { recursive: true, force: true });
  }
});

test('Relay image and Slidev adapters perform real configured HTTP calls', async () => {
  const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=', 'base64');
  let imageRequestId = '';
  let hmrCalls = 0;
  const server = http.createServer((request, response) => {
    if (request.url === '/v1/images') {
      imageRequestId = String(request.headers['x-request-id'] || '');
      response.writeHead(200, { 'Content-Type': 'application/json' });
      response.end(JSON.stringify({ data: [{ b64_json: png.toString('base64'), mime_type: 'image/png' }] }));
      return;
    }
    if (request.url === '/api/hmr') {
      hmrCalls += 1;
      response.writeHead(204);
      response.end();
      return;
    }
    if (request.url === '/v1/chat/completions') {
      response.writeHead(200, { 'Content-Type': 'application/json' });
      response.end(JSON.stringify({ choices: [{ message: { content: '{}' } }] }));
      return;
    }
    response.writeHead(404);
    response.end();
  });
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  const address = server.address();
  assert.ok(address && typeof address !== 'string');
  const baseUrl = `http://127.0.0.1:${address.port}`;
  const previous = { base: process.env.RELAY_BASE_URL, key: process.env.RELAY_API_KEY, hmr: process.env.SLIDEV_HMR_URL, data: process.env.DATA_DIR };
  process.env.RELAY_BASE_URL = baseUrl;
  process.env.RELAY_API_KEY = 'relay-test-key';
  process.env.SLIDEV_HMR_URL = `${baseUrl}/api/hmr`;
  try {
    const image = await new RelayModelAdapter().generateImage({ prompt: 'A test presentation visual', width: 1600, height: 900 });
    assert.equal(image.bytes.equals(png), true);
    assert.ok(imageRequestId.startsWith('relay_'));
    const worker = new SlidevQuickPreviewWorker();
    const project: any = { projectId: 'p_test', name: 'Test', currentDeckRevisionId: 'deckrev_1' };
    const page: any = { pageId: 'page_test', orderIndex: 0 };
    const preview = await worker.render(project, page, 'Title', 'Body', 'editorial');
    assert.equal(preview.engine, 'slidev_hmr');
    assert.equal(hmrCalls, 1);
    const { directory, store } = await tempStore();
    process.env.DATA_DIR = directory;
    try {
      const service = new OnlineService(store, new EventBus(store));
      const imageProject = await service.createProject('owner_image', { name: 'Image deck', themeId: 'test', themeVersion: '1', slidesMarkdown: '# Visual\n\nA visual page without locked metrics.' });
      const turn = await service.createChatTurn('owner_image', { projectId: imageProject.projectId, deckRevisionId: imageProject.currentDeckRevisionId, target: { mode: 'single', pageIds: [imageProject.pages[0].pageId] }, message: '把图片改为抽象几何视觉', clientRevision: 0 });
      assert.equal(turn.operation.status, 'completed');
      const relayArtifact = service.artifacts('owner_image', imageProject.projectId).find((artifact) => artifact.source === 'relay_image');
      assert.ok(relayArtifact?.assetPath?.startsWith(`projects/owner_image/${imageProject.projectId}/pages/`));
      const currentImageVersion = imageProject.pages[0].versions.find((version) => version.versionId === imageProject.pages[0].currentVersionId);
      assert.equal(currentImageVersion?.visualArtifactId, relayArtifact?.artifactId);
      assert.equal(service.usage('owner_image', imageProject.projectId).some((entry) => entry.operationId === turn.operation.operationId && entry.status === 'settled'), true);
    } finally {
      await fs.rm(directory, { recursive: true, force: true });
    }
  } finally {
    if (previous.base === undefined) delete process.env.RELAY_BASE_URL; else process.env.RELAY_BASE_URL = previous.base;
    if (previous.key === undefined) delete process.env.RELAY_API_KEY; else process.env.RELAY_API_KEY = previous.key;
    if (previous.hmr === undefined) delete process.env.SLIDEV_HMR_URL; else process.env.SLIDEV_HMR_URL = previous.hmr;
    if (previous.data === undefined) delete process.env.DATA_DIR; else process.env.DATA_DIR = previous.data;
    await new Promise<void>((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
  }
});

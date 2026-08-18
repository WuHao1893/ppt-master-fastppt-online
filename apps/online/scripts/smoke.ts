import assert from 'node:assert/strict';
import WebSocket from 'ws';

const apiBase = process.env.API_URL || 'http://127.0.0.1:8787';
const wsBase = process.env.WS_URL || 'ws://127.0.0.1:8788';

async function request<T>(token: string | undefined, path: string, init: RequestInit = {}): Promise<T> {
  const headers = new Headers(init.headers);
  if (init.body) headers.set('Content-Type', 'application/json');
  if (token) headers.set('Authorization', `Bearer ${token}`);
  const response = await fetch(`${apiBase}${path}`, { ...init, headers });
  const body = await response.json().catch(() => ({}));
  assert.equal(response.ok, true, `${init.method || 'GET'} ${path}: ${JSON.stringify(body)}`);
  return body as T;
}

async function waitForExport(token: string, projectId: string, exportId: string): Promise<any> {
  const deadline = Date.now() + Number(process.env.EXPORT_WAIT_MS || 120_000);
  while (Date.now() < deadline) {
    const result = await request<{ export: any }>(token, `/api/v1/projects/${projectId}/exports/${exportId}`);
    if (result.export.status === 'completed' || result.export.status === 'failed') return result.export;
    await new Promise((resolve) => setTimeout(resolve, 250));
  }
  throw new Error('Export did not finish in the smoke-test window.');
}

async function main(): Promise<void> {
  const health = await request<{ renderer: string; exporter: string }>(undefined, '/api/v1/health');
  assert.equal(health.exporter, 'ppt_master_svg_to_drawingml');
  const expectedPageStatus = health.renderer === 'powerpoint_com' ? 'authoritative' : 'svg_fallback';
  const login = await request<{ token: string }>(undefined, '/api/v1/auth/login', { method: 'POST', body: JSON.stringify({ email: `smoke-${Date.now()}@fastppt.local`, name: 'Smoke Runner' }) });
  const markdown = '# Cover 2026\n\n38% baseline and 2026 target.\n---\n# Two-column evidence\n\nKeep all numbers: 38% and 68%.\n---\n# Timeline delivery\n\nThree stages from plan to QA.';
  const created = await request<{ project: any }>(login.token, '/api/v1/projects', { method: 'POST', body: JSON.stringify({ name: 'Smoke Deck', slidesMarkdown: markdown }) });
  const project = created.project;
  assert.equal(project.pages.length, 3);
  assert.equal(new Set(project.pages.map((page: any) => page.pageId)).size, 3, 'page_id must be stable and unique');
  const [first, second, third] = project.pages;
  assert.equal(first.orderIndex, 0);

  const wsEvents: string[] = [];
  const wsSeqs: number[] = [];
  const wsTicket = await request<{ ticket: string }>(login.token, '/api/v1/auth/ws-ticket', { method: 'POST' });
  const socket = new WebSocket(`${wsBase}/?projectId=${project.projectId}&afterSeq=0`, [`fastppt-ticket.${wsTicket.ticket}`]);
  socket.on('message', (payload) => { try { const event = JSON.parse(payload.toString()); wsEvents.push(event.type); wsSeqs.push(event.seq); } catch { /* ignore */ } });
  await new Promise<void>((resolve, reject) => { socket.once('open', () => resolve()); socket.once('error', reject); });

  const single = await request<{ operation: any }>(login.token, `/api/v1/projects/${project.projectId}/chat/turns`, { method: 'POST', body: JSON.stringify({ projectId: project.projectId, deckRevisionId: project.currentDeckRevisionId, target: { mode: 'single', pageIds: [first.pageId] }, message: '标题改短，保留所有数字；右侧改为两列', clientRevision: 0 }) });
  assert.equal(single.operation.confirmationRequired, false);
  assert.equal(single.operation.status, 'completed');
  let refreshed = (await request<{ project: any }>(login.token, `/api/v1/projects/${project.projectId}`)).project;
  const refreshedFirst = refreshed.pages.find((page: any) => page.pageId === first.pageId);
  assert.equal(refreshedFirst.versions.length, 2);
  assert.equal(refreshedFirst.status, expectedPageStatus);

  const multiPlan = await request<{ operation: any }>(login.token, `/api/v1/projects/${project.projectId}/chat/turns`, { method: 'POST', body: JSON.stringify({ projectId: project.projectId, deckRevisionId: refreshed.currentDeckRevisionId, target: { mode: 'multi', pageIds: [first.pageId, second.pageId] }, message: '把标题改短并统一样式', clientRevision: 1 }) });
  assert.equal(multiPlan.operation.confirmationRequired, true);
  assert.equal(multiPlan.operation.status, 'planned');
  const confirmed = await request<{ operation: any }>(login.token, `/api/v1/projects/${project.projectId}/edit-operations/${multiPlan.operation.operationId}/confirm`, { method: 'POST', body: JSON.stringify({ operationId: multiPlan.operation.operationId }) });
  assert.equal(confirmed.operation.status, 'completed');
  const groupRollback = await request<{ operation: any }>(login.token, `/api/v1/projects/${project.projectId}/edit-operations/${multiPlan.operation.operationId}/rollback`, { method: 'POST' });
  assert.equal(groupRollback.operation.status, 'rolled_back', 'multi-page operations must support one-click group rollback');

  refreshed = (await request<{ project: any }>(login.token, `/api/v1/projects/${project.projectId}`)).project;
  const globalPlan = await request<{ operation: any }>(login.token, `/api/v1/projects/${project.projectId}/chat/turns`, { method: 'POST', body: JSON.stringify({ projectId: project.projectId, deckRevisionId: refreshed.currentDeckRevisionId, target: { mode: 'global', pageIds: [] }, message: '把所有标题过长的页改成一句话', clientRevision: 2 }) });
  assert.equal(globalPlan.operation.confirmationRequired, true);
  assert.ok(Object.keys(globalPlan.operation.structuredPlan.candidateReasons || {}).length > 0, 'global candidates need explanations');
  await request(login.token, `/api/v1/projects/${project.projectId}/edit-operations/${globalPlan.operation.operationId}/cancel`, { method: 'POST' });

  refreshed = (await request<{ project: any }>(login.token, `/api/v1/projects/${project.projectId}`)).project;
  const versions = refreshed.pages.find((page: any) => page.pageId === first.pageId).versions;
  const oldest = versions[0].versionId;
  await request(login.token, `/api/v1/projects/${project.projectId}/versions/${oldest}/rollback`, { method: 'POST' });
  const rolled = (await request<{ project: any }>(login.token, `/api/v1/projects/${project.projectId}`)).project;
  assert.equal(rolled.pages.find((page: any) => page.pageId === first.pageId).currentVersionId, oldest);
  const reordered = await request<{ pages: any[] }>(login.token, `/api/v1/projects/${project.projectId}/pages/reorder`, { method: 'POST', body: JSON.stringify({ pageIds: [third.pageId, second.pageId, first.pageId] }) });
  assert.deepEqual(reordered.pages.map((page) => page.pageId), [third.pageId, second.pageId, first.pageId]);
  const thirdVersionCount = rolled.pages.find((page: any) => page.pageId === third.pageId).versions.length;
  await request(login.token, `/api/v1/projects/${project.projectId}/pages/${third.pageId}/archive`, { method: 'POST' });
  const archivedState = (await request<{ project: any }>(login.token, `/api/v1/projects/${project.projectId}`)).project;
  assert.equal(archivedState.pages.find((page: any) => page.pageId === third.pageId).archived, true);
  assert.equal(archivedState.pages.find((page: any) => page.pageId === third.pageId).versions.length, thirdVersionCount, 'soft archive must retain page history');
  await request(login.token, `/api/v1/projects/${project.projectId}/pages/${third.pageId}/restore`, { method: 'POST' });

  const exportResponse = await request<{ export: any }>(login.token, `/api/v1/projects/${project.projectId}/render`, { method: 'POST' });
  const exportJob = await waitForExport(login.token, project.projectId, exportResponse.export.exportId);
  assert.equal(exportJob.status, 'completed', JSON.stringify(exportJob));
  assert.equal(exportJob.exportEngine, 'ppt_master_svg_to_drawingml');
  assert.ok(['passed', 'passed-with-warnings'].includes(exportJob.qaStatus), JSON.stringify(exportJob));
  assert.equal(exportJob.artifactPath, null, 'public export DTO must not expose an absolute artifact path');
  assert.equal(exportJob.artifactObjectKey, null, 'public export DTO must not expose an internal object key');
  assert.ok(exportJob.artifactName);
  const downloadResponse = await fetch(`${apiBase}/api/v1/projects/${project.projectId}/exports/${exportJob.exportId}/download`, { headers: { Authorization: `Bearer ${login.token}` } });
  assert.equal(downloadResponse.ok, true);
  assert.equal((await downloadResponse.arrayBuffer()).byteLength > 1000, true, 'download must contain a PPTX artifact');
  const imported = await request<{ project: any; importedInto: string }>(login.token, `/api/v1/projects/${project.projectId}/import`, { method: 'POST', body: JSON.stringify({ slidesMarkdown: markdown }) });
  assert.equal(imported.importedInto, project.projectId);
  assert.deepEqual(imported.project.pages.filter((page: any) => !page.archived).sort((a: any, b: any) => a.orderIndex - b.orderIndex).map((page: any) => page.pageId), [third.pageId, second.pageId, first.pageId], 'source import must preserve positional page IDs after reorder');
  const audit = await request<{ audit: any[] }>(login.token, `/api/v1/projects/${project.projectId}/audit`);
  assert.ok(audit.audit.length >= 5, 'audit log should record login, project, chat, rollback and export activity');
  const artifacts = await request<{ artifacts: any[] }>(login.token, `/api/v1/projects/${project.projectId}/artifacts`);
  assert.ok(artifacts.artifacts.some((artifact) => artifact.provenance && artifact.source === 'deterministic_svg'), 'preview provenance must be queryable');
  const snapshots = await request<{ promptSnapshots: any[] }>(login.token, `/api/v1/projects/${project.projectId}/prompt-snapshots`);
  assert.ok(snapshots.promptSnapshots.some((snapshot) => snapshot.prompt && snapshot.hash), 'complete prompt snapshots must be queryable');
  const beforeSplit = imported.project.pages.find((page: any) => page.pageId === second.pageId).versions.length;
  const splitPlan = await request<{ operation: any }>(login.token, `/api/v1/projects/${project.projectId}/pages/${second.pageId}/split`, { method: 'POST' });
  assert.equal(splitPlan.operation.status, 'planned');
  assert.equal(splitPlan.operation.confirmationRequired, true, 'page-count changes must require confirmation');
  const split = await request<{ operation: any }>(login.token, `/api/v1/projects/${project.projectId}/edit-operations/${splitPlan.operation.operationId}/confirm`, { method: 'POST', body: JSON.stringify({ operationId: splitPlan.operation.operationId }) });
  assert.equal(split.operation.resultVersionIds.length, 2);
  const afterSplit = (await request<{ project: any }>(login.token, `/api/v1/projects/${project.projectId}`)).project;
  const splitSource = afterSplit.pages.find((page: any) => page.pageId === second.pageId);
  const splitContinuation = afterSplit.pages.find((page: any) => page.pageId === split.operation.resolvedPageIds[1]);
  assert.equal(splitSource.versions.length, beforeSplit + 1, 'split must append a source-page version without replacing its page ID');
  assert.ok(splitContinuation && splitContinuation.pageId !== second.pageId);
  const factPage = afterSplit.pages.find((page: any) => page.factAnchors.some((fact: any) => fact.value === '38%'));
  assert.ok(factPage, 'smoke deck must retain the 38% fact anchor');
  const protectedTurn = await request<{ operation: any }>(login.token, `/api/v1/projects/${project.projectId}/chat/turns`, { method: 'POST', body: JSON.stringify({ projectId: project.projectId, deckRevisionId: afterSplit.currentDeckRevisionId, target: { mode: 'single', pageIds: [factPage.pageId] }, message: '正文改为：This replacement omits the locked metric.', clientRevision: 3 }) });
  assert.equal(protectedTurn.operation.confirmationRequired, true);
  assert.ok(protectedTurn.operation.factImpact.removed.some((value: string) => value.includes('38%')), 'fact impact must identify removed locked metrics');
  await request(login.token, `/api/v1/projects/${project.projectId}/edit-operations/${protectedTurn.operation.operationId}/cancel`, { method: 'POST' });
  await new Promise((resolve) => setTimeout(resolve, 180));
  assert.ok(wsEvents.includes('preview.quick.ready'), `expected quick preview event, got ${wsEvents.join(',')}`);
  if (health.renderer === 'powerpoint_com') assert.ok(wsEvents.includes('preview.pptx.ready'), `expected authoritative PowerPoint event, got ${wsEvents.join(',')}`);
  else assert.ok(wsEvents.includes('render.warning'), `expected SVG fallback warning, got ${wsEvents.join(',')}`);
  assert.ok(wsSeqs.every((seq, index) => index === 0 || seq > wsSeqs[index - 1]), `event sequence must be monotonic, got ${wsSeqs.join(',')}`);
  socket.close();
  console.log(JSON.stringify({ ok: true, projectId: project.projectId, pages: rolled.pages.length, events: wsEvents.slice(-8), export: exportJob.renderMode }, null, 2));
}

main().catch((error) => { console.error(error); process.exitCode = 1; });

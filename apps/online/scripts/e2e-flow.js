async (page) => {
  const assert = (condition, message) => {
    if (!condition) throw new Error(message);
  };
  const errors = [];
  page.on('pageerror', (error) => errors.push(error.message));
  const health = await page.evaluate(async () => (await fetch('/api/v1/health')).json());
  const unique = Date.now();
  await page.locator('input[type="email"]').fill(`e2e-${unique}@fastppt.local`);
  await page.locator('.auth-panel input:not([type="email"]):not([type="password"])').fill('E2E Editor');
  await page.getByRole('button', { name: '\u8fdb\u5165\u5de5\u4f5c\u53f0' }).click();
  await page.locator('.app-shell, .empty-workspace').first().waitFor({ state: 'visible' });

  const markdown = '# Metric page\n\n42% baseline must remain auditable.\n---\n# Second page with a deliberately long heading for batch editing\n\nStable page IDs survive grouped changes.';
  const projectName = `E2E Browser Deck ${unique}`;
  const created = await page.evaluate(async ({ name, slidesMarkdown }) => {
    const response = await fetch('/api/v1/projects', {
      method: 'POST',
      credentials: 'include',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, slidesMarkdown }),
    });
    return { ok: response.ok, body: await response.json() };
  }, { name: projectName, slidesMarkdown: markdown });
  assert(created.ok, `Project creation failed: ${JSON.stringify(created.body)}`);
  await page.reload();
  await page.locator('.app-shell').waitFor({ state: 'visible' });
  await page.getByText(projectName, { exact: true }).first().waitFor({ state: 'visible' });

  const cookies = await page.context().cookies();
  const cookieHeader = cookies.map((cookie) => `${cookie.name}=${cookie.value}`).join('; ');
  const webOrigin = page.url().replace(/\/$/, '');
  const firstPageId = created.body.project.pages[0].pageId;
  await page.context().setOffline(true);
  await page.waitForTimeout(900);
  const archiveResponse = await page.request.post(`${webOrigin}/api/v1/projects/${created.body.project.projectId}/pages/${firstPageId}/archive`, { headers: { Cookie: cookieHeader } });
  assert(archiveResponse.ok(), `Offline archive event setup failed: ${archiveResponse.status()}`);
  const restoreResponse = await page.request.post(`${webOrigin}/api/v1/projects/${created.body.project.projectId}/pages/${firstPageId}/restore`, { headers: { Cookie: cookieHeader } });
  assert(restoreResponse.ok(), `Offline restore event setup failed: ${restoreResponse.status()}`);
  await page.context().setOffline(false);
  await page.waitForFunction(() => document.querySelector('.event-note')?.textContent?.includes('page · restored'), undefined, { timeout: 20000 });
  assert(await page.locator('.page-row').filter({ hasText: 'Metric page' }).isVisible(), 'Missed-event replay did not restore the archived page.');

  const composer = page.locator('.composer textarea');
  await composer.fill('\u628a 42% \u6539\u4e3a 43%');
  await page.locator('.send-button').click();
  await page.locator('.confirmation-card').waitFor({ state: 'visible' });
  assert((await page.locator('.slide-render').innerText()).includes('42%'), 'Fact changed before confirmation.');
  await page.locator('.confirmation-actions .primary-button').click();
  await page.waitForFunction(async (projectId) => {
    const response = await fetch(`/api/v1/projects/${projectId}`, { credentials: 'include' });
    const body = await response.json();
    return body.project?.pages?.[0]?.body?.includes('43%');
  }, created.body.project.projectId);
  if (health.renderer === 'powerpoint_com') {
    const image = page.locator('.authoritative-render img[data-artifact-id]');
    await image.waitFor({ state: 'visible', timeout: 120000 });
    const imageEvidence = await image.evaluate(async (node) => {
      const response = await fetch(node.src, { credentials: 'include' });
      const bytes = new Uint8Array(await response.arrayBuffer());
      return {
        complete: node.complete,
        naturalWidth: node.naturalWidth,
        naturalHeight: node.naturalHeight,
        contentType: response.headers.get('content-type'),
        status: response.status,
        png: bytes.length > 8 && bytes.slice(0, 8).every((value, index) => value === [137, 80, 78, 71, 13, 10, 26, 10][index]),
      };
    });
    assert(imageEvidence.complete && imageEvidence.naturalWidth >= 1000 && imageEvidence.naturalHeight >= 500, `Authoritative image was not decoded: ${JSON.stringify(imageEvidence)}`);
    assert(imageEvidence.status === 200 && imageEvidence.contentType === 'image/png' && imageEvidence.png, `Authoritative content was not a PNG: ${JSON.stringify(imageEvidence)}`);
    await page.screenshot({ path: 'output/playwright/authority-proof-latest.png', fullPage: true });
  }
  const currentPageState = await page.evaluate(async (projectId) => {
    const response = await fetch(`/api/v1/projects/${projectId}`, { credentials: 'include' });
    const body = await response.json();
    return { versionId: body.project.pages[0].currentVersionId, status: body.project.pages[0].status };
  }, created.body.project.projectId);
  const currentVersionId = currentPageState.versionId;
  const currentStatus = currentPageState.status;
  await page.waitForFunction((versionId) => document.querySelector('.version-label')?.getAttribute('data-version-id') === versionId, currentVersionId);
  const initialVersion = created.body.project.pages[0].versions[0];
  await page.getByTitle('\u7248\u672c\u5386\u53f2').click();
  const initialHistoryRow = page.locator('.history-row').filter({ hasText: initialVersion.versionId });
  await initialHistoryRow.getByTitle('\u5bf9\u6bd4\u7248\u672c').click();
  await page.waitForFunction((versionId) => document.querySelector('.version-label')?.getAttribute('data-version-id') === versionId, initialVersion.versionId);
  assert(await page.locator('.version-label').getAttribute('data-version-id') === initialVersion.versionId, 'Comparison toolbar did not switch to the historical version id.');
  assert((await page.locator('.status-pill').innerText()).startsWith('\u5bf9\u6bd4 \u00b7 '), 'Comparison toolbar did not identify comparison mode.');
  const expectedComparisonStatus = initialVersion.previewKind === 'pptx_authoritative' ? 'authoritative' : initialVersion.previewKind;
  assert(await page.locator('.status-pill').getAttribute('data-page-status') === expectedComparisonStatus, 'Comparison toolbar did not use the historical preview status.');
  await page.screenshot({ path: 'output/playwright/compare-status-proof-latest.png', fullPage: true });
  await initialHistoryRow.getByTitle('\u5bf9\u6bd4\u7248\u672c').click();
  await page.waitForFunction((versionId) => document.querySelector('.version-label')?.getAttribute('data-version-id') === versionId, currentVersionId);
  assert(await page.locator('.version-label').getAttribute('data-version-id') === currentVersionId, 'Toolbar did not restore the current version id after comparison.');
  assert(await page.locator('.status-pill').getAttribute('data-page-status') === currentStatus, 'Toolbar did not restore the current page status after comparison.');
  const rollback = page.getByRole('button', { name: '\u64a4\u9500\u672c\u6b21\u64cd\u4f5c' }).last();
  await rollback.waitFor({ state: 'visible' });
  await rollback.click();
  await page.waitForFunction(async (projectId) => {
    const response = await fetch(`/api/v1/projects/${projectId}`, { credentials: 'include' });
    const body = await response.json();
    return body.project?.pages?.[0]?.body?.includes('42%');
  }, created.body.project.projectId);
  const rolledBackPage = await page.evaluate(async (projectId) => {
    const response = await fetch(`/api/v1/projects/${projectId}`, { credentials: 'include' });
    const body = await response.json();
    return body.project.pages[0];
  }, created.body.project.projectId);
  await page.waitForFunction((versionId) => document.querySelector(`.history-row[data-version-id="${versionId}"]`)?.getAttribute('data-current') === 'true', rolledBackPage.currentVersionId);
  const historyRows = page.locator('.history-row');
  const restoredCurrentRow = page.locator(`.history-row[data-version-id="${rolledBackPage.currentVersionId}"]`);
  assert(await page.locator('.history-row[data-current="true"]').count() === 1, 'History must mark exactly one current version after rollback.');
  assert((await restoredCurrentRow.locator('.history-info strong').innerText()) === '\u5f53\u524d\u7248\u672c', 'Restored version was not labeled as current.');
  assert(await restoredCurrentRow.getByTitle('\u6062\u590d\u6b64\u7248\u672c').count() === 0, 'Current version incorrectly offered a restore action.');
  const newestNonCurrentRow = historyRows.first();
  assert(await newestNonCurrentRow.getAttribute('data-current') === 'false', 'Newest immutable version was incorrectly marked current after rollback.');
  assert(await newestNonCurrentRow.getByTitle('\u6062\u590d\u6b64\u7248\u672c').count() === 1, 'Newest non-current version did not offer a restore action.');
  await page.screenshot({ path: 'output/playwright/history-current-label-proof-latest.png', fullPage: true });

  await page.locator('.check-box').nth(0).click();
  await page.locator('.check-box').nth(1).click();
  await composer.fill('\u6807\u9898\u6539\u4e3a E2E Batch');
  await page.locator('.send-button').click();
  await page.locator('.confirmation-card').waitFor({ state: 'visible' });
  assert((await page.locator('.candidate-list > div').count()) === 2, 'Multi-page plan did not list both target pages.');
  await page.locator('.confirmation-actions .primary-button').click();
  await page.waitForFunction(() => [...document.querySelectorAll('.page-row-title')].filter((node) => node.textContent === 'E2E Batch').length === 2);
  const groupRollback = page.getByRole('button', { name: '\u64a4\u9500\u672c\u6b21\u64cd\u4f5c' }).last();
  await groupRollback.waitFor({ state: 'visible' });
  await groupRollback.click();
  await page.waitForFunction(() => ![...document.querySelectorAll('.page-row-title')].some((node) => node.textContent === 'E2E Batch'));

  assert(errors.length === 0, `Browser errors: ${errors.join('; ')}`);
  return { ok: true, projectName };
}

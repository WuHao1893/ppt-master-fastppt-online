async (page) => {
  const assert = (condition, message) => {
    if (!condition) throw new Error(message);
  };
  const waitForCondition = async (label, callback, argument, options) => {
    try {
      await page.waitForFunction(callback, argument, options);
    } catch (error) {
      throw new Error(`${label}: ${error.message}`);
    }
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
  await waitForCondition('Missed-event replay', () => document.querySelector('.event-note')?.textContent?.includes('page · restored'), undefined, { timeout: 20000 });
  assert(await page.locator('.page-row').filter({ hasText: 'Metric page' }).isVisible(), 'Missed-event replay did not restore the archived page.');

  const composer = page.locator('.composer textarea');
  await composer.fill('\u628a 42% \u6539\u4e3a 43%');
  await page.locator('.send-button').click();
  await page.locator('.confirmation-card').waitFor({ state: 'visible' });
  assert((await page.locator('.slide-render').innerText()).includes('42%'), 'Fact changed before confirmation.');
  await Promise.all([
    page.waitForResponse((response) => response.url().includes('/edit-operations/') && response.url().endsWith('/confirm') && response.request().method() === 'POST'),
    page.locator('.confirmation-actions .primary-button').click(),
  ]);
  await waitForCondition('Confirmed fact replacement', async (projectId) => {
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
  await waitForCondition('Current version toolbar sync', (versionId) => document.querySelector('.version-label')?.getAttribute('data-version-id') === versionId, currentVersionId);
  const initialVersion = created.body.project.pages[0].versions[0];
  await page.getByTitle('\u7248\u672c\u5386\u53f2').click();
  const initialHistoryRow = page.locator('.history-row').filter({ hasText: initialVersion.versionId });
  await initialHistoryRow.getByTitle('\u5bf9\u6bd4\u7248\u672c').click();
  await waitForCondition('Historical comparison toolbar sync', (versionId) => document.querySelector('.version-label')?.getAttribute('data-version-id') === versionId, initialVersion.versionId);
  assert(await page.locator('.version-label').getAttribute('data-version-id') === initialVersion.versionId, 'Comparison toolbar did not switch to the historical version id.');
  assert((await page.locator('.status-pill').innerText()).startsWith('\u5bf9\u6bd4 \u00b7 '), 'Comparison toolbar did not identify comparison mode.');
  const expectedComparisonStatus = initialVersion.previewKind === 'pptx_authoritative' ? 'authoritative' : initialVersion.previewKind;
  assert(await page.locator('.status-pill').getAttribute('data-page-status') === expectedComparisonStatus, 'Comparison toolbar did not use the historical preview status.');
  await page.screenshot({ path: 'output/playwright/compare-status-proof-latest.png', fullPage: true });
  await initialHistoryRow.getByTitle('\u5bf9\u6bd4\u7248\u672c').click();
  await waitForCondition('Comparison exit toolbar sync', (versionId) => document.querySelector('.version-label')?.getAttribute('data-version-id') === versionId, currentVersionId);
  assert(await page.locator('.version-label').getAttribute('data-version-id') === currentVersionId, 'Toolbar did not restore the current version id after comparison.');
  assert(await page.locator('.status-pill').getAttribute('data-page-status') === currentStatus, 'Toolbar did not restore the current page status after comparison.');
  const rollback = page.getByRole('button', { name: '\u64a4\u9500\u672c\u6b21\u64cd\u4f5c' }).last();
  await rollback.waitFor({ state: 'visible' });
  await rollback.click();
  await waitForCondition('Single-page rollback', async (projectId) => {
    const response = await fetch(`/api/v1/projects/${projectId}`, { credentials: 'include' });
    const body = await response.json();
    return body.project?.pages?.[0]?.body?.includes('42%');
  }, created.body.project.projectId);
  const rolledBackPage = await page.evaluate(async (projectId) => {
    const response = await fetch(`/api/v1/projects/${projectId}`, { credentials: 'include' });
    const body = await response.json();
    return body.project.pages[0];
  }, created.body.project.projectId);
  await waitForCondition('Rollback history current marker', (versionId) => document.querySelector(`.history-row[data-version-id="${versionId}"]`)?.getAttribute('data-current') === 'true', rolledBackPage.currentVersionId);
  const historyRows = page.locator('.history-row');
  const restoredCurrentRow = page.locator(`.history-row[data-version-id="${rolledBackPage.currentVersionId}"]`);
  assert(await page.locator('.history-row[data-current="true"]').count() === 1, 'History must mark exactly one current version after rollback.');
  assert((await restoredCurrentRow.locator('.history-info strong').innerText()) === '\u5f53\u524d\u7248\u672c', 'Restored version was not labeled as current.');
  assert(await restoredCurrentRow.getByTitle('\u6062\u590d\u6b64\u7248\u672c').count() === 0, 'Current version incorrectly offered a restore action.');
  const newestNonCurrentRow = historyRows.first();
  assert(await newestNonCurrentRow.getAttribute('data-current') === 'false', 'Newest immutable version was incorrectly marked current after rollback.');
  assert(await newestNonCurrentRow.getByTitle('\u6062\u590d\u6b64\u7248\u672c').count() === 1, 'Newest non-current version did not offer a restore action.');
  await page.screenshot({ path: 'output/playwright/history-current-label-proof-latest.png', fullPage: true });
  await page.locator('.drawer-tabs button').filter({ hasText: '\u5bf9\u8bdd\u8bb0\u5f55' }).click();
  const historyOperation = page.locator('.history-operation').first();
  await historyOperation.waitFor({ state: 'visible' });
  assert(await historyOperation.locator('.history-audit-grid').count() === 1, 'History did not expose operation audit details.');
  const historyAuditText = await historyOperation.innerText();
  assert(historyAuditText.includes('\u64cd\u4f5c ID') && historyAuditText.includes('\u8017\u65f6') && historyAuditText.includes('\u6210\u672c') && historyAuditText.includes('\u6a21\u578b') && historyAuditText.includes('QA'), 'History audit fields are incomplete.');
  const statusSelect = page.getByLabel('\u6309\u72b6\u6001\u7b5b\u9009');
  const statusValues = await statusSelect.locator('option').evaluateAll((options) => options.map((option) => option.value));
  const terminalStatus = statusValues.includes('rolled_back') ? 'rolled_back' : 'completed';
  await statusSelect.selectOption(terminalStatus);
  assert(await page.locator(`.history-operation[data-operation-status="${terminalStatus}"]`).count() > 0, 'Status filter returned no records.');
  await page.getByLabel('\u6309\u9875\u9762\u7b5b\u9009').selectOption(firstPageId);
  assert(await page.locator('.history-operation').count() > 0, 'Page filter returned no records.');
  const conversationOption = await page.getByLabel('\u6309\u4f1a\u8bdd\u7b5b\u9009').locator('option').nth(1).getAttribute('value');
  assert(Boolean(conversationOption), 'Conversation filter did not expose a conversation.');
  await page.getByLabel('\u6309\u4f1a\u8bdd\u7b5b\u9009').selectOption(conversationOption);
  assert(await page.locator('.history-operation').count() > 0, 'Conversation filter returned no records.');
  await page.locator('.history-page-links button').first().click();
  assert(await page.locator('.version-label').getAttribute('data-version-id'), 'History page/version link did not locate a version.');
  await page.getByTitle('\u5173\u95ed\u5386\u53f2\u62bd\u5c49').click();

  await page.locator('.check-box').nth(0).click();
  await page.locator('.check-box').nth(1).click();
  await composer.fill('\u6807\u9898\u6539\u4e3a E2E Batch');
  await page.locator('.send-button').click();
  await page.locator('.confirmation-card').waitFor({ state: 'visible' });
  assert((await page.locator('.candidate-list > div').count()) === 2, 'Multi-page plan did not list both target pages.');
  await Promise.all([
    page.waitForResponse((response) => response.url().includes('/edit-operations/') && response.url().endsWith('/confirm') && response.request().method() === 'POST'),
    page.locator('.confirmation-actions .primary-button').click(),
  ]);
  await waitForCondition('Multi-page edit completion', () => [...document.querySelectorAll('.page-row-title')].filter((node) => node.textContent === 'E2E Batch').length === 2);
  const groupRollback = page.getByRole('button', { name: '\u64a4\u9500\u672c\u6b21\u64cd\u4f5c' }).last();
  await groupRollback.waitFor({ state: 'visible' });
  await groupRollback.click();
  await waitForCondition('Multi-page rollback', () => ![...document.querySelectorAll('.page-row-title')].some((node) => node.textContent === 'E2E Batch'));

  const importMode = page.getByTitle('\u5bfc\u5165\u6587\u6863');
  await Promise.all([
    page.waitForResponse((response) => response.url().includes('/api/v2/projects/') && response.url().endsWith('/work-sessions') && response.request().method() === 'POST'),
    importMode.click(),
  ]);
  await page.reload();
  await page.locator('.app-shell').waitFor({ state: 'visible' });
  await waitForCondition('Import mode restore', () => document.querySelector('[title="\u5bfc\u5165\u6587\u6863"]')?.classList.contains('active'));
  assert((await importMode.getAttribute('class'))?.includes('active'), 'Latest production mode was not restored after refresh.');
  const documentInput = page.locator('input[accept=".md,.docx,.pdf"]');
  await documentInput.setInputFiles([
    'tests/fixtures/e2e-metrics-a.md',
    'tests/fixtures/e2e-metrics-b.md',
  ]);
  await page.locator('.document-card').filter({ hasText: 'metrics-a.md' }).waitFor({ state: 'visible' });
  await page.locator('.document-card').filter({ hasText: 'metrics-b.md' }).waitFor({ state: 'visible' });
  await page.locator('.conflict-card').waitFor({ state: 'visible' });
  assert(await page.locator('.document-card').filter({ hasText: /SHA [a-f0-9]{12}/ }).count() === 2, 'Document pool did not show SHA summaries.');
  const contextToggles = page.locator('.document-context-toggle input');
  await Promise.all([
    page.waitForResponse((response) => response.url().endsWith('/context') && response.request().method() === 'PATCH'),
    contextToggles.first().evaluate((input) => input.click()),
  ]);
  await waitForCondition('Document context disabled', () => document.querySelector('.document-context-toggle input')?.checked === false);
  await page.waitForTimeout(1000);
  assert(await page.locator('.conflict-card').count() === 0, 'Conflict did not clear when a source left context.');
  await Promise.all([
    page.waitForResponse((response) => response.url().endsWith('/context') && response.request().method() === 'PATCH'),
    page.locator('.document-context-toggle input:not(:checked)').evaluate((input) => input.click()),
  ]);
  await waitForCondition('Document context enabled', () => document.querySelectorAll('.document-context-toggle input:checked').length === 2);
  await page.locator('.conflict-card').waitFor({ state: 'visible' });
  for (let index = 0; index < 10 && await page.locator('.conflict-card').count(); index += 1) {
    const beforeCount = await page.locator('.conflict-card').count();
    await Promise.all([
      page.waitForResponse((response) => response.url().endsWith('/resolve') && response.request().method() === 'POST'),
      page.locator('.conflict-card').first().getByRole('button', { name: '\u5ffd\u7565\u8be5\u5b57\u6bb5' }).click(),
    ]);
    await waitForCondition('Conflict count reduced', (count) => document.querySelectorAll('.conflict-card').length < count, beforeCount);
  }
  assert(await page.locator('.conflict-card').count() === 0, 'Conflict did not clear after selecting a source value.');
  await page.reload();
  await page.locator('.app-shell').waitFor({ state: 'visible' });
  await waitForCondition('Import mode restore with documents', () => document.querySelector('[title="\u5bfc\u5165\u6587\u6863"]')?.classList.contains('active'));
  assert(await page.locator('.document-card').count() === 2, 'Project documents were not restored after refresh.');
  assert(await page.locator('.document-context-toggle input:checked').count() === 2, 'Document context choices were not restored after refresh.');

  const pageEntryMode = page.getByTitle('\u6309\u9875\u5f55\u5165');
  await Promise.all([
    page.waitForResponse((response) => response.url().includes('/api/v2/projects/') && response.url().endsWith('/work-sessions') && response.request().method() === 'POST'),
    pageEntryMode.click(),
  ]);
  await page.reload();
  await page.locator('.app-shell').waitFor({ state: 'visible' });
  await waitForCondition('Page-entry mode restore', () => document.querySelector('[title="\u6309\u9875\u5f55\u5165"]')?.classList.contains('active'));
  assert(await page.locator('.draft-card').count() === 2, 'Page-entry mode or drafts were not restored after refresh.');
  const openChatCollapse = page.getByTitle('\u6536\u8d77\u804a\u5929\u9762\u677f');
  if (await openChatCollapse.count()) await openChatCollapse.click();
  const initialDraftTitles = await page.locator('.draft-card input').evaluateAll((inputs) => inputs.map((input) => input.value));
  const firstHandle = page.locator('.draft-drag-handle').first();
  const secondCard = page.locator('.draft-card').nth(1);
  await firstHandle.scrollIntoViewIfNeeded();
  const handleBox = await firstHandle.boundingBox();
  const secondBox = await secondCard.boundingBox();
  assert(handleBox && secondBox, 'Draft drag targets were not measurable.');
  const draftHitTest = await page.evaluate(({ start, end }) => ({
    start: document.elementFromPoint(start.x, start.y)?.closest('[data-draft-id]')?.getAttribute('data-draft-id'),
    end: document.elementFromPoint(end.x, end.y)?.closest('[data-draft-id]')?.getAttribute('data-draft-id'),
  }), {
    start: { x: handleBox.x + handleBox.width / 2, y: handleBox.y + handleBox.height / 2 },
    end: { x: secondBox.x + secondBox.width / 2, y: secondBox.y + secondBox.height / 2 },
  });
  assert(draftHitTest.start && draftHitTest.end && draftHitTest.start !== draftHitTest.end, `Draft drag points are occluded: ${JSON.stringify(draftHitTest)}`);
  await page.mouse.move(handleBox.x + handleBox.width / 2, handleBox.y + handleBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(secondBox.x + secondBox.width / 2, secondBox.y + secondBox.height / 2, { steps: 8 });
  await page.mouse.up();
  const reorderedDraftTitles = await page.locator('.draft-card input').evaluateAll((inputs) => inputs.map((input) => input.value));
  assert(reorderedDraftTitles[0] === initialDraftTitles[1] && reorderedDraftTitles[1] === initialDraftTitles[0], 'Pointer drag did not reorder draft pages.');

  const secondProjectName = `${projectName} secondary`;
  const secondary = await page.evaluate(async ({ name, slidesMarkdown }) => {
    const response = await fetch('/api/v1/projects', { method: 'POST', credentials: 'include', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name, slidesMarkdown }) });
    return (await response.json()).project;
  }, { name: secondProjectName, slidesMarkdown: '# Secondary\n\nRestore target.' });
  await page.reload();
  await page.locator('.app-shell').waitFor({ state: 'visible' });
  await page.getByTitle('\u6211\u7684\u9879\u76ee').click();
  await page.locator('.project-menu').getByRole('button', { name: secondProjectName, exact: true }).click();
  await waitForCondition('Secondary project selected', (name) => document.querySelector('.breadcrumbs strong')?.textContent === name, secondProjectName);
  await waitForCondition('Secondary project preference saved', (projectId) => {
    const key = Object.keys(localStorage).find((candidate) => candidate.startsWith('fastppt.workspace.v2.'));
    return key ? JSON.parse(localStorage.getItem(key)).lastProjectId === projectId : false;
  }, secondary.projectId);
  await page.reload();
  await waitForCondition('Secondary project restored', (name) => document.querySelector('.breadcrumbs strong')?.textContent === name, secondProjectName);
  await page.getByTitle('\u6211\u7684\u9879\u76ee').click();
  await page.locator('.project-menu').getByRole('button', { name: projectName, exact: true }).click();
  await waitForCondition('Primary project reselected', (name) => document.querySelector('.breadcrumbs strong')?.textContent === name, projectName);
  await waitForCondition('Primary project pages loaded', () => document.querySelectorAll('.page-row').length === 2);
  await page.waitForTimeout(500);
  await page.locator('.page-row').nth(1).click();
  await waitForCondition('Current page preference saved', ({ projectId, pageId }) => {
    const key = Object.keys(localStorage).find((candidate) => candidate.startsWith('fastppt.workspace.v2.'));
    if (!key) return false;
    const value = JSON.parse(localStorage.getItem(key));
    return value.projects?.[projectId]?.currentPageId === pageId;
  }, { projectId: created.body.project.projectId, pageId: created.body.project.pages[1].pageId });
  await page.getByTitle('\u6536\u8d77\u5bfc\u822a').click();
  const expandChat = page.getByTitle('\u5c55\u5f00\u804a\u5929\u9762\u677f');
  if (await expandChat.count()) await expandChat.click();
  await page.reload();
  await waitForCondition('Primary project restored', (name) => document.querySelector('.breadcrumbs strong')?.textContent === name, projectName);
  assert((await page.locator('.app-shell').getAttribute('class')).includes('nav-collapsed'), 'Navigation collapsed state was not restored.');
  assert(await page.locator('.chat-capsule-wrap').getAttribute('class').then((value) => value.includes('expanded')), 'Chat expanded state was not restored.');
  assert((await page.locator('.page-row.active .page-row-title').innerText()).includes('Second page'), 'Current page was not restored.');

  const beautifyProjectName = `E2E Beautify Deck ${unique}`;
  const beautifyProject = await page.evaluate(async ({ name }) => {
    const response = await fetch('/api/v1/projects', {
      method: 'POST',
      credentials: 'include',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, slidesMarkdown: '# PPTX intake\n\nAwaiting uploaded structure.' }),
    });
    return { ok: response.ok, body: await response.json() };
  }, { name: beautifyProjectName });
  assert(beautifyProject.ok, `Beautify project creation failed: ${JSON.stringify(beautifyProject.body)}`);
  await page.reload();
  await page.locator('.app-shell').waitFor({ state: 'visible' });
  const expandNavigation = page.getByTitle('\u5c55\u5f00\u5bfc\u822a');
  if (await expandNavigation.count()) await expandNavigation.click();
  await page.getByTitle('\u6211\u7684\u9879\u76ee').click();
  await page.locator('.project-menu').getByRole('button', { name: beautifyProjectName, exact: true }).click();
  await waitForCondition('Beautify project selected', (name) => document.querySelector('.breadcrumbs strong')?.textContent === name, beautifyProjectName);
  await Promise.all([
    page.waitForResponse((response) => response.url().includes('/api/v2/projects/') && response.url().endsWith('/work-sessions') && response.request().method() === 'POST'),
    page.getByTitle('\u0050\u0050\u0054 \u7f8e\u5316').click(),
  ]);
  const pptxInput = page.locator('input[accept=".pptx"]');
  const parseResponsePromise = page.waitForResponse((response) => response.url().endsWith('/parse') && response.request().method() === 'POST', { timeout: 120000 });
  await pptxInput.setInputFiles('tests/fixtures/pptx-structure.pptx');
  const parseResponse = await parseResponsePromise;
  assert(parseResponse.ok(), `PPTX parse failed: ${parseResponse.status()}`);
  const pptxDocument = page.locator('.document-card').filter({ hasText: 'pptx-structure.pptx' });
  await pptxDocument.getByText('\u5df2\u5c31\u7eea').waitFor({ state: 'visible', timeout: 120000 });
  assert((await pptxDocument.innerText()).includes('PPTX') && (await pptxDocument.innerText()).includes('SHA '), 'Parsed PPTX metadata was not shown.');
  await page.reload();
  await page.locator('.app-shell').waitFor({ state: 'visible' });
  await waitForCondition('Beautify mode restore', () => document.querySelector('[title="PPT \u7f8e\u5316"]')?.classList.contains('active'));
  assert(await page.locator('.document-card').filter({ hasText: 'pptx-structure.pptx' }).count() === 1, 'Parsed PPTX was not restored after refresh.');
  assert(await page.locator('.page-row').count() === 2, 'PPTX structure did not produce the expected two stable pages.');
  const beautifyCollapse = page.getByTitle('\u6536\u8d77\u804a\u5929\u9762\u677f');
  if (!(await beautifyCollapse.count())) await page.getByTitle('\u5c55\u5f00\u804a\u5929\u9762\u677f').click();
  await composer.fill('\u7edf\u4e00\u914d\u8272\u548c\u5e03\u5c40\uff0c\u4fdd\u7559\u539f\u59cb\u5185\u5bb9\u7ed3\u6784');
  await page.locator('.send-button').click();
  const visualConfirmation = page.locator('.confirmation-card');
  await visualConfirmation.waitFor({ state: 'visible', timeout: 120000 });
  const visualCandidate = visualConfirmation.locator('.visual-candidate');
  await visualCandidate.waitFor({ state: 'visible' });
  assert(await visualCandidate.locator('img').count() === 1, 'PPTX beautify did not expose a visual candidate image.');
  assert((await visualCandidate.innerText()).includes('\u5e03\u5c40\u5019\u9009') || (await visualCandidate.innerText()).includes('\u89c6\u89c9\u5019\u9009'), 'Visual candidate status was not shown.');
  const beautifyBefore = await page.evaluate(async (projectId) => {
    const response = await fetch(`/api/v1/projects/${projectId}`, { credentials: 'include' });
    const body = await response.json();
    return body.project.pages[0].currentVersionId;
  }, beautifyProject.body.project.projectId);
  await Promise.all([
    page.waitForResponse((response) => response.url().includes('/edit-operations/') && response.url().endsWith('/confirm') && response.request().method() === 'POST'),
    visualConfirmation.locator('.confirmation-actions .primary-button').click(),
  ]);
  await waitForCondition('PPTX beautify version created', async ({ projectId, previousVersionId }) => {
    const response = await fetch(`/api/v1/projects/${projectId}`, { credentials: 'include' });
    const body = await response.json();
    return body.project?.pages?.[0]?.currentVersionId !== previousVersionId;
  }, { projectId: beautifyProject.body.project.projectId, previousVersionId: beautifyBefore }, { timeout: 120000 });
  if (health.renderer === 'powerpoint_com') {
    const authorityImage = page.locator('.authoritative-render img[data-artifact-id]');
    await authorityImage.waitFor({ state: 'visible', timeout: 120000 });
    const authorityEvidence = await authorityImage.evaluate(async (node) => {
      const response = await fetch(node.src, { credentials: 'include' });
      const bytes = new Uint8Array(await response.arrayBuffer());
      return {
        status: response.status,
        contentType: response.headers.get('content-type'),
        width: node.naturalWidth,
        height: node.naturalHeight,
        png: bytes.length > 1000 && bytes.slice(0, 8).every((value, index) => value === [137, 80, 78, 71, 13, 10, 26, 10][index]),
      };
    });
    assert(authorityEvidence.status === 200 && authorityEvidence.contentType === 'image/png' && authorityEvidence.png && authorityEvidence.width >= 1000 && authorityEvidence.height >= 500, `Beautify authoritative PNG failed: ${JSON.stringify(authorityEvidence)}`);
  }
  await page.screenshot({ path: 'output/playwright/beautify-browser-flow-latest.png', fullPage: true });

  await page.getByRole('button', { name: '\u5bfc\u51fa \u0050\u0050\u0054\u0058' }).click();
  const exportBanner = page.locator('.export-banner');
  await exportBanner.getByText('\u0050\u0050\u0054\u0058 \u5bfc\u51fa\u5b8c\u6210').waitFor({ state: 'visible', timeout: 120000 });
  const officeButton = exportBanner.getByRole('button', { name: '\u0050\u006f\u0077\u0065\u0072\u0050\u006f\u0069\u006e\u0074 \u7f51\u9875\u9884\u89c8' });
  assert(await officeButton.count() === 1, 'Non-sensitive project did not expose the Office Viewer action.');
  await page.context().route('https://view.officeapps.live.com/**', (route) => route.fulfill({ status: 200, contentType: 'text/html', body: '<!doctype html><title>Office Viewer acceptance stub</title>' }));
  const [officeResponse, officePage] = await Promise.all([
    page.waitForResponse((response) => response.url().endsWith('/office-preview') && response.request().method() === 'POST'),
    page.context().waitForEvent('page'),
    officeButton.click(),
  ]);
  assert(officeResponse.ok(), `Office preview grant failed: ${officeResponse.status()}`);
  const officeGrant = await officeResponse.json();
  const officeExportId = officeResponse.url().split('?')[0].split('/').at(-2);
  assert(officeExportId, 'Office preview response did not identify its export.');
  await officePage.waitForURL((url) => url.href === officeGrant.officeViewerUrl, { timeout: 30000 });
  assert(officePage.url() === officeGrant.officeViewerUrl, 'Office Viewer did not open in a new tab.');
  assert(officePage.url().includes(`src=${encodeURIComponent(officeGrant.publicUrl)}`), 'Office Viewer did not receive the opaque public preview URL.');
  await officePage.close();
  await page.screenshot({ path: 'output/playwright/office-viewer-nonsensitive-latest.png', fullPage: true });

  await page.getByTitle('\u8bbe\u7f6e').click();
  const sensitiveToggle = page.locator('.settings-panel input[type="checkbox"]');
  await Promise.all([
    page.waitForResponse((response) => response.url().endsWith('/settings') && response.request().method() === 'PATCH'),
    sensitiveToggle.evaluate((input) => input.click()),
  ]);
  await waitForCondition('Sensitive mode enabled', () => document.querySelector('.settings-panel input[type="checkbox"]')?.checked === true);
  assert(await exportBanner.getByRole('button', { name: /PowerPoint/ }).count() === 0, 'Sensitive mode still exposed the Office Viewer action.');
  const sensitiveOfficeResponse = await page.evaluate(async ({ projectId, exportId }) => {
    const response = await fetch(`/api/v2/projects/${projectId}/exports/${exportId}/office-preview`, { method: 'POST', credentials: 'include' });
    return { status: response.status, body: await response.json() };
  }, { projectId: beautifyProject.body.project.projectId, exportId: officeExportId });
  assert(sensitiveOfficeResponse.status === 403, `Sensitive mode Office preview expected 403, got ${JSON.stringify(sensitiveOfficeResponse)}`);
  await page.screenshot({ path: 'output/playwright/office-viewer-sensitive-blocked-latest.png', fullPage: true });

  await page.setViewportSize({ width: 390, height: 844 });
  const collapseChat = page.getByTitle('\u6536\u8d77\u804a\u5929\u9762\u677f');
  if (await collapseChat.count()) await collapseChat.click();
  const mobileMetrics = await page.evaluate(() => {
    const rect = (selector) => document.querySelector(selector)?.getBoundingClientRect();
    return {
      capsule: rect('.chat-capsule-wrap'),
      slide: rect('.slide-render'),
      rangeButtons: [...document.querySelectorAll('.mode-tabs button')].map((node) => node.getBoundingClientRect()),
      send: rect('.send-button'),
      headerButtons: [...document.querySelectorAll('.chat-header-actions .icon-button')].map((node) => node.getBoundingClientRect()),
    };
  });
  assert(mobileMetrics.capsule.height <= 58, `Collapsed mobile chat is too tall: ${mobileMetrics.capsule.height}`);
  assert(mobileMetrics.slide.height > mobileMetrics.capsule.height * 2, 'Collapsed chat still obscures the key slide canvas.');
  await page.screenshot({ path: 'output/playwright/acceptance-mobile-collapsed-remediated.png', fullPage: true });
  if (await page.getByTitle('\u5c55\u5f00\u804a\u5929\u9762\u677f').count()) await page.getByTitle('\u5c55\u5f00\u804a\u5929\u9762\u677f').click();
  const touchMetrics = await page.evaluate(() => ({
    rangeButtons: [...document.querySelectorAll('.mode-tabs button')].map((node) => node.getBoundingClientRect()),
    send: document.querySelector('.send-button')?.getBoundingClientRect(),
    headerButtons: [...document.querySelectorAll('.chat-header-actions .icon-button')].map((node) => node.getBoundingClientRect()),
    capsule: document.querySelector('.chat-capsule-wrap')?.getBoundingClientRect(),
  }));
  assert(touchMetrics.rangeButtons.every((rect) => rect.height >= 44), 'A scope control is below the 44px touch target.');
  assert(touchMetrics.send.width >= 44 && touchMetrics.send.height >= 44, 'Send control is below the 44px touch target.');
  assert(touchMetrics.headerButtons.every((rect) => rect.width >= 44 && rect.height >= 44), 'A chat header control is below the 44px touch target.');
  assert(touchMetrics.capsule.height <= 844 * 0.47, 'Expanded mobile chat exceeds its viewport budget.');
  await page.screenshot({ path: 'output/playwright/acceptance-mobile-remediated.png', fullPage: true });

  assert(errors.length === 0, `Browser errors: ${errors.join('; ')}`);
  return { ok: true, projectName, secondaryProjectId: secondary.projectId };
}

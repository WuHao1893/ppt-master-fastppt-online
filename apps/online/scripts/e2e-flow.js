async (page) => {
  const assert = (condition, message) => {
    if (!condition) throw new Error(message);
  };
  const errors = [];
  page.on('pageerror', (error) => errors.push(error.message));
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

  const composer = page.locator('.composer textarea');
  await composer.fill('\u628a 42% \u6539\u4e3a 43%');
  await page.locator('.send-button').click();
  await page.locator('.confirmation-card').waitFor({ state: 'visible' });
  assert((await page.locator('.slide-render').innerText()).includes('42%'), 'Fact changed before confirmation.');
  await page.locator('.confirmation-actions .primary-button').click();
  await page.waitForFunction(() => document.querySelector('.slide-render')?.textContent?.includes('43%'));
  const rollback = page.getByRole('button', { name: '\u64a4\u9500\u672c\u6b21\u64cd\u4f5c' }).last();
  await rollback.waitFor({ state: 'visible' });
  await rollback.click();
  await page.waitForFunction(() => document.querySelector('.slide-render')?.textContent?.includes('42%'));

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

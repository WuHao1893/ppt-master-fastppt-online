import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import type { Project } from '../shared/models.js';
import { runPowerPointRender, runPptxExport } from '../server/workerBridge.js';

async function main(): Promise<void> {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'fastppt-powerpoint-'));
  try {
    const project = {
      projectId: 'p_powerpoint_smoke', ownerId: 'owner_powerpoint', name: 'PowerPoint Worker Smoke', themeId: 'test', themeVersion: '1', currentDeckRevisionId: 'deckrev_powerpoint', goalId: 'goal_powerpoint', status: 'ready', createdAt: new Date().toISOString(), updatedAt: new Date().toISOString(), sourceMarkdown: '# Authoritative render',
      pages: [{ pageId: 'page_powerpoint', projectId: 'p_powerpoint_smoke', currentVersionId: 'ver_powerpoint', orderIndex: 0, pageType: 'content', locked: false, factAnchorIds: ['fact_metric'], factAnchors: [{ factId: 'fact_metric', value: '42%', source: 'smoke', locked: true }], editableLevel: 'native_structure', status: 'authoritative', title: 'Authoritative render', body: '42% is rendered through Microsoft PowerPoint COM.', layout: 'editorial', sourceMarkdown: '# Authoritative render', contract: { conclusion: 'Authoritative render', evidence: ['42%'], mustKeep: ['42%'], canTrim: [], layout: 'editorial', density: 'airy', components: ['title', 'body'] }, versions: [] }],
    } as Project;
    const exported = await runPptxExport(project, directory, 'powerpoint-smoke.pptx');
    const rendered = await runPowerPointRender(exported.outputPath, path.join(directory, 'rendered'));
    assert.equal(rendered.renderer, 'microsoft-powerpoint-com');
    assert.equal(rendered.renders.length, 1);
    const info = await fs.stat(rendered.renders[0].path);
    assert.ok(info.size > 1_000);
    const qa = JSON.parse(await fs.readFile(rendered.qaPath, 'utf8')) as { status?: string };
    assert.equal(qa.status, 'passed');
    console.log(JSON.stringify({ ok: true, renderer: rendered.renderer, slides: rendered.renders.length }, null, 2));
  } finally {
    await fs.rm(directory, { recursive: true, force: true });
  }
}

main().catch((error) => { console.error(error); process.exitCode = 1; });

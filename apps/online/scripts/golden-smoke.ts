import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { parseSlidesMarkdown } from '../server/contracts.js';
import { resolvePythonBin } from '../server/workerBridge.js';

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const worker = path.resolve(scriptDir, '..', 'worker', 'export_pptx.py');
const fixture = path.resolve(scriptDir, '..', 'tests', 'fixtures', 'golden-deck.md');

async function runWorker(payload: unknown, outputPath: string): Promise<{ output: string; qa: string }> {
  const python = resolvePythonBin();
  return new Promise((resolve, reject) => {
    const child = spawn(python, [worker, '--output', outputPath], { cwd: path.dirname(worker), windowsHide: true, stdio: ['pipe', 'pipe', 'pipe'], env: { ...process.env, PYTHONUTF8: '1' } });
    let stdout = '';
    let stderr = '';
    const timeout = setTimeout(() => { child.kill(); reject(new Error('Golden Deck worker timed out.')); }, 180_000);
    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');
    child.stdout.on('data', (chunk: string) => { stdout += chunk; });
    child.stderr.on('data', (chunk: string) => { stderr += chunk; });
    child.on('error', reject);
    child.on('close', (code) => {
      clearTimeout(timeout);
      if (code !== 0) return reject(new Error(stderr.trim() || `Golden Deck worker exited with ${code}.`));
      try { resolve(JSON.parse(stdout.trim()) as { output: string; qa: string }); } catch (error) { reject(new Error(`Invalid Golden Deck worker response: ${(error as Error).message}`)); }
    });
    child.stdin.end(JSON.stringify(payload));
  });
}

async function main(): Promise<void> {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'fastppt-golden-smoke-'));
  try {
    const slides = parseSlidesMarkdown(await fs.readFile(fixture, 'utf8'));
    assert.equal(slides.length, 7);
    const imagePath = path.join(directory, 'field-study.png');
    const python = resolvePythonBin();
    await new Promise<void>((resolve, reject) => {
      const child = spawn(python, ['-c', 'from PIL import Image; import sys; Image.new("RGB", (960, 540), (37, 99, 235)).save(sys.argv[1], "PNG")', imagePath], { windowsHide: true, stdio: ['ignore', 'ignore', 'pipe'], env: { ...process.env, PYTHONUTF8: '1' } });
      let stderr = '';
      child.stderr.setEncoding('utf8');
      child.stderr.on('data', (chunk: string) => { stderr += chunk; });
      child.on('error', reject);
      child.on('close', (code) => code === 0 ? resolve() : reject(new Error(stderr || `Unable to create Golden Deck image (${code}).`)));
    });
    const outputPath = path.join(directory, 'golden-deck.pptx');
    const result = await runWorker({
      projectId: 'p_golden_deck', name: 'Golden Deck', themeId: 'fastppt-editorial', themeVersion: '1.0.0',
      pages: slides.map((slide, index) => ({
        pageId: `golden_page_${String(index + 1).padStart(2, '0')}`,
        pageType: slide.pageType,
        title: slide.title,
        body: slide.body,
        layout: /two-column/i.test(slide.title) ? 'two-column' : /timeline/i.test(slide.title) ? 'timeline' : /image region/i.test(slide.title) ? 'image-right' : /chart|data/i.test(slide.title) ? 'data-focus' : slide.layout,
        editableLevel: /complex flow/i.test(slide.title) ? 'native_partial' : 'native_structure',
        nonEditableRegions: /image region/i.test(slide.title) ? ['visual_anchor'] : /complex flow/i.test(slide.title) ? ['complex_flow_structure'] : [],
        visualAssetPath: /image region/i.test(slide.title) ? imagePath : undefined,
      })),
    }, outputPath);
    const qa = JSON.parse(await fs.readFile(result.qa, 'utf8')) as any;
    assert.equal(qa.export_engine, 'ppt_master_svg_to_drawingml');
    assert.equal(qa.page_count, 7);
    assert.equal(qa.full_slide_raster_count, 0);
    assert.equal(qa.static_structure_passed, true);
    assert.equal(qa.svg_quality?.summary?.errors, 0);
    assert.equal(qa.svg_quality?.summary?.warnings, 0);
    assert.ok(['passed', 'passed-with-warnings'].includes(qa.pptx_postflight?.status));
    assert.ok(['passed', 'passed-with-advisories'].includes(qa.delivery_check?.status));
    assert.equal(qa.pages.filter((page: any) => page.full_slide_raster).length, 0);
    assert.equal(qa.pages.find((page: any) => page.page_id === 'golden_page_06')?.local_image_region, 'visual-anchor');
    assert.equal(qa.pages.find((page: any) => page.page_id === 'golden_page_07')?.editable_level, 'native_partial');
    assert.ok(qa.pages.every((page: any) => page.text_shapes >= 1 && page.native_shapes >= 1));
    console.log(JSON.stringify({ ok: true, pages: qa.page_count, fullSlideRasters: qa.full_slide_raster_count, imageRegions: qa.pages.filter((page: any) => page.local_image_region).length }, null, 2));
  } finally {
    await fs.rm(directory, { recursive: true, force: true });
  }
}

main().catch((error) => { console.error(error); process.exitCode = 1; });

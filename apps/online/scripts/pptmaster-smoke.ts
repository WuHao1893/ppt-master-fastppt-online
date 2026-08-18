import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const worker = path.resolve(scriptDir, '..', 'worker', 'export_pptx.py');

async function makePng(outputPath: string): Promise<void> {
  const python = process.env.PYTHON_BIN || 'python';
  await new Promise<void>((resolve, reject) => {
    const child = spawn(python, [
      '-c',
      'from PIL import Image; import sys; Image.new("RGB", (960, 540), (37, 99, 235)).save(sys.argv[1], "PNG")',
      outputPath,
    ], { windowsHide: true, stdio: ['ignore', 'ignore', 'pipe'], env: { ...process.env, PYTHONUTF8: '1' } });
    let stderr = '';
    child.stderr.setEncoding('utf8');
    child.stderr.on('data', (chunk: string) => { stderr += chunk; });
    child.on('error', reject);
    child.on('close', (code) => {
      if (code === 0) resolve();
      else reject(new Error(stderr.trim() || `Unable to create test PNG; Python exited with ${code}.`));
    });
  });
}

async function runWorker(payload: unknown, outputPath: string): Promise<{ output: string; qa: string }> {
  const python = process.env.PYTHON_BIN || 'python';
  return new Promise((resolve, reject) => {
    const child = spawn(python, [worker, '--output', outputPath], {
      cwd: path.dirname(worker),
      windowsHide: true,
      stdio: ['pipe', 'pipe', 'pipe'],
      env: { ...process.env, PYTHONUTF8: '1' },
    });
    let stdout = '';
    let stderr = '';
    const timeout = setTimeout(() => {
      child.kill();
      reject(new Error('ppt-master smoke worker timed out.'));
    }, 150_000);
    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');
    child.stdout.on('data', (chunk: string) => { stdout += chunk; });
    child.stderr.on('data', (chunk: string) => { stderr += chunk; });
    child.on('error', reject);
    child.on('close', (code) => {
      clearTimeout(timeout);
      if (code !== 0) {
        reject(new Error(stderr.trim() || `ppt-master smoke worker exited with ${code}.`));
        return;
      }
      try {
        resolve(JSON.parse(stdout.trim()) as { output: string; qa: string });
      } catch (error) {
        reject(new Error(`ppt-master smoke worker returned invalid JSON: ${(error as Error).message}`));
      }
    });
    child.stdin.end(JSON.stringify(payload));
  });
}

async function main(): Promise<void> {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'fastppt-pptmaster-smoke-'));
  try {
    const visualPath = path.join(directory, 'visual.png');
    const outputPath = path.join(directory, 'native-editable.pptx');
    await makePng(visualPath);
    const result = await runWorker({
      projectId: 'p_pptmaster_smoke',
      name: 'Native Editable Smoke',
      pages: [
        {
          pageId: 'page_native_text',
          pageType: 'content',
          title: 'Editable text and shapes',
          body: 'The 42% fact remains native PowerPoint text.',
          layout: 'editorial',
          editableLevel: 'native_structure',
          nonEditableRegions: [],
        },
        {
          pageId: 'page_local_image',
          pageType: 'content',
          title: 'Local image region',
          body: 'The image is local while this sentence remains editable.',
          layout: 'image-right',
          editableLevel: 'native_partial',
          nonEditableRegions: ['visual_anchor'],
          visualAssetPath: visualPath,
        },
      ],
    }, outputPath);
    const qa = JSON.parse(await fs.readFile(result.qa, 'utf8')) as any;
    assert.equal(path.resolve(result.output), path.resolve(outputPath));
    assert.equal(qa.export_engine, 'ppt_master_svg_to_drawingml');
    assert.equal(qa.page_count, 2);
    assert.equal(qa.full_slide_raster_count, 0);
    assert.equal(qa.static_structure_passed, true);
    assert.equal(qa.svg_quality?.summary?.errors, 0);
    assert.equal(qa.svg_quality?.summary?.warnings, 0);
    assert.ok(['passed', 'passed-with-warnings'].includes(qa.pptx_postflight?.status));
    assert.ok(['passed', 'passed-with-advisories'].includes(qa.delivery_check?.status));
    assert.ok(qa.pages.every((page: any) => page.text_shapes >= 1 && page.native_shapes >= 1));
    const visualPage = qa.pages.find((page: any) => page.page_id === 'page_local_image');
    assert.equal(visualPage?.local_image_region, 'visual-anchor');
    assert.equal(visualPage?.picture_shapes, 1);
    assert.equal(visualPage?.full_slide_raster, false);
    assert.ok((await fs.stat(outputPath)).size > 10_000);
    console.log(JSON.stringify({
      ok: true,
      engine: qa.export_engine,
      pages: qa.page_count,
      quality: qa.pptx_postflight.status,
      pictureShapes: visualPage.picture_shapes,
    }, null, 2));
  } finally {
    await fs.rm(directory, { recursive: true, force: true });
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

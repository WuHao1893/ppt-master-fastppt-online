import { spawn } from 'node:child_process';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import type { Project } from '../shared/models.js';

const serverDir = path.dirname(fileURLToPath(import.meta.url));
const workerCandidates = [
  path.resolve(serverDir, '..', 'worker', 'export_pptx.py'),
  path.resolve(serverDir, '..', '..', 'worker', 'export_pptx.py'),
];

const renderWorkerCandidates = [
  path.resolve(serverDir, '..', 'worker', 'render_powerpoint.py'),
  path.resolve(serverDir, '..', '..', 'worker', 'render_powerpoint.py'),
];

export interface PptxExportResult {
  outputPath: string;
  qaPath: string | null;
}

export interface PptxVisualAsset {
  bytes: Buffer;
  extension: '.png' | '.jpg' | '.jpeg' | '.webp';
}

export interface PowerPointRenderResult {
  renderer: string;
  qaPath: string;
  renders: Array<{ slide_index: number; path: string; width: number; height: number }>;
}

export async function runPptxExport(
  project: Project,
  outputDir: string,
  fileName: string,
  visualAssets: Record<string, PptxVisualAsset> = {},
): Promise<PptxExportResult> {
  const resolvedOutputDir = path.resolve(outputDir);
  await fs.mkdir(resolvedOutputDir, { recursive: true });
  const outputPath = path.resolve(resolvedOutputDir, fileName);
  if (!outputPath.startsWith(`${resolvedOutputDir}${path.sep}`)) throw new Error('Export path escaped the project artifact directory.');
  const python = process.env.PYTHON_BIN || 'python';
  let workerScript = '';
  for (const candidate of workerCandidates) {
    try {
      await fs.access(candidate);
      workerScript = candidate;
      break;
    } catch {
      // Try the source-tree candidate when running from compiled output.
    }
  }
  if (!workerScript) throw new Error('PPTX export worker script is missing from the application package.');
  const pages = project.pages.filter((page) => !page.archived).sort((a, b) => a.orderIndex - b.orderIndex);
  let stagingDir: string | null = null;
  const visualAssetPaths: Record<string, string> = {};
  try {
    const suppliedAssets = Object.entries(visualAssets);
    if (suppliedAssets.length) {
      stagingDir = await fs.mkdtemp(path.join(resolvedOutputDir, '.fastppt-export-assets-'));
      for (const [index, [pageId, asset]] of suppliedAssets.entries()) {
        if (!pages.some((page) => page.pageId === pageId)) throw new Error(`Visual asset references an exported page that does not exist: ${pageId}.`);
        const fileName = `page-${String(index + 1).padStart(3, '0')}${asset.extension}`;
        const filePath = path.resolve(stagingDir, fileName);
        if (!filePath.startsWith(`${stagingDir}${path.sep}`)) throw new Error('Visual asset staging path escaped its directory.');
        await fs.writeFile(filePath, asset.bytes, { flag: 'wx' });
        visualAssetPaths[pageId] = filePath;
      }
    }
    const payload = {
      projectId: project.projectId,
      name: project.name,
      themeId: project.themeId,
      themeVersion: project.themeVersion,
      pages: pages.map((page) => {
        const version = page.versions.find((candidate) => candidate.versionId === page.currentVersionId);
        return {
          pageId: page.pageId,
          pageType: page.pageType,
          title: page.title,
          body: page.body,
          layout: page.layout,
          editableLevel: page.editableLevel,
          factAnchors: page.factAnchors,
          nonEditableRegions: version?.nonEditableRegions || [],
          visualAssetPath: visualAssetPaths[page.pageId] || null,
        };
      }),
    };
    return await new Promise<PptxExportResult>((resolve, reject) => {
      const child = spawn(python, [workerScript, '--output', outputPath], {
        cwd: path.dirname(workerScript),
        windowsHide: true,
        stdio: ['pipe', 'pipe', 'pipe'],
        env: { ...process.env, PYTHONUTF8: '1' },
      });
      let stdout = '';
      let stderr = '';
      const timeoutMs = Number(process.env.PPTX_EXPORT_TIMEOUT_MS || 120_000);
      const timeout = setTimeout(() => {
        child.kill();
        reject(new Error(`PPTX export worker timed out after ${timeoutMs} ms.`));
      }, timeoutMs);
      child.stdout.setEncoding('utf8');
      child.stderr.setEncoding('utf8');
      child.stdout.on('data', (chunk: string) => { stdout += chunk; });
      child.stderr.on('data', (chunk: string) => { stderr += chunk; });
      child.on('error', (error) => {
        clearTimeout(timeout);
        reject(new Error(`Unable to start PPTX export worker: ${error.message}. Set PYTHON_BIN to the ppt-master worker environment.`));
      });
      child.on('close', (code) => {
        clearTimeout(timeout);
        if (code !== 0) {
          reject(new Error(stderr.trim() || `PPTX export worker exited with code ${code}.`));
          return;
        }
        try {
          const result = JSON.parse(stdout.trim()) as { output: string; qa?: string };
          resolve({ outputPath: path.resolve(result.output), qaPath: result.qa ? path.resolve(result.qa) : null });
        } catch {
          reject(new Error('PPTX export worker returned invalid JSON.'));
        }
      });
      child.stdin.end(JSON.stringify(payload));
    });
  } finally {
    if (stagingDir) await fs.rm(stagingDir, { recursive: true, force: true });
  }
}

export async function runPowerPointRender(inputPath: string, outputDir: string): Promise<PowerPointRenderResult> {
  const resolvedInput = path.resolve(inputPath);
  const resolvedOutputDir = path.resolve(outputDir);
  await fs.mkdir(resolvedOutputDir, { recursive: true });
  let workerScript = '';
  for (const candidate of renderWorkerCandidates) {
    try {
      await fs.access(candidate);
      workerScript = candidate;
      break;
    } catch {
      // Try the source-tree candidate when running from compiled output.
    }
  }
  if (!workerScript) throw new Error('PowerPoint COM render worker is missing from the application package.');
  const python = process.env.PYTHON_BIN || 'python';
  return new Promise<PowerPointRenderResult>((resolve, reject) => {
    const child = spawn(python, [workerScript, '--input', resolvedInput, '--output-dir', resolvedOutputDir], {
      cwd: path.dirname(workerScript),
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'pipe'],
      env: { ...process.env, PYTHONUTF8: '1' },
    });
    let stdout = '';
    let stderr = '';
    const timeout = setTimeout(() => {
      child.kill();
      reject(new Error('PowerPoint COM render worker timed out after 60 seconds.'));
    }, 60_000);
    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');
    child.stdout.on('data', (chunk: string) => { stdout += chunk; });
    child.stderr.on('data', (chunk: string) => { stderr += chunk; });
    child.on('error', (error) => {
      clearTimeout(timeout);
      reject(new Error(`Unable to start PowerPoint COM renderer: ${error.message}`));
    });
    child.on('close', (code) => {
      clearTimeout(timeout);
      if (code !== 0) {
        reject(new Error(stderr.trim() || `PowerPoint COM renderer exited with code ${code}.`));
        return;
      }
      try {
        const result = JSON.parse(stdout.trim()) as { renderer?: string; qa?: string; renders?: PowerPointRenderResult['renders'] };
        if (!result.renderer || !result.qa || !Array.isArray(result.renders)) throw new Error('PowerPoint COM renderer response is incomplete.');
        resolve({ renderer: result.renderer, qaPath: path.resolve(result.qa), renders: result.renders });
      } catch {
        reject(new Error('PowerPoint COM renderer returned invalid JSON.'));
      }
    });
  });
}

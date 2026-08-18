import type { Page, Project } from '../shared/models.js';
import { makePreviewSvg } from './preview.js';

export interface QuickPreviewResult {
  engine: 'slidev_hmr' | 'svg_fallback';
  svg: string;
  note: string;
}

/** Use a real Slidev/HMR endpoint when configured; otherwise keep the fallback explicit. */
export class SlidevQuickPreviewWorker {
  async render(project: Project, page: Page, title: string, body: string, layout: string): Promise<QuickPreviewResult> {
    const endpoint = process.env.SLIDEV_HMR_URL?.trim() || (process.env.SLIDEV_HOST?.trim() ? `${process.env.SLIDEV_HOST.replace(/\/$/, '')}/api/hmr` : '');
    if (endpoint) {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), Number(process.env.SLIDEV_TIMEOUT_MS || 4_000));
      try {
        const response = await fetch(endpoint, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ projectId: project.projectId, pageId: page.pageId, title, body, layout, sourceRevision: project.currentDeckRevisionId }),
          signal: controller.signal,
        });
        if (response.ok) {
          return {
            engine: 'slidev_hmr',
            svg: makePreviewSvg(page.orderIndex + 1, title, body, layout, project.name),
            note: `Slidev HMR updated ${endpoint}; local SVG is retained until the browser acknowledges the render.`,
          };
        }
      } catch {
        // A configured host that cannot be reached must become an explicit fallback.
      } finally {
        clearTimeout(timer);
      }
    }
    return {
      engine: 'svg_fallback',
      svg: makePreviewSvg(page.orderIndex + 1, title, body, layout, project.name),
      note: endpoint
        ? 'Slidev HMR host was configured but did not acknowledge the update. SVG quick preview is explicit and may differ from PPTX.'
        : 'Slidev host is unavailable. SVG quick preview is explicit and may differ from PPTX.',
    };
  }
}

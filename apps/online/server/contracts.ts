import type { FactAnchor, PageContract, PageType } from '../shared/models.js';
import { sha256 } from './utils.js';

export interface ParsedSlide {
  title: string;
  body: string;
  sourceMarkdown: string;
  pageType: PageType;
  layout: string;
  facts: FactAnchor[];
  contract: PageContract;
}

const splitPattern = /^\s*---\s*$/m;

function detectPageType(index: number, title: string, total: number): PageType {
  const normalized = title.toLowerCase();
  if (index === 0) return 'cover';
  if (index === total - 1 || /结语|谢谢|ending|thank/.test(normalized)) return 'ending';
  if (/目录|contents|agenda|议程/.test(normalized)) return 'toc';
  if (/章节|chapter|section|第一|第二|第三|第四/.test(normalized)) return 'section';
  return 'content';
}

function detectLayout(text: string): string {
  if (/两列|双栏|two.?column|2.?column/i.test(text)) return 'two-column';
  if (/时间线|timeline|阶段|路线图|flow|流程/i.test(text)) return 'timeline';
  if (/图表|chart|数据|指标|kpi/i.test(text)) return 'data-focus';
  return 'editorial';
}

function extractFacts(text: string, pageIndex: number): FactAnchor[] {
  const matches = text.match(/(?<![\w])(?:\d+(?:\.\d+)?%?|\d{4}年?|第[一二三四五六七八九十]+[章节])/g) ?? [];
  return [...new Set(matches)].slice(0, 18).map((value, index) => ({
    factId: `fact_${String(pageIndex + 1).padStart(3, '0')}_${String(index + 1).padStart(2, '0')}`,
    value,
    source: `slides.md#page-${pageIndex + 1}`,
    locked: true,
  }));
}

export function parseSlidesMarkdown(markdown: string): ParsedSlide[] {
  const normalized = markdown.trim() || '# FastPPT Online\n\n从浏览器开始逐页精修。\n---\n# 交付链路\n\n事实锚点、视觉预览、可编辑导出与质量门。';
  const chunks = normalized.split(splitPattern).map((chunk) => chunk.trim()).filter(Boolean);
  return chunks.map((sourceMarkdown, index) => {
    const heading = sourceMarkdown.match(/^#{1,3}\s+(.+)$/m);
    const title = heading?.[1]?.trim() || `未命名页面 ${index + 1}`;
    const body = sourceMarkdown.replace(/^#{1,3}\s+.+$/m, '').trim() || '暂无正文内容';
    const facts = extractFacts(`${title}\n${body}`, index);
    const layout = detectLayout(`${title}\n${body}`);
    const pageType = detectPageType(index, title, chunks.length);
    return {
      title,
      body,
      sourceMarkdown,
      pageType,
      layout,
      facts,
      contract: {
        conclusion: title,
        evidence: body.split(/\r?\n/).filter(Boolean).slice(0, 5),
        mustKeep: facts.map((fact) => fact.value),
        canTrim: body.split(/\r?\n/).filter(Boolean).slice(5),
        layout,
        density: body.length > 500 ? 'dense' : body.length > 220 ? 'balanced' : 'airy',
        components: layout === 'two-column' ? ['title', 'body', 'right_panel'] : ['title', 'body', 'visual_anchor'],
      },
    };
  });
}

export function contractHash(contract: PageContract): string {
  return sha256(JSON.stringify(contract));
}

export function sourceHash(source: string): string {
  return sha256(source);
}

import type { EditMode, EditPlan, FactAnchor, Page, Project } from '../shared/models.js';
import { sha256 } from './utils.js';

const complexPatterns = /复杂图表|结构化重绘|流程图连接线|自动重排连接线|复杂流程图|sankey|复杂图表/i;
const factTokenPattern = /(?<![\w])(?:\d+(?:\.\d+)?%?|\d{4}年?|第[一二三四五六七八九十]+[章节])/g;
const factReplacementPattern = /(?<![\w])((?:\d+(?:\.\d+)?%?|\d{4}年?|第[一二三四五六七八九十]+[章节]))\s*(?:改为|改成|替换为|调整为|变成|to)\s*((?:\d+(?:\.\d+)?%?|\d{4}年?|第[一二三四五六七八九十]+[章节]))/giu;

export interface FactChange {
  oldValue: string;
  newValue: string;
}

export function factChanges(message: string): FactChange[] {
  const changes: FactChange[] = [];
  for (const match of message.matchAll(factReplacementPattern)) {
    const oldValue = match[1];
    const newValue = match[2];
    if (oldValue && newValue && oldValue !== newValue) changes.push({ oldValue, newValue });
  }
  return changes;
}

export function applyFactChanges(anchors: FactAnchor[], message: string, confirmed: boolean): FactAnchor[] {
  const changes = factChanges(message);
  if (!confirmed || !changes.length) return anchors.map((anchor) => ({ ...anchor }));
  return anchors.map((anchor) => {
    const change = changes.find((candidate) => candidate.oldValue === anchor.value);
    return change ? { ...anchor, value: change.newValue } : { ...anchor };
  });
}

export function resolveSimilarPages(project: Project, message: string): { pageIds: string[]; reasons: Record<string, string> } {
  const terms = message.toLowerCase().split(/\s+/).filter((term) => term.length > 1);
  const scored = project.pages.map((page) => {
    const haystack = `${page.title} ${page.body} ${page.layout} ${page.contract.components.join(' ')}`.toLowerCase();
    const hits = terms.filter((term) => haystack.includes(term)).length;
    const structural = /标题|title/.test(message) && page.title.length > 18 ? 2 : 0;
    const score = hits + structural;
    const reason = structural ? '标题长度符合请求条件' : hits ? `合同/标题中命中 ${hits} 个相关词` : `布局标签为 ${page.layout}，与请求目标相近`;
    return { page, score, reason };
  }).sort((a, b) => b.score - a.score || a.page.orderIndex - b.page.orderIndex);
  const selected = scored.filter((item) => item.score > 0).slice(0, 12);
  const fallback = selected.length ? selected : scored.filter((item) => item.page.pageType === 'content').slice(0, 6);
  return {
    pageIds: fallback.map((item) => item.page.pageId),
    reasons: Object.fromEntries(fallback.map((item) => [item.page.pageId, item.reason])),
  };
}

function addChange(changes: EditPlan['changes'], change: EditPlan['changes'][number]): void {
  changes.push(change);
}

export function previewTextChange(page: Page, message: string): { title: string; body: string } {
  let title = page.title;
  let body = page.body;
  const explicitTitle = message.match(/(?:标题|title)(?:改成|改为|设为|[:：])\s*[“"']?([^。；;\n”"']{2,60})/i)?.[1]?.trim();
  if (explicitTitle) title = explicitTitle;
  else if (/标题|title|一句话|改短|缩短|headline/i.test(message) && title.length > 18) title = `${title.slice(0, 17)}…`;
  const explicitBody = message.match(/(?:正文|内容|body)(?:改成|改为|设为|[:：])\s*[“"']?([^”"']{4,500})/i)?.[1]?.trim();
  if (explicitBody) body = explicitBody;
  for (const change of factChanges(message)) {
    title = title.split(change.oldValue).join(change.newValue);
    body = body.split(change.oldValue).join(change.newValue);
  }
  return { title, body };
}

function calculateFactImpact(pages: Page[], message: string): EditPlan['factImpact'] {
  const added = new Set<string>();
  const removed = new Set<string>();
  const changed = new Set<string>();
  const requestedChanges = factChanges(message);
  for (const page of pages) {
    const edited = previewTextChange(page, message);
    const proposed = `${edited.title}\n${edited.body}`;
    for (const fact of page.factAnchors) {
      if (!proposed.includes(fact.value)) removed.add(`${fact.factId}:${fact.value}`);
    }
    const proposedFacts = proposed.match(/(?<![\w])(?:\d+(?:\.\d+)?%?|\d{4}年?|第[一二三四五六七八九十]+[章节])/g) ?? [];
    for (const value of new Set(proposedFacts)) {
      if (!page.factAnchors.some((fact) => fact.value === value)) added.add(value);
    }
    for (const change of requestedChanges) {
      const anchor = page.factAnchors.find((fact) => fact.value === change.oldValue);
      if (!anchor || !proposed.includes(change.newValue)) continue;
      changed.add(`${anchor.factId}:${change.oldValue}->${change.newValue}`);
      removed.delete(`${anchor.factId}:${change.oldValue}`);
      added.delete(change.newValue);
    }
  }
  return { added: [...added], removed: [...removed], changed: [...changed] };
}

export function buildEditPlan(project: Project, pages: Page[], mode: EditMode, message: string): EditPlan {
  const changes: EditPlan['changes'] = [];
  const factIds = pages.flatMap((page) => page.factAnchors.map((fact) => fact.factId));
  factIds.forEach((factId) => addChange(changes, { kind: 'preserve_fact', factId, target: 'fact_anchor' }));
  const normalized = message.toLowerCase();
  const unsupported = complexPatterns.test(message) ? ['复杂图表/流程图结构化重绘不属于 MVP，原结构将被保留。'] : [];
  if (/标题|title|一句话|改短|缩短|headline/.test(normalized)) {
    addChange(changes, { kind: 'rewrite_text', target: 'title', constraint: 'one_line', value: '根据用户原意压缩为一句话' });
  }
  if (/两列|双栏|三列|三阶段|布局|右侧|左侧|column|layout|timeline|路线图/.test(normalized)) {
    const value = /三阶段|timeline|路线图/.test(normalized) ? 'timeline' : /三列/.test(normalized) ? 'three-column' : 'two-column';
    addChange(changes, { kind: 'layout_change', target: 'page_layout', value });
  }
  if (/颜色|配色|风格|样式|字体|style|color|theme/.test(normalized)) {
    addChange(changes, { kind: 'style_change', target: 'theme', value: '保留主题约束，仅调整局部视觉层级' });
  }
  if (/图片|配图|背景|视觉|image|photo|background/.test(normalized)) {
    addChange(changes, { kind: 'image_replace', target: 'visual_anchor', value: '先生成视觉预览，再进入矢量/可编辑重建' });
  }
  if (changes.every((change) => change.kind === 'preserve_fact')) {
    addChange(changes, { kind: 'rewrite_text', target: 'body', constraint: '保持事实锚点，优化表达密度', value: message });
  }
  const imageUnits = changes.some((change) => change.kind === 'image_replace') ? pages.length : 0;
  const factImpact = calculateFactImpact(pages, message);
  const requiresConfirmation = mode !== 'single' || unsupported.length > 0 || factImpact.removed.length > 0 || factImpact.changed.length > 0 || /删除事实|删除数字|替换主题|改变页数/.test(message);
  if (factImpact.removed.length > 0) {
    unsupported.push(`检测到将移除锁定事实：${factImpact.removed.join('、')}。确认前不会执行。`);
  }
  return {
    intent: imageUnits ? 'layout_and_visual_revision' : 'structured_page_revision',
    affectedPageIds: pages.map((page) => page.pageId),
    changes,
    factImpact,
    unsupported,
    requiresConfirmation,
    estimatedCost: { imageUnits, amount: imageUnits * 0.08, currency: 'USD' },
    summary: unsupported.length ? '已识别请求，但复杂结构保留原样并等待确认。' : `${pages.length} 个页面将按合同约束执行：${message}`,
  };
}

export function validateEditPlan(project: Project, pages: Page[], mode: EditMode, plan: EditPlan): string[] {
  const errors: string[] = [];
  const expectedIds = pages.map((page) => page.pageId);
  const expected = new Set(expectedIds);
  const actual = new Set(plan.affectedPageIds);
  if (actual.size !== expected.size || [...expected].some((pageId) => !actual.has(pageId))) errors.push('计划影响页必须与已解析的目标 page_id 完全一致。');
  if (plan.affectedPageIds.length !== actual.size) errors.push('计划影响页不能包含重复 page_id。');
  if (mode !== 'single' && !plan.requiresConfirmation) errors.push('多页和全局修改必须经过确认门。');
  if ((plan.factImpact.removed.length || plan.factImpact.changed.length || plan.unsupported.length) && !plan.requiresConfirmation) errors.push('事实变化或不支持项必须经过确认门。');
  const knownFacts = new Set(pages.flatMap((page) => page.factAnchors.map((fact) => `${fact.factId}:${fact.value}`)));
  for (const item of plan.factImpact.removed) {
    if (item.includes(':') && !knownFacts.has(item)) errors.push(`计划删除了不存在的事实锚点：${item}。`);
  }
  for (const item of plan.factImpact.changed) {
    const match = item.match(/^([^:]+):(.+)->(.+)$/);
    if (!match || !knownFacts.has(`${match[1]}:${match[2]}`) || !match[3]) errors.push(`事实替换格式或来源无效：${item}。`);
  }
  const imageChange = plan.changes.some((change) => change.kind === 'image_replace');
  const expectedUnits = imageChange ? pages.length : 0;
  if (plan.estimatedCost.imageUnits !== expectedUnits) errors.push(`图像成本预估不一致，应为 ${expectedUnits} image unit。`);
  if (plan.estimatedCost.amount < 0 || plan.estimatedCost.amount > 1_000_000) errors.push('成本预估超出允许范围。');
  return errors;
}

export function composePrompt(project: Project, page: Page, plan: EditPlan, message: string): { prompt: string; hash: string } {
  const prompt = [
    'You are the FastPPT Online page revision worker.',
    `Project purpose: ${project.name}; audience: small-team presentation reviewers; language: user chat language.`,
    `Theme: ${project.themeId}@${project.themeVersion}; page type: ${page.pageType}; layout: ${page.layout}; density: ${page.contract.density}.`,
    `Conclusion: ${page.contract.conclusion}. Evidence: ${page.contract.evidence.join(' | ')}.`,
    `Locked facts: ${page.factAnchors.map((fact) => `${fact.factId}=${fact.value}`).join(', ') || 'none'}.`,
    `Must keep: ${page.contract.mustKeep.join(', ') || 'all meaningful content'}. Can trim: ${page.contract.canTrim.join(' | ') || 'none'}.`,
    `Adjacent consistency: preserve typography, color tokens, and component vocabulary across the deck.`,
    `Current revision: ${project.currentDeckRevisionId}; requested change: ${message}.`,
    `Structured plan: ${JSON.stringify(plan)}. Never invent facts, dates, names, amounts, or references.`,
    'If visual regeneration is required, produce a preview artifact before any editable reconstruction.',
  ].join('\n');
  return { prompt, hash: sha256(prompt) };
}

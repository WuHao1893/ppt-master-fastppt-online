import type { EditPlan, Page, Project } from '../shared/models.js';

export interface ContentQaResult {
  passed: boolean;
  errors: string[];
  warnings: string[];
  numbersBefore: string[];
  numbersAfter: string[];
}

const numberPattern = /(?<![\w])(?:\d+(?:\.\d+)?%?|\d{4}年?)/g;

function numbers(value: string): string[] {
  return [...new Set(value.match(numberPattern) || [])];
}

export function runContentQa(project: Project, page: Page, next: { title: string; body: string }, plan: EditPlan, confirmed: boolean): ContentQaResult {
  const before = `${page.title}\n${page.body}`;
  const after = `${next.title}\n${next.body}`;
  const errors: string[] = [];
  const warnings: string[] = [];
  const declaredRemoved = new Set(plan.factImpact.removed.map((value) => value.includes(':') ? value.slice(value.indexOf(':') + 1) : value));
  const declaredChanged = plan.factImpact.changed.flatMap((value) => {
    const match = value.match(/^([^:]+):(.+)->(.+)$/);
    return match ? [{ factId: match[1], oldValue: match[2], newValue: match[3] }] : [];
  });
  const changedOldValues = new Set(declaredChanged.map((change) => change.oldValue));
  const changedNewValues = new Set(declaredChanged.map((change) => change.newValue));
  changedOldValues.forEach((value) => declaredRemoved.add(value));
  for (const change of declaredChanged) {
    if (!confirmed) errors.push(`事实 ${change.factId} 从 ${change.oldValue} 改为 ${change.newValue} 只能在确认后执行。`);
    else if (after.includes(change.oldValue) || !after.includes(change.newValue)) errors.push(`事实 ${change.factId} 的替换结果未完整落到正文。`);
  }
  for (const fact of page.factAnchors.filter((anchor) => anchor.locked)) {
    if (!after.includes(fact.value)) {
      if (changedOldValues.has(fact.value)) continue;
      if (!declaredRemoved.has(fact.value)) errors.push(`锁定事实 ${fact.factId}:${fact.value} 未在计划中声明却从正文消失。`);
      else if (!confirmed) errors.push(`锁定事实 ${fact.factId}:${fact.value} 只能在确认后删除。`);
    }
  }
  const numbersBefore = numbers(before);
  const numbersAfter = numbers(after);
  const removedNumbers = numbersBefore.filter((value) => !numbersAfter.includes(value));
  const addedNumbers = numbersAfter.filter((value) => !numbersBefore.includes(value));
  if (removedNumbers.some((value) => !declaredRemoved.has(value))) errors.push(`数字一致性检查失败，未声明删除：${removedNumbers.filter((value) => !declaredRemoved.has(value)).join('、')}。`);
  const undeclaredAddedNumbers = addedNumbers.filter((value) => !changedNewValues.has(value));
  if (undeclaredAddedNumbers.length > 0) warnings.push(`正文新增数字：${undeclaredAddedNumbers.join('、')}，来源为本轮用户指令，需人工复核。`);
  if (/\uFFFD|\?{3,}/u.test(after)) errors.push('正文包含疑似编码损坏字符。');
  if (/ {3,}|\t{2,}/.test(after)) warnings.push('正文包含连续空白，建议检查排版。');
  const canonicalTerms = ['PPTX', 'FastPPT', 'Slidev', 'PowerPoint'];
  for (const term of canonicalTerms) {
    const variants = project.pages.flatMap((candidate) => `${candidate.title}\n${candidate.body}`.match(new RegExp(term, 'ig')) || []);
    const nextVariants = after.match(new RegExp(term, 'ig')) || [];
    if ([...variants, ...nextVariants].some((value) => value !== term)) warnings.push(`术语 ${term} 存在大小写不一致。`);
  }
  return { passed: errors.length === 0, errors, warnings: [...new Set(warnings)], numbersBefore, numbersAfter };
}

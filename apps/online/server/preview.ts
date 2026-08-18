import { escapeXml } from './utils.js';

const palette = [
  { accent: '#f97316', soft: '#fff1e8', ink: '#152238' },
  { accent: '#0f766e', soft: '#e6f5f2', ink: '#102a2b' },
  { accent: '#2563eb', soft: '#eaf1ff', ink: '#14233f' },
  { accent: '#c026d3', soft: '#fbeafa', ink: '#32183a' },
];

export function makePreviewSvg(
  pageNumber: number,
  title: string,
  body: string,
  layout: string,
  projectName: string,
): string {
  const theme = palette[(pageNumber - 1) % palette.length];
  const lines = body.split(/\r?\n/).map((line) => line.trim()).filter(Boolean).slice(0, 8);
  const lineMarkup = lines.map((line, index) => {
    const y = 205 + index * 28;
    return `<text x="94" y="${y}" font-family="Arial,Microsoft YaHei,sans-serif" font-size="15" fill="${theme.ink}">${escapeXml(line.slice(0, 62))}</text>`;
  }).join('');
  const columnMarkup = layout.includes('column')
    ? `<rect x="560" y="158" width="300" height="250" rx="16" fill="${theme.soft}"/><rect x="590" y="190" width="120" height="10" rx="5" fill="${theme.accent}"/><rect x="590" y="220" width="220" height="8" rx="4" fill="#b9c4d5"/><rect x="590" y="245" width="180" height="8" rx="4" fill="#d4dce7"/><rect x="590" y="290" width="120" height="10" rx="5" fill="#8c9bb0"/><rect x="590" y="320" width="220" height="8" rx="4" fill="#d4dce7"/>`
    : `<path d="M585 330 C650 235 700 300 770 210" fill="none" stroke="${theme.accent}" stroke-width="7" stroke-linecap="round"/><circle cx="585" cy="330" r="10" fill="${theme.accent}"/><circle cx="770" cy="210" r="10" fill="${theme.accent}"/><rect x="590" y="366" width="190" height="8" rx="4" fill="#c5cfdd"/>`;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 960 540" role="img" aria-label="${escapeXml(title)}"><rect width="960" height="540" fill="#f8fafc"/><rect width="16" height="540" fill="${theme.accent}"/><rect x="62" y="46" width="160" height="8" rx="4" fill="${theme.accent}"/><text x="62" y="92" font-family="Arial,Microsoft YaHei,sans-serif" font-size="12" letter-spacing="1.5" fill="#64748b">${escapeXml(projectName.toUpperCase())}</text><text x="62" y="148" font-family="Arial,Microsoft YaHei,sans-serif" font-weight="700" font-size="31" fill="${theme.ink}">${escapeXml(title.slice(0, 42))}</text>${lineMarkup}${columnMarkup}<text x="62" y="500" font-family="Arial,sans-serif" font-size="12" fill="#64748b">${String(pageNumber).padStart(2, '0')}  /  FASTPPT ONLINE</text><text x="874" y="500" font-family="Arial,sans-serif" font-size="12" fill="#64748b">${escapeXml(layout)}</text></svg>`;
}

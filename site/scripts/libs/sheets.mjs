/**
 * スプレッドシート表示のセンチネル（sentinels.mjs から呼ばれる）。
 *
 *   ::::sheetdemo[キャプション]{book="data/sales.xlsx" sheet="売上" range="A1:F8"}
 *   F2:F31 = 1250, 360, 600
 *   J2 = 68580
 *   ::::
 *     … 実行前のシートを表示し、「▶ 実行」で中の行に書いたセルの値に切り替える。
 *       コードは実際には動かさない（何が起きるかのイメージを見せるだけ）。
 *       中の行は `セル = 値` または `範囲 = 値, 値, …`（範囲は左上から行の順に埋める）。
 *
 *   ::sheetcode{book="data/sales.xlsx" sheet="売上" range="A1:F8" code="SalesSummary" macro="CalcAmounts"}
 *     … VBA ／ GAS のコードと、そのコードが扱うシートを左右に並べる。
 *       code="SalesSummary" で vba/SalesSummary.bas と gas/SalesSummary.gs を対にする。
 *       macro="CalcAmounts" で VBA は Sub CalcAmounts、GAS は function calcAmounts だけを切り出す
 *       （省略するとファイル全体）。VBA ／ GAS の切り替えは `::::platform` と共有する。
 *
 * 共通の属性:
 *   book      … example/ から数えたブックのパス（必須）
 *   sheet     … 表示するシート名（省略時は先頭）
 *   range     … 表示する範囲（省略時はデータのある範囲全体）
 *   highlight … 強調する範囲。カンマ区切りで複数書ける
 *   src       … 別レクチャーの example/ を使うとき `<sec>/<lec>`
 *
 * 生成物: シートは生 HTML（空行を含めない）。::sheetcode はコードフェンスを含む
 * `::::sheetcode` コンテナになり、remark-sheetcode.mjs が左右に並べる。
 * クリック時の動き（実行・リセット・数式バー）は site/src/scripts/sheet-client.js、
 * 装飾は site/src/styles/sheet.css。
 */

import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import path from 'node:path';

const require = createRequire(import.meta.url);
const ExcelJS = require('exceljs');

export const SHEETDEMO_OPEN_RE = /^::::sheetdemo(?:\[([^\]]*)\])?(?:\{([^}]*)\})?\s*$/;
export const SHEETCODE_RE = /^::sheetcode(?:\[([^\]]*)\])?(?:\{([^}]*)\})?\s*$/;
const CLOSE_RE = /^::::\s*$/;

const books = new Map();

async function loadBook(abs) {
  if (!books.has(abs)) {
    books.set(abs, (async () => {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(abs);
      return wb;
    })());
  }
  return books.get(abs);
}

// ---- セル番地 ----

function colToNum(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function numToCol(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function parseRef(ref) {
  const m = String(ref).trim().match(/^([A-Za-z]+)(\d+)$/);
  return m ? { col: colToNum(m[1]), row: Number(m[2]) } : null;
}

/** `A1:F8` / `F2` を { top, left, bottom, right } に。読めなければ null。 */
function parseRange(s) {
  const [a, b] = String(s).split(':');
  const p = parseRef(a);
  const q = b ? parseRef(b) : p;
  if (!p || !q) return null;
  return {
    top: Math.min(p.row, q.row),
    left: Math.min(p.col, q.col),
    bottom: Math.max(p.row, q.row),
    right: Math.max(p.col, q.col),
  };
}

function inRange(r, row, col) {
  return row >= r.top && row <= r.bottom && col >= r.left && col <= r.right;
}

// ---- 値の表示 ----

const pad = (n) => String(n).padStart(2, '0');

/** ブックの表示形式に合わせて、画面に出す文字列を作る。使っている形式だけに対応する。 */
function formatValue(value, numFmt) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) {
    const y = value.getUTCFullYear();
    const m = pad(value.getUTCMonth() + 1);
    const d = pad(value.getUTCDate());
    return `${y}/${m}/${d}`;
  }
  if (typeof value === 'object') {
    // 数式のセル。openpyxl で作ったブックは結果の値を持たないので、あれば出す。
    if ('formula' in value || 'sharedFormula' in value) return formatValue(value.result, numFmt);
    if ('richText' in value) return value.richText.map((t) => t.text).join('');
    if ('text' in value) return String(value.text);
    return '';
  }
  if (typeof value === 'number' && numFmt && /#,##0/.test(numFmt)) {
    const s = Math.round(value).toLocaleString('en-US');
    return /"¥"|\\¥|¥/.test(numFmt) ? `¥${s}` : s;
  }
  return String(value);
}

/** 数式バーに出す文字列（数式ならその式）。 */
function rawValue(value) {
  if (value && typeof value === 'object' && 'formula' in value) return `=${value.formula}`;
  if (value instanceof Date) return formatValue(value);
  return formatValue(value);
}

function isNumberLike(value) {
  if (typeof value === 'number' || value instanceof Date) return true;
  if (value && typeof value === 'object' && 'result' in value) return isNumberLike(value.result);
  return false;
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function argbToCss(color) {
  const argb = color?.argb;
  if (!argb || argb.length !== 8) return null;
  return `#${argb.slice(2)}`;
}

// Excel の列幅（文字数）をおおよその px に。
const colWidthPx = (w) => Math.round((w ?? 8.43) * 7 + 5);

// ---- シートの HTML ----

/**
 * シート 1 枚を表示する HTML（1 行。空行を含めない）。
 * after = Map<"F2", 値>。指定したセルには実行後の表示を data-after に持たせる。
 */
function renderSheet(wb, ws, { range, highlight, after, fileName }) {
  const dims = ws.dimensions;
  const view = parseRange(range) ?? {
    top: 1,
    left: 1,
    bottom: Math.max(dims?.bottom ?? 1, 1),
    right: Math.max(dims?.right ?? 1, 1),
  };
  const marks = String(highlight ?? '')
    .split(',')
    .map((s) => parseRange(s))
    .filter(Boolean);

  // 結合セル: 左上は colspan/rowspan、残りは描かない。
  const merges = [];
  for (const m of ws.model.merges ?? []) {
    const r = parseRange(m);
    if (r) merges.push(r);
  }
  const mergeAt = (row, col) => merges.find((r) => r.top === row && r.left === col);
  const hidden = (row, col) => merges.some((r) => inRange(r, row, col) && !(r.top === row && r.left === col));

  const cols = [];
  cols.push('<col class="sheet__rowhead-col">');
  for (let c = view.left; c <= view.right; c++) {
    cols.push(`<col style="width:${colWidthPx(ws.getColumn(c).width)}px">`);
  }

  const head = ['<th class="sheet__corner"></th>'];
  for (let c = view.left; c <= view.right; c++) head.push(`<th class="sheet__colhead">${numToCol(c)}</th>`);

  const rows = [];
  for (let r = view.top; r <= view.bottom; r++) {
    const tds = [`<th class="sheet__rowhead">${r}</th>`];
    for (let c = view.left; c <= view.right; c++) {
      if (hidden(r, c)) continue;
      const cell = ws.getCell(r, c);
      const ref = `${numToCol(c)}${r}`;
      const cls = ['sheet__cell'];
      const style = [];
      const attrs = [`data-ref="${ref}"`];
      const fill = cell.fill?.type === 'pattern' ? argbToCss(cell.fill.fgColor) : null;
      if (fill) style.push(`background-color:${fill}`);
      if (cell.font?.bold) cls.push('is-bold');
      if (cell.font?.size && cell.font.size >= 13) style.push(`font-size:${(cell.font.size / 11).toFixed(2)}em`);
      const align = cell.alignment?.horizontal;
      if (align === 'center' || align === 'right' || align === 'left') style.push(`text-align:${align}`);
      else if (isNumberLike(cell.value)) cls.push('is-number');
      if (marks.some((m) => inRange(m, r, c))) cls.push('is-highlight');
      const merge = mergeAt(r, c);
      if (merge) {
        const cs = Math.min(merge.right, view.right) - c + 1;
        const rs = Math.min(merge.bottom, view.bottom) - r + 1;
        if (cs > 1) attrs.push(`colspan="${cs}"`);
        if (rs > 1) attrs.push(`rowspan="${rs}"`);
      }
      const shown = formatValue(cell.value, cell.numFmt);
      const raw = rawValue(cell.value);
      if (raw !== shown) attrs.push(`data-raw="${escapeHtml(raw)}"`);
      if (after?.has(ref)) {
        const v = after.get(ref);
        attrs.push(`data-after="${escapeHtml(formatValue(v, cell.numFmt))}"`);
        attrs.push(`data-after-raw="${escapeHtml(rawValue(v))}"`);
        if (typeof v === 'number' && !align) attrs.push('data-after-number');
      }
      if (style.length) attrs.push(`style="${style.join(';')}"`);
      tds.push(`<td class="${cls.join(' ')}" ${attrs.join(' ')}>${escapeHtml(shown)}</td>`);
    }
    rows.push(`<tr>${tds.join('')}</tr>`);
  }

  const tabs = wb.worksheets
    .map((s) => `<span class="sheet__tab${s === ws ? ' is-active' : ''}">${escapeHtml(s.name)}</span>`)
    .join('');

  return (
    `<div class="sheet not-content">` +
    `<div class="sheet__titlebar">` +
    `<span class="sheet__app sheet__app--excel">Excel</span>` +
    `<span class="sheet__app sheet__app--sheets">スプレッドシート</span>` +
    `<span class="sheet__file">${escapeHtml(fileName)}</span>` +
    `</div>` +
    `<div class="sheet__formulabar">` +
    `<span class="sheet__namebox"></span><span class="sheet__fx">fx</span><span class="sheet__formula"></span>` +
    `</div>` +
    `<div class="sheet__grid">` +
    `<table class="sheet__table"><colgroup>${cols.join('')}</colgroup>` +
    `<thead><tr>${head.join('')}</tr></thead><tbody>${rows.join('')}</tbody></table>` +
    `</div>` +
    `<div class="sheet__tabs">${tabs}</div>` +
    `</div>`
  );
}

// ---- 属性・ファイル ----

/** ブックや コードを読むレクチャーディレクトリ。src="<sec>/<lec>" なら別レクチャー。 */
function lectureDirFor(attrs, lectureAbsDir) {
  if (!attrs.src) return lectureAbsDir;
  const parts = attrs.src.split('/');
  if (parts.length !== 2 || parts.some((p) => !/^[\w-]+$/.test(p))) return null;
  return path.join(lectureAbsDir, '..', '..', parts[0], parts[1]);
}

function safeRel(rel) {
  const s = String(rel ?? '').trim().replace(/^\/+/, '');
  if (!s || s.split('/').some((seg) => seg === '..' || seg === '.')) return null;
  return s;
}

async function sheetHtml(attrs, lectureAbsDir, after) {
  const dir = lectureDirFor(attrs, lectureAbsDir);
  const rel = safeRel(attrs.book);
  if (!dir || !rel) return { error: `invalid book="${attrs.book ?? ''}"` };
  const abs = path.join(dir, 'example', ...rel.split('/'));
  let wb;
  try {
    wb = await loadBook(abs);
  } catch (e) {
    return { error: `cannot read ${path.relative(path.join(lectureAbsDir, '..', '..'), abs)} (${e.code ?? e.message})` };
  }
  const ws = attrs.sheet ? wb.getWorksheet(attrs.sheet) : wb.worksheets[0];
  if (!ws) return { error: `no sheet "${attrs.sheet}" in ${rel}` };
  return {
    html: renderSheet(wb, ws, {
      range: attrs.range,
      highlight: attrs.highlight,
      after,
      fileName: path.basename(rel),
    }),
  };
}

// ---- ::::sheetdemo ----

function parseValue(s) {
  const t = s.trim();
  if (/^-?\d+(\.\d+)?$/.test(t)) return Number(t);
  // 数式も書ける（数式バーに出す）。セルにはそのまま表示する。
  if (t.startsWith('=')) return { formula: t.slice(1), result: undefined };
  return t.replace(/^"(.*)"$/, '$1');
}

/** `F2:F31 = 1, 2, 3` の行を Map<ref, 値> に足す。読めなければ false。 */
function parseAfterLine(line, after) {
  const m = line.match(/^\s*([A-Za-z]+\d+(?::[A-Za-z]+\d+)?)\s*=\s*(.*)$/);
  if (!m) return false;
  const r = parseRange(m[1]);
  const values = m[2].split(',').map(parseValue);
  let i = 0;
  for (let row = r.top; row <= r.bottom; row++) {
    for (let col = r.left; col <= r.right; col++) {
      if (i >= values.length) return true;
      after.set(`${numToCol(col)}${row}`, values[i++]);
    }
  }
  return true;
}

export async function buildSheetDemo({ caption, attrs, lines, lectureAbsDir, where }) {
  const after = new Map();
  for (const line of lines) {
    if (!line.trim()) continue;
    if (!parseAfterLine(line, after)) console.warn(`[sheets] ::::sheetdemo: cannot read "${line}" in ${where}`);
  }
  const { html, error } = await sheetHtml(attrs, lectureAbsDir, after);
  if (error) {
    console.warn(`[sheets] ::::sheetdemo ${error} in ${where} (skipped)`);
    return null;
  }
  const cap = caption ? `<figcaption class="sheetdemo__caption">${escapeHtml(caption)}</figcaption>` : '';
  const buttons = after.size
    ? `<div class="sheetdemo__bar">` +
      `<button type="button" class="sheetdemo__run">▶ 実行</button>` +
      `<button type="button" class="sheetdemo__reset" disabled>↺ リセット</button>` +
      `<span class="sheetdemo__note">ブラウザ上での再現です</span>` +
      `</div>`
    : '';
  return `<figure class="sheetdemo not-content">${buttons}${html}${cap}</figure>`;
}

// ---- ::sheetcode ----

/** VBA の `Sub Name(` 〜 `End Sub` を、直前のコメント行ごと切り出す。 */
function extractVba(src, name) {
  const lines = src.split('\n');
  const start = lines.findIndex((l) => new RegExp(`^(Public |Private )?(Sub|Function) ${name}\\b`, 'i').test(l));
  if (start === -1) return null;
  const end = lines.findIndex((l, i) => i > start && /^End (Sub|Function)\b/i.test(l));
  if (end === -1) return null;
  let top = start;
  while (top > 0 && /^\s*'/.test(lines[top - 1])) top--;
  return lines.slice(top, end + 1).join('\n');
}

/** GAS の `function name(` 〜 行頭の `}` を、直前のコメント行ごと切り出す。 */
function extractGas(src, name) {
  const lines = src.split('\n');
  const start = lines.findIndex((l) => new RegExp(`^function ${name}\\s*\\(`).test(l));
  if (start === -1) return null;
  const end = lines.findIndex((l, i) => i > start && /^}/.test(l));
  if (end === -1) return null;
  let top = start;
  while (top > 0 && /^\s*\/\//.test(lines[top - 1])) top--;
  return lines.slice(top, end + 1).join('\n');
}

// VBA の CalcAmounts に対する GAS は calcAmounts（AGENTS.md の命名ルール）。
const gasNameOf = (vbaName) => vbaName.charAt(0).toLowerCase() + vbaName.slice(1);

async function readCode(abs) {
  try {
    return (await readFile(abs, 'utf8')).replace(/\r\n/g, '\n').replace(/\n+$/, '');
  } catch (e) {
    if (e.code === 'ENOENT') return null;
    throw e;
  }
}

export async function buildSheetCode({ attrs, lectureAbsDir, where }) {
  const dir = lectureDirFor(attrs, lectureAbsDir);
  const code = safeRel(attrs.code);
  if (!dir || !code) {
    console.warn(`[sheets] ::sheetcode needs code="<name>" in ${where} (skipped)`);
    return null;
  }
  const { html, error } = await sheetHtml(attrs, lectureAbsDir, null);
  if (error) {
    console.warn(`[sheets] ::sheetcode ${error} in ${where} (skipped)`);
    return null;
  }

  // Excel が主なので VBA を先に並べる。
  const panes = [
    { key: 'excel', lang: 'vb', file: `vba/${code}.bas`, pick: (s) => extractVba(s, attrs.macro) },
    { key: 'sheets', lang: 'js', file: `gas/${code}.gs`, pick: (s) => extractGas(s, gasNameOf(attrs.macro)) },
  ];
  const out = [`::::sheetcode`];
  for (const p of panes) {
    const src = await readCode(path.join(dir, 'example', ...p.file.split('/')));
    if (src === null) {
      console.warn(`[sheets] ::sheetcode: ${p.file} not found in ${where}`);
      continue;
    }
    const body = attrs.macro ? p.pick(src) : src;
    if (body === null) {
      console.warn(`[sheets] ::sheetcode: macro "${attrs.macro}" not found in ${p.file} (${where})`);
      continue;
    }
    out.push('```' + p.lang + ` data-platform="${p.key}" data-file="${p.file}"`);
    out.push(body);
    out.push('```');
  }
  if (out.length === 1) return null;
  out.push(html);
  out.push('::::');
  return out.join('\n');
}

/**
 * 本文中の `::::sheetdemo` 〜 `::::` と `::sheetcode` を展開する。
 * レクチャーの外（example/ を持たない docs）では呼ばない。
 */
export async function expandSheetSentinels(body, { lectureAbsDir, where, parseAttrs }) {
  if (!/^::(::)?sheet(demo|code)\b/m.test(body)) return body;
  const lines = body.split('\n');
  const out = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const demo = line.match(SHEETDEMO_OPEN_RE);
    if (demo) {
      const inner = [];
      let j = i + 1;
      while (j < lines.length && !CLOSE_RE.test(lines[j])) inner.push(lines[j++]);
      i = j;
      const block = await buildSheetDemo({
        caption: demo[1]?.trim(),
        attrs: parseAttrs(demo[2]),
        lines: inner,
        lectureAbsDir,
        where,
      });
      if (block) out.push(block);
      continue;
    }
    const sc = line.match(SHEETCODE_RE);
    if (sc) {
      const block = await buildSheetCode({ attrs: parseAttrs(sc[2]), lectureAbsDir, where });
      if (block) out.push(block);
      continue;
    }
    out.push(line);
  }
  return out.join('\n');
}

/**
 * `::::debugdemo` センチネル（sentinels.mjs から呼ばれる）。デバッガーの操作を 1 手ずつ見せる。
 *
 *   ::::debugdemo[キャプション]{code="SalesSummary" macro="SumTotal" book="sales.xlsm"}
 *   break: total = total +
 *     止めたい行にブレークポイントを置きます。
 *   run: total = total +
 *     vars: lastRow = 31, i = 2, total = 0
 *     vars.sheets: sheet = Sheet
 *     黄色い行は「これから実行する行」です。
 *   step: Next i | for (let i
 *     vars: total = 1250
 *   continue: total = total +
 *   stop:
 *   ::::
 *
 *   … Excel では VBE（ブレークポイントの茶色の丸・黄色い行・ローカル ウィンドウ）を、
 *     スプレッドシートでは Apps Script エディタ（デバッグ実行・デバッガのパネルの変数）を描き、
 *     「次へ ▶」か、光っているボタン（行）を押すと 1 手ずつ進む。コードは実際には動かさない。
 *
 * 手順の行（行頭から書く）:
 *   break: <行>      … その行のブレークポイントを置く（置いてあれば外す）
 *   run: <行>        … 実行して、その行で止まる（<行> を書かなければ最後まで動いて終わる）
 *   step: <行>       … 1 行進めて、その行で止まる（Excel は F8、GAS はステップ オーバー）
 *   continue: <行>   … 続きを実行して、次に止まる行（書かなければ最後まで動いて終わる）
 *   stop:            … 実行を止める
 *   <行> は、その行に含まれる文字列（最初に見つかった行）。`VBA の文字列 | GAS の文字列` で環境ごとに書ける。
 *
 * 手順の下に字下げして書く行:
 *   vars: 名前 = 値, …   … 止まったときの変数の値（変わったものだけでよい。前の手順の値を引き継ぐ）
 *                         `vars.excel:` ／ `vars.sheets:` で環境ごとに書ける。値は書いたとおりに出す（文字列は "…" で囲む）
 *   log: 文言            … イミディエイト ウィンドウ ／ 実行ログに 1 行足す（`log.excel:` ／ `log.sheets:` も）
 *   excel: 説明 ／ sheets: 説明 … その環境でだけ出す説明
 *   それ以外             … 両方に出す説明（`式` はコードの書体になる）
 *
 * 属性:
 *   code  … example/vba/<code>.bas と example/gas/<code>.gs を描く（必須）
 *   macro … VBA の Sub（GAS は先頭を小文字にした function）だけを切り出す（省略するとファイル全体）
 *   book  … VBE のタイトルに出すブックの名前（省略すると出さない）
 *   src   … 別レクチャーの example/ を使うとき `<sec>/<lec>`
 *   stage … example/ の代わりに stages/<stage>/ のコードを描く（完成の手前の状態を見せるとき。sheets.mjs の codeDirFor）
 *
 * 各手順のあとの状態（ブレークポイント・止まっている行・変数・ログ）はここで計算して data-dstates に持たせ、
 * site/src/scripts/debug-client.js がそれを描く。装飾は site/src/styles/debug.css。
 */

import path from 'node:path';
import {
  escapeHtml,
  extractGas,
  extractVba,
  gasNameOf,
  highlightGas,
  inlineHtml,
  codeDirFor,
  readCode,
} from './sheets.mjs';

export const DEBUGDEMO_OPEN_RE = /^::::debugdemo(?:\[([^\]]*)\])?(?:\{([^}]*)\})?\s*$/;
const CLOSE_RE = /^::::\s*$/;

const PLATFORMS = ['excel', 'sheets'];
const ACTION_RE = /^(break|run|step|continue|stop)\s*:?\s*(.*)$/;
const DETAIL_RE = /^(vars|log|excel|sheets)(?:\.(excel|sheets))?\s*:\s*(.*)$/;

// 手順の言い方。label は手順の案内（ステッパー）に出す短い印、title は説明の見出し。
const ACTIONS = {
  excel: {
    break: { label: 'F9', title: '左の灰色の帯をクリックして、ブレークポイントを置く（F9）', off: 'もう一度クリックして、ブレークポイントを外す' },
    run: { label: 'F5', title: 'F5 で実行する' },
    step: { label: 'F8', title: 'F8 で 1 行進める（ステップ イン）' },
    continue: { label: 'F5', title: 'F5 で続きを実行する（継続）' },
    stop: { label: '■', title: '■（リセット）で実行を止める' },
  },
  sheets: {
    break: { label: '●', title: '行番号の左をクリックして、ブレークポイントを置く', off: 'もう一度クリックして、ブレークポイントを外す' },
    run: { label: 'デバッグ', title: '「デバッグ」で実行する' },
    step: { label: '↷', title: 'ステップ オーバー（↷）で 1 行進める' },
    continue: { label: '▶', title: '再開（▶）で続きを実行する' },
    stop: { label: '■', title: '■（停止）で実行を止める' },
  },
};

// VBA の Dim の型ごとの、代入する前の値（ローカル ウィンドウの表示）。
const VBA_DEFAULTS = { Long: '0', Integer: '0', Double: '0', String: '""', Date: '0:00:00', Boolean: 'False', Worksheet: 'Nothing', Range: 'Nothing', Variant: 'Empty' };

const VBA_KEYWORDS = new Set(
  ('Option Explicit Sub End Function Dim As Long Integer String Double Date Boolean Variant Worksheet Range ' +
    'For To Step Next If Then Else ElseIf Each In Set Nothing True False Do Loop While Wend Until With Exit ' +
    'Const Private Public And Or Not Call ByVal ByRef Debug Print Stop')
    .split(' '),
);

const VBA_TOKEN_RE = /("(?:[^"]|"")*")|('.*$)|\b([A-Za-z_]\w*)\b|\b(\d+(?:\.\d+)?)\b/g;

function highlightVba(line) {
  let out = '';
  let last = 0;
  for (const m of line.matchAll(VBA_TOKEN_RE)) {
    let cls = null;
    if (m[1]) cls = 'string';
    else if (m[2]) cls = 'comment';
    else if (m[3]) cls = VBA_KEYWORDS.has(m[3]) ? 'keyword' : null;
    else cls = 'number';
    if (!cls) continue;
    out += escapeHtml(line.slice(last, m.index));
    out += `<span class="ddemo__tok--${cls}">${escapeHtml(m[0])}</span>`;
    last = m.index + m[0].length;
  }
  return out + escapeHtml(line.slice(last));
}

/** コード（既定は example/）を読み、描く範囲（macro の Sub ／ function、無ければ全体）と、ファイルでの開始行を返す。 */
async function loadCode(codeDir, file, pick) {
  const src = await readCode(path.join(codeDir, ...file.split('/')));
  if (src === null) return { error: `${file} not found` };
  const body = pick ? pick(src) : src;
  if (body === null) return { error: `macro not found in ${file}` };
  const start = src.slice(0, src.indexOf(body)).split('\n').length;
  return { lines: body.split('\n'), start };
}

/** 宣言している変数を、書いた順に { name, type } で返す（VBA は Dim、GAS は const / let / var）。 */
function declaredVars(platform, lines) {
  const out = [];
  const re = platform === 'excel' ? /^\s*Dim\s+(\w+)\s+As\s+(\w+)/i : /\b(?:const|let|var)\s+(\w+)/;
  for (const l of lines) {
    const m = l.match(re);
    if (m && !out.some((v) => v.name === m[1])) out.push({ name: m[1], type: m[2] ?? '' });
  }
  return out;
}

/** `a = 1, b = "x, y"` を [[名前, 値], …] に。値の中のカンマで切らないよう、次の `名前 =` の前でだけ切る。 */
function parseVars(text) {
  return text
    .split(/,\s*(?=[A-Za-z_]\w*\s*=)/)
    .map((part) => part.match(/^\s*([A-Za-z_]\w*)\s*=\s*(.*?)\s*$/))
    .filter(Boolean)
    .map((m) => [m[1], m[2]]);
}

/** 手順と、その下に字下げして書いた行を読む。 */
function parseSteps(lines, where) {
  const steps = [];
  for (const raw of lines) {
    if (!raw.trim()) continue;
    const indented = /^\s/.test(raw);
    const line = raw.trim();
    if (!indented) {
      const a = line.match(ACTION_RE);
      if (!a) {
        console.warn(`[debugdemo] cannot read "${line}" in ${where}`);
        continue;
      }
      const [excel, sheets] = a[2].split('|').map((s) => s.trim());
      steps.push({
        action: a[1],
        target: { excel: excel || null, sheets: (sheets ?? excel) || null },
        vars: { excel: [], sheets: [] },
        log: { excel: [], sheets: [] },
        notes: [],
      });
      continue;
    }
    const step = steps.at(-1);
    if (!step) {
      console.warn(`[debugdemo] "${line}" comes before any step in ${where}`);
      continue;
    }
    const d = line.match(DETAIL_RE);
    if (d && (d[1] === 'vars' || d[1] === 'log')) {
      const targets = d[2] ? [d[2]] : PLATFORMS;
      for (const p of targets) {
        if (d[1] === 'vars') step.vars[p].push(...parseVars(d[3]));
        else step.log[p].push(d[3].replace(/^"(.*)"$/, '$1'));
      }
    } else if (d && !d[2]) {
      step.notes.push({ platform: d[1], text: d[3] });
    } else {
      step.notes.push({ platform: null, text: line });
    }
  }
  return steps;
}

/** 文字列を含む最初の行の番号（0 始まり）。見つからなければ null。 */
function findLine(lines, needle) {
  if (!needle) return null;
  const i = lines.findIndex((l) => l.includes(needle));
  return i === -1 ? null : i;
}

/**
 * 手順を順に当てはめて、各手順のあとの状態を作る（0 番目は何もしていない状態）。
 * 状態は環境ごとに { bps, cur, paused, vars: [[名前, 値, 型, 変わったか]], log }。
 */
function buildStates(steps, code, where) {
  const fresh = () => ({ bps: [], cur: null, paused: false, vars: new Map(), log: [] });
  const st = { excel: fresh(), sheets: fresh() };
  const snapshot = (changed) =>
    Object.fromEntries(
      PLATFORMS.map((p) => {
        const s = st[p];
        const types = new Map(code[p].decl.map((v) => [v.name, v.type]));
        const vars = s.paused
          ? [...s.vars].map(([name, value]) => [name, value, types.get(name) ?? '', changed[p].has(name)])
          : [];
        return [p, { bps: [...s.bps], cur: s.cur, paused: s.paused, vars, log: [...s.log] }];
      }),
    );
  const states = [snapshot({ excel: new Set(), sheets: new Set() })];
  const lines = steps.map((step, k) => {
    const changed = { excel: new Set(), sheets: new Set() };
    const at = {};
    for (const p of PLATFORMS) {
      const s = st[p];
      const c = code[p];
      if (!c) continue;
      const line = findLine(c.lines, step.target[p]);
      if (step.target[p] && line === null) console.warn(`[debugdemo] step ${k + 1}: "${step.target[p]}" not found in ${c.file} (${where})`);
      at[p] = line;
      const end = () => {
        s.paused = false;
        s.cur = null;
        s.vars = new Map();
      };
      if (step.action === 'break') {
        if (line === null) continue;
        s.bps = s.bps.includes(line) ? s.bps.filter((b) => b !== line) : [...s.bps, line].sort((a, b) => a - b);
        at[p + 'Off'] = !s.bps.includes(line);
      } else if (step.action === 'run') {
        // 実行を始めると、宣言した変数が代入前の値で並ぶ（GAS は undefined）。
        s.vars = new Map(c.decl.map((v) => [v.name, p === 'excel' ? VBA_DEFAULTS[v.type] ?? 'Empty' : 'undefined']));
        s.log = [];
        if (line === null) end();
        else {
          s.paused = true;
          s.cur = line;
        }
      } else if (step.action === 'step' || step.action === 'continue') {
        if (line === null) end();
        else s.cur = line;
      } else if (step.action === 'stop') {
        end();
      }
      if (s.paused) {
        for (const [name, value] of step.vars[p]) {
          if (s.vars.get(name) !== value) changed[p].add(name);
          s.vars.set(name, value);
        }
      }
      s.log.push(...step.log[p]);
    }
    states.push(snapshot(changed));
    return at;
  });
  return { states, lines };
}

// ---- 描画 ----

function codeLinesHtml(platform, c) {
  return c.lines
    .map((l, i) => {
      const text = platform === 'excel' ? highlightVba(l) : highlightGas(l);
      const no = platform === 'sheets' ? `<span class="ddemo__no">${c.start + i}</span>` : '';
      return (
        `<div class="ddemo__line" data-line="${i}">` +
        `<button type="button" class="ddemo__gutter" data-line="${i}" aria-label="${c.start + i} 行目のブレークポイント"></button>` +
        no +
        `<span class="ddemo__text">${text || ' '}</span></div>`
      );
    })
    .join('');
}

function logHtml(title) {
  return (
    `<div class="ddemo__log" hidden>` +
    `<div class="ddemo__panel-title">${title}</div>` +
    `<div class="ddemo__log-body" role="log"></div>` +
    `</div>`
  );
}

function vbeHtml(c, attrs, hasLog) {
  const menus = ['ファイル', '編集', '表示', '挿入', '書式', 'デバッグ', '実行', 'ツール', 'アドイン', 'ウィンドウ', 'ヘルプ']
    .map((t) => `<span>${t}</span>`)
    .join('');
  const book = attrs.book ? ` - ${escapeHtml(attrs.book)}` : '';
  return (
    `<div class="ddemo__app ddemo__vbe sheet__only--excel" data-dplatform="excel">` +
    `<div class="ddemo__vbe-title">Microsoft Visual Basic for Applications${book}<span class="ddemo__mode" hidden> [中断]</span></div>` +
    `<div class="ddemo__vbe-menu" aria-hidden="true">${menus}</div>` +
    `<div class="ddemo__vbe-toolbar">` +
    `<button type="button" class="ddemo__tool" data-act="run continue" title="Sub/ユーザー フォームの実行 ／ 継続 (F5)"><span class="ddemo__icon ddemo__icon--run" aria-hidden="true">▶</span>実行 (F5)</button>` +
    `<span class="ddemo__tool is-disabled" aria-hidden="true"><span class="ddemo__icon" aria-hidden="true">❚❚</span>中断</span>` +
    `<button type="button" class="ddemo__tool" data-act="stop" title="リセット"><span class="ddemo__icon ddemo__icon--stop" aria-hidden="true">■</span>リセット</button>` +
    `<span class="ddemo__sep" aria-hidden="true"></span>` +
    `<button type="button" class="ddemo__tool" data-act="step" title="ステップ イン (F8)"><span class="ddemo__icon" aria-hidden="true">⤵</span>ステップ イン (F8)</button>` +
    `</div>` +
    `<div class="ddemo__vbe-window">` +
    `<div class="ddemo__vbe-wintitle">${escapeHtml(c.module)} (コード)</div>` +
    `<div class="ddemo__vbe-procs" aria-hidden="true"><span>(General)</span><span>${escapeHtml(attrs.macro || '(Declarations)')}</span></div>` +
    `<div class="ddemo__code" tabindex="0" role="region" aria-label="${escapeHtml(c.file)}">${codeLinesHtml('excel', c)}</div>` +
    `</div>` +
    `<div class="ddemo__panes">` +
    `<div class="ddemo__locals">` +
    `<div class="ddemo__panel-title">ローカル</div>` +
    `<table class="ddemo__vars"><thead><tr><th>式</th><th>値</th><th>型</th></tr></thead><tbody></tbody></table>` +
    `<p class="ddemo__empty">&lt;実行中ではありません&gt;</p>` +
    `</div>` +
    (hasLog ? logHtml('イミディエイト') : '') +
    `</div>` +
    `</div>`
  );
}

function gasHtml(c, attrs, hasLog) {
  const fn = attrs.macro ? gasNameOf(attrs.macro) : '';
  return (
    `<div class="ddemo__app ddemo__gas sheet__only--sheets" data-dplatform="sheets">` +
    `<div class="ddemo__gas-bar"><span class="ddemo__gas-logo" aria-hidden="true"></span><span class="ddemo__gas-name">Apps Script</span>` +
    `<span class="ddemo__gas-project">無題のプロジェクト</span></div>` +
    `<div class="ddemo__gas-toolbar">` +
    `<span class="ddemo__gbtn is-disabled" aria-hidden="true">▶ 実行</span>` +
    `<button type="button" class="ddemo__gbtn" data-act="run">デバッグ</button>` +
    (fn ? `<span class="ddemo__gas-fn">${escapeHtml(fn)}</span>` : '') +
    `<span class="ddemo__sep" aria-hidden="true"></span>` +
    `<span class="ddemo__gas-logtoggle" aria-hidden="true">実行ログ</span>` +
    `</div>` +
    `<div class="ddemo__gas-main">` +
    `<div class="ddemo__gas-left">` +
    `<div class="ddemo__gas-file">${escapeHtml(path.basename(c.file))}</div>` +
    `<div class="ddemo__code" tabindex="0" role="region" aria-label="${escapeHtml(c.file)}">${codeLinesHtml('sheets', c)}</div>` +
    (hasLog ? logHtml('実行ログ') : '') +
    `</div>` +
    `<div class="ddemo__debugger" hidden>` +
    `<div class="ddemo__panel-title">デバッガ</div>` +
    `<div class="ddemo__debugger-tools">` +
    `<button type="button" class="ddemo__dbtn" data-act="continue" title="再開">▶</button>` +
    `<button type="button" class="ddemo__dbtn" data-act="step" title="ステップ オーバー">↷</button>` +
    `<span class="ddemo__dbtn is-disabled" title="ステップ イン" aria-hidden="true">↓</span>` +
    `<span class="ddemo__dbtn is-disabled" title="ステップ アウト" aria-hidden="true">↑</span>` +
    `<button type="button" class="ddemo__dbtn" data-act="stop" title="停止">■</button>` +
    `</div>` +
    `<div class="ddemo__scope">変数</div>` +
    `<div class="ddemo__scope-name">Local</div>` +
    `<ul class="ddemo__gvars"></ul>` +
    `</div>` +
    `</div>` +
    `</div>`
  );
}

export async function buildDebugDemo({ caption, attrs, lines, lectureAbsDir, where }) {
  const dir = codeDirFor(attrs, lectureAbsDir);
  if (!dir || !attrs.code || !/^[\w-]+$/.test(attrs.code)) {
    console.warn(`[debugdemo] needs code="<name>" in ${where} (skipped)`);
    return null;
  }
  const code = {};
  for (const p of PLATFORMS) {
    const file = p === 'excel' ? `vba/${attrs.code}.bas` : `gas/${attrs.code}.gs`;
    const pick = attrs.macro ? (s) => (p === 'excel' ? extractVba(s, attrs.macro) : extractGas(s, gasNameOf(attrs.macro))) : null;
    const c = await loadCode(dir, file, pick);
    if (c.error) {
      console.warn(`[debugdemo] ${c.error} (${where})`);
      code[p] = null;
      continue;
    }
    code[p] = { ...c, file, module: attrs.code, decl: declaredVars(p, c.lines) };
  }
  if (!code.excel && !code.sheets) return null;
  const steps = parseSteps(lines, where);
  if (!steps.length) return null;
  const { states, lines: at } = buildStates(steps, code, where);

  // 次に押すところ。break は行の帯、ほかはツールバーのボタン（data-act）を光らせる。
  const plan = steps.map((s, k) => ({ act: s.action, line: { excel: at[k].excel ?? null, sheets: at[k].sheets ?? null } }));
  const hasLog = (p) => steps.some((s) => s.log[p].length);

  const label = (p, s, k) => (s.action === 'break' && at[k][p + 'Off'] ? '○' : ACTIONS[p][s.action].label);
  const title = (p, s, k) => (s.action === 'break' && at[k][p + 'Off'] ? ACTIONS[p].break.off : ACTIONS[p][s.action].title);
  const both = (fn) => PLATFORMS.map((p) => `<span class="sheet__only--${p}">${fn(p)}</span>`).join('');

  const bar =
    `<div class="sheetdemo__bar">` +
    `<button type="button" class="fdemo__prev ddemo__prev" disabled>◀ 前へ</button>` +
    `<button type="button" class="fdemo__next ddemo__next">次へ ▶</button>` +
    `<button type="button" class="ddemo__reset">↺ 最初から</button>` +
    `<span class="sheetdemo__note">ブラウザ上での再現です</span>` +
    `<button type="button" class="sheetdemo__fullscreen" aria-pressed="false">⛶ 全画面</button>` +
    `</div>`;
  const guide =
    `<ol class="sheetdemo__guide fdemo__steps ddemo__steps" aria-label="手順">` +
    steps
      .map(
        (s, k) =>
          `<li data-dstep="${k + 1}"><button type="button" class="fdemo__goto" title="${escapeHtml(title('excel', s, k))}">` +
          both((p) => escapeHtml(label(p, s, k))) +
          `</button></li>`,
      )
      .join('') +
    `</ol>`;
  const notesHtml = (notes) =>
    notes
      .map((n) => {
        const cls = n.platform ? ` sheet__only--${n.platform}` : '';
        return `<p class="fdemo__say-note${cls}">${inlineHtml(n.text)}</p>`;
      })
      .join('');
  const explain =
    `<div class="fdemo__explain" aria-live="polite">` +
    `<div class="ddemo__say" data-dstep="0"><p class="fdemo__say-title">「次へ ▶」か、光っているところを押して進めます</p>` +
    `<p class="fdemo__say-note">止まっている行と変数の値が、1 手ごとに変わっていきます。</p></div>` +
    steps
      .map(
        (s, k) =>
          `<div class="ddemo__say" data-dstep="${k + 1}" hidden>` +
          `<p class="fdemo__say-title">${both((p) => escapeHtml(title(p, s, k)))}</p>` +
          notesHtml(s.notes) +
          `</div>`,
      )
      .join('') +
    `</div>`;
  const apps = (code.excel ? vbeHtml(code.excel, attrs, hasLog('excel')) : '') + (code.sheets ? gasHtml(code.sheets, attrs, hasLog('sheets')) : '');
  const cap = caption ? `<figcaption class="sheetdemo__caption">${escapeHtml(caption)}</figcaption>` : '';
  const data = escapeHtml(JSON.stringify({ states, plan }));
  return (
    `<figure class="sheetdemo ddemo not-content" data-dsteps="${steps.length}" data-dstep="0" data-ddemo="${data}">` +
    bar +
    guide +
    explain +
    `<div class="ddemo__stage">${apps}</div>` +
    cap +
    `</figure>`
  );
}

/** 本文中の `::::debugdemo` 〜 `::::` を展開する。 */
export async function expandDebugSentinels(body, { lectureAbsDir, where, parseAttrs }) {
  if (!/^::::debugdemo\b/m.test(body)) return body;
  const lines = body.split('\n');
  const out = [];
  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(DEBUGDEMO_OPEN_RE);
    if (!m) {
      out.push(lines[i]);
      continue;
    }
    const inner = [];
    let j = i + 1;
    while (j < lines.length && !CLOSE_RE.test(lines[j])) inner.push(lines[j++]);
    i = j;
    const block = await buildDebugDemo({ caption: m[1]?.trim(), attrs: parseAttrs(m[2]), lines: inner, lectureAbsDir, where });
    if (block) out.push(block);
  }
  return out.join('\n');
}

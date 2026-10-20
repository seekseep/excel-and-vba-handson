/**
 * `::::editor{title="..." book="URL" sheet="URL" open="main.js"}` コンテナディレクティブを、簡易エディタ UI
 * `<section class="editor">`（左=ファイル一覧・右=コード）に変換する remark プラグイン。
 *
 * 属性:
 *   title … ヘッダの文言（省略時は「このステップの完成例」）
 *   book  … 節を終えた状態のマクロ入りブック（.xlsm）のダウンロード URL。Excel を選んでいるときだけボタンを出す。
 *           省略するとボタンを出さない（デモなど配るものが無い場合）
 *   sheet … 節を終えた状態のサンプルのスプレッドシート（Apps Script 入り）の「コピーを作成」URL。
 *           スプレッドシートを選んでいるときだけボタンを出す（sentinels.mjs が spreadsheets.json から渡す）
 *   open  … 最初に開くファイル名（省略時は先頭）
 *
 * コンテナの中身は sync-lectures.mjs（libs/sentinels.mjs）が実ファイルから生成した
 * コードフェンス（各 fence の meta に `data-file="<相対パス>"`）。このプラグインは:
 *   - 各コードフェンスからファイル名を読み取り（meta は消して Expressive Code の枠を出さない）
 *   - 左のファイル切り替えボタン、右のコードペインを組み立てる
 *   - ヘッダにブックのダウンロード・スプレッドシートのコピーのボタンを付ける
 * タブ切り替えの挙動は site/src/scripts/editor-client.js（全ページに inline 注入）。
 * 装飾は site/src/styles/editor.css。
 *
 * NOTE: 中でボタンを自前生成するため remark-download には依存しない。
 * コンテナは 4 コロン（::::）で開く（中にコードフェンスを含むため 3 コロンでもよいが統一）。
 */

const NAME = 'editor';
const FILE_META_RE = /data-file="([^"]*)"/;
const DEFAULT_TITLE = 'このステップの完成例';

// 一覧に出す名前。Office マクロ系の vba/ ・ gas/ は、どちらの環境かを本文で選んでいるので出さない。
// （data-file には元のパスを残す。環境の出し分けと syncEditors がそれを見る）
const displayName = (file) => file.replace(/^(vba|gas)\//, '');

function visit(node, callback) {
  if (!node || !Array.isArray(node.children)) return;
  for (const child of node.children) {
    callback(child);
    visit(child, callback);
  }
}

function el(hName, className, children, extraProps = {}) {
  return {
    type: 'element',
    data: { hName, hProperties: { className, ...extraProps } },
    children: children || [],
  };
}

function text(value) {
  return { type: 'text', value };
}

/** 完成版のサンプルのスプレッドシートを「コピーを作成」するボタン（新しいタブで開く）。 */
function sheetButton(sheetUrl) {
  return {
    type: 'element',
    data: {
      hName: 'a',
      hProperties: {
        className: ['download-button', 'platform-only--sheets'],
        href: sheetUrl,
        target: '_blank',
        rel: ['noopener', 'noreferrer'],
      },
    },
    children: [
      el('span', ['download-button__icon', 'download-button__icon--copy'], [], { 'aria-hidden': 'true' }),
      el('span', ['download-button__label'], [text('完成版のスプレッドシートをコピー')]),
    ],
  };
}

/** 完成ブックのダウンロードボタン（remark-download と同じ見た目）。 */
function bookButton(bookUrl) {
  return {
    type: 'element',
    data: {
      hName: 'a',
      hProperties: { className: ['download-button', 'platform-only--excel'], href: bookUrl, download: true },
    },
    children: [
      el('span', ['download-button__icon'], [], { 'aria-hidden': 'true' }),
      el('span', ['download-button__label'], [text('完成ブックをダウンロード')]),
    ],
  };
}

export default function remarkEditor() {
  return (tree) => {
    visit(tree, (node) => {
      if (node.type !== 'containerDirective') return;
      if (node.name !== NAME) return;

      const attrs = node.attributes || {};

      // コードフェンスを集め、ファイル名を取り出す。
      const codeNodes = node.children.filter((c) => c.type === 'code');
      const files = [];
      codeNodes.forEach((code, i) => {
        const m = (code.meta || '').match(FILE_META_RE);
        const name = m ? m[1] : code.lang || `file${i}`;
        code.meta = null; // Expressive Code にファイル名枠を出させない
        files.push({ name, code });
      });

      // デフォルトで開くファイル（open 属性で指定）。無ければ先頭。
      let activeIndex = attrs.open ? files.findIndex((f) => f.name === attrs.open) : -1;
      if (activeIndex === -1) activeIndex = 0;

      const fileButtons = files.map((f, i) =>
        el('button', ['editor__file', ...(i === activeIndex ? ['is-active'] : [])], [text(displayName(f.name))], {
          type: 'button',
          'data-file': f.name,
        }),
      );

      const panes = files.map((f, i) =>
        el('div', ['editor__pane', ...(i === activeIndex ? ['is-active'] : [])], [f.code], {
          'data-file': f.name,
        }),
      );

      const headerChildren = [el('span', ['editor__title'], [text(attrs.title || DEFAULT_TITLE)])];
      const actions = [];
      if (attrs.sheet) actions.push(sheetButton(attrs.sheet));
      if (attrs.book) actions.push(bookButton(attrs.book));
      if (actions.length > 0) headerChildren.push(el('span', ['editor__actions'], actions));

      const header = el('div', ['editor__header'], headerChildren);
      const body = el('div', ['editor__body'], [
        el('div', ['editor__files'], fileButtons),
        el('div', ['editor__code'], panes),
      ]);

      node.type = 'editorContainer';
      node.data = { hName: 'section', hProperties: { className: ['editor'] } };
      node.children = [header, body];
    });
  };
}

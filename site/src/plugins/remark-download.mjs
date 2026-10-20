/**
 * `:::download` コンテナディレクティブを「大きなダウンロードボタン」に変換する remark プラグイン。
 *
 * 中に書いた通常のリンクの URL とラベルを取り出し、`<a class="download-button">` に変換する。
 * URL の解決（`./start-book` → 実際のブックの URL など）は sync-lectures.mjs 側が済ませているため、
 * ここではリンクをボタンの見た目に変えるだけ。装飾は site/src/styles/download.css で当てる。
 *
 * 使い方（Markdown）:
 *   :::download
 *   [練習用ブックをダウンロード](./start-book)
 *   [サンプルのスプレッドシートをコピー](./start-spreadsheet)
 *   :::
 *
 * Google スプレッドシートの「コピーを作成」へのリンクはスプレッドシートを選んでいるときだけ、
 * それ以外（Excel のブック）は、スプレッドシートのリンクが並んでいるときは Excel を選んでいるときだけ出す。
 */

/** containerDirective を再帰的に拾うシンプルな visitor。 */
function visit(node, callback) {
  if (!node || !Array.isArray(node.children)) return;
  for (const child of node.children) {
    callback(child);
    visit(child, callback);
  }
}

// `:::download` の中に書いた Google スプレッドシートのリンク（sync-lectures.mjs が `./spreadsheet` を差し替えたもの）。
const SPREADSHEET_RE = /^https:\/\/docs\.google\.com\/spreadsheets\//;

/** 子ノードを再帰的にたどって link ノードをすべて返す。 */
function findLinks(node, out = []) {
  if (!node) return out;
  if (node.type === 'link') out.push(node);
  else if (Array.isArray(node.children)) for (const child of node.children) findLinks(child, out);
  return out;
}

/** ボタン 1 つ。label は表示する文字のノード、icon はアイコンの種類（CSS で描く）。 */
function button({ href, label, className = [], icon = 'download', download = true }) {
  const properties = { className: ['download-button', ...className], href };
  if (download) properties.download = true;
  else Object.assign(properties, { target: '_blank', rel: ['noopener', 'noreferrer'] });
  return {
    type: 'downloadButton',
    data: { hName: 'a', hProperties: properties },
    children: [
      {
        type: 'downloadButtonIcon',
        data: {
          hName: 'span',
          hProperties: { className: ['download-button__icon', `download-button__icon--${icon}`], 'aria-hidden': 'true' },
        },
      },
      {
        type: 'downloadButtonLabel',
        data: { hName: 'span', hProperties: { className: ['download-button__label'] } },
        // ボタンを複数作るときに同じノードを共有しないよう、毎回複製する。
        children: structuredClone(label),
      },
    ],
  };
}

export default function remarkDownload() {
  return (tree) => {
    visit(tree, (node) => {
      if (node.type !== 'containerDirective') return;
      if (node.name !== 'download') return;

      const links = findLinks(node);
      if (links.length === 0) return;
      const isSheet = (l) => SPREADSHEET_RE.test(l.url);
      const hasSheet = links.some(isSheet);

      const buttons = links.map((l) =>
        isSheet(l)
          ? button({ href: l.url, label: l.children, className: ['platform-only--sheets'], icon: 'copy', download: false })
          : button({ href: l.url, label: l.children, className: hasSheet ? ['platform-only--excel'] : [] }),
      );

      if (buttons.length === 1) {
        Object.assign(node, buttons[0]);
        return;
      }
      node.type = 'downloadButtons';
      node.data = { hName: 'span', hProperties: { className: ['download-buttons', 'download-buttons--row'] } };
      node.children = buttons;
    });
  };
}

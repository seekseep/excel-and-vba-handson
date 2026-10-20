/**
 * `:::download` コンテナディレクティブを「大きなダウンロードボタン」に変換する remark プラグイン。
 *
 * 中に書いた通常のリンクの URL とラベルを取り出し、`<a class="download-button">` に変換する。
 * URL の解決（`./project.zip` → 実際の配布 ZIP URL）は sync-lectures.mjs 側が済ませているため、
 * ここではリンクをボタンの見た目に変えるだけ。装飾は site/src/styles/download.css で当てる。
 *
 * 使い方（Markdown）:
 *   :::download
 *   [サンプルコードをダウンロード](./project.zip)
 *   [サンプルのスプレッドシートをコピー](./spreadsheet)
 *   :::
 *
 * 2 行目（任意）は Google スプレッドシートの「コピーを作成」へのリンク。スプレッドシートを選んでいるときだけ出し、
 * そのときの ZIP のボタンは「コードをダウンロード」にする（スプレッドシート用の ZIP はコードだけ）。
 */

import { platformZips } from './lib/platform-zip.mjs';

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

const textNodes = (value) => [{ type: 'text', value }];

export default function remarkDownload() {
  return (tree) => {
    visit(tree, (node) => {
      if (node.type !== 'containerDirective') return;
      if (node.name !== 'download') return;

      const links = findLinks(node);
      const zip = links.find((l) => !SPREADSHEET_RE.test(l.url));
      const sheet = links.find((l) => SPREADSHEET_RE.test(l.url));
      if (!zip && !sheet) return;

      const buttons = [];
      // スプレッドシートは、サンプルのスプレッドシートを自分のドライブにコピーして始める（ブックの ZIP は要らない）。
      if (sheet) {
        buttons.push(button({ href: sheet.url, label: sheet.children, className: ['platform-only--sheets'], icon: 'copy', download: false }));
      }
      if (zip) {
        // Office マクロ系で環境ごとの ZIP があるときは、Excel 用とスプレッドシート用のボタンを両方作り、
        // 選んでいる環境のほうだけを見せる（platform.css の platform-only--*）。
        const variants = platformZips(zip.url);
        if (variants) {
          for (const v of variants) {
            // スプレッドシート用の ZIP はコードだけなので、サンプルのスプレッドシートがある節ではそう書く。
            const label = v.platform === 'sheets' && sheet ? textNodes('コードをダウンロード') : zip.children;
            buttons.push(button({ href: v.url, label, className: [`platform-only--${v.platform}`] }));
          }
        } else {
          // 練習用ブックだけの ZIP。サンプルのスプレッドシートがあれば、スプレッドシートでは出さない。
          buttons.push(button({ href: zip.url, label: zip.children, className: sheet ? ['platform-only--excel'] : [] }));
        }
      }

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

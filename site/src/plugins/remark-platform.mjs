/**
 * `::::platform` コンテナディレクティブを「Excel（VBA）／ Google スプレッドシート（GAS）」の
 * 切り替えタブに変換する remark プラグイン。
 *
 * この教材は Excel を主に書き、同じ手順をスプレッドシートでも進められるようにしている。
 * 手順やコードが環境で分かれるところだけをこのコンテナで囲む。
 *
 * 使い方（Markdown）:
 *   ::::platform
 *   :::excel
 *   Excel での手順・VBA のコード
 *   :::
 *   :::sheets
 *   スプレッドシートでの手順・GAS のコード
 *   :::
 *   ::::
 *
 * どちらか片方だけでもよい（例: Excel にしか無い操作は :::excel だけを書く）。
 * 選んだ環境はページをまたいで共有し、1 か所で切り替えると全タブがそろう。
 * 切り替えの挙動は site/src/scripts/platform-client.js、装飾は site/src/styles/platform.css。
 */

const NAME = 'platform';
const PANELS = [
  { key: 'excel', label: 'Excel（VBA）' },
  { key: 'sheets', label: 'スプレッドシート（GAS）' },
];

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

export default function remarkPlatform() {
  return (tree) => {
    visit(tree, (node) => {
      if (node.type !== 'containerDirective' || node.name !== NAME) return;

      const found = PANELS.map((p) => ({
        ...p,
        node: node.children.find((c) => c.type === 'containerDirective' && c.name === p.key),
      })).filter((p) => p.node);
      if (found.length === 0) return;

      const tabs = found.map((p) =>
        el('button', ['platform__tab', `platform__tab--${p.key}`], [{ type: 'text', value: p.label }], {
          type: 'button',
          'data-platform': p.key,
        }),
      );
      const panels = found.map((p) =>
        el('div', ['platform__panel', `platform__panel--${p.key}`], p.node.children, {
          'data-platform': p.key,
        }),
      );

      node.type = 'platformContainer';
      // 片方しか無いときは「この環境だけの手順」と分かるよう印を付ける（CSS で注記を出す）。
      const only = found.length === 1 ? ` platform--only-${found[0].key}` : '';
      node.data = { hName: 'div', hProperties: { className: [`platform${only}`] } };
      node.children = [el('div', ['platform__tabs'], tabs, { role: 'tablist' }), ...panels];
    });
  };
}

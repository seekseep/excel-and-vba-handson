---
docs: true
title: この資料の作り方
---

# この資料の作り方

この教材は、AI エージェントの [Claude Code](https://docs.claude.com/ja/docs/claude-code/overview) と、2 つの自作スキル（Claude Code に作業の手順を覚えさせる拡張）を使って作りました。講師が課題と構成を決め、AI が本文・サンプルコード・練習用ブック・図の下書きを作り、講師が読んで直す、という分担です。

どちらのスキルも公開しているので、自分の勉強会の教材づくりにも使えます。Claude Code のスキルは、リポジトリを `~/.claude/skills/<スキル名>` に置くだけで使えるようになります。

## genfig ― 説明文から図を作る

[genfig](https://github.com/seekseep/genfig-skill) は、「ブラウザがサーバーにリクエストを送る」のような説明文から、FigJam 風の図（SVG）を作るスキルです。説明に含まれる関係（流れ・入れ物・分かれる・くり返す など）を読み取り、絵文字と図形を並べて矢印でつなぎます。

![月末の 5 つの作業を、genfig で図にしたもの](./01-why-automate/01-monthly-closing/images/01-monthly-work.svg)

この教材の図（各節の先頭の図や、上の「月末の 5 つの作業」の図）は、すべて genfig で作りました。図は 1 枚ずつ Python のスクリプトとして各節の `images/` に残してあるので、文言を直して作り直すこともできます。

```sh
git clone https://github.com/seekseep/genfig-skill.git ~/.claude/skills/genfig
```

入れたあとは、Claude Code に「この説明に図を入れて」「〜の図を作って」と頼むと使われます。

## handson-builder ― ハンズオン教材のリポジトリを組み立てる

[handson-builder](https://github.com/seekseep/handson-builder) は、勉強会・ハンズオン教材のリポジトリを丸ごと組み立てるスキルです。「章 ＋ 節」の 2 階層の本文、節ごとに動かせるサンプル、それを取り込んだ解説サイト（Astro Starlight）、GitHub Pages への自動公開までを、決まった形でそろえます。

この教材で使っている、次のような仕組みもこのスキルに入っています。

- Excel ／ スプレッドシートをナビゲーションバーで選ぶと、全ページの手順とコードが切り替わる
- 練習用ブックを描いて、マクロを実行したときの様子をその場で再現するデモ
- 節ごとの完成コードの表示と、環境ごとの受け取り方（Excel はブックのダウンロード、スプレッドシートは Google ドライブのサンプルのコピー）
- よくある質問、確認クイズ、注意書きなどの書き方

Office マクロ（VBA ＋ GAS）のほかに、Cloudflare、Docker と PostgreSQL、ブラウザだけで動くアプリ（ゲームなど）の教材にも対応しています。

```sh
git clone https://github.com/seekseep/handson-builder.git ~/.claude/skills/handson-builder
```

入れたあとは、Claude Code に「○○の勉強会教材を作って」「VBA のハンズオンを作って」と頼むと使われます。

## AI に任せたこと、任せなかったこと

本文やコードの下書き、図、デモの仕組みづくりは AI に任せました。一方で、何を題材にするか、どの順番で教えるか、どこまでを関数で済ませてどこから VBA を使うか、といった **決めること** は講師がしています。AI が書いたものは、実際に読んで、動かして、直しています。これは、この教材で伝えたい「AI と一緒に作るが、判断は自分でする」という考え方そのものです。

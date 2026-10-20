---
docs: true
title: はじめての VBA ハンズオン
---

# はじめての VBA ハンズオン

**Excel 作業を自動化しよう。** VBA をまったく触ったことがない人、少し触ったことがある人のためのハンズオンです。

文法を順番に覚える講座ではありません。「毎月の請求書づくりが面倒」という身近な課題から始めて、手作業 → Excel 関数 → VBA と道具を持ち替えながら解決していきます。どこまでが関数で十分で、どこから VBA が効くのか。その見極めも一緒に持ち帰ってください。

教材は **Excel（VBA）** で書いていますが、同じ課題を **Google スプレッドシート（Apps Script）** でも進められます。ページ内の「Excel（VBA）／スプレッドシート（GAS）」のタブを切り替えると、手順とコードがそろって切り替わります。

## 持ち帰ってほしいこと

- VBA の基本的なコードを読んで、何をしているか分かる
- 自分で（あるいは AI が）書いたコードが動かないとき、デバッガーで原因を調べられる
- Excel 関数と VBA を使い分けられる
- ボタンひとつで帳票を作る、という業務の自動化を一度やり切った経験

## 中身

マクロの基本から始めて、Excel 関数でやっていた集計を VBA で再現し、最後はボタンひとつで請求書を作るところまで進みます。章の一覧は [進め方](./sections/README.md) を見てください。

## 始め方

1. [進め方](./sections/README.md) を読む
2. [Excel の準備](./sections/EXCEL.md) または [スプレッドシートの準備](./sections/SHEETS.md) を済ませる
3. [第 1 章](./sections/01-why-automate/01-monthly-closing/LECTURE.md) から進める

## 過去のハンズオン

| ハンズオン | 中身 |
|---|---|
| [リアルタイムお絵かきハンズオン](https://seekseep.github.io/web-rtc-handson/) | PeerJS（WebRTC）でブラウザ同士を直接つなぐ、リアルタイムお絵かき |
| [URL生成ページを作るハンズオン](https://seekseep.github.io/url-generator-handson/) | HTML と JavaScript だけで、検索 URL を組み立てるページを作る |
| [AI エージェント業務活用ハンズオン](https://seekseep.github.io/ai-adoption-handson/) | Codex で座席表とシフト表を作りながら、AI エージェントの使い方と安全な線引きを身につける |
| [パチンコ物理ゲーム ハンズオン](https://seekseep.github.io/slingshot-game-handson/) | Phaser 3 + Matter.js で、アングリーバード風の物理ゲームを作る |
| [スイカゲーム風ハンズオン](https://seekseep.github.io/physics-based-game-handson/) | 物理演算で、スイカゲーム風のゲームを作る |
| [Cloudflare 公開・運用ハンズオン](https://seekseep.github.io/cloudflare-introduction-handson/) | Cloudflare の Pages / Workers / D1 / R2 で Web アプリを公開・運用する |
| [AI Webapp Security ハンズオン 2026](https://seekseep.github.io/ai-webapp-security-handson-2026/) | Web アプリのセキュリティ・パフォーマンスを手を動かして学ぶ |

## このリポジトリについて

- 教材本文は `sections/` に、解説サイトは `site/` にあります（`site/` は `sections/` から生成）
- 練習用ブック（`.xlsx`）は `tools/build_workbooks.py` で生成しています
- ドキュメントサイト: <https://seekseep.github.io/excel-and-vba-handson/>

学習用の教材です。登場する会社・人物・住所はすべて架空のものです。

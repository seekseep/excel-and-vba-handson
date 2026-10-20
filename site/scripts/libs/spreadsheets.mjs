/**
 * サンプルの Google スプレッドシート（sections/spreadsheets.json）の「コピーを作成」URL を引く。
 *
 * spreadsheets.json は「<sec>/<lec>」ごとに Google ドライブのファイル ID を 2 つ持つ:
 *   start  … 節を始める状態（データだけ）。本文の `./spreadsheet` がこれを指す（links.mjs）
 *   done   … 節を終えた状態（Apps Script 入り）。`::codeview` の完成コード欄にボタンを出す（sentinels.mjs）
 *   script … done に付けた Apps Script のスクリプト ID（tools/push_gas.mjs が書く。サイトでは使わない）
 */

import { readFile } from 'node:fs/promises';
import path from 'node:path';

import { ROOT } from './paths.mjs';

let cache;

async function load() {
  if (!cache) {
    try {
      cache = JSON.parse(await readFile(path.join(ROOT, 'sections', 'spreadsheets.json'), 'utf8'));
    } catch (e) {
      if (e.code !== 'ENOENT') throw e;
      cache = {};
    }
  }
  return cache;
}

/** kind（'start' | 'done'）のスプレッドシートの「コピーを作成」URL。無ければ null。 */
export async function spreadsheetCopyUrl(sec, lec, kind) {
  const id = (await load())[`${sec}/${lec}`]?.[kind];
  return id ? `https://docs.google.com/spreadsheets/d/${id}/copy` : null;
}

/**
 * 練習用ブック（sections/<sec>/<lec>/book/、tools/build_workbooks.py が作る）の配布ファイル名と URL。
 *
 *   book/<name>.xlsm … 節を終えた状態（マクロ入り）。Excel の受講者がダウンロードする
 *   book/<name>.xlsx … 節を始める状態（マクロなし）。表計算デモが描く。配るのは章の最初の節のものだけ
 *
 * 配布は ZIP にせず、build-downloads.mjs がブックを 1 つずつ downloads/ にコピーする。
 * どの節のブックも同じ名前（sales.xlsm など）なので、ダウンロードしたときに区別できるよう
 * 章と節の番号を付ける（sales-3-1.xlsm、始める状態は sales-3-1-start.xlsx）。
 * スプレッドシートの受講者は Google ドライブのサンプルをコピーするので、ここは Excel だけの話（spreadsheets.mjs）。
 */

import { readdir } from 'node:fs/promises';
import path from 'node:path';

import { ROOT } from './paths.mjs';

const BOOK_DIR = 'book';

async function listDir(abs) {
  try {
    return await readdir(abs, { withFileTypes: true });
  } catch (e) {
    if (e.code === 'ENOENT') return [];
    throw e;
  }
}

/** 節の book/ にある、拡張子 ext（'.xlsm' | '.xlsx'）のブックのファイル名。無ければ null。 */
async function bookFileOf(sec, lec, ext) {
  const entries = await listDir(path.join(ROOT, 'sections', sec, lec, BOOK_DIR));
  return entries.find((e) => e.isFile() && path.extname(e.name) === ext)?.name ?? null;
}

/** 同じ章の節のディレクトリ名（番号順）。 */
async function lecturesOf(sec) {
  const entries = await listDir(path.join(ROOT, 'sections', sec));
  return entries.filter((e) => e.isDirectory() && /^\d\d-/.test(e.name)).map((e) => e.name).sort();
}

function downloadName(sec, lec, file) {
  const ext = path.extname(file);
  const stem = path.basename(file, ext);
  const suffix = ext === '.xlsx' ? '-start' : '';
  return `${stem}-${Number.parseInt(sec, 10)}-${Number.parseInt(lec, 10)}${suffix}${ext}`;
}

/** 配るブック 1 つ。{ src: リポジトリルートからの相対パス, name: 配布ファイル名 }。 */
function bookEntry(sec, lec, file) {
  return { src: `sections/${sec}/${lec}/${BOOK_DIR}/${file}`, name: downloadName(sec, lec, file) };
}

/** 節を終えた状態のブック（.xlsm）。無ければ null。 */
export async function doneBookOf(sec, lec) {
  const file = await bookFileOf(sec, lec, '.xlsm');
  return file ? bookEntry(sec, lec, file) : null;
}

/**
 * 節を始める状態のブック。同じ章に前の節があれば、その節を終えた状態の .xlsm。
 * 章の最初の節なら、その節の .xlsx。無ければ null。
 */
export async function startBookOf(sec, lec) {
  const lectures = await lecturesOf(sec);
  for (let i = lectures.indexOf(lec) - 1; i >= 0; i--) {
    const done = await doneBookOf(sec, lectures[i]);
    if (done) return done;
  }
  const file = await bookFileOf(sec, lec, '.xlsx');
  return file ? bookEntry(sec, lec, file) : null;
}

/** 配るブックをすべて（重複なし）。build-downloads.mjs が downloads/ にコピーする。 */
export async function allDownloadBooks() {
  const books = new Map();
  for (const sec of (await listDir(path.join(ROOT, 'sections'))).filter((e) => e.isDirectory()).map((e) => e.name).sort()) {
    for (const lec of await lecturesOf(sec)) {
      for (const book of [await startBookOf(sec, lec), await doneBookOf(sec, lec)]) {
        if (book) books.set(book.name, book);
      }
    }
  }
  return [...books.values()];
}

/** 配るブックのサイト上の URL（`<base>/downloads/<配布ファイル名>`）。 */
export function bookUrl(base, book) {
  return book ? `${base}/downloads/${book.name}` : null;
}

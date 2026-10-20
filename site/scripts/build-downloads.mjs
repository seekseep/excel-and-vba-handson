#!/usr/bin/env node
/**
 * 各レクチャーの `example/` を ZIP 化し、site/public/downloads/<sec>-<lec>.zip に出力する。
 *
 * 中身: example/ 配下すべて（作業ツリー直読み。git 追跡は問わない）。
 *   vba/*.bas   … VBE の「ファイル → ファイルのインポート」で読める標準モジュール
 *   gas/*.gs    … Apps Script エディタに貼り付けるコード
 *   data/*.xlsx … 練習用ブック（Excel でもスプレッドシートでも開ける）
 * ZIP 内のルートフォルダは `<lec>/`（展開すると名前付きフォルダができる）。
 *
 * .bas だけは書き出し時に VBE が読める形へ変換する。
 *   - 先頭に `Attribute VB_Name = "<ファイル名>"` を足す（インポート後のモジュール名になる）。
 *     リポジトリの .bas に書かないのは、本文の完成例からコピーして VBE に貼ると
 *     この行が構文エラーになるため
 *   - Shift_JIS（CP932）＋ CRLF にする。リポジトリでは GitHub やエディタで読めるよう
 *     UTF-8 / LF で持っているが、日本語版 Windows の VBE は .bas を CP932 として読むので、
 *     UTF-8 のままだとインポートした日本語が文字化けする
 *
 * `example/assets/` を持つレクチャーは、素材だけの `<sec>-<lec>-assets.zip` も出す
 * （handson-builder の雛形の機能。この教材では今のところ使っていない）。
 *
 * ダウンロード URL の規約 (<sec>-<lec>.zip) は libs/naming.mjs に集約し、
 * sync-lectures.mjs のリンク生成と共有している。
 */

import { createWriteStream } from 'node:fs';
import { glob, mkdir, readFile, rm } from 'node:fs/promises';
import path from 'node:path';

import { ZipArchive } from 'archiver';
import iconv from 'iconv-lite';

import { walkFiles } from './libs/fs-walk.mjs';
import {
  assetsZipBasenameFor,
  exampleDirOf,
  parseLectureRel,
  platformZipBasenameFor,
  startZipBasenameFor,
  zipBasenameFor,
} from './libs/naming.mjs';
import { DOWNLOADS_DIR, ROOT } from './libs/paths.mjs';

const IGNORE_DIRS = new Set(['node_modules', '.git']);
const IGNORE_NAMES = new Set(['.DS_Store']);

/** VBE に読ませる形式へ変換する。それ以外のファイルは null（そのまま入れる）。 */
async function convertForImport(absPath) {
  if (path.extname(absPath).toLowerCase() !== '.bas') return null;
  const moduleName = path.basename(absPath, path.extname(absPath));
  const source = (await readFile(absPath, 'utf8')).replace(/\r\n/g, '\n');
  const text = `Attribute VB_Name = "${moduleName}"\n${source}`.replace(/\n/g, '\r\n');
  return iconv.encode(text, 'cp932');
}

/** files（<absDir> 相対）を outPath に ZIP 化する。ZIP 内は rootDir/ 配下に置く。 */
async function zipFiles(absDir, files, outPath, rootDir) {
  const output = createWriteStream(outPath);
  const archive = new ZipArchive({ zlib: { level: 9 } });
  const done = new Promise((resolve, reject) => {
    output.on('close', () => resolve(files.length));
    archive.on('warning', reject);
    archive.on('error', reject);
  });
  archive.pipe(output);
  for (const rel of files) {
    const abs = path.join(absDir, ...rel.split('/'));
    const name = path.posix.join(rootDir, rel);
    const converted = await convertForImport(abs);
    if (converted) archive.append(converted, { name });
    else archive.file(abs, { name });
  }
  archive.finalize();
  return done;
}

async function findExampleLectures() {
  const dirs = new Set();
  const it = glob('sections/*/*/example/**/*', { cwd: ROOT });
  for await (const rel of it) {
    const parts = rel.split(path.sep);
    dirs.add(parts.slice(0, 3).join('/'));
  }
  return [...dirs].sort();
}

async function zipExample(lectureRel) {
  const { sec, lec } = parseLectureRel(lectureRel);
  const exampleAbs = path.join(ROOT, ...exampleDirOf(lectureRel).split('/'));
  const files = await walkFiles(exampleAbs, { ignoreDirs: IGNORE_DIRS, ignoreNames: IGNORE_NAMES });
  const outPath = path.join(DOWNLOADS_DIR, zipBasenameFor(sec, lec));
  const count = await zipFiles(exampleAbs, files, outPath, lec);
  // Office マクロ系は、環境ごとの ZIP も出す（Excel は gas/ を、スプレッドシートは vba/ を入れない）。
  // 本文のダウンロードボタンは、選んでいる環境の ZIP を指す（remark-download / remark-editor）。
  // Excel は、同じ名前のマクロ入り（.xlsm）があるブックは .xlsx を入れず .xlsm だけにする。
  // スプレッドシートはマクロを使わないので .xlsm を入れない（.xlsx をアップロードして使う）。
  if (files.some((f) => f.startsWith('vba/') || f.startsWith('gas/'))) {
    const hasXlsm = (f) => files.includes(f.replace(/\.xlsx$/, '.xlsm'));
    const pick = {
      excel: files.filter((f) => !f.startsWith('gas/') && !(f.endsWith('.xlsx') && hasXlsm(f))),
      sheets: files.filter((f) => !f.startsWith('vba/') && !f.endsWith('.xlsm')),
    };
    for (const platform of ['excel', 'sheets']) {
      await zipFiles(exampleAbs, pick[platform], path.join(DOWNLOADS_DIR, platformZipBasenameFor(sec, lec, platform)), lec);
    }
  }
  return count;
}

/**
 * 章の最初の節には、その節を始める状態（練習用ブック .xlsx だけ）の ZIP を作る。
 * 2 つ目以降の節は、前の節を終えた状態の ZIP が始める状態になる（sync-lectures.mjs の `./start.zip`）。
 */
async function zipStart(lectureRel, isFirstInSection) {
  if (!isFirstInSection) return null;
  const { sec, lec } = parseLectureRel(lectureRel);
  const exampleAbs = path.join(ROOT, ...exampleDirOf(lectureRel).split('/'));
  const files = (await walkFiles(exampleAbs, { ignoreDirs: IGNORE_DIRS, ignoreNames: IGNORE_NAMES })).filter(
    (f) => f.startsWith('data/') && f.endsWith('.xlsx'),
  );
  if (files.length === 0) return null;
  await zipFiles(exampleAbs, files, path.join(DOWNLOADS_DIR, startZipBasenameFor(sec, lec)), lec);
  return files.length;
}

async function zipExampleAssets(lectureRel) {
  const { sec, lec } = parseLectureRel(lectureRel);
  const assetsAbs = path.join(ROOT, ...exampleDirOf(lectureRel).split('/'), 'assets');
  let files;
  try {
    files = await walkFiles(assetsAbs, { ignoreDirs: IGNORE_DIRS, ignoreNames: IGNORE_NAMES });
  } catch (e) {
    if (e.code === 'ENOENT') return null;
    throw e;
  }
  if (files.length === 0) return null;
  const outPath = path.join(DOWNLOADS_DIR, assetsZipBasenameFor(sec, lec));
  return zipFiles(assetsAbs, files, outPath, 'assets');
}

async function main() {
  await rm(DOWNLOADS_DIR, { recursive: true, force: true });
  await mkdir(DOWNLOADS_DIR, { recursive: true });

  const lectures = await findExampleLectures();
  const seenSections = new Set();
  for (const lectureRel of lectures) {
    const { sec, lec } = parseLectureRel(lectureRel);
    const count = await zipExample(lectureRel);
    const startCount = await zipStart(lectureRel, !seenSections.has(sec));
    seenSections.add(sec);
    if (startCount !== null) {
      console.log(`[build-downloads] ${lectureRel}/example/data -> site/public/downloads/${startZipBasenameFor(sec, lec)} (${startCount} files)`);
    }
    console.log(`[build-downloads] ${lectureRel}/example -> site/public/downloads/${zipBasenameFor(sec, lec)} (${count} files)`);
    const assetCount = await zipExampleAssets(lectureRel);
    if (assetCount !== null) {
      console.log(`[build-downloads] ${lectureRel}/example/assets -> site/public/downloads/${assetsZipBasenameFor(sec, lec)} (${assetCount} files)`);
    }
  }

  if (lectures.length === 0) {
    console.warn('[build-downloads] no lectures found (sections/*/*/example/)');
  }
  console.log('[build-downloads] done');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});

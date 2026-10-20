#!/usr/bin/env node
/**
 * Excel の受講者がダウンロードする練習用ブックを site/public/downloads/ にコピーする。
 *
 * 配るのは 1 節につき、節を終えた状態のマクロ入りブック（book/*.xlsm）1 つと、
 * 章の最初の節だけ、節を始める状態のブック（book/*.xlsx）。ZIP にはしない（中身が 1 ファイルなので）。
 * ファイル名と URL の規約は libs/books.mjs にまとめ、sync-lectures.mjs のリンク生成と共有している。
 *
 * コード（example/vba/*.bas・example/gas/*.gs）は配らない。本文の完成コードの欄に表示し、
 * Excel では .xlsm に、スプレッドシートでは Google ドライブの完成版に入っている。
 *
 * `example/assets/` を持つレクチャーは、素材だけの `<sec>-<lec>-assets.zip` を出す
 * （handson-builder の雛形の機能。この教材では今のところ使っていない）。
 */

import { createWriteStream } from 'node:fs';
import { copyFile, glob, mkdir, rm } from 'node:fs/promises';
import path from 'node:path';

import { ZipArchive } from 'archiver';

import { allDownloadBooks } from './libs/books.mjs';
import { walkFiles } from './libs/fs-walk.mjs';
import { assetsZipBasenameFor, parseLectureRel } from './libs/naming.mjs';
import { DOWNLOADS_DIR, ROOT } from './libs/paths.mjs';

const IGNORE_DIRS = new Set(['node_modules', '.git']);
const IGNORE_NAMES = new Set(['.DS_Store']);

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
    archive.file(path.join(absDir, ...rel.split('/')), { name: path.posix.join(rootDir, rel) });
  }
  archive.finalize();
  return done;
}

async function findAssetsLectures() {
  const dirs = new Set();
  for await (const rel of glob('sections/*/*/example/assets/**/*', { cwd: ROOT })) {
    dirs.add(rel.split(path.sep).slice(0, 3).join('/'));
  }
  return [...dirs].sort();
}

async function zipExampleAssets(lectureRel) {
  const { sec, lec } = parseLectureRel(lectureRel);
  const assetsAbs = path.join(ROOT, ...lectureRel.split('/'), 'example', 'assets');
  const files = await walkFiles(assetsAbs, { ignoreDirs: IGNORE_DIRS, ignoreNames: IGNORE_NAMES });
  if (files.length === 0) return;
  await zipFiles(assetsAbs, files, path.join(DOWNLOADS_DIR, assetsZipBasenameFor(sec, lec)), 'assets');
  console.log(`[build-downloads] ${lectureRel}/example/assets -> site/public/downloads/${assetsZipBasenameFor(sec, lec)} (${files.length} files)`);
}

async function main() {
  await rm(DOWNLOADS_DIR, { recursive: true, force: true });
  await mkdir(DOWNLOADS_DIR, { recursive: true });

  const books = await allDownloadBooks();
  for (const book of books) {
    await copyFile(path.join(ROOT, ...book.src.split('/')), path.join(DOWNLOADS_DIR, book.name));
    console.log(`[build-downloads] ${book.src} -> site/public/downloads/${book.name}`);
  }
  if (books.length === 0) {
    console.warn('[build-downloads] no workbooks found (sections/*/*/book/)');
  }

  for (const lectureRel of await findAssetsLectures()) {
    await zipExampleAssets(lectureRel);
  }
  console.log('[build-downloads] done');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});

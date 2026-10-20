#!/usr/bin/env node
/**
 * リポジトリルート配下のマークダウンから frontmatter に `docs: true` を持つものだけを
 * 拾い、src/content/docs/ 配下のサイトページを生成する。
 *
 * 採用される source/destination の対応 (規約。パスは ROOT=リポジトリルート相対):
 *   README.md                       → docs/index.md             (/)
 *   sections/README.md              → docs/getting-started.md   (/getting-started/)
 *   sections/<name>.md              → docs/<name>.md            (/<name>/)        例: DOCKER.md → /docker/
 *   sections/<sec>/README.md        → docs/<sec>/index.md       (/<sec>/)
 *   sections/<sec>/<lec>/LECTURE.md → docs/<sec>/<lec>.md       (/<sec>/<lec>/)   サイドバーには 1 ファイルとして並ぶ
 *
 * frontmatter で title / sidebar.label / sidebar.order を上書きできる。
 * sidebar.order は LECTURE.md に限り、ディレクトリ名 (`01-...`) からデフォルト値を導出する。
 * 本文中の相対リンク・画像の書き換えは libs/links.mjs が担う。
 *
 * REPO（GitHub URL）と base は astro.config.mjs から読む (libs/astro-config.mjs)。
 */

import { cp, glob, mkdir, readdir, readFile, writeFile, rm, stat } from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

import { loadAstroConfig } from './libs/astro-config.mjs';
import { parseFrontmatter } from './libs/frontmatter.mjs';
import { PUBLIC_ASSETS, publicTargetFor, transformLinks } from './libs/links.mjs';
import {
  buildFrontmatter,
  extractDescription,
  parseLectureOrder,
  stripLeadingH1,
} from './libs/markdown.mjs';
import { bookUrl, doneBookOf, startBookOf } from './libs/books.mjs';
import { parseLectureRel } from './libs/naming.mjs';
import { DOCS_DIR, PUBLIC_DIR, REPO_SUBDIR, ROOT } from './libs/paths.mjs';
import { expandSentinels } from './libs/sentinels.mjs';
import { spreadsheetCopyUrl } from './libs/spreadsheets.mjs';

const MAX_FILE_BYTES = 200 * 1024;
const PUBLISH_FLAG = 'docs';

/**
 * 本文の手書きセンチネルの差し替え先（links.mjs の transformLinks に渡す）。レクチャー以外では全部 null。
 *   startBook       … `./start-book`        その節を始める状態のブック（Excel。books.mjs）
 *   doneBook        … `./done-book`         その節を終えた状態のマクロ入りブック（Excel）
 *   startSpreadsheet … `./start-spreadsheet` 始める状態のサンプルのスプレッドシートの「コピーを作成」URL
 *   doneSpreadsheet  … `./done-spreadsheet`  終えた状態（Apps Script 入り）の「コピーを作成」URL
 */
async function resolveLinkTargets(sourceDir, base) {
  const parsed = parseLectureRel(sourceDir);
  if (!parsed) return {};
  const { sec, lec } = parsed;
  return {
    startBook: bookUrl(base, await startBookOf(sec, lec)),
    doneBook: bookUrl(base, await doneBookOf(sec, lec)),
    startSpreadsheet: await spreadsheetCopyUrl(sec, lec, 'start'),
    doneSpreadsheet: await spreadsheetCopyUrl(sec, lec, 'done'),
  };
}

async function syncFile({
  sourceFile,
  sourceDir,
  sourceSiteUrl,
  outputFile,
  defaultSidebarOrder,
  editPath,
  config,
}) {
  const stats = await stat(sourceFile);
  if (stats.size > MAX_FILE_BYTES) {
    console.warn(`[sync-lectures] skipping oversize: ${path.relative(ROOT, sourceFile)} (${stats.size} bytes)`);
    return;
  }

  const raw = await readFile(sourceFile, 'utf8');
  const { data: srcFm, body: bodyAfterFm } = parseFrontmatter(raw);
  const title = srcFm.title || '名称不明';
  const description = srcFm.description || extractDescription(bodyAfterFm);
  let body = transformLinks(stripLeadingH1(bodyAfterFm), {
    sourceDir,
    sourceSiteUrl,
    repo: config.repo,
    base: config.base,
    targets: await resolveLinkTargets(sourceDir, config.base),
  });

  // `::codeview` / `::assets` / `::preview` センチネルを、実ファイルから生成したブロックに展開する。
  // レクチャー以外（トップページ・章の概要）でも呼ぶ。そちらで使えるのは行き先を明示した
  // `::preview{src="<sec>/<lec>"}` だけで、他は警告を出して消える。
  const lecture = parseLectureRel(sourceDir);
  body = await expandSentinels(body, {
    lectureAbsDir: path.join(ROOT, ...sourceDir.split('/')),
    sec: lecture?.sec,
    lec: lecture?.lec,
    base: config.base,
    where: sourceDir,
  });

  const sidebar = srcFm.sidebar || {};
  const sidebarOrder = sidebar.order ?? defaultSidebarOrder;
  const sidebarLabel = sidebar.label;

  const fm = buildFrontmatter({
    title,
    description,
    sidebarOrder,
    sidebarLabel,
    editUrl: editPath && config.repo ? `${config.repo}/edit/main/${editPath}` : undefined,
  });

  await mkdir(path.dirname(outputFile), { recursive: true });
  await writeFile(outputFile, fm + body, 'utf8');
  console.log(`[sync-lectures] ${path.relative(ROOT, sourceFile)} -> ${path.relative(ROOT, outputFile)}`);
}

async function syncSiblingDocsDir(sourceDir, outputFile) {
  const sourceDocsDir = path.join(ROOT, ...sourceDir.split('/'), 'docs');
  try {
    const stats = await stat(sourceDocsDir);
    if (!stats.isDirectory()) return;
  } catch (e) {
    if (e.code === 'ENOENT') return;
    throw e;
  }
  const outputDocsDir = path.join(path.dirname(outputFile), 'docs');
  await cp(sourceDocsDir, outputDocsDir, {
    recursive: true,
    filter: (src) => !PUBLIC_ASSETS.has(path.basename(src)),
  });
  console.log(`[sync-lectures] ${path.relative(ROOT, sourceDocsDir)}/ -> ${path.relative(ROOT, outputDocsDir)}/`);
}

async function copyPublicAssets(base) {
  const it = glob('sections/*/*/docs/**', { cwd: ROOT });
  for await (const rel of it) {
    const relPosix = rel.split(path.sep).join('/');
    if (!PUBLIC_ASSETS.has(path.posix.basename(relPosix))) continue;
    const abs = path.join(ROOT, relPosix);
    const stats = await stat(abs);
    if (!stats.isFile()) continue;
    const t = publicTargetFor(relPosix, base);
    if (!t) {
      console.warn(`[sync-lectures] PUBLIC_ASSETS basename matched but path shape unsupported: ${relPosix}`);
      continue;
    }
    const dest = path.join(PUBLIC_DIR, ...t.destRel.split('/'));
    await mkdir(path.dirname(dest), { recursive: true });
    await cp(abs, dest);
    console.log(`[sync-lectures] ${relPosix} -> public/${t.destRel}`);
  }
}

// section 内の画像 (`sections/<sec>/<lec>/images/*`) を public/<sec>/<lec>/images/ に
// コピーする。links.mjs が画像参照を base 相対 URL に書き換えるので、GitHub に push
// しなくてもローカル dev・本番の両方で画像が表示される。genfig の .py などは除く。
const IMAGE_EXT = new Set(['.svg', '.png', '.jpg', '.jpeg', '.gif', '.webp', '.avif']);

async function copyLectureImages() {
  const it = glob('sections/*/*/images/**', { cwd: ROOT });
  for await (const rel of it) {
    const relPosix = rel.split(path.sep).join('/');
    if (!IMAGE_EXT.has(path.posix.extname(relPosix).toLowerCase())) continue;
    const abs = path.join(ROOT, relPosix);
    const stats = await stat(abs);
    if (!stats.isFile()) continue;
    // sections/<sec>/<lec>/images/... → public/<sec>/<lec>/images/...
    const destRel = relPosix.replace(/^sections\//, '');
    const dest = path.join(PUBLIC_DIR, ...destRel.split('/'));
    await mkdir(path.dirname(dest), { recursive: true });
    await cp(abs, dest);
  }
}

async function cleanDocsDir() {
  try {
    const entries = await readdir(DOCS_DIR, { withFileTypes: true });
    for (const entry of entries) {
      await rm(path.join(DOCS_DIR, entry.name), { recursive: true, force: true });
    }
  } catch (e) {
    if (e.code === 'ENOENT') return;
    throw e;
  }
}

/**
 * source path (POSIX, ROOT 相対) から出力先・サイト URL・既定の sidebar order を導出する。
 * 規約に当てはまらない場合は null を返し、呼び出し側で警告して skip する。
 */
function deriveDestination(srcRel) {
  const parts = srcRel.split('/');
  const file = parts[parts.length - 1];

  if (srcRel === 'README.md') {
    return { outputPath: 'index.md', slug: '/' };
  }
  if (srcRel === 'sections/README.md') {
    return { outputPath: 'getting-started.md', slug: '/getting-started/' };
  }
  if (parts[0] === 'sections' && parts.length === 2) {
    const name = file.replace(/\.md$/i, '').toLowerCase();
    return { outputPath: `${name}.md`, slug: `/${name}/` };
  }
  if (parts[0] === 'sections' && file === 'README.md' && parts.length >= 3) {
    const inner = parts.slice(1, -1);
    return {
      outputPath: path.posix.join(...inner, 'index.md'),
      slug: '/' + inner.join('/') + '/',
    };
  }
  if (parts[0] === 'sections' && file === 'LECTURE.md' && parts.length >= 3) {
    const inner = parts.slice(1, -1);
    return {
      outputPath: inner.join('/') + '.md',
      slug: '/' + inner.join('/') + '/',
      defaultSidebarOrder: parseLectureOrder(inner[inner.length - 1]),
    };
  }
  return null;
}

async function findPublishableMarkdown() {
  // ROOT 直下のアプリ・生成物ディレクトリは走査しない。
  // 対象は実質 README.md と sections/** のみ。
  const SKIP_TOP = new Set(['node_modules', '.git', 'site', 'dist', '.astro']);
  const it = glob('**/*.md', {
    cwd: ROOT,
    exclude: (rel) => {
      const top = rel.split(path.sep)[0];
      return SKIP_TOP.has(top);
    },
  });
  const all = [];
  for await (const rel of it) all.push(rel.split(path.sep).join('/'));
  all.sort();

  const result = [];
  for (const rel of all) {
    const raw = await readFile(path.join(ROOT, rel), 'utf8');
    const { data } = parseFrontmatter(raw);
    if (data[PUBLISH_FLAG] !== true) continue;
    result.push(rel);
  }
  return result;
}

/**
 * ROOT 配下の `docs: true` マークダウンを src/content/docs/ に一括同期する。
 * CLI (`node scripts/sync-lectures.mjs`) からも、dev 監視インテグレーション
 * (src/integrations/lecture-sync.mjs) からも呼べるよう export する。
 */
export async function syncLectures() {
  const config = await loadAstroConfig();
  await cleanDocsDir();
  await mkdir(DOCS_DIR, { recursive: true });
  await copyPublicAssets(config.base);
  await copyLectureImages();

  for (const srcRel of await findPublishableMarkdown()) {
    const dest = deriveDestination(srcRel);
    if (!dest) {
      console.warn(`[sync-lectures] no destination convention for ${srcRel} (skipped)`);
      continue;
    }
    const sourceDir = path.posix.dirname(srcRel);
    const outputFile = path.join(DOCS_DIR, dest.outputPath);
    await syncFile({
      sourceFile: path.join(ROOT, srcRel),
      sourceDir,
      sourceSiteUrl: dest.slug,
      outputFile,
      defaultSidebarOrder: dest.defaultSidebarOrder,
      // editUrl は実ファイルのリポジトリ相対パスを指すので REPO_SUBDIR を前置する。
      editPath: path.posix.join(REPO_SUBDIR, srcRel),
      config,
    });
    await syncSiblingDocsDir(sourceDir, outputFile);
  }

  console.log('[sync-lectures] done');
}

// CLI として直接実行されたときだけ走らせる（import 時に副作用を出さない）。
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  syncLectures().catch((err) => {
    console.error(err);
    process.exit(1);
  });
}

#!/usr/bin/env node
/**
 * Converts a .pptx into a positioned-HTML deck with a localization string table.
 *
 *   node tools/pptx-to-html.mjs "C:\path\to\talk.pptx" --slug my-talk
 *
 * Writes decks/<slug>/{slides.html, strings.json, deck.json, images/}. Run
 * `node tools/build-manifest.mjs` afterwards to list it on the landing page.
 * Requires Node 18+. No dependencies.
 */
import { readFileSync } from 'node:fs';
import { mkdir, writeFile, rm, access } from 'node:fs/promises';
import { inflateRawSync } from 'node:zlib';
import { basename, dirname, join, resolve, extname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { listZipEntries, slugify, SLUG_PATTERN } from './pptx-core.mjs';
import { pptxToHtml } from './pptx-html.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');

function inflateAll(bytes) {
  const files = new Map();
  for (const entry of listZipEntries(bytes)) {
    const raw = bytes.subarray(entry.dataStart, entry.dataStart + entry.compSize);
    if (entry.method === 0) files.set(entry.name, raw);
    else if (entry.method === 8) files.set(entry.name, inflateRawSync(raw));
    else throw new Error(`Unsupported compression (method ${entry.method}) for ${entry.name}`);
  }
  return files;
}

const exists = (p) => access(p).then(() => true, () => false);

/** "1-5" -> [1, 5]; "3" -> [3, 3]. */
function parseRange(text) {
  const m = /^(\d+)(?:-(\d+))?$/.exec(text);
  if (!m) throw new Error(`--slides expects a number or a range like 1-5, got "${text}"`);
  return [Number(m[1]), Number(m[2] ?? m[1])];
}

function parseArgs(argv) {
  const opts = { _: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--force') opts.force = true;
    else if (a === '--no-images') opts.noImages = true;
    else if (a.startsWith('--')) opts[a.slice(2)] = argv[++i];
    else opts._.push(a);
  }
  return opts;
}

async function main() {
  const opts = parseArgs(process.argv.slice(2));
  const file = opts._[0];
  if (!file) {
    console.error('Usage: node tools/pptx-to-html.mjs <file.pptx> [--slug name] [--out dir] [--title t] [--slides 1-5] [--force]');
    process.exit(1);
  }

  const slug = opts.slug ?? slugify(basename(file, extname(file)));
  if (!SLUG_PATTERN.test(slug)) throw new Error(`"${slug}" is not a valid slug.`);
  const outDir = resolve(opts.out ?? join(ROOT, 'decks', slug));
  if (await exists(outDir) && !opts.force) throw new Error(`${outDir} already exists. Pass --force to overwrite it.`);

  const started = Date.now();
  const files = inflateAll(new Uint8Array(readFileSync(file)));
  const result = pptxToHtml(files, {
    slug, title: opts.title, author: opts.author, date: opts.date, theme: opts.theme,
    description: opts.description ?? `Imported from ${basename(file)}.`,
    slideRange: opts.slides ? parseRange(opts.slides) : undefined,
  });

  await rm(outDir, { recursive: true, force: true });
  await mkdir(join(outDir, 'images'), { recursive: true });
  await writeFile(join(outDir, 'slides.html'), result.html, 'utf8');
  await writeFile(join(outDir, 'strings.json'), `${JSON.stringify(result.strings, null, 2)}\n`, 'utf8');
  await writeFile(join(outDir, 'deck.json'), `${JSON.stringify(result.deck, null, 2)}\n`, 'utf8');
  if (!opts.noImages) {
    for (const { zipPath, name } of result.images) await writeFile(join(outDir, 'images', name), files.get(zipPath));
  }

  const { stats } = result;
  console.log(`Wrote ${outDir} in ${((Date.now() - started) / 1000).toFixed(1)}s`);
  console.log(`  ${stats.slides} slides, ${stats.strings} strings, ${stats.shapes} shapes, ${stats.pictures} pictures, ${stats.connectors} connectors, ${stats.tables} tables, ${stats.codeParagraphs} code paragraphs`);
  if (result.warnings.length) {
    console.log(`  ${result.warnings.length} warning(s):`);
    for (const w of result.warnings.slice(0, 40)) console.log(`    - ${w}`);
    if (result.warnings.length > 40) console.log(`    ... and ${result.warnings.length - 40} more`);
  }
}

main().catch((error) => { console.error(error.message); process.exit(1); });

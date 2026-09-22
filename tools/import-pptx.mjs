/**
 * Turn a .pptx into a deck folder: slides.md, deck.json and images/.
 *
 * A .pptx is a ZIP full of XML, so this needs no dependencies and no
 * PowerPoint. The conversion itself lives in pptx-core.mjs, which the import
 * page in the browser uses too; this file only deals with the file system.
 *
 *   node tools/import-pptx.mjs ../talk.pptx --slug my-talk
 */
import { readFileSync } from 'node:fs';
import { mkdir, writeFile, rm, access } from 'node:fs/promises';
import { inflateRawSync } from 'node:zlib';
import { basename, dirname, join, resolve, extname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { listZipEntries, pptxToDeck, slugify, SLUG_PATTERN } from './pptx-core.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');

/** Inflate every entry up front so the shared core can stay synchronous. */
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

function parseArgs(argv) {
  const opts = { images: true };
  const rest = [];
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--force') opts.force = true;
    else if (a === '--dry-run') opts.dryRun = true;
    else if (a === '--no-images') opts.images = false;
    else if (a.startsWith('--')) {
      const key = a.slice(2);
      const value = argv[++i];
      if (value === undefined) throw new Error(`Missing value for --${key}`);
      opts[key] = value;
    } else rest.push(a);
  }
  return { opts, rest };
}

async function exists(p) {
  try { await access(p); return true; } catch { return false; }
}

async function main() {
  const { opts, rest } = parseArgs(process.argv.slice(2));
  const source = rest[0];
  if (!source) {
    console.error('Usage: node tools/import-pptx.mjs <deck.pptx> [--slug name] [--title "..."]');
    console.error('       [--author "..."] [--date YYYY-MM-DD] [--theme night] [--force] [--no-images] [--dry-run]');
    process.exit(2);
  }

  const file = resolve(source);
  if (!await exists(file)) throw new Error(`No such file: ${file}`);

  const slug = opts.slug ?? slugify(basename(file, extname(file)));
  if (!SLUG_PATTERN.test(slug)) {
    throw new Error(`Invalid slug "${slug}". Use letters, digits, dot, dash or underscore.`);
  }

  const outDir = opts.out ? resolve(opts.out) : join(ROOT, 'decks', slug);
  if (await exists(outDir) && !opts.force && !opts.dryRun) {
    throw new Error(`${outDir} already exists. Pass --force to overwrite it.`);
  }

  const files = inflateAll(new Uint8Array(readFileSync(file)));
  const result = pptxToDeck(files, {
    ...opts,
    slug,
    description: opts.description ?? `Imported from ${basename(file)}.`,
  });

  console.log(`Read ${result.stats.slidesRead} slide(s) from ${basename(file)}`);
  console.log(`  slides written:  ${result.stats.slidesWritten}`);
  console.log(`  speaker notes:   ${result.stats.notes}`);
  console.log(`  images:          ${result.stats.images}`);

  if (opts.dryRun) {
    console.log('\n--- dry run, nothing written ---\n');
    console.log(result.markdown.split('\n').slice(0, 40).join('\n'));
  } else {
    if (opts.force) await rm(outDir, { recursive: true, force: true });
    await mkdir(outDir, { recursive: true });
    await writeFile(join(outDir, 'slides.md'), result.markdown, 'utf8');
    await writeFile(join(outDir, 'deck.json'), `${JSON.stringify(result.deck, null, 2)}\n`, 'utf8');
    if (result.images.length) {
      await mkdir(join(outDir, 'images'), { recursive: true });
      for (const image of result.images) {
        await writeFile(join(outDir, 'images', image.name), files.get(image.zipPath));
      }
    }
    console.log(`\nWrote ${outDir}`);
  }

  if (result.warnings.length) {
    console.log(`\nNeeds a look (${result.warnings.length}):`);
    for (const w of result.warnings.slice(0, 25)) console.log(`  - ${w}`);
    if (result.warnings.length > 25) console.log(`  ... and ${result.warnings.length - 25} more`);
  }
  if (!opts.dryRun) console.log('\nNext: node tools/build-manifest.mjs, then open the deck and edit it down.');
}

main().catch((err) => {
  console.error(`import-pptx: ${err.message}`);
  process.exit(1);
});

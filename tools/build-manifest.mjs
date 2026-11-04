#!/usr/bin/env node
/**
 * Generates decks.json by scanning decks/<slug>/ for slides.md.
 *
 * The landing page reads that manifest, so this is what makes the index
 * "auto-generated": add a folder, re-run this, and the deck shows up.
 *
 *   node tools/build-manifest.mjs           # write decks.json
 *   node tools/build-manifest.mjs --check   # fail if decks.json is stale (CI)
 *
 * Requires Node 18+. No dependencies.
 */

import { readdir, readFile, writeFile, stat } from 'node:fs/promises';
import { join, dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const DECKS_DIR = join(ROOT, 'decks');
const MANIFEST_PATH = join(ROOT, 'decks.json');
const SLUG_PATTERN = /^[a-z0-9][a-z0-9._-]*$/i;

async function exists(path) {
  try {
    await stat(path);
    return true;
  } catch {
    return false;
  }
}

/** Falls back to the first Markdown heading so a deck without deck.json still gets a real name. */
function titleFromSlides(markdown, slug) {
  const heading = markdown.match(/^#{1,3}\s+(.+?)\s*$/m);
  return heading ? heading[1].trim() : slug;
}

async function readMeta(deckDir, slug) {
  const metaPath = join(deckDir, 'deck.json');
  if (!(await exists(metaPath))) return {};
  try {
    return JSON.parse(await readFile(metaPath, 'utf8'));
  } catch (error) {
    throw new Error(`decks/${slug}/deck.json is not valid JSON: ${error.message}`);
  }
}

async function buildManifest() {
  if (!(await exists(DECKS_DIR))) {
    throw new Error(`No decks directory at ${DECKS_DIR}`);
  }

  const entries = await readdir(DECKS_DIR, { withFileTypes: true });
  const decks = [];
  const warnings = [];

  for (const entry of entries) {
    if (!entry.isDirectory()) continue;

    const slug = entry.name;
    if (!SLUG_PATTERN.test(slug)) {
      warnings.push(`skipped "${slug}": folder name is not a valid deck slug`);
      continue;
    }

    const deckDir = join(DECKS_DIR, slug);
    const hasHtml = await exists(join(deckDir, 'slides.html'));
    const slidesPath = join(deckDir, 'slides.md');
    if (!hasHtml && !(await exists(slidesPath))) {
      warnings.push(`skipped "${slug}": no slides.md or slides.html`);
      continue;
    }

    const meta = await readMeta(deckDir, slug);
    const title = meta.title || (hasHtml ? slug : titleFromSlides(await readFile(slidesPath, 'utf8'), slug));

    decks.push({
      slug,
      title,
      description: meta.description || '',
      author: meta.author || '',
      date: meta.date || '',
      tags: Array.isArray(meta.tags) ? meta.tags : [],
      theme: meta.theme || 'night',
      ...(hasHtml ? { format: 'html' } : {}),
    });
  }

  // Newest first, then alphabetical so undated decks stay in a stable order.
  decks.sort((a, b) => (b.date || '').localeCompare(a.date || '') || a.title.localeCompare(b.title));

  return {
    manifest: { generatedBy: 'tools/build-manifest.mjs', count: decks.length, decks },
    warnings,
  };
}

const { manifest, warnings } = await buildManifest();
const serialized = `${JSON.stringify(manifest, null, 2)}\n`;
const checkOnly = process.argv.includes('--check');

for (const warning of warnings) console.warn(`warning: ${warning}`);

if (checkOnly) {
  const current = (await exists(MANIFEST_PATH)) ? await readFile(MANIFEST_PATH, 'utf8') : '';
  // Compare normalized: git may check decks.json out with CRLF on Windows,
  // which would otherwise fail this check for a purely cosmetic difference.
  const normalize = (text) => text.replace(/\r\n/g, '\n');
  if (normalize(current) !== normalize(serialized)) {
    console.error('decks.json is out of date. Run: node tools/build-manifest.mjs');
    process.exit(1);
  }
  console.log(`decks.json is up to date (${manifest.count} deck(s)).`);
} else {
  await writeFile(MANIFEST_PATH, serialized, 'utf8');
  console.log(`Wrote decks.json with ${manifest.count} deck(s): ${manifest.decks.map((d) => d.slug).join(', ') || 'none'}`);
}

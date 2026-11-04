#!/usr/bin/env node
/**
 * Builds an offline copy of the deck site into dist/.
 *
 *   node tools/build-site.mjs
 *
 * Produces:
 *   dist/                       the site with every dependency vendored, so it
 *                               works on a network with no CDN access. Serve it
 *                               over HTTP from any static host or file share.
 *   dist/standalone/<slug>.html one self-contained file per deck. No server and
 *                               no internet: double-click it, or email it.
 *
 * dist/ is a plain folder, so sharing decks does not depend on GitHub Pages.
 *
 * Requires Node 18+. No dependencies.
 */

import { readFile, writeFile, mkdir, rm, readdir, stat } from 'node:fs/promises';
import { join, dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { assembleStandaloneHtml } from './standalone-html.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const DIST = join(ROOT, 'dist');
const CACHE = join(ROOT, '.asset-cache');
const CDN = 'https://cdn.jsdelivr.net/npm';

const REVEAL_VERSION = '5.2.1';
const MERMAID_VERSION = '11.17.1';
const HLJS_VERSION = '11.12.0';

const asset = {
  revealCss: `reveal.js@${REVEAL_VERSION}/dist/reveal.css`,
  revealJs: `reveal.js@${REVEAL_VERSION}/dist/reveal.js`,
  markdownJs: `reveal.js@${REVEAL_VERSION}/plugin/markdown/markdown.js`,
  highlightJs: `reveal.js@${REVEAL_VERSION}/plugin/highlight/highlight.js`,
  notesJs: `reveal.js@${REVEAL_VERSION}/plugin/notes/notes.js`,
  hljsCss: `highlight.js@${HLJS_VERSION}/styles/atom-one-dark.css`,
  mermaidJs: `mermaid@${MERMAID_VERSION}/dist/mermaid.min.js`,
  theme: (name) => `reveal.js@${REVEAL_VERSION}/dist/theme/${name}.css`,
};

async function exists(path) {
  try {
    await stat(path);
    return true;
  } catch {
    return false;
  }
}

/** Downloads once, then reuses .asset-cache so repeat builds are offline too. */
async function fetchAsset(path) {
  const cached = join(CACHE, path.replace(/[/@]/g, '_'));
  if (await exists(cached)) return readFile(cached, 'utf8');

  const url = `${CDN}/${path}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Failed to download ${url}: HTTP ${res.status}`);
  const text = await res.text();

  await mkdir(CACHE, { recursive: true });
  await writeFile(cached, text, 'utf8');
  return text;
}

/**
 * Reveal's themes pull Montserrat and Open Sans from Google Fonts. assets/theme.css
 * overrides both font stacks anyway, so dropping the imports removes the last
 * external request without changing how anything looks.
 */
const stripRemoteFonts = (css) =>
  css.replace(/@import\s+url\(https?:\/\/fonts\.googleapis\.com[^)]*\);?/g, '');

async function writeVendored(path) {
  const contents = path.endsWith('.css') ? stripRemoteFonts(await fetchAsset(path)) : await fetchAsset(path);
  const target = join(DIST, 'vendor', path);
  await mkdir(dirname(target), { recursive: true });
  await writeFile(target, contents, 'utf8');
  return contents;
}

async function copyInto(relPath) {
  const target = join(DIST, relPath);
  await mkdir(dirname(target), { recursive: true });
  await writeFile(target, await readFile(join(ROOT, relPath), 'utf8'), 'utf8');
}

/** Images and other deck assets are binary, so they must not go through utf8. */
async function copyBinaryInto(relPath) {
  const target = join(DIST, relPath);
  await mkdir(dirname(target), { recursive: true });
  await writeFile(target, await readFile(join(ROOT, relPath)));
}

const MIME_BY_EXT = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.svg': 'image/svg+xml',
  '.webp': 'image/webp',
  '.avif': 'image/avif',
  '.bmp': 'image/bmp',
};

/** Every file in a deck folder except the two we handle by name. */
async function deckAssetFiles(deckDir, sub = '') {
  const here = join(ROOT, deckDir, sub);
  const found = [];
  for (const entry of await readdir(here, { withFileTypes: true })) {
    const rel = sub ? `${sub}/${entry.name}` : entry.name;
    if (entry.isDirectory()) found.push(...await deckAssetFiles(deckDir, rel));
    else if (!['slides.md', 'deck.json'].includes(rel)) found.push(rel);
  }
  return found;
}

/**
 * Rewrite image references to data URIs so a standalone file really is
 * standalone. Covers Markdown `![](images/x.png)` and raw `src="images/x.png"`.
 */
async function inlineDeckImages(markdown, deckDir, onMissing) {
  const cache = new Map();
  const toDataUri = async (ref) => {
    const clean = ref.split(/[?#]/)[0];
    if (/^(https?:|data:|\/\/)/i.test(clean)) return null;
    if (cache.has(clean)) return cache.get(clean);

    const ext = clean.slice(clean.lastIndexOf('.')).toLowerCase();
    const mime = MIME_BY_EXT[ext];
    let uri = null;
    if (mime) {
      try {
        const bytes = await readFile(join(ROOT, deckDir, clean));
        uri = `data:${mime};base64,${bytes.toString('base64')}`;
      } catch {
        onMissing?.(clean);
      }
    }
    cache.set(clean, uri);
    return uri;
  };

  const replacements = [];
  const patterns = [
    /!\[([^\]]*)\]\(([^)\s]+)(\s+"[^"]*")?\)/g,
    /\bsrc="([^"]+)"/g,
  ];
  for (const [i, pattern] of patterns.entries()) {
    for (const m of markdown.matchAll(pattern)) {
      const ref = i === 0 ? m[2] : m[1];
      replacements.push({ match: m[0], ref, index: m.index });
    }
  }

  let out = '';
  let cursor = 0;
  for (const r of replacements.sort((a, b) => a.index - b.index)) {
    const uri = await toDataUri(r.ref);
    out += markdown.slice(cursor, r.index);
    out += uri ? r.match.replace(r.ref, uri) : r.match;
    cursor = r.index + r.match.length;
  }
  return out + markdown.slice(cursor);
}

/** Markdown goes inside a <textarea>, so a literal closing tag would end it early. */
const escapeForTextarea = (markdown) => markdown.replace(/<\/textarea>/gi, '&lt;/textarea&gt;');

const escapeHtml = (value) =>
  String(value ?? '').replace(/[&<>"']/g, (ch) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  })[ch]);

async function buildStandalone(deck, slides, sharedTheme) {
  const usesMermaid = slides.includes('class="mermaid"');

  const [revealCss, themeCss, hljsCss, revealJs, markdownJs, highlightJs] = await Promise.all([
    fetchAsset(asset.revealCss),
    fetchAsset(asset.theme(deck.theme || 'night')),
    fetchAsset(asset.hljsCss),
    fetchAsset(asset.revealJs),
    fetchAsset(asset.markdownJs),
    fetchAsset(asset.highlightJs),
  ]);

  const mermaidJs = usesMermaid ? await fetchAsset(asset.mermaidJs) : '';

  // The Notes plugin is deliberately omitted: its speaker view opens a popup that
  // loads plugin/notes/notes.html, which does not exist in a single file. Speaker
  // notes still work on the served site.
  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>${escapeHtml(deck.title || deck.slug)}</title>
    <!--
      Self-contained deck: Reveal.js, highlighting${usesMermaid ? ', Mermaid' : ''} and the slides
      are all inlined below. No server and no internet needed.
      Generated by tools/build-site.mjs — edit decks/${deck.slug}/slides.md instead.
    -->
    <style>${stripRemoteFonts(revealCss)}</style>
    <style>${stripRemoteFonts(themeCss)}</style>
    <style>${hljsCss}</style>
    <style>${sharedTheme}</style>
  </head>
  <body>
    <div class="reveal">
      <div class="slides">
        <section data-markdown
                 data-separator="^\\r?\\n---\\r?\\n"
                 data-separator-vertical="^\\r?\\n--\\r?\\n"
                 data-separator-notes="^Note:">
          <textarea data-template>
${escapeForTextarea(slides)}
          </textarea>
        </section>
      </div>
    </div>

    <script>${revealJs}<\/script>
    <script>${markdownJs}<\/script>
    <script>${highlightJs}<\/script>
${usesMermaid ? `    <script>${mermaidJs}<\/script>` : ''}
    <script>
      (async () => {
        await Reveal.initialize({
          plugins: [RevealMarkdown, RevealHighlight],
          hash: true,
          slideNumber: 'c/t',
          transition: 'slide',
          width: 1280,
          height: 720,
          margin: 0.06,
          minScale: 0.2,
          maxScale: 1.0,
        });
${usesMermaid ? `        mermaid.initialize({ startOnLoad: false, theme: 'dark' });
        const diagrams = document.querySelectorAll('.mermaid');
        if (diagrams.length) {
          await mermaid.run({ nodes: diagrams });
          Reveal.layout();
        }` : ''}
        function fitCurrentSlide() {
          const section = Reveal.getCurrentSlide();
          if (!section) return;
          const limit = Reveal.getConfig().height;
          const shrinkTo = (floor) => {
            let percent = 100;
            section.style.fontSize = '';
            while (section.scrollHeight > limit && percent > floor) {
              percent -= 5;
              section.style.fontSize = percent + '%';
            }
          };
          section.classList.remove('is-two-column');
          shrinkTo(60);
          if (section.scrollHeight > limit) {
            section.classList.add('is-two-column');
            shrinkTo(30);
            if (section.scrollHeight > limit) {
              section.classList.remove('is-two-column');
              shrinkTo(30);
            }
          }
        }
        Reveal.on('slidechanged', fitCurrentSlide);
        Reveal.on('resize', fitCurrentSlide);
        fitCurrentSlide();
        window.DECK_READY = true;
      })();
    <\/script>
  </body>
</html>
`;
}

/**
 * Standalone build for decks imported from PowerPoint (format "html"): the
 * slides are already HTML, so they go in as-is with pictures inlined. Any
 * strings.<lang>.json files ride along, and ?lang=<code> picks one.
 */
async function buildStandaloneHtml(deck, meta, slidesHtml, tables, images, sharedTheme) {
  const [reveal, theme, revealJs, pptx, i18n] = await Promise.all([
    fetchAsset(asset.revealCss),
    fetchAsset(asset.theme(deck.theme || 'white')),
    fetchAsset(asset.revealJs),
    readFile(join(ROOT, 'assets/pptx-slide.css'), 'utf8'),
    readFile(join(ROOT, 'assets/deck-i18n.js'), 'utf8'),
  ]);
  return assembleStandaloneHtml({
    title: deck.title || deck.slug,
    source: `decks/${deck.slug}/slides.html by tools/build-site.mjs`,
    lang: meta.lang,
    canvas: meta.canvas,
    slidesHtml,
    images,
    tables,
    css: { reveal, theme, shared: sharedTheme, pptx },
    js: { reveal: revealJs, i18n },
  });
}

async function main() {
  const manifest = JSON.parse(await readFile(join(ROOT, 'decks.json'), 'utf8'));
  const decks = manifest.decks ?? [];
  if (!decks.length) throw new Error('decks.json lists no decks. Run: node tools/build-manifest.mjs');

  await rm(DIST, { recursive: true, force: true }).catch((error) => {
    // Common on Windows when a static server is still running inside dist/.
    if (error.code === 'EBUSY' || error.code === 'EPERM') {
      throw new Error(
        `Cannot clear ${DIST} because something is using it ` +
          `(${error.code}). Stop any server running out of dist/ and try again.`,
      );
    }
    throw error;
  });
  await mkdir(DIST, { recursive: true });

  // 1. Vendor every dependency the served site needs.
  const themes = [...new Set(['night', ...decks.map((d) => d.theme || 'night')])];
  const vendorPaths = [
    asset.revealCss,
    asset.revealJs,
    asset.markdownJs,
    asset.highlightJs,
    asset.notesJs,
    asset.hljsCss,
    asset.mermaidJs,
    ...themes.map(asset.theme),
  ];
  for (const path of vendorPaths) await writeVendored(path);

  // The Notes plugin popup loads this at runtime, so it has to ship too.
  await writeVendored(`reveal.js@${REVEAL_VERSION}/plugin/notes/speaker-view.html`);

  // 2. Copy the site and repoint it at the vendored copies.
  await copyInto('index.html');
  await copyInto('decks.json');
  await copyInto('assets/theme.css');
  await copyInto('assets/pptx-slide.css');
  await copyInto('assets/deck-i18n.js');
  const deckHtml = await readFile(join(ROOT, 'deck.html'), 'utf8');
  const rewritten = deckHtml.replace(
    /window\.DECK_ASSET_BASE = '[^']*';/,
    "window.DECK_ASSET_BASE = './vendor';",
  );
  if (rewritten === deckHtml) throw new Error('Could not rewrite DECK_ASSET_BASE in deck.html');
  await writeFile(join(DIST, 'deck.html'), rewritten, 'utf8');

  // The import page runs the same converter as the CLI, in the browser.
  const importHtml = await readFile(join(ROOT, 'import.html'), 'utf8');
  const importRewritten = importHtml.replace(
    /window\.DECK_ASSET_BASE = '[^']*';/,
    "window.DECK_ASSET_BASE = './vendor';",
  );
  if (importRewritten === importHtml) throw new Error('Could not rewrite DECK_ASSET_BASE in import.html');
  await writeFile(join(DIST, 'import.html'), importRewritten, 'utf8');
  await copyInto('tools/pptx-core.mjs');
  for (const file of ['xml-lite', 'pptx-style', 'pptx-geom', 'pptx-html', 'zip-tools', 'standalone-html']) {
    await copyInto(`tools/${file}.mjs`);
  }
  const convertHtml = await readFile(join(ROOT, 'convert.html'), 'utf8');
  const convertRewritten = convertHtml.replace(
    /window\.DECK_ASSET_BASE = '[^']*';/,
    "window.DECK_ASSET_BASE = './vendor';",
  );
  if (convertRewritten === convertHtml) throw new Error('Could not rewrite DECK_ASSET_BASE in convert.html');
  await writeFile(join(DIST, 'convert.html'), convertRewritten, 'utf8');

  // 3. Copy each deck's sources, and build its standalone single file.
  const sharedTheme = await readFile(join(ROOT, 'assets/theme.css'), 'utf8');
  await mkdir(join(DIST, 'standalone'), { recursive: true });

  for (const deck of decks) {
    const deckDir = join('decks', deck.slug);
    const posixDir = deckDir.replace(/\\/g, '/');

    if (deck.format === 'html') {
      const slidesHtml = await readFile(join(ROOT, deckDir, 'slides.html'), 'utf8');
      const meta = JSON.parse(await readFile(join(ROOT, deckDir, 'deck.json'), 'utf8'));

      // Everything in the folder ships: slides.html, string tables, images.
      const assets = await deckAssetFiles(deckDir);
      for (const file of assets) await copyBinaryInto(`${posixDir}/${file}`);
      await copyInto(`${posixDir}/deck.json`);

      const tables = {};
      for (const file of assets) {
        const m = /^strings\.([A-Za-z0-9-]+)\.json$/.exec(file);
        if (m) tables[m[1]] = JSON.parse(await readFile(join(ROOT, deckDir, file), 'utf8'));
      }

      const images = {};
      for (const file of assets) {
        const m = /^images\/([^/]+)$/.exec(file);
        const mime = m && MIME_BY_EXT[m[1].slice(m[1].lastIndexOf('.')).toLowerCase()];
        if (mime) images[m[1]] = `data:${mime};base64,${(await readFile(join(ROOT, deckDir, file))).toString('base64')}`;
      }

      const html = await buildStandaloneHtml(deck, meta, slidesHtml, tables, images, sharedTheme);
      await writeFile(join(DIST, 'standalone', `${deck.slug}.html`), html, 'utf8');
      console.log(`standalone/${deck.slug}.html  ${(Buffer.byteLength(html) / 1024 / 1024).toFixed(1)} MB  (html deck, ${Object.keys(tables).length} translation(s))`);
      continue;
    }

    const slides = await readFile(join(ROOT, deckDir, 'slides.md'), 'utf8');

    await copyInto(join(deckDir, 'slides.md').replace(/\\/g, '/'));
    if (await exists(join(ROOT, deckDir, 'deck.json'))) {
      await copyInto(join(deckDir, 'deck.json').replace(/\\/g, '/'));
    }

    // Images and anything else the deck ships alongside its slides.
    const assets = await deckAssetFiles(deckDir);
    for (const asset of assets) {
      await copyBinaryInto(`${deckDir.replace(/\\/g, '/')}/${asset}`);
    }

    const missing = [];
    const inlined = await inlineDeckImages(slides, deckDir, (ref) => missing.push(ref));
    for (const ref of missing) {
      console.warn(`  warning: ${deck.slug} references ${ref}, which is missing`);
    }

    const html = await buildStandalone(deck, inlined, sharedTheme);
    await writeFile(join(DIST, 'standalone', `${deck.slug}.html`), html, 'utf8');
    const assetNote = assets.length ? `  (+${assets.length} asset${assets.length === 1 ? '' : 's'})` : '';
    console.log(`standalone/${deck.slug}.html  ${(Buffer.byteLength(html) / 1024 / 1024).toFixed(1)} MB${assetNote}`);
  }

  await writeFile(
    join(DIST, 'README.txt'),
    [
      'Decks as Code — built output',
      '',
      'standalone/<name>.html',
      '  A whole deck in one file. Double-click it. No server, no internet.',
      '  The easiest thing to email or drop in a chat.',
      '',
      'index.html and deck.html',
      '  The full site with a landing page listing every deck. These two read',
      '  slides.md at runtime, which browsers block on a file:// URL, so they',
      '  must be served over HTTP rather than opened by double-click:',
      '',
      '    npx --yes http-server . -p 8931',
      '    python -m http.server 8931',
      '',
      '  Everything is vendored, so it needs no internet access.',
      '',
      'import.html',
      '  Convert a PowerPoint file into a Markdown deck in the browser. Also needs to be',
      '  served rather than double-clicked. The file never leaves your machine.',
      '',
      'convert.html',
      '  Convert a PowerPoint file into localizable HTML (positioned slides plus a',
      '  string table), preview it, and download a single HTML file. Also needs a server.',
      '',
      'Arrow keys navigate, S opens speaker notes (served site only),',
      'Esc shows the slide overview, ? lists every shortcut.',
      '',
    ].join('\n'),
    'utf8',
  );

  const files = await readdir(DIST);
  console.log(`\nBuilt dist/ (${files.length} entries) for ${decks.length} deck(s).`);
}

await main();

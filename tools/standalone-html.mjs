/**
 * Assembles one self-contained HTML file for a deck converted from PowerPoint.
 *
 * Pure string work with no Node APIs, so build-site.mjs and the browser
 * converter (convert.html) produce the same file. Callers supply the Reveal,
 * theme and slide assets already read into strings.
 */

const escapeHtml = (value) =>
  String(value ?? '').replace(/[&<>"']/g, (ch) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  })[ch]);

/** Reveal's themes import Google Fonts; assets/theme.css replaces the fonts anyway. */
export const stripRemoteFonts = (css) =>
  css.replace(/@import\s+url\(https?:\/\/fonts\.googleapis\.com[^)]*\);?/g, '');

/**
 * @param {object} o
 * @param {string} o.title
 * @param {string} [o.source]       where it was generated from, for the header comment
 * @param {string} [o.lang]         source language tag
 * @param {{width:number,height:number}} [o.canvas]
 * @param {string} o.slidesHtml     slides.html as written by the converter (pictures as images/<name>)
 * @param {Record<string,string>} [o.images]  picture name -> data URI; embedded once each
 * @param {Record<string, object>} [o.tables]  translation tables keyed by language
 * @param {{reveal:string, theme:string, shared:string, pptx:string}} o.css
 * @param {{reveal:string, i18n:string}} o.js
 */
export function assembleStandaloneHtml(o) {
  const canvas = o.canvas ?? { width: 1280, height: 720 };
  // `<` is escaped so no string in a table can close the script element early.
  const json = JSON.stringify(o.tables ?? {}).replace(/</g, '\\u003c');

  // Each picture is embedded once, however many slides show it. Layout pictures
  // repeat on every slide, and inlining them per use would multiply the file size.
  const available = o.images ?? {};
  const used = new Set();
  const cssVars = new Map();
  let slidesHtml = o.slidesHtml
    .replace(/\ssrc="images\/([^"]+)"/g, (m, name) => {
      if (!(name in available)) return m;
      used.add(name);
      return ` data-img="${escapeHtml(name)}"`;
    })
    .replace(/url\(images\/([^)"']+)\)/g, (m, name) => {
      if (!(name in available)) return m;
      used.add(name);
      if (!cssVars.has(name)) cssVars.set(name, `--img-${cssVars.size}`);
      return `var(${cssVars.get(name)})`;
    });
  const rootCss = cssVars.size ? `:root{${[...cssVars].map(([name, v]) => `${v}:url(${available[name]})`).join(';')}}` : '';
  const imagesJson = JSON.stringify(Object.fromEntries([...used].map((n) => [n, available[n]])));
  slidesHtml = slidesHtml.trim();

  return `<!doctype html>
<html lang="${escapeHtml(o.lang || 'en')}">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>${escapeHtml(o.title)}</title>
    <!--
      Self-contained deck converted from PowerPoint. Reveal.js, the slides and
      their pictures are inlined below. No server and no internet needed.${o.source ? `\n      Generated from ${escapeHtml(o.source)}.` : ''}
      Add ?lang=<code> to the address to show an embedded translation.
    -->
    <style>${stripRemoteFonts(o.css.reveal)}</style>
    <style>${stripRemoteFonts(o.css.theme)}</style>
    <style>${o.css.shared}</style>
    <style>${o.css.pptx}</style>
    ${rootCss ? `<style>${rootCss}</style>` : ''}
  </head>
  <body>
    <div class="reveal">
      <div class="slides">
${slidesHtml}
      </div>
    </div>

    <script id="deck-strings" type="application/json">${json}</script>
    <script id="deck-images" type="application/json">${imagesJson}</script>
    <script>${o.js.reveal}</script>
    <script>${o.js.i18n}</script>
    <script>
      (async () => {
        const images = JSON.parse(document.getElementById('deck-images').textContent);
        document.querySelectorAll('img[data-img]').forEach((img) => { img.src = images[img.dataset.img]; });
        const tables = JSON.parse(document.getElementById('deck-strings').textContent);
        const lang = DeckI18n.langFromUrl();
        if (lang && tables[lang]) DeckI18n.apply(tables[lang], lang);
        await Reveal.initialize({
          hash: true,
          slideNumber: 'c/t',
          transition: 'none',
          width: ${Number(canvas.width) || 1280},
          height: ${Number(canvas.height) || 720},
          margin: 0,
          minScale: 0.2,
          maxScale: 3,
          center: false,
        });
        DeckI18n.autoFit(Reveal);
        DeckI18n.mountPicker({
          source: ${JSON.stringify(o.lang || 'en')},
          languages: Object.keys(tables),
          reveal: Reveal,
          load: async (tag) => tables[tag],
        });
        window.DECK_READY = true;
      })();
    </script>
  </body>
</html>
`;
}

/**
 * Converts PresentationML (the XML inside a .pptx) into deck Markdown.
 *
 * This module runs unchanged in Node and in the browser, so the CLI
 * (tools/import-pptx.mjs) and the import page (import.html) cannot drift
 * apart. It therefore uses no Node APIs and no dependencies. Callers do their
 * own decompression and hand in the already-inflated entries.
 */

const textDecoder = new TextDecoder('utf-8');

export const SLUG_PATTERN = /^[a-z0-9][a-z0-9._-]*$/i;

export function slugify(value) {
  const out = String(value).toLowerCase().normalize('NFKD')
    .replace(/[^\w\s.-]/g, '')
    .trim()
    .replace(/[\s_]+/g, '-')
    .replace(/-{2,}/g, '-')
    .replace(/^[-.]+|[-.]+$/g, '');
  return out || 'imported-deck';
}

/* -------------------------------------------------------------- zip ---- */

/**
 * Read a ZIP central directory and describe each entry. Decompression is left
 * to the caller, because Node does it synchronously and browsers do not.
 */
export function listZipEntries(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  const floor = Math.max(0, bytes.length - 66000);
  for (let i = bytes.length - 22; i >= floor; i--) {
    if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('Not a .pptx file (no ZIP end-of-central-directory record).');

  const count = view.getUint16(eocd + 10, true);
  let off = view.getUint32(eocd + 16, true);
  if (off === 0xffffffff || count === 0xffff) {
    throw new Error('ZIP64 archives are not supported. Re-save the deck from PowerPoint and try again.');
  }

  const entries = [];
  for (let n = 0; n < count; n++) {
    if (off + 46 > bytes.length || view.getUint32(off, true) !== 0x02014b50) break;
    const method = view.getUint16(off + 10, true);
    const compSize = view.getUint32(off + 20, true);
    const nameLen = view.getUint16(off + 28, true);
    const extraLen = view.getUint16(off + 30, true);
    const commentLen = view.getUint16(off + 32, true);
    const localOff = view.getUint32(off + 42, true);
    const name = textDecoder.decode(bytes.subarray(off + 46, off + 46 + nameLen));

    if (view.getUint32(localOff, true) !== 0x04034b50) throw new Error(`Corrupt ZIP entry: ${name}`);
    const dataStart = localOff + 30
      + view.getUint16(localOff + 26, true)
      + view.getUint16(localOff + 28, true);

    entries.push({ name, method, dataStart, compSize });
    off += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

/* -------------------------------------------------------------- xml ---- */

const decodeXml = (s) => s
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
  .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
  .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
  .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
  .replace(/&amp;/g, '&');

const attr = (tag, name) => {
  const m = new RegExp(`\\b${name.replace(':', '\\:')}="([^"]*)"`).exec(tag);
  return m ? decodeXml(m[1]) : null;
};

const inner = (el, name) => {
  const open = el.indexOf('>');
  const close = el.lastIndexOf(`</${name}>`);
  return close > open ? el.slice(open + 1, close) : '';
};

/**
 * Find every `<name>` element, tracking nesting so each returned slice is a
 * complete element. Results are in document order with their source offsets.
 */
function findElements(xml, name) {
  const out = [];
  const openTag = new RegExp(`<${name}(?=[\\s/>])`, 'g');
  let m;
  while ((m = openTag.exec(xml))) {
    const start = m.index;
    const headEnd = xml.indexOf('>', start);
    if (headEnd < 0) break;
    if (xml[headEnd - 1] === '/') {
      out.push({ start, xml: xml.slice(start, headEnd + 1) });
      openTag.lastIndex = headEnd + 1;
      continue;
    }
    const scan = new RegExp(`<${name}(?=[\\s/>])|</${name}>`, 'g');
    scan.lastIndex = headEnd + 1;
    let depth = 1;
    let end = -1;
    let s;
    while ((s = scan.exec(xml))) {
      if (s[0][1] === '/') {
        if (--depth === 0) { end = s.index + s[0].length; break; }
      } else {
        const gt = xml.indexOf('>', s.index);
        if (gt > 0 && xml[gt - 1] === '/') { scan.lastIndex = gt + 1; continue; }
        depth++;
      }
    }
    if (end < 0) break;
    out.push({ start, xml: xml.slice(start, end) });
    openTag.lastIndex = end;
  }
  return out;
}

/* ---------------------------------------------------------- markdown ---- */

/**
 * Escape only what would otherwise change the Markdown structure.
 *
 * `*` and `_` are deliberately left alone so that star notation typed into the
 * original slides still reads as emphasis.
 */
const escapeInline = (s) => s
  .replace(/([\\`[\]])/g, '\\$1')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;');

/**
 * Slides paste code into a text box and set a monospace font. That font is the
 * only reliable signal that the content is code rather than prose, and getting
 * it right matters: run through the normal path, a leading "#" comment turns
 * into an H1 and every indent is collapsed.
 */
const MONO_FONT = /^(cascadia|consolas|courier|menlo|monaco|lucida console|fira (code|mono)|source code pro|jetbrains mono|sf mono|roboto mono|ibm plex mono|ubuntu mono|dejavu sans mono|liberation mono)/i;

function isCodeShape(spXml) {
  // A monospace font on a bulleted or numbered shape means a styled list, not code.
  if (/<a:buAutoNum|<a:buChar/.test(spXml)) return false;
  const faces = [...spXml.matchAll(/typeface="([^"]+)"/g)].map((m) => m[1]);
  return faces.length > 0 && faces.every((f) => MONO_FONT.test(f));
}

/** Code text, kept verbatim: no escaping, no collapsing, indentation intact. */
function codeLines(spXml) {
  const out = [];
  for (const p of findElements(spXml, 'a:p')) {
    let line = '';
    const tokens = [
      ...findElements(p.xml, 'a:r').map((r) => ({ ...r, br: false })),
      ...findElements(p.xml, 'a:br').map((b) => ({ ...b, br: true })),
    ].sort((a, b) => a.start - b.start);
    for (const t of tokens) {
      if (t.br) { out.push(line); line = ''; continue; }
      line += findElements(t.xml, 'a:t').map((x) => decodeXml(inner(x.xml, 'a:t'))).join('');
    }
    out.push(line);
  }
  while (out.length && !out[0].trim()) out.shift();
  while (out.length && !out[out.length - 1].trim()) out.pop();
  return out.map((l) => l.replace(/[\u200b-\u200d\ufeff]/g, '').replace(/\s+$/, ''));
}

/** A lone "---" or "--" line would be read as a slide separator. */
const deSeparator = (line) => (/^\s*-{2,3}\s*$/.test(line) ? line.replace('-', '\\-') : line);

/** Inline runs of one paragraph, keeping bold, italic and hyperlinks. */
function runsToMarkdown(paragraphXml, rels) {
  const tokens = [];
  for (const r of findElements(paragraphXml, 'a:r')) tokens.push({ ...r, kind: 'run' });
  for (const f of findElements(paragraphXml, 'a:fld')) tokens.push({ ...f, kind: 'run' });
  for (const b of findElements(paragraphXml, 'a:br')) tokens.push({ ...b, kind: 'br' });
  tokens.sort((a, b) => a.start - b.start);

  let out = '';
  for (const token of tokens) {
    if (token.kind === 'br') { out += ' '; continue; }
    const raw = findElements(token.xml, 'a:t').map((t) => decodeXml(inner(t.xml, 'a:t'))).join('');
    if (!raw) continue;

    // Strip the zero-width characters PowerPoint sprinkles into text.
    let piece = raw.replace(/[\u200b-\u200d\ufeff]/g, '').replace(/\s+/g, ' ');
    if (!piece) continue;

    const props = findElements(token.xml, 'a:rPr')[0]?.xml ?? '';
    const lead = /^\s/.test(piece) ? ' ' : '';
    const tail = /\s$/.test(piece) ? ' ' : '';
    const core = piece.trim();
    if (core) {
      let styled = escapeInline(core);
      if (attr(props, 'b') === '1') styled = `**${styled}**`;
      if (attr(props, 'i') === '1') styled = `*${styled}*`;
      const linkId = attr(findElements(token.xml, 'a:hlinkClick')[0]?.xml ?? '', 'r:id');
      const href = linkId ? rels.get(linkId) : null;
      if (href && /^https?:/i.test(href)) styled = `[${styled}](${href})`;
      piece = lead + styled + tail;
    }
    out += piece;
  }
  return out.replace(/\s+/g, ' ').trim();
}

/**
 * PowerPoint often arranges several images side by side. One Markdown
 * paragraph each would stack them vertically and blow past the bottom of the
 * slide, so a run of images becomes a single paragraph that flows and wraps.
 */
function mergeImageRuns(lines) {
  const isImage = (l) => /^!\[[^\]]*\]\([^)]*\)$/.test(l.trim());
  const out = [];

  for (let i = 0; i < lines.length; i += 1) {
    if (!isImage(lines[i])) { out.push(lines[i]); continue; }

    const group = [];
    let j = i;
    let last = i;
    while (j < lines.length && (isImage(lines[j]) || !lines[j].trim())) {
      if (isImage(lines[j])) { group.push(lines[j].trim()); last = j; }
      j += 1;
    }
    out.push(group.join(' '), '');
    i = last;
  }
  return out;
}

/** One paragraph -> { text, level, bullet } */
function readParagraph(pXml, rels, bulletDefault) {
  const text = runsToMarkdown(pXml, rels);
  if (!text) return null;
  const props = findElements(pXml, 'a:pPr')[0]?.xml ?? '';
  const level = Number(attr(props, 'lvl') ?? 0);
  let bullet = bulletDefault;
  if (/<a:buNone\b/.test(props)) bullet = 'none';
  else if (/<a:buAutoNum\b/.test(props)) bullet = 'number';
  else if (/<a:buChar\b/.test(props)) bullet = 'bullet';
  return { text, level, bullet };
}

/** a:tbl -> a Markdown table. */
function tableToMarkdown(frameXml, rels) {
  const tbl = findElements(frameXml, 'a:tbl')[0];
  if (!tbl) return null;
  const rows = findElements(tbl.xml, 'a:tr').map((tr) =>
    findElements(tr.xml, 'a:tc').map((tc) =>
      findElements(tc.xml, 'a:p')
        .map((p) => runsToMarkdown(p.xml, rels))
        .filter(Boolean)
        .join(' ')
        .replace(/\|/g, '\\|') || ' '));
  if (!rows.length) return null;

  const width = Math.max(...rows.map((r) => r.length));
  const pad = (r) => [...r, ...Array(width - r.length).fill(' ')];
  const [head, ...body] = rows;
  return [
    `| ${pad(head).join(' | ')} |`,
    `| ${Array(width).fill('---').join(' | ')} |`,
    ...body.map((r) => `| ${pad(r).join(' | ')} |`),
  ].join('\n');
}

/* ------------------------------------------------------------ shapes ---- */

const SKIP_PLACEHOLDERS = new Set(['ftr', 'sldNum', 'dt']);
const isTitle = (type) => type === 'title' || type === 'ctrTitle';

/** Ordered, non-overlapping shapes; recurses into groups. */
function collectShapes(spTreeXml) {
  const groups = findElements(spTreeXml, 'p:grpSp');
  const covered = groups.map((g) => [g.start, g.start + g.xml.length]);
  const insideGroup = (pos) => covered.some(([a, b]) => pos > a && pos < b);

  const items = [];
  for (const kind of ['p:sp', 'p:pic', 'p:graphicFrame']) {
    for (const el of findElements(spTreeXml, kind)) {
      if (!insideGroup(el.start)) items.push({ kind, xml: el.xml, start: el.start });
    }
  }
  for (const g of groups) {
    // Recurse into the group's contents, not the group element itself, or a
    // nested group would match itself forever.
    items.push({ kind: 'group', start: g.start, children: collectShapes(inner(g.xml, 'p:grpSp')) });
  }
  items.sort((a, b) => a.start - b.start);
  return items.flatMap((i) => (i.kind === 'group' ? i.children : [i]));
}

function shapeInfo(spXml) {
  const nv = findElements(spXml, 'p:nvSpPr')[0]?.xml ?? '';
  const ph = findElements(nv, 'p:ph')[0]?.xml;
  return { placeholder: ph ? (attr(ph, 'type') ?? 'body') : null };
}

/* ------------------------------------------------------------- build ---- */

/** Resolve a relationship target against the folder of the part that owns it. */
function resolvePath(baseDir, target) {
  const stack = [];
  for (const part of `${baseDir}/${target}`.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') stack.pop();
    else stack.push(part);
  }
  return stack.join('/');
}

const basename = (p) => p.split('/').pop();

/**
 * Convert an inflated .pptx into deck sources.
 *
 * @param {Map<string, Uint8Array>} files  every part of the package, inflated
 * @param {object} [options]  title, description, author, date, theme, images
 * @returns {{markdown: string, deck: object, images: Array, warnings: string[], stats: object}}
 */
export function pptxToDeck(files, options = {}) {
  const includeImages = options.images !== false;
  const text = (name) => {
    const bytes = files.get(name);
    return bytes ? textDecoder.decode(bytes) : '';
  };

  const relsFor = (part) => {
    const dir = part.split('/').slice(0, -1).join('/');
    const xml = text(`${dir}/_rels/${basename(part)}.rels`);
    const map = new Map();
    for (const { xml: tag } of findElements(xml, 'Relationship')) {
      const id = attr(tag, 'Id');
      const target = attr(tag, 'Target');
      if (!id || !target) continue;
      map.set(id, attr(tag, 'TargetMode') === 'External' ? target : resolvePath(dir, target));
    }
    return map;
  };

  const presRels = relsFor('ppt/presentation.xml');
  const slidePaths = [...text('ppt/presentation.xml').matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"/g)]
    .map((m) => presRels.get(m[1]))
    .filter((p) => p && files.has(p));
  if (!slidePaths.length) throw new Error('No slides found. Is this really a .pptx?');

  const warnings = [];
  const images = new Map();
  const chunks = [];
  let notesCount = 0;

  slidePaths.forEach((slidePath, index) => {
    const xml = text(slidePath);
    const rels = relsFor(slidePath);
    const spTree = findElements(xml, 'p:spTree')[0]?.xml ?? xml;
    const label = `slide ${index + 1}`;

    let title = null;
    const lines = [];
    let slideImages = 0;

    for (const shape of collectShapes(spTree)) {
      if (shape.kind === 'p:sp') {
        const { placeholder } = shapeInfo(shape.xml);
        if (placeholder && SKIP_PLACEHOLDERS.has(placeholder)) continue;

        if (!isTitle(placeholder) && isCodeShape(shape.xml)) {
          const code = codeLines(shape.xml);
          if (code.length) {
            // A longer fence than anything inside keeps embedded backticks safe.
            const longest = Math.max(0, ...code.map((l) => (l.match(/`+/g) ?? []).map((b) => b.length)).flat());
            const fence = '`'.repeat(Math.max(3, longest + 1));
            lines.push(fence, ...code, fence, '');
            continue;
          }
        }

        const bulletDefault = placeholder && !isTitle(placeholder) && placeholder !== 'subTitle'
          ? 'bullet'
          : 'none';
        const paras = findElements(shape.xml, 'a:p')
          .map((p) => readParagraph(p.xml, rels, bulletDefault))
          .filter(Boolean);
        if (!paras.length) continue;

        if (isTitle(placeholder) && title === null) {
          title = paras.map((p) => p.text).join(' ');
          continue;
        }
        for (const p of paras) {
          const indent = '  '.repeat(p.level);
          if (p.bullet === 'none') lines.push(deSeparator(p.level ? `${indent}${p.text}` : p.text));
          else if (p.bullet === 'number') lines.push(`${indent}1. ${p.text}`);
          else lines.push(`${indent}- ${p.text}`);
        }
        lines.push('');
      } else if (shape.kind === 'p:pic' && includeImages) {
        const blip = findElements(shape.xml, 'a:blip')[0]?.xml ?? '';
        // Modern Office stores SVG in an extension, keeping a raster fallback.
        const svgId = attr(findElements(blip, 'asvg:svgBlip')[0]?.xml ?? '', 'r:embed');
        const id = svgId ?? attr(blip, 'r:embed');
        const mediaPath = id ? rels.get(id) : null;
        if (!mediaPath || !files.has(mediaPath)) continue;

        const cNvPr = findElements(shape.xml, 'p:cNvPr')[0]?.xml ?? '';
        const alt = attr(cNvPr, 'descr') ?? attr(cNvPr, 'name') ?? '';
        if (!images.has(mediaPath)) images.set(mediaPath, basename(mediaPath));
        lines.push(`![${escapeInline(alt).replace(/\|/g, '')}](images/${images.get(mediaPath)})`, '');
        slideImages += 1;
      } else if (shape.kind === 'p:graphicFrame') {
        const table = tableToMarkdown(shape.xml, rels);
        if (table) { lines.push(table, ''); continue; }
        if (/<c:chart\b|charts\/chart/.test(shape.xml)) warnings.push(`${label}: chart dropped, rebuild it by hand`);
        else if (/diagramData|<dgm:/.test(shape.xml)) warnings.push(`${label}: SmartArt dropped, rebuild it by hand`);
      }
    }

    // Speaker notes live in a separate part, linked from the slide's rels.
    const notesPath = [...rels.values()].find((v) => /notesSlide\d+\.xml$/.test(v));
    let notes = '';
    if (notesPath) {
      const notesXml = text(notesPath);
      const notesTree = findElements(notesXml, 'p:spTree')[0]?.xml ?? notesXml;
      const notesRels = relsFor(notesPath);
      notes = collectShapes(notesTree)
        .filter((s) => s.kind === 'p:sp' && shapeInfo(s.xml).placeholder === 'body')
        .flatMap((s) => findElements(s.xml, 'a:p').map((p) => runsToMarkdown(p.xml, notesRels)))
        .filter((t) => t && !/^\d+$/.test(t))
        .join('\n');
    }
    if (notes) notesCount++;

    if (!title && !lines.join('').trim() && !notes) {
      warnings.push(`${label}: empty, skipped`);
      return;
    }

    const heading = title
      ? `${chunks.length === 0 ? '#' : '##'} ${title}`
      : `<!-- ${label}: no title in the original -->`;
    if (!title) warnings.push(`${label}: no title placeholder, add a heading`);

    /*
     * A diagram in PowerPoint is often a dozen separate icons arranged by
     * position. Markdown has no positions, so it stacks them and the meaning is
     * lost. Nothing downstream can repair that, so name the slide here.
     */
    if (slideImages >= 4) {
      warnings.push(`${label}: ${slideImages} images, probably a diagram - replace them with one exported picture`);
    }

    const block = [heading, '', ...mergeImageRuns(lines)].join('\n').replace(/\n{3,}/g, '\n\n').trimEnd();
    chunks.push(notes ? `${block}\n\nNote: ${notes}` : block);
  });

  const core = text('docProps/core.xml');
  const coreTitle = inner(findElements(core, 'dc:title')[0]?.xml ?? '', 'dc:title');
  const coreAuthor = inner(findElements(core, 'dc:creator')[0]?.xml ?? '', 'dc:creator');

  const deck = {
    title: options.title || coreTitle || options.slug || 'Imported deck',
    description: options.description ?? 'Imported from PowerPoint.',
    author: options.author ?? coreAuthor ?? '',
    date: options.date ?? new Date().toISOString().slice(0, 10),
    tags: ['imported'],
    theme: options.theme ?? 'night',
  };

  return {
    markdown: `${chunks.join('\n\n---\n\n')}\n`,
    deck,
    images: [...images].map(([zipPath, name]) => ({ zipPath, name })),
    warnings,
    stats: {
      slidesRead: slidePaths.length,
      slidesWritten: chunks.length,
      notes: notesCount,
      images: images.size,
    },
  };
}

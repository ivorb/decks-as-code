/**
 * Converts an inflated .pptx into positioned HTML plus a string table.
 *
 *   pptxToHtml(files, options) -> { html, strings, images, deck, warnings, stats }
 *
 * Every slide becomes a fixed-size canvas (the deck's own size, 1 CSS px per
 * 9525 EMU) with shapes placed absolutely. All text stays live HTML, and each
 * paragraph is tagged with data-i18n so it can be swapped for a translation.
 * See docs/string-table.md for the contract this module writes.
 *
 * Like pptx-core.mjs this uses no Node APIs and no dependencies; callers do
 * their own decompression and pass in the entries as a Map<string, Uint8Array>.
 */
import { parseXml, child, kids, path, descendants, attrOf } from './xml-lite.mjs';
import {
  emuToPx, ptToPx, px, readColor, colorToCss, readTheme, readFill, styleFill, readLine,
  readRunProps, readParaProps, mergeProps, levelProps, fontStack, fillFromElement, MONO_FONT,
} from './pptx-style.mjs';
import {
  presetGeometry, connectorPath, customGeometry, readAdjust,
} from './pptx-geom.mjs';

const decoder = new TextDecoder('utf-8');
const num = (v, d = 0) => (v == null || v === '' || Number.isNaN(Number(v)) ? d : Number(v));
const truthy = (v) => v === '1' || v === 'true';
const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const escAttr = (s) => esc(s).replace(/"/g, '&quot;');
const round = (n) => Math.round(n * 100) / 100;

/** Inline style string from an object, skipping empty values. */
const style = (o) => Object.entries(o).filter(([, v]) => v != null && v !== '').map(([k, v]) => `${k}:${v}`).join(';');

/* ----------------------------------------------------------- package ---- */

function resolvePath(baseDir, target) {
  if (target.startsWith('/')) return target.slice(1);
  const stack = [];
  for (const part of `${baseDir}/${target}`.split('/')) {
    if (part === '..') stack.pop();
    else if (part && part !== '.') stack.push(part);
  }
  return stack.join('/');
}

function openPackage(files) {
  const xmlCache = new Map();
  const relCache = new Map();
  const text = (name) => (files.has(name) ? decoder.decode(files.get(name)) : '');
  const xml = (name) => {
    if (!files.has(name)) return null;
    if (!xmlCache.has(name)) xmlCache.set(name, parseXml(text(name)));
    return xmlCache.get(name);
  };
  const rels = (part) => {
    if (relCache.has(part)) return relCache.get(part);
    const dir = part.split('/').slice(0, -1).join('/');
    const map = new Map();
    const root = xml(`${dir}/_rels/${part.split('/').pop()}.rels`);
    for (const r of kids(root, 'Relationship')) {
      const external = r.attrs.TargetMode === 'External';
      map.set(r.attrs.Id, {
        type: r.attrs.Type ?? '',
        external,
        target: external ? r.attrs.Target : resolvePath(dir, r.attrs.Target ?? ''),
      });
    }
    relCache.set(part, map);
    return map;
  };
  return { files, text, xml, rels };
}

const relOfType = (rels, suffix) => [...rels.values()].find((r) => r.type.endsWith(suffix))?.target ?? null;

/* ------------------------------------------------------ master/layout ---- */

const phType = (t) => (t === 'ctrTitle' ? 'title' : ['subTitle', 'obj', 'body', 'chart', 'tbl', 'media', 'pic'].includes(t) ? 'body' : t);

function readPh(node) {
  const nv = child(node, node.name === 'p:pic' ? 'p:nvPicPr' : 'p:nvSpPr');
  const ph = path(nv, 'p:nvPr', 'p:ph');
  return ph ? { type: ph.attrs.type ?? 'body', idx: ph.attrs.idx ?? null } : null;
}

function placeholdersOf(tree) {
  const out = [];
  for (const n of tree?.children ?? []) {
    if (n.name !== 'p:sp' && n.name !== 'p:pic') continue;
    const ph = readPh(n);
    if (ph) out.push({ ...ph, node: n });
  }
  return out;
}

function buildChain(pkg, slidePath, cache) {
  const slideRels = pkg.rels(slidePath);
  const layoutPath = relOfType(slideRels, '/slideLayout');
  if (cache.has(layoutPath)) return cache.get(layoutPath);

  const layoutXml = layoutPath ? pkg.xml(layoutPath) : null;
  const layoutRels = layoutPath ? pkg.rels(layoutPath) : new Map();
  const masterPath = relOfType(layoutRels, '/slideMaster');
  const masterXml = masterPath ? pkg.xml(masterPath) : null;
  const masterRels = masterPath ? pkg.rels(masterPath) : new Map();
  const themePath = relOfType(masterRels, '/theme');
  const theme = readTheme(themePath ? pkg.xml(themePath) : null);

  const clrMap = { ...(child(masterXml, 'p:clrMap')?.attrs ?? {}) };
  const override = child(child(layoutXml, 'p:clrMapOvr'), 'a:overrideClrMapping');
  if (override) Object.assign(clrMap, override.attrs);

  const layoutTree = path(layoutXml, 'p:cSld', 'p:spTree');
  const masterTree = path(masterXml, 'p:cSld', 'p:spTree');
  const chain = {
    layoutPath, masterPath, layoutXml, masterXml, layoutRels, masterRels, theme, clrMap,
    env: { theme, clrMap },
    layoutTree,
    masterTree,
    layoutPhs: placeholdersOf(layoutTree),
    masterPhs: placeholdersOf(masterTree),
    txStyles: child(masterXml, 'p:txStyles'),
  };
  cache.set(layoutPath, chain);
  return chain;
}

/** The layout and master placeholder a slide placeholder inherits from. */
function inheritFor(chain, ph) {
  if (!ph) return [];
  const sameType = (m) => phType(m.type) === phType(ph.type);
  const layout = (ph.idx != null ? chain.layoutPhs.find((m) => m.idx === ph.idx) : null)
    ?? chain.layoutPhs.find(sameType) ?? null;
  const master = chain.masterPhs.find((m) => phType(m.type) === phType(layout?.type ?? ph.type)) ?? null;
  return [layout?.node, master?.node].filter(Boolean);
}

/* ------------------------------------------------------------ helpers ---- */

function readXfrm(node) {
  const x = node && (node.name === 'a:xfrm' || node.name === 'p:xfrm' ? node : child(node, 'a:xfrm'));
  const off = child(x, 'a:off');
  const ext = child(x, 'a:ext');
  if (!x || !off || !ext) return null;
  const chOff = child(x, 'a:chOff');
  const chExt = child(x, 'a:chExt');
  return {
    x: num(off.attrs.x), y: num(off.attrs.y), cx: num(ext.attrs.cx), cy: num(ext.attrs.cy),
    rot: num(x.attrs.rot) / 60000, flipH: truthy(x.attrs.flipH), flipV: truthy(x.attrs.flipV),
    chOff: chOff ? { x: num(chOff.attrs.x), y: num(chOff.attrs.y) } : null,
    chExt: chExt ? { cx: num(chExt.attrs.cx), cy: num(chExt.attrs.cy) } : null,
  };
}

/*
 * Transforms are 2D affine matrices in EMU: X = a*x + c*y + e, Y = b*x + d*y + f.
 * Groups compose one over the other, so nested groups, scaling, flips and
 * rotation all fall out of the same arithmetic.
 */
const IDENTITY = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };

const mul = (P, Q) => ({
  a: P.a * Q.a + P.c * Q.b, b: P.b * Q.a + P.d * Q.b,
  c: P.a * Q.c + P.c * Q.d, d: P.b * Q.c + P.d * Q.d,
  e: P.a * Q.e + P.c * Q.f + P.e, f: P.b * Q.e + P.d * Q.f + P.f,
});

const scaleX = (T) => Math.hypot(T.a, T.b);
const scaleY = (T) => Math.hypot(T.c, T.d);
const translate = (T, x, y) => mul(T, { a: 1, b: 0, c: 0, d: 1, e: x, f: y });

/** Where a shape's box lands on the slide, in px, with its net rotation and flips. */
function place(xf, T) {
  const cx = xf.x + xf.cx / 2;
  const cy = xf.y + xf.cy / 2;
  const centerX = T.a * cx + T.c * cy + T.e;
  const centerY = T.b * cx + T.d * cy + T.f;
  const w = emuToPx(xf.cx * scaleX(T));
  const h = emuToPx(xf.cy * scaleY(T));
  const mirrored = T.a * T.d - T.b * T.c < 0;
  const turn = mirrored ? Math.atan2(-T.b, -T.a) : Math.atan2(T.b, T.a);
  return {
    x: emuToPx(centerX) - w / 2,
    y: emuToPx(centerY) - h / 2,
    w,
    h,
    rot: xf.rot + (turn * 180) / Math.PI,
    flipH: xf.flipH !== mirrored,
    flipV: xf.flipV,
  };
}

function composeGroup(T, xf) {
  const sx = xf.chExt?.cx ? xf.cx / xf.chExt.cx : 1;
  const sy = xf.chExt?.cy ? xf.cy / xf.chExt.cy : 1;
  const into = { a: sx, b: 0, c: 0, d: sy, e: xf.x - sx * (xf.chOff?.x ?? 0), f: xf.y - sy * (xf.chOff?.y ?? 0) };
  const cx = xf.x + xf.cx / 2;
  const cy = xf.y + xf.cy / 2;
  const fx = xf.flipH ? -1 : 1;
  const fy = xf.flipV ? -1 : 1;
  const flip = { a: fx, b: 0, c: 0, d: fy, e: cx * (1 - fx), f: cy * (1 - fy) };
  const th = (xf.rot * Math.PI) / 180;
  const cos = Math.cos(th);
  const sin = Math.sin(th);
  const turn = { a: cos, b: sin, c: -sin, d: cos, e: cx - cos * cx + sin * cy, f: cy - sin * cx - cos * cy };
  return mul(T, mul(turn, mul(flip, into)));
}

const firstOf = (list, fn) => { for (const item of list) { const v = fn(item); if (v != null) return v; } return null; };

function rotFlip(box, flip = false) {
  const parts = [];
  if (box.rot) parts.push(`rotate(${round(box.rot)}deg)`);
  if (flip && (box.flipH || box.flipV)) parts.push(`scale(${box.flipH ? -1 : 1},${box.flipV ? -1 : 1})`);
  return parts.length ? parts.join(' ') : null;
}

function warn(S, message) {
  const text = `slide ${S.slideNo}: ${message}`;
  if (!S.warnings.includes(text)) S.warnings.push(text);
}

function imageRef(S, zipPath) {
  if (!S.images.has(zipPath)) {
    let name = zipPath.split('/').pop();
    const taken = new Set(S.images.values());
    let n = 1;
    const base = name;
    while (taken.has(name)) name = `${++n}-${base}`;
    S.images.set(zipPath, name);
  }
  return `images/${S.images.get(zipPath)}`;
}

function bgCss(fill, S, part) {
  if (!fill) return null;
  if (fill.type === 'solid' || fill.type === 'grad') return { background: fill.css };
  if (fill.type === 'pic') {
    const target = part.rels.get(fill.relId)?.target;
    if (!target || !S.pkg.files.has(target)) return null;
    if (/\.(emf|wmf)$/i.test(target)) { warn(S, 'EMF/WMF picture fill not rendered'); return null; }
    return { 'background-image': `url(${imageRef(S, target)})`, 'background-size': fill.tile ? 'auto' : '100% 100%', 'background-repeat': fill.tile ? 'repeat' : 'no-repeat' };
  }
  return null;
}

const DASH = { dash: 'dashed', sysDash: 'dashed', lgDash: 'dashed', dashDot: 'dashed', lgDashDot: 'dashed', sysDot: 'dotted', dot: 'dotted', sysDashDot: 'dashed', sysDashDotDot: 'dashed' };

/* --------------------------------------------------------------- text ---- */

const WINGDINGS = {
  '§': '▪', '\uF0A7': '▪', Ø: '➢', '\uF0D8': '➢', ü: '✓', '\uF0FC': '✓',
  l: '●', '\uF06C': '●', n: '■', '\uF06E': '■', v: '❖', '\uF076': '❖', q: '❑', u: '◆', '\uF0B7': '•',
};
const SYMBOLISH = /wingdings|symbol/i;

function toRoman(n) {
  const map = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  let out = '';
  let rest = n;
  for (const [v, s] of map) while (rest >= v) { out += s; rest -= v; }
  return out;
}

function autoNumber(n, scheme) {
  const alpha = (v) => { let s = ''; let k = v; while (k > 0) { s = String.fromCharCode(97 + ((k - 1) % 26)) + s; k = Math.floor((k - 1) / 26); } return s; };
  const core = /^alpha/.test(scheme) ? alpha(n) : /^roman/.test(scheme) ? toRoman(n) : String(n);
  const text = /Uc/.test(scheme) ? core.toUpperCase() : core;
  if (/ParenBoth$/.test(scheme)) return `(${text})`;
  if (/ParenR$/.test(scheme)) return `${text})`;
  return `${text}.`;
}

/** Resolve a typeface (including +mj-lt/+mn-lt) to a plain name. */
const faceName = (face, theme) => (face === '+mj-lt' ? theme.major : face === '+mn-lt' ? theme.minor : face);
const isMono = (rp, theme) => MONO_FONT.test(faceName(rp.font, theme) ?? '');

function safeHref(S, part, relId) {
  const rel = relId ? part.rels.get(relId) : null;
  if (!rel?.external) return null;
  return /^(https?:|mailto:)/i.test(rel.target) ? rel.target : null;
}

/**
 * One paragraph -> { html, string }. `string` is the inline HTML that goes in
 * the table (null for blank or code paragraphs, which are not translated).
 */
function renderParagraph(p, pIndex, tx) {
  const { S, part, env } = tx;
  const pPr = child(p, 'a:pPr');
  const level = Math.min(8, num(pPr?.attrs.lvl, 0));
  const props = mergeProps(tx.baseProps(level), readParaProps(pPr, env));

  const runs = [];
  for (const c of p.children) {
    if (c.name === 'a:br') runs.push({ br: true });
    else if (c.name === 'a:r' || c.name === 'a:fld') {
      const t = child(c, 'a:t')?.text ?? '';
      if (t !== '') runs.push({ text: t, rp: { ...props.r, ...readRunProps(child(c, 'a:rPr'), env) } });
    }
  }
  const textRuns = runs.filter((r) => !r.br);
  const plain = textRuns.map((r) => r.text).join('');
  const sizeOf = (rp) => (rp.sz ?? 1800) * tx.fontScale;

  const common = {
    'text-align': { l: 'left', ctr: 'center', r: 'right', just: 'justify' }[props.algn] ?? null,
  };
  const lnScale = 1 - tx.lnReduce;
  common['line-height'] = props.lnSpc?.pct ? round(props.lnSpc.pct * 1.2 * lnScale) : props.lnSpc?.pt ? px(ptToPx(props.lnSpc.pt)) : lnScale !== 1 ? round(1.2 * lnScale) : null;
  const space = (s) => (s?.pct != null ? `${round(s.pct * 1.2)}em` : s?.pt != null ? px(ptToPx(s.pt)) : null);
  common['margin-top'] = space(props.spcBef);
  common['margin-bottom'] = space(props.spcAft);

  if (!plain.trim()) {
    tx.numbering = {};
    const rp = { ...props.r, ...readRunProps(child(p, 'a:endParaRPr'), env) };
    return { html: `<p class="e" style="${style({ ...common, 'font-size': px(ptToPx(sizeOf(rp) / 100)) })}"></p>`, string: null };
  }

  const dominant = textRuns.reduce((a, b) => (b.text.length > a.text.length ? b : a), textRuns[0]).rp;
  const codeParagraph = textRuns.filter((r) => r.text.trim()).every((r) => isMono(r.rp, env.theme));

  // A paragraph is bold, italic or underlined as a whole when every run agrees;
  // when they differ, the inherited style is the baseline and the exceptions
  // become tags. Colors and sizes come from the dominant run, so the table
  // stays translator-friendly.
  const visible = textRuns.filter((r) => r.text.trim());
  const base = {};
  for (const k of ['b', 'i', 'u']) {
    const values = new Set(visible.map((r) => Boolean(r.rp[k])));
    base[k] = values.size === 1 ? [...values][0] : Boolean(props.r?.[k]);
  }

  const segs = [];
  for (const r of runs) {
    if (r.br) { segs.push({ br: true }); continue; }
    const tags = [];
    const href = safeHref(S, part, r.rp.link);
    if (!codeParagraph && isMono(r.rp, env.theme)) tags.push('code');
    if (r.rp.b && !base.b) tags.push('b');
    if (r.rp.i && !base.i) tags.push('i');
    if (r.rp.u && !base.u) tags.push('u');
    if (r.rp.baseline > 0) tags.push('sup');
    else if (r.rp.baseline < 0) tags.push('sub');
    const key = `${tags.join(',')}|${href ?? ''}`;
    const last = segs[segs.length - 1];
    if (last && !last.br && last.key === key) last.text += r.text;
    else segs.push({ key, tags, href, text: r.text });
  }
  const inline = segs.map((s) => {
    if (s.br) return '<br>';
    let out = esc(s.text);
    for (const t of [...s.tags].reverse()) out = `<${t}>${out}</${t}>`;
    return s.href ? `<a href="${escAttr(s.href)}">${out}</a>` : out;
  }).join('');

  // Bullets and indentation.
  let marL = props.marL ?? 0;
  let indent = props.indent ?? 0;
  const bu = props.bu ?? { type: 'none' };
  let bullet = null;
  if (bu.type === 'char') {
    const raw = bu.char ?? '•';
    bullet = SYMBOLISH.test(props.buFont ?? '') ? (WINGDINGS[raw] ?? '•') : raw;
  } else if (bu.type === 'auto') {
    const n = (tx.numbering[level] ?? bu.startAt - 1) + 1;
    tx.numbering[level] = n;
    bullet = autoNumber(n, bu.scheme);
  }
  for (const k of Object.keys(tx.numbering)) if (Number(k) > level) delete tx.numbering[k];
  if (bu.type !== 'auto') delete tx.numbering[level];
  if (bullet && indent >= 0) { indent = -285750; marL += 285750; }

  const mono = codeParagraph;
  const font = dominant.font ? fontStack(dominant.font, env.theme) : null;
  const css = {
    ...common,
    'font-size': `calc(var(--fit,1) * ${px(ptToPx(sizeOf(dominant) / 100))})`,
    color: dominant.color,
    'font-family': mono ? (font ?? "Consolas, 'Cascadia Mono', monospace") : font,
    'font-weight': base.b ? 700 : null,
    'font-style': base.i ? 'italic' : null,
    'text-decoration': [base.u ? 'underline' : null, dominant.strike ? 'line-through' : null].filter(Boolean).join(' ') || null,
    'text-transform': dominant.cap === 'all' ? 'uppercase' : null,
    'letter-spacing': dominant.spc ? px(ptToPx(dominant.spc / 100)) : null,
    'padding-left': marL ? px(emuToPx(marL)) : null,
    'text-indent': !bullet && indent ? px(emuToPx(indent)) : null,
  };
  let attrs = '';
  if (bullet) {
    css['--bl'] = px(emuToPx(marL + indent));
    if (props.buColor) css['--buc'] = props.buColor;
    if (props.buSize) css['--bus'] = `${round(props.buSize * 100)}%`;
    attrs = ` data-bu="${escAttr(bullet)}"`;
  }

  let string = null;
  let id = '';
  if (!codeParagraph) {
    id = `${tx.idPrefix}.p${pIndex}`;
    string = { id, kind: tx.kind, text: inline };
  } else S.stats.codeParagraphs += 1;

  const idAttr = string ? ` data-i18n="${id}"` : ' data-nt';
  return { html: `<p${idAttr}${attrs} style="${style(css)}">${inline}</p>`, string };
}

/** Merge bodyPr attributes from the placeholder chain (master -> layout -> slide). */
function bodyPrOf(txBody, inh) {
  const own = child(txBody, 'a:bodyPr');
  const chain = [...inh].reverse().map((n) => child(child(n, 'p:txBody'), 'a:bodyPr')).filter(Boolean);
  if (own) chain.push(own);
  const attrs = Object.assign({}, ...chain.map((b) => b.attrs));
  const fit = [...chain].reverse().map((b) => b.children.find((c) => c.name.startsWith('a:') && /AutoFit|noAutofit/i.test(c.name))).find(Boolean);
  return { attrs, fit };
}

/** Paragraphs of a text body -> HTML for the .tx container (and string pushes). */
function renderTextBody(txBody, S, part, opts) {
  const { env } = S.chain;
  const ph = opts.ph ?? null;
  const inh = opts.inh ?? [];
  const body = bodyPrOf(txBody, inh);
  const fontScale = body.fit?.name === 'a:normAutofit' ? num(body.fit.attrs.fontScale, 100000) / 100000 : 1;
  const lnReduce = body.fit?.name === 'a:normAutofit' ? num(body.fit.attrs.lnSpcReduction, 0) / 100000 : 0;
  const fontRef = readColor(child(opts.styleNode, 'a:fontRef'), env);

  const baseProps = (level) => {
    let p = levelProps(S.pres.defaultTextStyle, level, env);
    if (ph) {
      const key = phType(ph.type) === 'title' ? 'p:titleStyle' : phType(ph.type) === 'body' ? 'p:bodyStyle' : 'p:otherStyle';
      p = mergeProps(p, levelProps(child(S.chain.txStyles, key), level, env));
    }
    if (fontRef) p = mergeProps(p, { r: { color: colorToCss(fontRef) } });
    for (const n of [...inh].reverse()) p = mergeProps(p, levelProps(child(child(n, 'p:txBody'), 'a:lstStyle'), level, env));
    if (opts.extraLevels) p = mergeProps(p, opts.extraLevels(level));
    return mergeProps(p, levelProps(child(txBody, 'a:lstStyle'), level, env));
  };

  const tx = { S, part, env, baseProps, fontScale, lnReduce, numbering: {}, idPrefix: opts.idPrefix, kind: opts.kind ?? 'text' };
  const paragraphs = kids(txBody, 'a:p');
  const out = [];
  paragraphs.forEach((p, i) => {
    const { html, string } = renderParagraph(p, i + 1, tx);
    out.push(html);
    if (string) S.addString(string);
  });
  return { html: out.join(''), body, hasText: paragraphs.some((p) => descendants(p, 'a:t').some((t) => t.text.trim())) };
}

function txContainerStyle(body, box) {
  const a = body.attrs;
  const anchor = { t: 'flex-start', ctr: 'center', b: 'flex-end' }[a.anchor] ?? 'flex-start';
  const css = {
    padding: [num(a.tIns, 45720), num(a.rIns, 91440), num(a.bIns, 45720), num(a.lIns, 91440)].map((v) => px(emuToPx(v))).join(' '),
    'justify-content': anchor,
    'min-height': px(box.h),
    'white-space': a.wrap === 'none' ? 'pre' : null,
  };
  if (a.vert === 'vert' || a.vert === 'eaVert') css['writing-mode'] = 'vertical-rl';
  else if (a.vert === 'vert270') { css['writing-mode'] = 'vertical-rl'; css.transform = 'rotate(180deg)'; }
  return style(css);
}

/* ------------------------------------------------------------- shapes ---- */

function geometryOf(holders, box, S) {
  const prstGeom = firstOf(holders, (h) => child(h, 'a:prstGeom'));
  const cust = firstOf(holders, (h) => child(h, 'a:custGeom'));
  if (prstGeom) {
    const prst = prstGeom.attrs.prst ?? 'rect';
    const geo = presetGeometry(prst, readAdjust(prstGeom), box.w, box.h);
    if (!geo) { S.unknown.set(prst, (S.unknown.get(prst) ?? 0) + 1); return { kind: 'css', radius: null, prst }; }
    return { ...geo, prst };
  }
  if (cust) {
    const geo = customGeometry(cust, box.w, box.h);
    if (geo) return { ...geo, prst: 'custGeom' };
  }
  return { kind: 'css', radius: null, prst: 'rect' };
}

function lineCss(ln) {
  if (!ln) return null;
  return `${px(Math.max(1, ln.w))} ${DASH[ln.dash] ?? 'solid'} ${ln.color}`;
}

function clipRef(S, d) {
  let id = S.clips.get(d);
  if (!id) {
    id = `cp${S.slideNo}-${S.clips.size + 1}`;
    S.clips.set(d, id);
    S.defs.push(`<clipPath id="${id}" clipPathUnits="objectBoundingBox"><path d="${d}"/></clipPath>`);
  }
  return id;
}

function renderSp(sp, S, part, T) {
  const cNv = path(sp, 'p:nvSpPr', 'p:cNvPr');
  if (truthy(cNv?.attrs.hidden)) return '';
  const id = cNv?.attrs.id ?? '0';
  const ph = readPh(sp);
  if (part.decor && ph) return '';
  const spPr = child(sp, 'p:spPr');
  const styleNode = child(sp, 'p:style');
  const txBody = child(sp, 'p:txBody');
  const inh = part.decor ? [] : inheritFor(S.chain, ph);
  const holders = [spPr, ...inh.map((n) => child(n, 'p:spPr'))].filter(Boolean);
  const env = S.chain.env;

  const prst = child(spPr, 'a:prstGeom')?.attrs.prst;
  const hasText = txBody && descendants(txBody, 'a:t').some((t) => t.text.trim());
  if (prst && /^(line|straightConnector1|bentConnector\d|curvedConnector\d)$/.test(prst) && !hasText) {
    return renderConnector(sp, S, part, T);
  }

  const xf = firstOf([spPr, ...inh.map((n) => child(n, 'p:spPr'))], readXfrm);
  let box;
  if (xf) box = place(xf, T);
  else {
    box = { x: 0, y: 0, w: S.width, h: 0 };
    warn(S, `shape ${id} has no position; placed at the top-left`);
  }

  const geo = geometryOf(holders, box, S);
  let fill = firstOf(holders, (h) => readFill(h, env));
  if (!fill && styleNode) fill = styleFill(styleNode, env);
  const lnHolder = holders.find((h) => child(h, 'a:ln')) ?? spPr;
  const ln = readLine(lnHolder, styleNode, env);

  const lineBox = txBody ? renderTextBody(txBody, S, part, {
    ph, inh, styleNode, idPrefix: `${part.scope}.sh${id}`, kind: 'text',
  }) : null;

  const bg = bgCss(fill, S, part);
  const hasBg = Boolean(bg) || (ln && geo.kind === 'css');
  if (!lineBox?.hasText && !bg && !ln) return '';
  S.stats.shapes += 1;

  const flipT = xf && (box.flipH || box.flipV) ? `scale(${box.flipH ? -1 : 1},${box.flipV ? -1 : 1})` : null;
  const parts = [];
  if (geo.kind === 'css') {
    if (hasBg) {
      parts.push(`<div class="bg" style="${style({
        ...bg,
        border: lineCss(ln),
        'border-radius': geo.radius === '50%' ? '50%' : geo.radius ? px(geo.radius) : null,
        transform: flipT,
      })}"></div>`);
    }
  } else {
    if (bg && !geo.open) parts.push(`<div class="bg" style="${style({ ...bg, 'clip-path': `url(#${clipRef(S, geo.d)})`, transform: flipT })}"></div>`);
    if (ln) {
      const dash = DASH[ln.dash] ? `stroke-dasharray="${ln.dash.includes('ot') || ln.dash === 'sysDot' ? '1 3' : '6 4'}"` : '';
      parts.push(`<svg class="ln" viewBox="0 0 1 1" preserveAspectRatio="none" style="${style({ transform: flipT })}"><path d="${geo.d}" fill="none" stroke="${ln.color}" stroke-width="${round(ln.w)}" vector-effect="non-scaling-stroke" ${dash}/></svg>`);
    }
  }
  if (lineBox?.hasText || lineBox?.html) parts.push(`<div class="tx" style="${txContainerStyle(lineBox.body, box)}">${lineBox.html}</div>`);

  const shapeStyle = style({
    left: px(box.x), top: px(box.y), width: px(box.w), 'min-height': px(box.h), transform: xf ? rotFlip(box) : null,
  });
  return `<div class="sh" data-id="${id}" style="${shapeStyle}">${parts.join('')}</div>`;
}

function markerDefs(S, id, ln, which, end) {
  const kind = end.type === 'oval' ? 'oval' : end.type === 'diamond' ? 'diamond' : 'arrow';
  const size = { sm: 2, med: 3, lg: 5 };
  const mw = size[end.w] ?? 3;
  const mh = size[end.len] ?? 3;
  const shape = kind === 'oval' ? '<circle cx="5" cy="5" r="4"/>' : kind === 'diamond' ? '<path d="M5,0 L10,5 L5,10 L0,5 Z"/>' : '<path d="M0,0 L10,5 L0,10 Z"/>';
  return `<marker id="${id}" viewBox="0 0 10 10" refX="${kind === 'arrow' ? 8 : 5}" refY="5" markerUnits="strokeWidth" markerWidth="${mw}" markerHeight="${mh}" orient="auto-start-reverse" fill="${ln.color}">${shape}</marker>`;
}

function renderConnector(node, S, part, T) {
  const isSp = node.name === 'p:sp';
  const cNv = path(node, isSp ? 'p:nvSpPr' : 'p:nvCxnSpPr', 'p:cNvPr');
  if (truthy(cNv?.attrs.hidden)) return '';
  const id = cNv?.attrs.id ?? '0';
  const spPr = child(node, 'p:spPr');
  const xf = readXfrm(spPr);
  const ln = readLine(spPr, child(node, 'p:style'), S.chain.env);
  if (!xf || !ln) return '';
  const box = place(xf, T);
  const prstGeom = child(spPr, 'a:prstGeom');
  const w = Math.max(box.w, 0.01);
  const h = Math.max(box.h, 0.01);
  const d = connectorPath(prstGeom?.attrs.prst ?? 'line', readAdjust(prstGeom), w, h);
  S.stats.connectors += 1;

  const defs = [];
  let startM = '';
  let endM = '';
  if (ln.head) { const mid = `mk${S.slideNo}-${S.scopeKey}${id}h`; defs.push(markerDefs(S, mid, ln, 'h', ln.head)); startM = ` marker-start="url(#${mid})"`; }
  if (ln.tail) { const mid = `mk${S.slideNo}-${S.scopeKey}${id}t`; defs.push(markerDefs(S, mid, ln, 't', ln.tail)); endM = ` marker-end="url(#${mid})"`; }
  const dash = DASH[ln.dash] ? ` stroke-dasharray="${/ot/i.test(ln.dash) ? '1 3' : '6 4'}"` : '';
  const transform = [box.rot ? `rotate(${round(box.rot)}deg)` : null, box.flipH || box.flipV ? `scale(${box.flipH ? -1 : 1},${box.flipV ? -1 : 1})` : null].filter(Boolean).join(' ');
  const css = style({ left: px(box.x), top: px(box.y), width: px(w), height: px(h), transform: transform || null });
  return `<svg class="cx" data-id="${id}" style="${css}" viewBox="0 0 ${round(w)} ${round(h)}" overflow="visible"><defs>${defs.join('')}</defs><path d="${d}" fill="none" stroke="${ln.color}" stroke-width="${round(ln.w)}"${dash}${startM}${endM}/></svg>`;
}

function renderPic(pic, S, part, T) {
  const cNv = path(pic, 'p:nvPicPr', 'p:cNvPr');
  if (truthy(cNv?.attrs.hidden)) return '';
  const id = cNv?.attrs.id ?? '0';
  const ph = readPh(pic);
  if (part.decor && ph) return '';
  const inh = part.decor ? [] : inheritFor(S.chain, ph);
  const spPr = child(pic, 'p:spPr');
  const xf = firstOf([spPr, ...inh.map((n) => child(n, 'p:spPr'))], readXfrm);
  if (!xf) return '';
  const box = place(xf, T);

  const blipFill = child(pic, 'p:blipFill');
  const blip = child(blipFill, 'a:blip');
  const svgBlip = descendants(blip ?? { children: [] }, 'asvg:svgBlip')[0];
  const relId = svgBlip?.attrs['r:embed'] ?? blip?.attrs['r:embed'];
  const target = relId ? part.rels.get(relId)?.target : null;
  if (!target || !S.pkg.files.has(target)) return '';

  const env = S.chain.env;
  const geo = presetGeometry(child(spPr, 'a:prstGeom')?.attrs.prst ?? 'rect', readAdjust(child(spPr, 'a:prstGeom')), box.w, box.h);
  const radius = geo?.kind === 'css' ? (geo.radius === '50%' ? '50%' : geo.radius ? px(geo.radius) : null) : null;
  const ln = readLine(spPr, null, env);
  S.stats.pictures += 1;

  const outer = style({
    left: px(box.x), top: px(box.y), width: px(box.w), height: px(box.h), transform: rotFlip(box), 'border-radius': radius, border: lineCss(ln),
  });
  if (/\.(emf|wmf)$/i.test(target)) {
    warn(S, `picture ${id} is EMF/WMF and cannot be shown in a browser`);
    return `<div class="pic miss" style="${outer}" title="Unsupported picture format"></div>`;
  }

  const crop = child(blipFill, 'a:srcRect');
  const l = num(crop?.attrs.l) / 100000; const t = num(crop?.attrs.t) / 100000;
  const r = num(crop?.attrs.r) / 100000; const b = num(crop?.attrs.b) / 100000;
  const vw = 1 - l - r || 1; const vh = 1 - t - b || 1;
  const imgCss = crop ? style({
    width: `${round(100 / vw)}%`, height: `${round(100 / vh)}%`, left: `${round((-l / vw) * 100)}%`, top: `${round((-t / vh) * 100)}%`,
  }) : '';
  const flip = box.flipH || box.flipV ? `transform:scale(${box.flipH ? -1 : 1},${box.flipV ? -1 : 1})` : '';

  const descr = cNv?.attrs.descr?.trim();
  let altAttr = ' alt=""';
  if (descr) {
    const altId = `${part.scope}.sh${id}.alt`;
    S.addString({ id: altId, kind: 'alt', text: esc(descr) });
    altAttr = ` alt="${escAttr(descr)}" data-i18n-alt="${altId}"`;
  }
  return `<div class="pic" data-id="${id}" style="${outer}"><img src="${imageRef(S, target)}"${altAttr} style="${[imgCss, flip].filter(Boolean).join(';')}"></div>`;
}

/* ------------------------------------------------------------- tables ---- */

function tableStyleOf(S, styleId) {
  if (!styleId) return null;
  if (!S.tableStyles) S.tableStyles = S.pkg.xml('ppt/tableStyles.xml');
  return kids(S.tableStyles, 'a:tblStyle').find((s) => s.attrs.styleId === styleId) ?? null;
}

function renderTable(frame, tbl, S, part, T) {
  const xf = readXfrm(child(frame, 'p:xfrm'));
  if (!xf) return '';
  const box = place(xf, T);
  const env = S.chain.env;
  const id = path(frame, 'p:nvGraphicFramePr', 'p:cNvPr')?.attrs.id ?? '0';
  const tblPr = child(tbl, 'a:tblPr');
  const tstyle = tableStyleOf(S, child(tblPr, 'a:tableStyleId')?.text.trim());
  const firstRow = truthy(tblPr?.attrs.firstRow);
  const bandRow = truthy(tblPr?.attrs.bandRow);

  const part2 = (name) => child(tstyle, name);
  const cellStyle = (rowIdx) => {
    const picks = [part2('wholeTbl'), bandRow ? part2(rowIdx % 2 === (firstRow ? 1 : 0) ? 'band1H' : 'band2H') : null, firstRow && rowIdx === 0 ? part2('firstRow') : null].filter(Boolean);
    let fill = null; let color = null; let bold = false;
    for (const p of picks) {
      const f = readFill(child(child(p, 'a:tcStyle'), 'a:fill'), env);
      if (f) fill = f;
      const tx = child(p, 'a:tcTxStyle');
      const c = readColor(tx, env);
      if (c) color = colorToCss(c);
      if (tx?.attrs.b === 'on') bold = true;
    }
    return { fill, color, bold };
  };

  const cols = kids(child(tbl, 'a:tblGrid'), 'a:gridCol').map((c) => emuToPx(num(c.attrs.w) * scaleX(T)));
  const cells = [];
  const rows = kids(tbl, 'a:tr').map((tr, r) => {
    const cs = cellStyle(r);
    const tds = kids(tr, 'a:tc').map((tc, c) => {
      if (truthy(tc.attrs.hMerge) || truthy(tc.attrs.vMerge)) return '';
      const tcPr = child(tc, 'a:tcPr');
      const txBody = child(tc, 'a:txBody');
      const content = txBody ? renderTextBody(txBody, S, part, {
        idPrefix: `${part.scope}.sh${id}.r${r + 1}c${c + 1}`,
        kind: 'table',
        extraLevels: () => ({ r: { ...(cs.color ? { color: cs.color } : {}), ...(cs.bold || (firstRow && r === 0) ? { b: true } : {}) } }),
      }) : { html: '' };
      const own = readFill(tcPr, env) ?? cs.fill;
      const border = (name) => { const l = readLine(child(tcPr, name) ? { children: [child(tcPr, name)] } : null, null, env); return l ? lineCss(l) : null; };
      const css = style({
        ...(bgCss(own, S, part) ?? {}),
        'vertical-align': { ctr: 'middle', b: 'bottom' }[tcPr?.attrs.anchor] ?? 'top',
        padding: [num(tcPr?.attrs.marT, 45720), num(tcPr?.attrs.marR, 91440), num(tcPr?.attrs.marB, 45720), num(tcPr?.attrs.marL, 91440)].map((v) => px(emuToPx(v))).join(' '),
        'border-left': border('a:lnL'), 'border-right': border('a:lnR'), 'border-top': border('a:lnT'), 'border-bottom': border('a:lnB'),
      });
      const span = `${num(tc.attrs.gridSpan, 1) > 1 ? ` colspan="${tc.attrs.gridSpan}"` : ''}${num(tc.attrs.rowSpan, 1) > 1 ? ` rowspan="${tc.attrs.rowSpan}"` : ''}`;
      return `<td${span} style="${css}">${content.html}</td>`;
    });
    cells.push(tds.length);
    return `<tr style="height:${px(emuToPx(num(tr.attrs.h) * scaleY(T)))}">${tds.join('')}</tr>`;
  });
  S.stats.tables += 1;
  return `<table class="tb" data-id="${id}" style="left:${px(box.x)};top:${px(box.y)};width:${px(box.w)}"><colgroup>${cols.map((w) => `<col style="width:${px(w)}">`).join('')}</colgroup>${rows.join('')}</table>`;
}

/* --------------------------------------------------- frames and groups ---- */

function renderFrame(frame, S, part, T) {
  const graphic = path(frame, 'a:graphic', 'a:graphicData');
  const tbl = child(graphic, 'a:tbl');
  if (tbl) return renderTable(frame, tbl, S, part, T);

  const uri = graphic?.attrs.uri ?? '';
  const xf = readXfrm(child(frame, 'p:xfrm'));
  const box = xf ? place(xf, T) : null;
  const ghost = (label) => (box ? `<div class="miss" style="left:${px(box.x)};top:${px(box.y)};width:${px(box.w)};height:${px(box.h)}" title="${label}"></div>` : '');

  if (/diagram/.test(uri) && box) {
    const dm = child(graphic, 'dgm:relIds');
    const drawing = [...part.rels.values()].find((r) => r.type.endsWith('/diagramDrawing'))?.target;
    if (drawing && dm) {
      // The drawing part holds SmartArt already laid out as plain shapes.
      const raw = S.pkg.text(drawing).replace(/<(\/?)dsp:/g, '<$1p:');
      const tree = parseXml(raw);
      const inner = path(tree, 'p:spTree');
      if (inner) {
        S.stats.smartart += 1;
        return renderTree(inner, S, part, translate(T, xf.x, xf.y));
      }
    }
    warn(S, 'SmartArt could not be rendered');
    return ghost('SmartArt not rendered');
  }
  if (/chart/.test(uri)) { warn(S, 'chart not rendered (placeholder shown)'); return ghost('Chart not rendered'); }
  if (!/ole|table/i.test(uri)) warn(S, `graphic frame (${uri.split('/').pop() || 'unknown'}) not rendered`);
  return ghost('Content not rendered');
}

function renderTree(tree, S, part, T) {
  const out = [];
  const visit = (node, t) => {
    switch (node.name) {
      case 'p:sp': out.push(renderSp(node, S, part, t)); break;
      case 'p:pic': out.push(renderPic(node, S, part, t)); break;
      case 'p:cxnSp': out.push(renderConnector(node, S, part, t)); break;
      case 'p:graphicFrame': out.push(renderFrame(node, S, part, t)); break;
      case 'p:grpSp': {
        if (truthy(path(node, 'p:nvGrpSpPr', 'p:cNvPr')?.attrs.hidden)) break;
        const xf = readXfrm(child(node, 'p:grpSpPr'));
        S.stats.groups += 1;
        const inner = xf ? composeGroup(t, xf) : t;
        node.children.forEach((c) => visit(c, inner));
        break;
      }
      case 'mc:AlternateContent': {
        const pick = child(node, 'mc:Fallback') ?? child(node, 'mc:Choice');
        pick?.children.forEach((c) => visit(c, t));
        break;
      }
      default: break;
    }
  };
  tree.children.forEach((c) => visit(c, T));
  return out.filter(Boolean).join('\n');
}

/* ---------------------------------------------------- backgrounds, notes ---- */

function backgroundOf(S, sources) {
  const env = S.chain.env;
  for (const { xml, part } of sources) {
    const bg = path(xml, 'p:cSld', 'p:bg');
    if (!bg) continue;
    const bgPr = child(bg, 'p:bgPr');
    if (bgPr) {
      const css = bgCss(readFill(bgPr, env), S, part);
      if (css) return css;
    }
    const ref = child(bg, 'p:bgRef');
    if (ref) {
      const phClr = readColor(ref, env);
      const idx = num(ref.attrs.idx);
      const entry = S.chain.theme.bgFills[idx - 1001];
      const fill = entry ? fillFromElement(entry, { ...env, phClr }) : phClr ? { type: 'solid', css: colorToCss(phClr) } : null;
      const css = bgCss(fill, S, part);
      if (css) return css;
    }
  }
  return { background: '#fff' };
}

function renderNotes(S, slidePath, slideRels) {
  const notesPath = relOfType(slideRels, '/notesSlide');
  const notesXml = notesPath ? S.pkg.xml(notesPath) : null;
  const tree = path(notesXml, 'p:cSld', 'p:spTree');
  if (!tree) return '';
  const part = { rels: S.pkg.rels(notesPath), scope: `s${S.slideNo}.notes` };
  const out = [];
  for (const sp of kids(tree, 'p:sp')) {
    if (readPh(sp)?.type !== 'body') continue;
    const txBody = child(sp, 'p:txBody');
    if (!txBody) continue;
    kids(txBody, 'a:p').forEach((p, i) => {
      const tx = {
        S, part, env: S.chain.env, baseProps: () => ({}), fontScale: 1, lnReduce: 0, numbering: {}, idPrefix: part.scope, kind: 'notes',
      };
      const { html, string } = renderParagraph(p, i + 1, tx);
      if (string) { S.addString(string); out.push(html.replace(/ style="[^"]*"/, '')); }
    });
  }
  return out.length ? `<aside class="notes">${out.join('')}</aside>` : '';
}

/* --------------------------------------------------------------- entry ---- */

export function pptxToHtml(files, options = {}) {
  const pkg = openPackage(files);
  const presXml = pkg.xml('ppt/presentation.xml');
  if (!presXml) throw new Error('No slides found. Is this really a .pptx?');
  const presRels = pkg.rels('ppt/presentation.xml');
  const slidePaths = kids(child(presXml, 'p:sldIdLst'), 'p:sldId')
    .map((s) => presRels.get(s.attrs['r:id'])?.target)
    .filter((p) => p && files.has(p));
  if (!slidePaths.length) throw new Error('No slides found. Is this really a .pptx?');

  const sz = child(presXml, 'p:sldSz');
  const width = Math.round(emuToPx(num(sz?.attrs.cx, 12192000)));
  const height = Math.round(emuToPx(num(sz?.attrs.cy, 6858000)));

  const langCounts = {};
  const warnings = [];
  const strings = [];
  const images = new Map();
  const unknown = new Map();
  const stats = { slides: 0, shapes: 0, pictures: 0, groups: 0, connectors: 0, tables: 0, smartart: 0, codeParagraphs: 0, strings: 0 };
  const chainCache = new Map();
  const sections = [];

  slidePaths.forEach((slidePath, index) => {
    const slideNo = index + 1;
    const sld = pkg.xml(slidePath);
    for (const m of pkg.text(slidePath).matchAll(/\blang="([A-Za-z-]+)"/g)) langCounts[m[1]] = (langCounts[m[1]] ?? 0) + 1;

    const chain = buildChain(pkg, slidePath, chainCache);
    const S = {
      pkg, chain, slideNo, width, height, warnings, images, unknown, stats, strings,
      pres: { defaultTextStyle: child(presXml, 'p:defaultTextStyle') },
      defs: [], clips: new Map(), title: '', scopeKey: '',
      addString(s) {
        const entry = { id: s.id, slide: slideNo, kind: s.kind, text: s.text };
        if (this.title && this.title !== s.text) entry.context = this.title;
        strings.push(entry);
      },
    };

    const tree = path(sld, 'p:cSld', 'p:spTree');
    const titleSp = kids(tree, 'p:sp').find((sp) => ['title', 'ctrTitle'].includes(readPh(sp)?.type));
    S.title = descendants(titleSp ?? { children: [] }, 'a:t').map((t) => t.text).join('').trim();

    const slideRels = pkg.rels(slidePath);
    const slidePart = { rels: slideRels, scope: `s${slideNo}`, decor: false };
    const showLayout = sld.attrs.showMasterSp !== '0';
    const showMaster = showLayout && chain.layoutXml?.attrs.showMasterSp !== '0';

    const html = [];
    if (showMaster && chain.masterTree) {
      S.scopeKey = 'M';
      html.push(renderTree(chain.masterTree, S, { rels: chain.masterRels, scope: `s${slideNo}.M`, decor: true }, IDENTITY));
    }
    if (showLayout && chain.layoutTree) {
      S.scopeKey = 'L';
      html.push(renderTree(chain.layoutTree, S, { rels: chain.layoutRels, scope: `s${slideNo}.L`, decor: true }, IDENTITY));
    }
    S.scopeKey = '';
    if (tree) html.push(renderTree(tree, S, slidePart, IDENTITY));

    const bg = backgroundOf(S, [
      { xml: sld, part: slidePart },
      { xml: chain.layoutXml, part: { rels: chain.layoutRels } },
      { xml: chain.masterXml, part: { rels: chain.masterRels } },
    ]);
    const hl = chain.theme.colors.hlink ? `#${chain.theme.colors.hlink}` : null;
    const notes = renderNotes(S, slidePath, slideRels);
    const defs = S.defs.length ? `<svg class="defs" width="0" height="0" aria-hidden="true"><defs>${S.defs.join('')}</defs></svg>` : '';
    // Hidden slides stay visible here so every string can be reviewed and translated.
    const hidden = sld.attrs.show === '0' ? ' data-hidden-slide="true"' : '';
    if (hidden) warn(S, 'hidden in the original deck (shown here)');
    sections.push(`<section class="pptx-slide" data-slide="${slideNo}"${hidden}><div class="pptx-canvas" style="${style({ ...bg, width: px(width), height: px(height), '--hl': hl })}">${defs}${html.filter(Boolean).join('\n')}
</div>${notes}</section>`);
    stats.slides += 1;
  });

  for (const [prst, count] of unknown) warnings.push(`${count} shape(s) use preset "${prst}", drawn as a rectangle`);
  stats.strings = strings.length;

  const sourceLang = Object.entries(langCounts).sort((a, b) => b[1] - a[1])[0]?.[0] ?? 'en';
  const core = pkg.xml('docProps/core.xml');
  const coreTitle = child(core, 'dc:title')?.text.trim();
  const coreAuthor = child(core, 'dc:creator')?.text.trim();

  return {
    html: `${sections.join('\n')}\n`,
    strings: { version: 1, sourceLang, strings },
    images: [...images].map(([zipPath, name]) => ({ zipPath, name })),
    deck: {
      title: options.title || coreTitle || options.slug || 'Imported deck',
      description: options.description ?? 'Imported from PowerPoint.',
      author: options.author ?? coreAuthor ?? '',
      date: options.date ?? new Date().toISOString().slice(0, 10),
      tags: options.tags ?? [],
      theme: options.theme ?? 'white',
      format: 'html',
      lang: sourceLang,
      canvas: { width, height },
    },
    warnings,
    stats,
  };
}

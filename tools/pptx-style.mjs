/**
 * Colors, fills, lines, themes and text-style resolution for the PPTX to HTML
 * converter. Everything here is pure and runs in Node and the browser.
 */
import { child, kids, attrOf } from './xml-lite.mjs';

export const EMU_PER_PX = 9525;
export const emuToPx = (v) => Number(v) / EMU_PER_PX;
/** Points to CSS pixels at 96 dpi. */
export const ptToPx = (pt) => (Number(pt) * 96) / 72;
const round = (n, d = 2) => Math.round(n * 10 ** d) / 10 ** d;
export const px = (n) => `${round(n)}px`;

/* ------------------------------------------------------------- color ---- */

const clamp = (v, lo = 0, hi = 255) => Math.min(hi, Math.max(lo, v));

function rgbToHsl({ r, g, b }) {
  const rn = r / 255; const gn = g / 255; const bn = b / 255;
  const max = Math.max(rn, gn, bn); const min = Math.min(rn, gn, bn);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h;
  if (max === rn) h = (gn - bn) / d + (gn < bn ? 6 : 0);
  else if (max === gn) h = (bn - rn) / d + 2;
  else h = (rn - gn) / d + 4;
  return { h: h / 6, s, l };
}

function hslToRgb({ h, s, l }) {
  if (s === 0) { const v = Math.round(l * 255); return { r: v, g: v, b: v }; }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t) => {
    let x = t;
    if (x < 0) x += 1;
    if (x > 1) x -= 1;
    if (x < 1 / 6) return p + (q - p) * 6 * x;
    if (x < 1 / 2) return q;
    if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
    return p;
  };
  return { r: Math.round(f(h + 1 / 3) * 255), g: Math.round(f(h) * 255), b: Math.round(f(h - 1 / 3) * 255) };
}

const hexToRgb = (hex) => {
  const h = hex.replace('#', '');
  const v = h.length === 3 ? [...h].map((c) => c + c).join('') : h.padEnd(6, '0');
  return { r: parseInt(v.slice(0, 2), 16), g: parseInt(v.slice(2, 4), 16), b: parseInt(v.slice(4, 6), 16) };
};

const PRESET_COLORS = {
  black: '000000', white: 'ffffff', red: 'ff0000', green: '008000', blue: '0000ff', yellow: 'ffff00',
  gray: '808080', grey: '808080', orange: 'ffa500', purple: '800080', cyan: '00ffff', magenta: 'ff00ff',
  darkBlue: '00008b', darkGray: 'a9a9a9', lightGray: 'd3d3d3', silver: 'c0c0c0',
};

const DEFAULT_CLR_MAP = {
  bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2',
  accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4', accent5: 'accent5', accent6: 'accent6',
  hlink: 'hlink', folHlink: 'folHlink',
};

const COLOR_ELEMENTS = new Set(['a:srgbClr', 'a:schemeClr', 'a:sysClr', 'a:prstClr', 'a:scrgbClr']);

/**
 * Resolve a color element (or a parent holding one) to {r,g,b,a}.
 * `env` carries the theme, the colour map and the current placeholder colour.
 */
export function readColor(node, env) {
  if (!node) return null;
  const el = COLOR_ELEMENTS.has(node.name) ? node : node.children.find((c) => COLOR_ELEMENTS.has(c.name));
  if (!el) return null;

  let rgb = null;
  const val = attrOf(el, 'val');
  if (el.name === 'a:srgbClr') rgb = hexToRgb(val ?? '000000');
  else if (el.name === 'a:sysClr') rgb = hexToRgb(attrOf(el, 'lastClr') ?? (val === 'window' ? 'ffffff' : '000000'));
  else if (el.name === 'a:prstClr') rgb = hexToRgb(PRESET_COLORS[val] ?? '000000');
  else if (el.name === 'a:scrgbClr') {
    const ch = (n) => Math.round(clamp((Number(attrOf(el, n, 0)) / 100000) * 255));
    rgb = { r: ch('r'), g: ch('g'), b: ch('b') };
  } else if (el.name === 'a:schemeClr') {
    if (val === 'phClr') rgb = env.phClr ? { ...env.phClr } : null;
    else {
      const mapped = (env.clrMap ?? DEFAULT_CLR_MAP)[val] ?? val;
      const hex = env.theme?.colors?.[mapped] ?? env.theme?.colors?.[val];
      rgb = hex ? hexToRgb(hex) : null;
    }
  }
  if (!rgb) return null;

  let color = { ...rgb, a: 1 };
  for (const t of el.children) {
    const v = Number(attrOf(t, 'val', 0)) / 100000;
    if (t.name === 'a:alpha') color.a = v;
    else if (['a:lumMod', 'a:lumOff', 'a:satMod', 'a:satOff'].includes(t.name)) {
      const hsl = rgbToHsl(color);
      if (t.name === 'a:lumMod') hsl.l = clamp(hsl.l * v, 0, 1);
      else if (t.name === 'a:lumOff') hsl.l = clamp(hsl.l + v, 0, 1);
      else if (t.name === 'a:satMod') hsl.s = clamp(hsl.s * v, 0, 1);
      else hsl.s = clamp(hsl.s + v, 0, 1);
      color = { ...hslToRgb(hsl), a: color.a };
    } else if (t.name === 'a:tint') {
      color = { r: Math.round(color.r * v + 255 * (1 - v)), g: Math.round(color.g * v + 255 * (1 - v)), b: Math.round(color.b * v + 255 * (1 - v)), a: color.a };
    } else if (t.name === 'a:shade') {
      color = { r: Math.round(color.r * v), g: Math.round(color.g * v), b: Math.round(color.b * v), a: color.a };
    }
  }
  return color;
}

export const colorToCss = (c) => (c.a < 1
  ? `rgba(${c.r},${c.g},${c.b},${round(c.a, 3)})`
  : `#${[c.r, c.g, c.b].map((n) => n.toString(16).padStart(2, '0')).join('')}`);

/* ------------------------------------------------------------- theme ---- */

export function readTheme(themeRoot) {
  const theme = { colors: {}, major: null, minor: null, fills: [], bgFills: [], lines: [] };
  if (!themeRoot) return theme;
  const elements = child(themeRoot, 'a:themeElements');
  const scheme = child(elements, 'a:clrScheme');
  for (const c of scheme?.children ?? []) {
    const inner = c.children[0];
    if (!inner) continue;
    const hex = inner.name === 'a:sysClr' ? attrOf(inner, 'lastClr', '000000') : attrOf(inner, 'val', '000000');
    theme.colors[c.name.replace('a:', '')] = hex;
  }
  const fonts = child(elements, 'a:fontScheme');
  theme.major = attrOf(child(child(fonts, 'a:majorFont'), 'a:latin'), 'typeface');
  theme.minor = attrOf(child(child(fonts, 'a:minorFont'), 'a:latin'), 'typeface');
  const fmt = child(elements, 'a:fmtScheme');
  theme.fills = child(fmt, 'a:fillStyleLst')?.children ?? [];
  theme.bgFills = child(fmt, 'a:bgFillStyleLst')?.children ?? [];
  theme.lines = child(fmt, 'a:lnStyleLst')?.children ?? [];
  return theme;
}

/* -------------------------------------------------------------- fill ---- */

const FILL_NAMES = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'];

/** Turn one fill element into { type, css?, relId? }. */
export function fillFromElement(el, env) {
  if (!el) return null;
  if (el.name === 'a:noFill') return { type: 'none' };
  if (el.name === 'a:solidFill') {
    const c = readColor(el, env);
    return c ? { type: 'solid', css: colorToCss(c) } : null;
  }
  if (el.name === 'a:gradFill') {
    const stops = (child(el, 'a:gsLst')?.children ?? [])
      .map((gs) => ({ pos: Number(attrOf(gs, 'pos', 0)) / 1000, color: readColor(gs, env) }))
      .filter((s) => s.color)
      .sort((a, b) => a.pos - b.pos);
    if (!stops.length) return null;
    if (stops.length === 1) return { type: 'solid', css: colorToCss(stops[0].color) };
    const list = stops.map((s) => `${colorToCss(s.color)} ${round(s.pos, 1)}%`).join(', ');
    const lin = child(el, 'a:lin');
    if (child(el, 'a:path')) return { type: 'grad', css: `radial-gradient(circle at center, ${list})` };
    const angle = lin ? Number(attrOf(lin, 'ang', 0)) / 60000 : 0;
    return { type: 'grad', css: `linear-gradient(${round(angle + 90, 1)}deg, ${list})` };
  }
  if (el.name === 'a:blipFill') {
    const blip = child(el, 'a:blip');
    return { type: 'pic', relId: attrOf(blip, 'r:embed'), tile: Boolean(child(el, 'a:tile')) };
  }
  if (el.name === 'a:pattFill') {
    const fg = readColor(child(el, 'a:fgClr'), env);
    return fg ? { type: 'solid', css: colorToCss(fg) } : null;
  }
  return null;
}

/** First fill child of a shape-properties-like element, or null if unspecified. */
export function readFill(parent, env) {
  const el = parent?.children.find((c) => FILL_NAMES.includes(c.name));
  return el ? fillFromElement(el, env) : null;
}

/** Fill implied by a shape's <p:style><a:fillRef>, using the theme's fill list. */
export function styleFill(styleNode, env) {
  const ref = child(styleNode, 'a:fillRef');
  if (!ref) return null;
  const idx = Number(attrOf(ref, 'idx', 0));
  if (idx === 0) return { type: 'none' };
  const phClr = readColor(ref, env);
  const list = idx >= 1001 ? env.theme?.bgFills : env.theme?.fills;
  const entry = list?.[(idx >= 1001 ? idx - 1001 : idx - 1)];
  if (!entry) return phClr ? { type: 'solid', css: colorToCss(phClr) } : null;
  return fillFromElement(entry, { ...env, phClr });
}

/* -------------------------------------------------------------- line ---- */

/** Resolve a shape outline to { w (px), color, dash, head, tail } or null. */
export function readLine(spPr, styleNode, env) {
  const ln = child(spPr, 'a:ln');
  const ref = child(styleNode, 'a:lnRef');
  const refIdx = Number(attrOf(ref, 'idx', 0));
  const themeLine = refIdx > 0 ? env.theme?.lines?.[refIdx - 1] : null;

  let w = ln && attrOf(ln, 'w') != null ? Number(attrOf(ln, 'w')) : null;
  if (w == null && themeLine) w = Number(attrOf(themeLine, 'w', 12700));
  if (w == null) w = 12700;

  let fill = ln ? readFill(ln, env) : null;
  if (!fill && ref && refIdx > 0) {
    const phClr = readColor(ref, env);
    fill = phClr ? { type: 'solid', css: colorToCss(phClr) } : null;
  }
  if (!fill || fill.type === 'none' || fill.type === 'pic') return null;
  const color = fill.type === 'solid' ? fill.css : '#000';

  const end = (name) => {
    const e = child(ln, name);
    const type = attrOf(e, 'type', 'none');
    return type === 'none' ? null : { type, w: attrOf(e, 'w', 'med'), len: attrOf(e, 'len', 'med') };
  };
  return {
    w: Math.max(0.5, emuToPx(w)),
    color,
    dash: attrOf(child(ln, 'a:prstDash'), 'val', 'solid'),
    head: end('a:headEnd'),
    tail: end('a:tailEnd'),
  };
}

/* ------------------------------------------------------- text styles ---- */

const truthy = (v) => v === '1' || v === 'true';

/** Run properties from an a:rPr / a:defRPr / a:endParaRPr element. */
export function readRunProps(node, env) {
  const out = {};
  if (!node) return out;
  const a = node.attrs;
  if (a.sz) out.sz = Number(a.sz);
  if (a.b != null) out.b = truthy(a.b);
  if (a.i != null) out.i = truthy(a.i);
  if (a.u != null) out.u = a.u !== 'none';
  if (a.strike != null) out.strike = a.strike !== 'noStrike';
  if (a.baseline != null) out.baseline = Number(a.baseline);
  if (a.cap != null) out.cap = a.cap;
  if (a.spc != null) out.spc = Number(a.spc);
  const fill = readFill(node, env);
  if (fill?.type === 'solid') out.color = fill.css;
  const latin = child(node, 'a:latin');
  if (latin?.attrs.typeface) out.font = latin.attrs.typeface;
  const ea = child(node, 'a:ea');
  if (ea?.attrs.typeface) out.eaFont = ea.attrs.typeface;
  const link = child(node, 'a:hlinkClick');
  if (link?.attrs['r:id']) out.link = link.attrs['r:id'];
  return out;
}

const spacing = (node) => {
  const pct = child(node, 'a:spcPct');
  if (pct) return { pct: Number(pct.attrs.val) / 100000 };
  const pts = child(node, 'a:spcPts');
  if (pts) return { pt: Number(pts.attrs.val) / 100 };
  return undefined;
};

/** Paragraph properties from an a:lvlNpPr / a:pPr element. */
export function readParaProps(node, env) {
  const out = {};
  if (!node) return out;
  const a = node.attrs;
  if (a.algn) out.algn = a.algn;
  if (a.marL != null) out.marL = Number(a.marL);
  if (a.indent != null) out.indent = Number(a.indent);
  if (a.rtl != null) out.rtl = truthy(a.rtl);
  const ln = child(node, 'a:lnSpc');
  if (ln) out.lnSpc = spacing(ln);
  const bef = child(node, 'a:spcBef');
  if (bef) out.spcBef = spacing(bef);
  const aft = child(node, 'a:spcAft');
  if (aft) out.spcAft = spacing(aft);

  if (child(node, 'a:buNone')) out.bu = { type: 'none' };
  else if (child(node, 'a:buChar')) out.bu = { type: 'char', char: attrOf(child(node, 'a:buChar'), 'char', '•') };
  else if (child(node, 'a:buAutoNum')) {
    const n = child(node, 'a:buAutoNum');
    out.bu = { type: 'auto', scheme: attrOf(n, 'type', 'arabicPeriod'), startAt: Number(attrOf(n, 'startAt', 1)) };
  } else if (child(node, 'a:buBlip')) out.bu = { type: 'char', char: '•' };

  const buFont = child(node, 'a:buFont');
  if (buFont?.attrs.typeface) out.buFont = buFont.attrs.typeface;
  const buClr = child(node, 'a:buClr');
  if (buClr) { const c = readColor(buClr, env); if (c) out.buColor = colorToCss(c); }
  const buSz = child(node, 'a:buSzPct');
  if (buSz) out.buSize = Number(buSz.attrs.val) / 100000;

  const def = child(node, 'a:defRPr');
  if (def) out.r = readRunProps(def, env);
  return out;
}

/** Overlay `over` on `base`, merging the nested run-props object. */
export function mergeProps(base, over) {
  if (!over) return base;
  const out = { ...base };
  for (const [k, v] of Object.entries(over)) {
    if (v === undefined) continue;
    out[k] = k === 'r' ? { ...(base.r ?? {}), ...v } : v;
  }
  return out;
}

/** Level N (0-based) properties from an a:lstStyle-like element. */
export function levelProps(listStyle, level, env) {
  return readParaProps(child(listStyle, `a:lvl${level + 1}pPr`), env);
}

/* ------------------------------------------------------------- fonts ---- */

export const MONO_FONT = /^(cascadia|consolas|courier|menlo|monaco|lucida console|fira (code|mono)|source code pro|jetbrains mono|sf mono|roboto mono|ibm plex mono|ubuntu mono|dejavu sans mono|liberation mono)/i;

const quote = (name) => `'${name.replace(/'/g, '')}'`;

/** A CSS font stack for a PPTX typeface, resolving theme fonts. */
export function fontStack(typeface, theme) {
  let name = typeface;
  if (name === '+mj-lt') name = theme?.major;
  else if (name === '+mn-lt') name = theme?.minor;
  if (!name) return null;
  if (MONO_FONT.test(name)) return `${quote(name)}, Consolas, 'Cascadia Mono', Menlo, monospace`;
  return `${quote(name)}, 'Segoe UI', system-ui, -apple-system, 'Helvetica Neue', Arial, sans-serif`;
}

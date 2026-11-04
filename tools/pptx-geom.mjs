/**
 * Preset geometry for the PPTX to HTML converter.
 *
 * Shapes come back in one of two forms:
 *   { kind: 'css', radius }  a rectangle or ellipse the browser can draw itself
 *   { kind: 'path', d, open } a path in unit coordinates (0..1), used as a
 *                             clip-path for the fill and as the outline
 *
 * Presets are approximations of the ECMA-376 definitions. They cover what real
 * decks use most; anything else returns null and the caller falls back to a
 * plain rectangle and says so in the import report.
 */

const KAPPA = 0.5522847498;

/** Collects path commands in pixel space, then normalizes to the unit square. */
class Path {
  constructor() { this.cmds = []; }

  M(x, y) { this.cmds.push(['M', x, y]); return this; }

  L(x, y) { this.cmds.push(['L', x, y]); return this; }

  Q(cx, cy, x, y) { this.cmds.push(['Q', cx, cy, x, y]); return this; }

  C(a, b, c, d, x, y) { this.cmds.push(['C', a, b, c, d, x, y]); return this; }

  Z() { this.cmds.push(['Z']); return this; }

  poly(points, close = true) {
    points.forEach(([x, y], i) => (i ? this.L(x, y) : this.M(x, y)));
    return close ? this.Z() : this;
  }

  /** Normalized to 0..1; `w`/`h` are the shape size in px. */
  toUnit(w, h) {
    const f = (n) => Math.round(n * 1e4) / 1e4;
    return this.cmds.map(([c, ...n]) => c + n.map((v, i) => f(v / (i % 2 === 0 ? w : h))).join(' ')).join(' ');
  }

  /** Absolute pixel path string. */
  toPx() {
    const f = (n) => Math.round(n * 100) / 100;
    return this.cmds.map(([c, ...n]) => c + n.map(f).join(' ')).join(' ');
  }
}

const gd = (adj, name, fallback) => (adj[name] != null ? adj[name] : fallback);

/** Read <a:avLst> guides into { adj: 50000, ... }. */
export function readAdjust(prstGeom) {
  const out = {};
  const list = prstGeom?.children.find((c) => c.name === 'a:avLst');
  for (const g of list?.children ?? []) {
    const m = /val\s+(-?\d+)/.exec(g.attrs.fmla ?? '');
    if (m && g.attrs.name) out[g.attrs.name] = Number(m[1]);
  }
  return out;
}

export function presetGeometry(prst, adj, w, h) {
  const ss = Math.min(w, h);
  const hw = w / 2;
  const hh = h / 2;
  const P = () => new Path();
  const unit = (p, open = false) => ({ kind: 'path', d: p.toUnit(w, h), open });

  switch (prst) {
    case 'rect':
    case 'flowChartProcess':
    case 'flowChartPredefinedProcess':
    case 'flowChartInternalStorage':
    case 'flowChartDocument':
    case 'flowChartMultidocument':
    case 'flowChartMagneticDisk':
    case 'flowChartPunchedCard':
      return { kind: 'css', radius: null };
    case 'roundRect':
      return { kind: 'css', radius: (ss * gd(adj, 'adj', 16667)) / 100000 };
    case 'flowChartAlternateProcess':
      return { kind: 'css', radius: ss / 6 };
    case 'flowChartTerminator':
      return { kind: 'css', radius: ss / 2 };
    case 'ellipse':
    case 'flowChartConnector':
      return { kind: 'css', radius: '50%' };

    case 'triangle': {
      const x = (w * gd(adj, 'adj', 50000)) / 100000;
      return unit(P().poly([[x, 0], [w, h], [0, h]]));
    }
    case 'rtTriangle': return unit(P().poly([[0, 0], [0, h], [w, h]]));
    case 'diamond':
    case 'flowChartDecision':
      return unit(P().poly([[hw, 0], [w, hh], [hw, h], [0, hh]]));
    case 'parallelogram':
    case 'flowChartInputOutput': {
      const x = (ss * gd(adj, 'adj', 25000)) / 100000;
      return unit(P().poly([[0, h], [x, 0], [w, 0], [w - x, h]]));
    }
    case 'trapezoid': {
      const x = (ss * gd(adj, 'adj', 25000)) / 100000;
      return unit(P().poly([[0, h], [x, 0], [w - x, 0], [w, h]]));
    }
    case 'homePlate': {
      const dx = (ss * gd(adj, 'adj', 50000)) / 100000;
      return unit(P().poly([[0, 0], [w - dx, 0], [w, hh], [w - dx, h], [0, h]]));
    }
    case 'chevron': {
      const dx = (ss * gd(adj, 'adj', 50000)) / 100000;
      return unit(P().poly([[0, 0], [w - dx, 0], [w, hh], [w - dx, h], [0, h], [dx, hh]]));
    }
    case 'hexagon': {
      const dx = (ss * gd(adj, 'adj', 25000)) / 100000;
      return unit(P().poly([[0, hh], [dx, 0], [w - dx, 0], [w, hh], [w - dx, h], [dx, h]]));
    }
    case 'octagon': {
      const d = (ss * gd(adj, 'adj', 29289)) / 100000;
      return unit(P().poly([[d, 0], [w - d, 0], [w, d], [w, h - d], [w - d, h], [d, h], [0, h - d], [0, d]]));
    }
    case 'plus': {
      const d = (ss * gd(adj, 'adj', 25000)) / 100000;
      return unit(P().poly([[d, 0], [w - d, 0], [w - d, d], [w, d], [w, h - d], [w - d, h - d], [w - d, h], [d, h], [d, h - d], [0, h - d], [0, d], [d, d]]));
    }

    case 'rightArrow': {
      const dx = (ss * gd(adj, 'adj2', 50000)) / 100000;
      const dy = (h * gd(adj, 'adj1', 50000)) / 200000;
      const x1 = w - dx;
      return unit(P().poly([[0, hh - dy], [x1, hh - dy], [x1, 0], [w, hh], [x1, h], [x1, hh + dy], [0, hh + dy]]));
    }
    case 'leftArrow': {
      const dx = (ss * gd(adj, 'adj2', 50000)) / 100000;
      const dy = (h * gd(adj, 'adj1', 50000)) / 200000;
      return unit(P().poly([[w, hh - dy], [dx, hh - dy], [dx, 0], [0, hh], [dx, h], [dx, hh + dy], [w, hh + dy]]));
    }
    case 'upArrow': {
      const dy = (ss * gd(adj, 'adj2', 50000)) / 100000;
      const dx = (w * gd(adj, 'adj1', 50000)) / 200000;
      return unit(P().poly([[hw, 0], [w, dy], [hw + dx, dy], [hw + dx, h], [hw - dx, h], [hw - dx, dy], [0, dy]]));
    }
    case 'downArrow': {
      const dy = (ss * gd(adj, 'adj2', 50000)) / 100000;
      const dx = (w * gd(adj, 'adj1', 50000)) / 200000;
      return unit(P().poly([[hw - dx, 0], [hw + dx, 0], [hw + dx, h - dy], [w, h - dy], [hw, h], [0, h - dy], [hw - dx, h - dy]]));
    }
    case 'leftRightArrow': {
      const dx = (ss * gd(adj, 'adj2', 50000)) / 100000;
      const dy = (h * gd(adj, 'adj1', 50000)) / 200000;
      return unit(P().poly([[0, hh], [dx, 0], [dx, hh - dy], [w - dx, hh - dy], [w - dx, 0], [w, hh], [w - dx, h], [w - dx, hh + dy], [dx, hh + dy], [dx, h]]));
    }

    case 'flowChartOffpageConnector': return unit(P().poly([[0, 0], [w, 0], [w, h * 0.8], [hw, h], [0, h * 0.8]]));
    case 'flowChartDelay':
      return unit(P().M(0, 0).L(hw, 0).C(hw + hw * KAPPA, 0, w, hh - hh * KAPPA, w, hh)
        .C(w, hh + hh * KAPPA, hw + hw * KAPPA, h, hw, h).L(0, h).Z());

    case 'round1Rect': {
      const r = (ss * gd(adj, 'adj', 16667)) / 100000;
      return unit(P().M(0, 0).L(w - r, 0).Q(w, 0, w, r).L(w, h).L(0, h).Z());
    }
    case 'round2SameRect': {
      const r1 = (ss * gd(adj, 'adj1', 16667)) / 100000;
      const r2 = (ss * gd(adj, 'adj2', 0)) / 100000;
      return unit(P().M(r1, 0).L(w - r1, 0).Q(w, 0, w, r1).L(w, h - r2).Q(w, h, w - r2, h)
        .L(r2, h).Q(0, h, 0, h - r2).L(0, r1).Q(0, 0, r1, 0).Z());
    }
    case 'snip1Rect': {
      const d = (ss * gd(adj, 'adj', 16667)) / 100000;
      return unit(P().poly([[0, 0], [w - d, 0], [w, d], [w, h], [0, h]]));
    }
    case 'snip2SameRect': {
      const d1 = (ss * gd(adj, 'adj1', 16667)) / 100000;
      const d2 = (ss * gd(adj, 'adj2', 0)) / 100000;
      return unit(P().poly([[d1, 0], [w - d1, 0], [w, d1], [w, h - d2], [w - d2, h], [d2, h], [0, h - d2], [0, d1]]));
    }
    case 'cube': {
      const d = (ss * gd(adj, 'adj', 25000)) / 100000;
      return unit(P().poly([[0, d], [d, 0], [w, 0], [w, h - d], [w - d, h], [0, h]]));
    }
    case 'can': {
      const d = (h * gd(adj, 'adj', 25000)) / 200000;
      return unit(P().M(0, d).Q(0, 0, hw, 0).Q(w, 0, w, d).L(w, h - d).Q(w, h, hw, h).Q(0, h, 0, h - d).Z());
    }
    case 'star4': {
      const k = (gd(adj, 'adj', 12500) / 50000) * 0.7071;
      const ix = hw * k;
      const iy = hh * k;
      return unit(P().poly([[0, hh], [hw - ix, hh - iy], [hw, 0], [hw + ix, hh - iy], [w, hh], [hw + ix, hh + iy], [hw, h], [hw - ix, hh + iy]]));
    }

    case 'rightBrace':
    case 'leftBrace': {
      const r = Math.min((ss * gd(adj, 'adj1', 8333)) / 100000, hh);
      const ym = (h * gd(adj, 'adj2', 50000)) / 100000;
      const mx = (x) => (prst === 'leftBrace' ? w - x : x);
      return unit(P().M(mx(0), 0).Q(mx(hw), 0, mx(hw), r).L(mx(hw), ym - r).Q(mx(hw), ym, mx(w), ym)
        .Q(mx(hw), ym, mx(hw), ym + r).L(mx(hw), h - r).Q(mx(hw), h, mx(0), h), true);
    }
    case 'leftBracket': return unit(P().poly([[w, 0], [0, 0], [0, h], [w, h]], false), true);
    case 'rightBracket': return unit(P().poly([[0, 0], [w, 0], [w, h], [0, h]], false), true);

    case 'wedgeRectCallout':
    case 'wedgeRoundRectCallout': {
      const xt = hw + (w * gd(adj, 'adj1', -20833)) / 100000;
      const yt = hh + (h * gd(adj, 'adj2', 62500)) / 100000;
      const p = P().M(0, 0);
      const vertical = Math.abs((yt - hh) / h) >= Math.abs((xt - hw) / w);
      const left = xt < hw;
      const above = yt < hh;
      const x1 = left ? w / 6 : (w * 7) / 12;
      const x2 = left ? (w * 5) / 12 : (w * 5) / 6;
      const y1 = above ? h / 6 : (h * 7) / 12;
      const y2 = above ? (h * 5) / 12 : (h * 5) / 6;
      if (vertical && above) p.L(x1, 0).L(xt, yt).L(x2, 0);
      p.L(w, 0);
      if (!vertical && !left) p.L(w, y1).L(xt, yt).L(w, y2);
      p.L(w, h);
      if (vertical && !above) p.L(x2, h).L(xt, yt).L(x1, h);
      p.L(0, h);
      if (!vertical && left) p.L(0, y2).L(xt, yt).L(0, y1);
      return unit(p.Z());
    }
    default:
      return null;
  }
}

/**
 * Connector (and plain line) path in pixel space for a w x h box, drawn from the
 * top-left to the bottom-right. The caller mirrors it for flipH/flipV.
 */
export function connectorPath(prst, adj, w, h) {
  const a1 = gd(adj, 'adj1', 50000) / 100000;
  const a2 = gd(adj, 'adj2', 50000) / 100000;
  const a3 = gd(adj, 'adj3', 50000) / 100000;
  const p = new Path();
  switch (prst) {
    case 'bentConnector2': return p.M(0, 0).L(w, 0).L(w, h).toPx();
    case 'bentConnector3': return p.M(0, 0).L(w * a1, 0).L(w * a1, h).L(w, h).toPx();
    case 'bentConnector4': return p.M(0, 0).L(w * a1, 0).L(w * a1, h * a2).L(w, h * a2).L(w, h).toPx();
    case 'bentConnector5': return p.M(0, 0).L(w * a1, 0).L(w * a1, h * a2).L(w * a3, h * a2).L(w * a3, h).L(w, h).toPx();
    case 'curvedConnector2': return p.M(0, 0).C(w / 2, 0, w, h / 2, w, h).toPx();
    case 'curvedConnector3': return p.M(0, 0).C(w * a1, 0, w * a1, h / 2, w * a1, h / 2).C(w * a1, h / 2, w * a1, h, w, h).toPx();
    default: return p.M(0, 0).L(w, h).toPx();
  }
}

/** Unit-square path for a:custGeom, built from its first drawable path(s). */
export function customGeometry(custGeom, w, h) {
  const list = custGeom.children.find((c) => c.name === 'a:pathLst');
  const parts = [];
  let open = true;
  for (const path of list?.children ?? []) {
    if (path.name !== 'a:path' || path.attrs.fill === 'none' && path.attrs.stroke === '0') continue;
    const pw = Number(path.attrs.w) || w;
    const ph = Number(path.attrs.h) || h;
    const pt = (n) => `${Math.round((Number(n.attrs.x) / pw) * 1e4) / 1e4} ${Math.round((Number(n.attrs.y) / ph) * 1e4) / 1e4}`;
    for (const cmd of path.children) {
      const pts = cmd.children.filter((c) => c.name === 'a:pt');
      if (cmd.name === 'a:moveTo') parts.push(`M${pt(pts[0])}`);
      else if (cmd.name === 'a:lnTo') parts.push(`L${pt(pts[0])}`);
      else if (cmd.name === 'a:cubicBezTo') parts.push(`C${pts.map(pt).join(' ')}`);
      else if (cmd.name === 'a:quadBezTo') parts.push(`Q${pts.map(pt).join(' ')}`);
      else if (cmd.name === 'a:close') { parts.push('Z'); open = false; }
    }
  }
  return parts.length ? { kind: 'path', d: parts.join(' '), open } : null;
}

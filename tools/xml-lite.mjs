/**
 * A small XML parser for OOXML parts. Runs unchanged in Node and the browser
 * (no DOMParser), and returns a plain tree that is cheap to walk:
 *
 *   { name: 'a:p', attrs: { ... }, children: [...elements], text: 'own text' }
 *
 * Prefixes are kept as written. PowerPoint always uses the same ones (p:, a:,
 * r:, mc:), and every part is read through them.
 */

const ENTITIES = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

export const decodeEntities = (s) => (s.indexOf('&') < 0 ? s : s.replace(
  /&(#x[0-9a-f]+|#\d+|lt|gt|amp|quot|apos);/gi,
  (m, e) => {
    if (e[0] !== '#') return ENTITIES[e.toLowerCase()];
    const code = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : Number(e.slice(1));
    return Number.isFinite(code) && code <= 0x10ffff ? String.fromCodePoint(code) : m;
  },
));

const ATTR = /([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

function parseAttrs(source) {
  const attrs = {};
  ATTR.lastIndex = 0;
  let m;
  while ((m = ATTR.exec(source))) attrs[m[1]] = decodeEntities(m[2] ?? m[3]);
  return attrs;
}

export function parseXml(xml) {
  const root = { name: '#root', attrs: {}, children: [], text: '' };
  const stack = [root];
  let i = 0;
  const len = xml.length;

  while (i < len) {
    const lt = xml.indexOf('<', i);
    const top = stack[stack.length - 1];
    if (lt < 0) { top.text += decodeEntities(xml.slice(i)); break; }
    if (lt > i) top.text += decodeEntities(xml.slice(i, lt));

    if (xml.startsWith('<!--', lt)) {
      const end = xml.indexOf('-->', lt + 4);
      i = end < 0 ? len : end + 3;
    } else if (xml.startsWith('<![CDATA[', lt)) {
      const end = xml.indexOf(']]>', lt + 9);
      top.text += xml.slice(lt + 9, end < 0 ? len : end);
      i = end < 0 ? len : end + 3;
    } else if (xml[lt + 1] === '?' || xml[lt + 1] === '!') {
      const end = xml.indexOf('>', lt + 2);
      i = end < 0 ? len : end + 1;
    } else if (xml[lt + 1] === '/') {
      const end = xml.indexOf('>', lt + 2);
      if (stack.length > 1) stack.pop();
      i = end < 0 ? len : end + 1;
    } else {
      // Find the end of the tag, skipping any '>' inside a quoted attribute.
      let j = lt + 1;
      let quote = '';
      while (j < len) {
        const c = xml[j];
        if (quote) { if (c === quote) quote = ''; } else if (c === '"' || c === "'") quote = c;
        else if (c === '>') break;
        j++;
      }
      const selfClosing = xml[j - 1] === '/';
      const body = xml.slice(lt + 1, selfClosing ? j - 1 : j);
      const space = body.search(/\s/);
      const name = space < 0 ? body : body.slice(0, space);
      const node = { name, attrs: space < 0 ? {} : parseAttrs(body.slice(space)), children: [], text: '' };
      top.children.push(node);
      if (!selfClosing) stack.push(node);
      i = j + 1;
    }
  }
  return root.children[0] ?? root;
}

/** First direct child with this name. */
export const child = (node, name) => node?.children.find((c) => c.name === name) ?? null;

/** Every direct child with this name. */
export const kids = (node, name) => (node ? node.children.filter((c) => c.name === name) : []);

/** Follow a path of child names: path(n, 'p:nvSpPr', 'p:cNvPr'). */
export function path(node, ...names) {
  let cur = node;
  for (const n of names) {
    cur = child(cur, n);
    if (!cur) return null;
  }
  return cur;
}

/** Depth-first search for every descendant with this name. */
export function descendants(node, name, out = []) {
  for (const c of node?.children ?? []) {
    if (c.name === name) out.push(c);
    descendants(c, name, out);
  }
  return out;
}

export const attrOf = (node, name, fallback = null) => node?.attrs[name] ?? fallback;

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseXml, child, kids, path } from '../tools/xml-lite.mjs';
import { listZipEntries } from '../tools/pptx-core.mjs';
import { makeZip } from '../tools/zip-tools.mjs';
import { pptxToHtml } from '../tools/pptx-html.mjs';
import { presetGeometry } from '../tools/pptx-geom.mjs';

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const enc = new TextEncoder();

/** Builds a minimal but valid .pptx package as the Map pptxToHtml expects. */
async function buildPptx(parts) {
  const blob = makeZip(Object.entries(parts).map(([name, text]) => ({ name, bytes: enc.encode(text) })));
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const files = new Map();
  for (const entry of listZipEntries(bytes)) {
    files.set(entry.name, bytes.subarray(entry.dataStart, entry.dataStart + entry.compSize));
  }
  return files;
}

const rels = (items) => `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.map(([id, type, target, ext]) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"${ext ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`;

const slide = (shapes) => `<p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes}</p:spTree></p:cSld></p:sld>`;

const textBox = (id, x, y, cx, cy, paragraphs, hidden = false) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Box ${id}"${hidden ? ' hidden="1"' : ''}/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;

const para = (...runs) => `<a:p>${runs.join('')}</a:p>`;
const run = (text, attrs = '', inner = '') => `<a:r><a:rPr lang="en-US" ${attrs}>${inner}</a:rPr><a:t>${text}</a:t></a:r>`;

async function sample() {
  const shapes = [
    textBox(2, 914400, 914400, 4572000, 914400, para(run('Hello '), run('bold &amp; &lt;tags&gt;', 'b="1"'), run(' end'))),
    textBox(3, 914400, 2286000, 4572000, 914400, para(run('x = 1', '', '<a:latin typeface="Consolas"/>'))),
    textBox(4, 914400, 3200400, 4572000, 457200, para(run('See '), run('docs', '', '<a:hlinkClick r:id="rId9"/>'))),
    `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="5400000"><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/><a:chOff x="0" y="0"/><a:chExt cx="1000000" cy="500000"/></a:xfrm></p:grpSpPr>${textBox(6, 0, 0, 1000000, 500000, para(run('Rotated')))}</p:grpSp>`,
    textBox(7, 0, 0, 100, 100, para(run('Hidden text')), true),
  ].join('');
  return buildPptx({
    'ppt/presentation.xml': `<p:presentation ${NS}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`,
    'ppt/_rels/presentation.xml.rels': rels([['rId1', 'slide', 'slides/slide1.xml']]),
    'ppt/slides/slide1.xml': slide(shapes),
    'ppt/slides/_rels/slide1.xml.rels': rels([['rId9', 'hyperlink', 'https://example.com/docs', true], ['rId8', 'notesSlide', '../notesSlides/notesSlide1.xml']]),
    'ppt/notesSlides/notesSlide1.xml': `<p:notes ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${para(run('Say this first.'))}</p:txBody></p:sp></p:spTree></p:cSld></p:notes>`,
  });
}

test('xml-lite keeps nesting, attributes and entities', () => {
  const root = parseXml('<?xml version="1.0"?><a:p x="1 &amp; 2"><a:r><a:t> hi &lt;there&gt; </a:t></a:r><a:br/></a:p>');
  assert.equal(root.name, 'a:p');
  assert.equal(root.attrs.x, '1 & 2');
  assert.equal(path(root, 'a:r', 'a:t').text, ' hi <there> ');
  assert.equal(kids(root, 'a:br').length, 1);
  assert.equal(child(root, 'a:missing'), null);
});

test('every paragraph becomes a string with a stable id, in document order', async () => {
  const { strings, warnings } = pptxToHtml(await sample());
  assert.deepEqual(strings.strings.map((s) => s.id), ['s1.sh2.p1', 's1.sh4.p1', 's1.sh6.p1', 's1.notes.p1']);
  assert.equal(strings.version, 1);
  assert.equal(strings.sourceLang, 'en-US');
  assert.deepEqual(warnings, []);
});

test('text is kept verbatim and escaped, with only contract tags inline', async () => {
  const { strings } = pptxToHtml(await sample());
  const byId = Object.fromEntries(strings.strings.map((s) => [s.id, s]));
  assert.equal(byId['s1.sh2.p1'].text, 'Hello <b>bold &amp; &lt;tags&gt;</b> end');
  assert.equal(byId['s1.sh4.p1'].text, 'See <a href="https://example.com/docs">docs</a>');
  assert.equal(byId['s1.notes.p1'].text, 'Say this first.');
  assert.equal(byId['s1.notes.p1'].kind, 'notes');
});

test('monospace paragraphs and hidden shapes stay out of the table', async () => {
  const { html, strings, stats } = pptxToHtml(await sample());
  assert.ok(!strings.strings.some((s) => s.text.includes('x = 1')));
  assert.ok(html.includes('x = 1'), 'code is still in the slide');
  assert.ok(html.includes('data-nt'), 'code paragraphs are marked no-translate');
  assert.ok(!html.includes('Hidden text'));
  assert.equal(stats.codeParagraphs, 1);
});

test('every translatable paragraph carries data-i18n matching the table', async () => {
  const { html, strings } = pptxToHtml(await sample());
  for (const s of strings.strings) assert.ok(html.includes(`data-i18n="${s.id}"`), s.id);
});

test('group rotation reaches the shapes inside it', async () => {
  const { html } = pptxToHtml(await sample());
  assert.match(html, /data-id="6"[^>]*transform:rotate\(90deg\)/);
});

test('slide canvas follows the deck size', async () => {
  const { html, deck } = pptxToHtml(await sample());
  assert.deepEqual(deck.canvas, { width: 1280, height: 720 });
  assert.match(html, /class="pptx-canvas" style="[^"]*width:1280px;height:720px/);
});

test('preset geometry returns unit paths and falls back to null when unknown', () => {
  const arrow = presetGeometry('rightArrow', {}, 200, 100);
  assert.equal(arrow.kind, 'path');
  assert.ok(arrow.d.startsWith('M'));
  assert.equal(presetGeometry('roundRect', {}, 200, 100).kind, 'css');
  assert.equal(presetGeometry('noSuchShape', {}, 200, 100), null);
});

test('rejects files that are not a pptx', () => {
  assert.throws(() => pptxToHtml(new Map()), /Is this really a \.pptx/);
});

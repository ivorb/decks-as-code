// Pure(ish) translation pipeline for decks/<slug>/strings.json. See docs/string-table.md.
// No I/O except translateLanguage's optional logger; the CLI owns files and exit codes.

const ALLOWED_TAGS = new Set(['b', 'i', 'u', 'sup', 'sub', 'br', 'a', 'code']);
const DEFAULT_MAX_ITEMS = 30;
const DEFAULT_MAX_CHARS = 6000;

// FNV-1a 32-bit over UTF-8 bytes, lowercase hex, 8 chars.
export function fnv1a(text) {
  let hash = 0x811c9dc5;
  for (const byte of Buffer.from(String(text), 'utf8')) {
    hash ^= byte;
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return hash.toString(16).padStart(8, '0');
}

// ---- inline markup validation -------------------------------------------------

const TAG_RE = /<(\/?)([a-zA-Z][a-zA-Z0-9]*)([^<>]*)>/g;
const HREF_RE = /^\s+href\s*=\s*(?:"([^"]*)"|'([^']*)')\s*$/;

// Returns { tags, hrefs, errors } for an HTML fragment. `tags` entries look like
// "b", "/b" or "br", sorted, so two fragments can be compared as multisets.
export function analyzeMarkup(html) {
  const tags = [];
  const hrefs = [];
  const errors = [];
  const stack = [];

  for (const match of html.matchAll(TAG_RE)) {
    const closing = match[1] === '/';
    const name = match[2].toLowerCase();
    const attrs = match[3];

    if (!ALLOWED_TAGS.has(name)) {
      errors.push(`disallowed tag <${closing ? '/' : ''}${match[2]}>`);
      continue;
    }
    if (name === 'br') {
      if (closing) errors.push('</br> is not valid');
      else if (!/^\s*\/?\s*$/.test(attrs)) errors.push('<br> must not have attributes');
      else tags.push('br');
      continue;
    }
    if (closing) {
      if (attrs.trim()) errors.push(`</${name}> must not have attributes`);
      tags.push(`/${name}`);
      if (stack.pop() !== name) errors.push(`unbalanced or misnested </${name}>`);
      continue;
    }
    if (name === 'a') {
      const href = HREF_RE.exec(attrs);
      if (!href) errors.push('<a> must have exactly one attribute: href="..."');
      else hrefs.push(href[1] ?? href[2]);
    } else if (attrs.trim()) {
      errors.push(`<${name}> must not have attributes`);
    }
    tags.push(name);
    stack.push(name);
  }
  if (stack.length) errors.push(`unclosed tag(s): ${stack.map((t) => `<${t}>`).join('')}`);

  if (html.replace(TAG_RE, '').includes('<')) errors.push('contains an unescaped "<" (use &lt;)');

  return { tags: tags.sort(), hrefs: hrefs.sort(), errors };
}

function countBy(list) {
  const counts = new Map();
  for (const item of list) counts.set(item, (counts.get(item) ?? 0) + 1);
  return counts;
}

function diffCounts(expected, actual, label) {
  const errors = [];
  const e = countBy(expected);
  const a = countBy(actual);
  for (const key of new Set([...e.keys(), ...a.keys()])) {
    if ((e.get(key) ?? 0) !== (a.get(key) ?? 0)) {
      errors.push(`${label} ${JSON.stringify(key)}: expected ${e.get(key) ?? 0}, got ${a.get(key) ?? 0}`);
    }
  }
  return errors;
}

// Returns a list of problems; empty means the translation is acceptable.
export function validateTranslation(sourceText, translated) {
  if (typeof translated !== 'string' || !translated.trim()) return ['text is empty or not a string'];

  const source = analyzeMarkup(sourceText);
  const result = analyzeMarkup(translated);
  const errors = [];

  // Only enforce well-formedness when the source itself is well-formed.
  if (!source.errors.length) errors.push(...result.errors);
  errors.push(...diffCounts(source.tags, result.tags, 'tag'));
  errors.push(...diffCounts(source.hrefs, result.hrefs, 'href'));
  return errors;
}

// ---- planning and batching ----------------------------------------------------

// Decides which source strings need translating.
// Returns { kept: Map<id, {text, src}>, todo: SourceString[], dropped: string[] }.
export function planLanguage(table, existing, { force = false } = {}) {
  const kept = new Map();
  const todo = [];
  const prior = existing?.strings ?? {};

  for (const item of table.strings) {
    if (!item.text.trim()) continue;
    const src = fnv1a(item.text);
    const entry = prior[item.id];
    const reusable =
      !force &&
      entry &&
      entry.src === src &&
      typeof entry.text === 'string' &&
      validateTranslation(item.text, entry.text).length === 0;
    if (reusable) kept.set(item.id, { text: entry.text, src });
    else todo.push(item);
  }
  const sourceIds = new Set(table.strings.map((s) => s.id));
  const dropped = Object.keys(prior).filter((id) => !sourceIds.has(id));
  return { kept, todo, dropped };
}

// Groups consecutive strings (document order keeps a slide together) into batches.
export function makeBatches(items, { maxItems = DEFAULT_MAX_ITEMS, maxChars = DEFAULT_MAX_CHARS } = {}) {
  const batches = [];
  let current = [];
  let chars = 0;
  for (const item of items) {
    const size = item.text.length + (item.context?.length ?? 0) + item.id.length;
    if (current.length && (current.length >= maxItems || chars + size > maxChars)) {
      batches.push(current);
      current = [];
      chars = 0;
    }
    current.push(item);
    chars += size;
  }
  if (current.length) batches.push(current);
  return batches;
}

// ---- prompting ----------------------------------------------------------------

function languageName(code) {
  try {
    return new Intl.DisplayNames(['en'], { type: 'language' }).of(code) ?? code;
  } catch {
    return code;
  }
}

function payloadOf(items, sourceLang, targetLang) {
  return {
    sourceLang,
    targetLang,
    strings: items.map(({ id, slide, kind, context, text }) => {
      const out = { id, slide, kind };
      if (context) out.context = context;
      out.text = text;
      return out;
    }),
  };
}

function systemPrompt(sourceLang, targetLang) {
  return [
    `You are a professional translator localizing slide deck text from ${languageName(sourceLang)} (${sourceLang}) into ${languageName(targetLang)} (${targetLang}).`,
    'You receive a JSON object whose "strings" array holds slide text. Each item has: id, slide (1-based), kind, optional context (usually the slide title) and text.',
    'Use slide, context and neighbouring items to pick the right sense of short fragments (e.g. "Save", "Tool"), but translate only the "text" of each item.',
    'Kinds: "text" = slide text; "table" = table cell (keep it short, match the column style); "alt" = image alt text (a concise description for screen readers); "notes" = speaker notes (natural spoken style).',
    'The text is an HTML fragment. Keep every tag (<b>, <i>, <u>, <sup>, <sub>, <code>, <br>, <a href="...">) and close each one; move tags to follow the new word order if needed. Never translate the contents of <code>.',
    'Never change, translate or drop an href value. Keep entities (&lt; &gt; &amp;) escaped. Do not add tags, markdown or commentary.',
    'Leave product names, code, URLs and untranslatable proper nouns as they are. Do not leave a string empty.',
    'Reply with JSON only, exactly: {"translations": {"<id>": "<translated text>", ...}} with one entry for every requested id and no other ids.',
  ].join('\n');
}

export function buildMessages(items, sourceLang, targetLang) {
  return [
    { role: 'system', content: systemPrompt(sourceLang, targetLang) },
    { role: 'user', content: `Translate these strings.\n${JSON.stringify(payloadOf(items, sourceLang, targetLang))}` },
  ];
}

function buildRetryMessages(messages, rawReply, items, failures, sourceLang, targetLang) {
  const problems = items.map((item) => `- ${item.id}: ${failures.get(item.id).join('; ')}`).join('\n');
  const out = [...messages];
  if (rawReply != null) out.push({ role: 'assistant', content: rawReply });
  out.push({
    role: 'user',
    content:
      `Your previous reply had problems:\n${problems}\n` +
      'Fix them and reply with JSON only, containing translations for exactly these strings.\n' +
      JSON.stringify(payloadOf(items, sourceLang, targetLang)),
  });
  return out;
}

// Extracts {id: text} from a model reply. Throws Error with a message fit for feeding back.
export function parseResponse(raw) {
  let body = String(raw).trim();
  const fence = /^```(?:json)?\s*([\s\S]*?)\s*```$/i.exec(body);
  if (fence) body = fence[1];
  let data;
  try {
    data = JSON.parse(body);
  } catch {
    const start = body.indexOf('{');
    const end = body.lastIndexOf('}');
    try {
      data = JSON.parse(body.slice(start, end + 1));
    } catch {
      throw new Error('reply is not valid JSON');
    }
  }
  const translations = data?.translations;
  if (!translations || typeof translations !== 'object' || Array.isArray(translations)) {
    throw new Error('reply must be {"translations": {"<id>": "<text>"}}');
  }
  return translations;
}

// Checks a parsed reply against the requested items.
// Returns { ok: Map<id,text>, bad: Map<id,string[]>, batchError?: string }.
function evaluate(translations, items) {
  const requested = new Set(items.map((i) => i.id));
  const extras = Object.keys(translations).filter((id) => !requested.has(id));
  if (extras.length) {
    const message = `unexpected id(s) not requested: ${extras.join(', ')}`;
    return { ok: new Map(), bad: new Map(items.map((i) => [i.id, [message]])), batchError: message };
  }
  const ok = new Map();
  const bad = new Map();
  for (const item of items) {
    if (!Object.hasOwn(translations, item.id)) {
      bad.set(item.id, ['missing from reply']);
      continue;
    }
    const problems = validateTranslation(item.text, translations[item.id]);
    if (problems.length) bad.set(item.id, problems);
    else ok.set(item.id, translations[item.id]);
  }
  return { ok, bad };
}

// Translates one batch; failed strings get exactly one retry with the errors fed back.
// Returns { translated: Map<id,text>, failed: Map<id,string[]> }. Fatal provider errors propagate.
export async function translateBatch({ provider, items, sourceLang, targetLang, signal }) {
  const messages = buildMessages(items, sourceLang, targetLang);
  const translated = new Map();
  let pending = items;
  let failures = new Map();
  let conversation = messages;
  let rawReply = null;

  for (let attempt = 0; attempt < 2 && pending.length; attempt++) {
    if (attempt === 1) {
      conversation = buildRetryMessages(messages, rawReply, pending, failures, sourceLang, targetLang);
    }
    rawReply = null;
    try {
      rawReply = await provider.complete({ messages: conversation, signal });
      const result = evaluate(parseResponse(rawReply), pending);
      for (const [id, text] of result.ok) translated.set(id, text);
      failures = result.bad;
    } catch (err) {
      if (err?.fatal) throw err;
      failures = new Map(pending.map((i) => [i.id, [err.message]]));
    }
    pending = pending.filter((i) => failures.has(i.id));
  }
  return { translated, failed: new Map(pending.map((i) => [i.id, failures.get(i.id)])) };
}

// ---- one language end to end --------------------------------------------------

// Returns { file, stats } where `file` is the new strings.<lang>.json content (object).
export async function translateLanguage({ table, existing, lang, provider, force = false, batchOptions, log = () => {} }) {
  const { kept, todo, dropped } = planLanguage(table, existing, { force });
  const entries = new Map(kept);
  const failed = new Map();
  const batches = makeBatches(todo, batchOptions);

  for (const [index, items] of batches.entries()) {
    log(`  batch ${index + 1}/${batches.length} (${items.length} strings)`);
    const result = await translateBatch({
      provider,
      items,
      sourceLang: table.sourceLang,
      targetLang: lang,
    });
    for (const item of items) {
      if (result.translated.has(item.id)) {
        entries.set(item.id, { text: result.translated.get(item.id), src: fnv1a(item.text) });
      }
    }
    for (const [id, errors] of result.failed) failed.set(id, errors);
  }

  // Source order, so diffs stay stable.
  const strings = {};
  for (const item of table.strings) {
    if (entries.has(item.id)) strings[item.id] = entries.get(item.id);
  }
  return {
    file: { version: 1, lang, sourceLang: table.sourceLang, strings },
    stats: {
      kept: kept.size,
      translated: todo.length - failed.size,
      failed,
      dropped,
      batches: batches.length,
    },
  };
}

// One entry per line keeps diffs reviewable (matches docs/examples/strings.de.json).
export function serializeTranslation(file) {
  const ids = Object.keys(file.strings);
  const lines = ids.map(
    (id, i) =>
      `    ${JSON.stringify(id)}: { "text": ${JSON.stringify(file.strings[id].text)}, "src": ${JSON.stringify(file.strings[id].src)} }${i < ids.length - 1 ? ',' : ''}`,
  );
  return [
    '{',
    `  "version": ${file.version},`,
    `  "lang": ${JSON.stringify(file.lang)},`,
    `  "sourceLang": ${JSON.stringify(file.sourceLang)},`,
    ids.length ? '  "strings": {' : '  "strings": {}',
    ...(ids.length ? [...lines, '  }'] : []),
    '}',
    '',
  ].join('\n');
}

// Throws a descriptive Error if the source table breaks the v1 contract.
export function validateSourceTable(table) {
  if (!table || typeof table !== 'object') throw new Error('strings.json is not an object');
  if (table.version !== 1) throw new Error(`Unsupported strings.json version: ${table.version} (expected 1)`);
  if (typeof table.sourceLang !== 'string' || !table.sourceLang) throw new Error('strings.json: missing sourceLang');
  if (!Array.isArray(table.strings)) throw new Error('strings.json: "strings" must be an array');
  const seen = new Set();
  for (const item of table.strings) {
    if (!item || typeof item.id !== 'string' || typeof item.text !== 'string') {
      throw new Error('strings.json: every string needs a string "id" and "text"');
    }
    if (seen.has(item.id)) throw new Error(`strings.json: duplicate id ${item.id}`);
    seen.add(item.id);
  }
}

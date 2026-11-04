import { test } from 'node:test';
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import { mkdtemp, mkdir, readFile, writeFile, copyFile, rm, access } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import { fnv1a, validateTranslation, planLanguage, makeBatches, translateBatch, translateLanguage, parseResponse } from './translate-core.mjs';
import { createProvider, parseRetryAfter, ProviderError } from './llm-provider.mjs';
import { run } from './translate.mjs';

const EXAMPLES = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'docs', 'examples');
const SOURCE_TABLE = JSON.parse(await readFile(path.join(EXAMPLES, 'strings.json'), 'utf8'));

async function makeDecksDir({ withGerman = false } = {}) {
  const decksDir = await mkdtemp(path.join(os.tmpdir(), 'translate-test-'));
  const deck = path.join(decksDir, 'demo');
  await mkdir(deck);
  await copyFile(path.join(EXAMPLES, 'strings.json'), path.join(deck, 'strings.json'));
  if (withGerman) await copyFile(path.join(EXAMPLES, 'strings.de.json'), path.join(deck, 'strings.de.json'));
  return { decksDir, deck, cleanup: () => rm(decksDir, { recursive: true, force: true }) };
}

function cli(argv, deps) {
  const stdout = [];
  const stderr = [];
  return run(argv, { ...deps, stdout: (l) => stdout.push(l), stderr: (l) => stderr.push(l) }).then((code) => ({ code, stdout, stderr }));
}

// Provider stub that answers from a script of functions(items) => reply string.
function scripted(...replies) {
  const calls = [];
  return {
    name: 'scripted',
    calls,
    async complete({ messages }) {
      calls.push(messages);
      const reply = replies[Math.min(calls.length - 1, replies.length - 1)];
      return typeof reply === 'function' ? reply(messages) : reply;
    },
  };
}
const payloadOf = (messages) => JSON.parse(messages.at(-1).content.split('\n').at(-1));
const reply = (map) => JSON.stringify({ translations: map });

// ---- hash ---------------------------------------------------------------------

test('fnv1a matches reference vectors and the docs/examples pair', () => {
  assert.equal(fnv1a(''), '811c9dc5');
  assert.equal(fnv1a('a'), 'e40c292c');
  assert.equal(fnv1a('foobar'), 'bf9cf968');
  assert.equal(fnv1a('Decks as <b>Code</b>'), '2218a577');
  assert.equal(fnv1a('Replacing PPTX with plain-text presentations'), '2463d0d5');
});

test('fnv1a hashes UTF-8 bytes, not UTF-16 code units', () => {
  // "é" is bytes C3 A9: fnv1a = (((811c9dc5 ^ c3) * 01000193) ^ a9) * 01000193
  let h = 0x811c9dc5;
  for (const b of [0xc3, 0xa9]) h = Math.imul(h ^ b, 0x01000193) >>> 0;
  assert.equal(fnv1a('é'), h.toString(16).padStart(8, '0'));
  assert.equal(fnv1a('x').length, 8);
});

// ---- validation ---------------------------------------------------------------

test('validation accepts reordered tags and rejects broken ones', () => {
  const src = 'Read the <a href="https://example.com/docs">docs</a> <b>first</b>.<br>Then try it.';
  assert.deepEqual(validateTranslation(src, '<b>Zuerst</b> die <a href="https://example.com/docs">Doku</a> lesen.<br>Dann ausprobieren.'), []);
  assert.deepEqual(validateTranslation(src, 'Zuerst <b>die <a href="https://example.com/docs">Doku</a></b> lesen.<br/>Dann.'), []);

  assert.ok(validateTranslation(src, 'Lesen. Dann.').length, 'dropped tags');
  assert.ok(validateTranslation('<b>x</b>', '<b>x').length, 'unclosed');
  assert.ok(validateTranslation('<b>x</b> <i>y</i>', '<b>x <i>y</b></i>').length, 'misnested');
  assert.ok(validateTranslation('<b>x</b>', '<b>x</b><i>y</i>').length, 'extra tag');
  assert.ok(validateTranslation('<b>x</b>', '<b onclick="1">x</b>').length, 'attributes');
  assert.ok(validateTranslation('x', '<script>x</script>').length, 'disallowed tag');
  assert.ok(validateTranslation('1 &lt; 2', '1 < 2').length, 'unescaped <');
});

test('<code> is an allowed inline tag and must be kept; new-style ids are opaque', async () => {
  assert.deepEqual(validateTranslation('Run <code>npm test</code> now', 'Jetzt <code>npm test</code> ausführen'), []);
  assert.ok(validateTranslation('Run <code>npm test</code> now', 'Jetzt npm test ausführen').length);

  const table = { version: 1, sourceLang: 'en', strings: [{ id: 's3.L.sh7.p2', slide: 3, kind: 'text', text: 'Use <code>x</code>' }] };
  const { file } = await translateLanguage({ table, existing: null, lang: 'de', provider: createProvider('mock') });
  assert.equal(file.strings['s3.L.sh7.p2'].text, '[de: Use <code>x</code>]');
});

test('validation requires href values to be unchanged', () => {
  const src = '<a href="https://example.com/docs">docs</a>';
  assert.ok(validateTranslation(src, '<a href="https://example.com/de/docs">Doku</a>').length);
  assert.ok(validateTranslation(src, '<a href="HTTPS://EXAMPLE.COM/docs">Doku</a>').length);
  assert.ok(validateTranslation(src, '<a>Doku</a>').length);
  assert.deepEqual(validateTranslation(src, '<a href="https://example.com/docs">Doku</a>'), []);
});

test('validation rejects empty and non-string text', () => {
  for (const bad of ['', '   ', null, undefined, 42]) {
    assert.ok(validateTranslation('Save', bad).length, String(bad));
  }
});

// ---- incremental planning -----------------------------------------------------

test('plan skips fresh translations and re-queues stale or missing ones', () => {
  const existing = JSON.parse(JSON.stringify({
    version: 1, lang: 'de', sourceLang: 'en',
    strings: {
      's1.sh2.p1': { text: 'Decks als <b>Code</b>', src: '2218a577' },
      's1.sh3.p1': { text: 'alt', src: 'deadbeef' },
      'gone.p1': { text: 'orphan', src: '00000000' },
    },
  }));
  const plan = planLanguage(SOURCE_TABLE, existing);
  assert.deepEqual([...plan.kept.keys()], ['s1.sh2.p1']);
  assert.equal(plan.todo.length, SOURCE_TABLE.strings.length - 1);
  assert.ok(plan.todo.some((s) => s.id === 's1.sh3.p1'), 'stale hash re-queued');
  assert.deepEqual(plan.dropped, ['gone.p1']);

  assert.equal(planLanguage(SOURCE_TABLE, existing, { force: true }).todo.length, SOURCE_TABLE.strings.length);
});

test('plan re-queues a matching hash whose stored text is invalid', () => {
  const existing = { strings: { 's1.sh2.p1': { text: 'Decks als Code', src: '2218a577' } } };
  assert.ok(planLanguage(SOURCE_TABLE, existing).todo.some((s) => s.id === 's1.sh2.p1'));
});

test('batches stay in document order and respect the size limit', () => {
  const batches = makeBatches(SOURCE_TABLE.strings, { maxItems: 3 });
  assert.deepEqual(batches.map((b) => b.length), [3, 3, 2]);
  assert.deepEqual(batches.flat().map((s) => s.id), SOURCE_TABLE.strings.map((s) => s.id));
});

test('CLI only translates what is missing or stale, then is a no-op', async () => {
  const { decksDir, deck, cleanup } = await makeDecksDir({ withGerman: true });
  try {
    const provider = scripted((m) => reply(Object.fromEntries(payloadOf(m).strings.map((s) => [s.id, `DE ${s.text}`]))));
    const first = await cli(['demo', '--to', 'de'], { decksDir, provider });
    assert.equal(first.code, 0);
    const requested = provider.calls.flatMap((m) => payloadOf(m).strings.map((s) => s.id));
    assert.equal(requested.length, SOURCE_TABLE.strings.length - 2);
    assert.ok(!requested.includes('s1.sh2.p1') && !requested.includes('s1.sh3.p1'));

    const written = JSON.parse(await readFile(path.join(deck, 'strings.de.json'), 'utf8'));
    assert.equal(written.strings['s1.sh2.p1'].text, 'Decks als <b>Code</b>');
    assert.equal(Object.keys(written.strings).length, SOURCE_TABLE.strings.length);

    const before = provider.calls.length;
    const second = await cli(['demo', '--to', 'de'], { decksDir, provider });
    assert.equal(second.code, 0);
    assert.equal(provider.calls.length, before, 'nothing left to translate, no LLM call');

    // Edit one source string: only it is retranslated.
    const table = JSON.parse(await readFile(path.join(deck, 'strings.json'), 'utf8'));
    table.strings[2].text = 'Why bother?';
    await writeFile(path.join(deck, 'strings.json'), JSON.stringify(table));
    await cli(['demo', '--to', 'de'], { decksDir, provider });
    assert.deepEqual(payloadOf(provider.calls.at(-1)).strings.map((s) => s.id), ['s2.sh2.p1']);

    await cli(['demo', '--to', 'de', '--force'], { decksDir, provider });
    const forced = provider.calls.slice(before + 1).flatMap((m) => payloadOf(m).strings);
    assert.equal(forced.length, SOURCE_TABLE.strings.length);
  } finally {
    await cleanup();
  }
});

// ---- retry then skip ----------------------------------------------------------

test('a failed string is retried once with the error fed back, then succeeds', async () => {
  const items = SOURCE_TABLE.strings.slice(0, 2);
  const provider = scripted(
    reply({ 's1.sh2.p1': 'Decks als Code', 's1.sh3.p1': 'Klartext' }), // first drops the <b> tags
    (m) => reply({ 's1.sh2.p1': 'Decks als <b>Code</b>' }),
  );
  const result = await translateBatch({ provider, items, sourceLang: 'en', targetLang: 'de' });
  assert.equal(provider.calls.length, 2);
  assert.deepEqual([...result.translated.keys()].sort(), ['s1.sh2.p1', 's1.sh3.p1']);
  assert.equal(result.failed.size, 0);

  const retry = provider.calls[1];
  const feedback = retry.at(-1).content;
  assert.match(feedback, /s1\.sh2\.p1/);
  assert.match(feedback, /tag "b"/);
  assert.equal(retry.at(-2).role, 'assistant');
  assert.deepEqual(payloadOf(retry).strings.map((s) => s.id), ['s1.sh2.p1'], 'only the failed string is retried');
});

test('strings that still fail after one retry are skipped, reported, and the file stays valid', async () => {
  const { decksDir, deck, cleanup } = await makeDecksDir();
  try {
    const provider = scripted((m) => {
      const map = Object.fromEntries(payloadOf(m).strings.map((s) => [s.id, `X ${s.text}`]));
      map['s1.sh2.p1'] = 'Decks als Code'; // always loses the tags
      return reply(map);
    });
    const result = await cli(['demo', '--to', 'de'], { decksDir, provider });
    assert.equal(result.code, 2);
    assert.match(result.stderr.join('\n'), /skipped s1\.sh2\.p1/);

    const file = JSON.parse(await readFile(path.join(deck, 'strings.de.json'), 'utf8'));
    assert.equal(file.version, 1);
    assert.ok(!('s1.sh2.p1' in file.strings));
    assert.equal(Object.keys(file.strings).length, SOURCE_TABLE.strings.length - 1);
    for (const [id, entry] of Object.entries(file.strings)) {
      const src = SOURCE_TABLE.strings.find((s) => s.id === id);
      assert.equal(entry.src, fnv1a(src.text));
      assert.deepEqual(validateTranslation(src.text, entry.text), []);
    }
  } finally {
    await cleanup();
  }
});

test('extra ids, invalid JSON and provider errors fail the batch and then retry', async () => {
  const items = SOURCE_TABLE.strings.slice(1, 2);
  const ok = reply({ 's1.sh3.p1': 'Klartext' });

  const extras = scripted(reply({ 's1.sh3.p1': 'Klartext', bogus: 'x' }), ok);
  assert.equal((await translateBatch({ provider: extras, items, sourceLang: 'en', targetLang: 'de' })).failed.size, 0);

  const junk = scripted('not json at all', ok);
  assert.equal((await translateBatch({ provider: junk, items, sourceLang: 'en', targetLang: 'de' })).failed.size, 0);

  let n = 0;
  const flaky = { name: 'flaky', complete: async () => { if (n++ === 0) throw new ProviderError('HTTP 500'); return ok; } };
  assert.equal((await translateBatch({ provider: flaky, items, sourceLang: 'en', targetLang: 'de' })).failed.size, 0);

  const dead = scripted('nope', 'still nope');
  const result = await translateBatch({ provider: dead, items, sourceLang: 'en', targetLang: 'de' });
  assert.equal(dead.calls.length, 2, 'exactly one retry');
  assert.equal(result.failed.size, 1);
});

test('fatal provider errors abort without writing a file', async () => {
  const { decksDir, deck, cleanup } = await makeDecksDir();
  try {
    const provider = { name: 'bad-auth', complete: async () => { throw new ProviderError('HTTP 401', { status: 401, fatal: true }); } };
    const result = await cli(['demo', '--to', 'de'], { decksDir, provider });
    assert.equal(result.code, 1);
    await assert.rejects(access(path.join(deck, 'strings.de.json')));
  } finally {
    await cleanup();
  }
});

test('parseResponse tolerates code fences and surrounding prose', () => {
  assert.deepEqual(parseResponse('```json\n{"translations":{"a":"b"}}\n```'), { a: 'b' });
  assert.deepEqual(parseResponse('Here you go: {"translations":{"a":"b"}} Done.'), { a: 'b' });
  assert.throws(() => parseResponse('{"a":"b"}'));
});

// ---- RTL / fallback -----------------------------------------------------------

test('RTL languages use the same file shape; skipped strings are simply absent (viewer falls back to source)', async () => {
  const { decksDir, deck, cleanup } = await makeDecksDir();
  try {
    const provider = scripted((m) => {
      const map = Object.fromEntries(payloadOf(m).strings.map((s) => [s.id, `ع ${s.text}`]));
      delete map['s2.notes.p1'];
      return reply(map);
    });
    const result = await cli(['demo', '--to', 'ar'], { decksDir, provider });
    assert.equal(result.code, 2);
    const file = JSON.parse(await readFile(path.join(deck, 'strings.ar.json'), 'utf8'));
    assert.deepEqual(Object.keys(file).sort(), ['lang', 'sourceLang', 'strings', 'version'], 'no dir/rtl field; the viewer derives it from lang');
    assert.equal(file.lang, 'ar');
    assert.ok(!('s2.notes.p1' in file.strings));
    assert.ok('s2.sh3.p2' in file.strings);
  } finally {
    await cleanup();
  }
});

// ---- mock provider end to end -------------------------------------------------

test('mock provider end to end against docs/examples/strings.json', async () => {
  const { decksDir, deck, cleanup } = await makeDecksDir();
  try {
    const result = await cli(['demo', '--to', 'de,pt-BR', '--provider', 'mock'], { decksDir });
    assert.equal(result.code, 0, result.stderr.join('\n'));

    for (const lang of ['de', 'pt-BR']) {
      const file = JSON.parse(await readFile(path.join(deck, `strings.${lang}.json`), 'utf8'));
      assert.equal(file.version, 1);
      assert.equal(file.lang, lang);
      assert.equal(file.sourceLang, 'en');
      assert.deepEqual(Object.keys(file.strings), SOURCE_TABLE.strings.map((s) => s.id), 'source order');
      for (const s of SOURCE_TABLE.strings) {
        const entry = file.strings[s.id];
        assert.equal(entry.text, `[${lang}: ${s.text}]`);
        assert.equal(entry.src, fnv1a(s.text));
        assert.match(entry.src, /^[0-9a-f]{8}$/);
        assert.deepEqual(validateTranslation(s.text, entry.text), []);
      }
    }

    const again = await cli(['demo', '--to', 'de', '--provider', 'mock'], { decksDir });
    assert.match(again.stdout.join('\n'), /0 translated, 8 kept/);
  } finally {
    await cleanup();
  }
});

test('mock translations also pass through the provider via the retry conversation', async () => {
  const mock = createProvider('mock');
  const items = SOURCE_TABLE.strings.slice(0, 2);
  const result = await translateBatch({ provider: mock, items, sourceLang: 'en', targetLang: 'fr' });
  assert.equal(result.translated.get('s1.sh2.p1'), '[fr: Decks as <b>Code</b>]');
});

test('--dry-run makes no provider calls and writes nothing', async () => {
  const { decksDir, deck, cleanup } = await makeDecksDir({ withGerman: true });
  try {
    const result = await cli(['demo', '--to', 'de,fr', '--dry-run'], { decksDir, env: {} });
    assert.equal(result.code, 0);
    assert.match(result.stdout.join('\n'), /de: 6 to translate/);
    assert.match(result.stdout.join('\n'), /fr: 8 to translate/);
    await assert.rejects(access(path.join(deck, 'strings.fr.json')));
    assert.equal(await readFile(path.join(deck, 'strings.de.json'), 'utf8'), await readFile(path.join(EXAMPLES, 'strings.de.json'), 'utf8'));
  } finally {
    await cleanup();
  }
});

test('CLI rejects bad arguments', async () => {
  const { decksDir, cleanup } = await makeDecksDir();
  try {
    for (const argv of [[], ['demo'], ['demo', '--to', 'de/../x'], ['../demo', '--to', 'de'], ['demo', '--to', 'de', '--bogus']]) {
      assert.equal((await cli(argv, { decksDir })).code, 1, argv.join(' '));
    }
    assert.equal((await cli(['missing', '--to', 'de', '--provider', 'mock'], { decksDir })).code, 1);
  } finally {
    await cleanup();
  }
});

// ---- openai-compatible provider (fake fetch, no network) ----------------------

const KEY = 'sk-test-secret-123';
const okResponse = (content) => new Response(JSON.stringify({ choices: [{ message: { content }, finish_reason: 'stop' }] }), { status: 200 });

test('openai provider honours retry-after on 429 and never leaks the key', async () => {
  const sleeps = [];
  const requests = [];
  const responses = [
    new Response('slow down', { status: 429, headers: { 'retry-after': '7' } }),
    new Response('slow down', { status: 429, headers: { 'retry-after': '2' } }),
    okResponse('hello'),
  ];
  const provider = createProvider('openai', {
    env: { LLM_API_KEY: KEY, LLM_MODEL: 'm' },
    baseUrl: 'https://llm.example.com/v1/',
    sleep: async (ms) => sleeps.push(ms),
    fetch: async (url, init) => { requests.push({ url: String(url), init }); return responses.shift(); },
  });
  assert.equal(await provider.complete({ messages: [{ role: 'user', content: 'hi' }] }), 'hello');
  assert.deepEqual(sleeps, [7000, 2000]);
  assert.equal(requests[0].url, 'https://llm.example.com/v1/chat/completions');
  assert.equal(requests[0].init.headers.authorization, `Bearer ${KEY}`);
  const body = JSON.parse(requests[0].init.body);
  assert.equal(body.model, 'm');
  assert.deepEqual(body.response_format, { type: 'json_object' });
});

test('openai provider: api-key header, 401 is fatal, error text is redacted', async () => {
  const seen = [];
  const provider = createProvider('openai', {
    env: { LLM_API_KEY: KEY, LLM_MODEL: 'm', LLM_AUTH_HEADER: 'api-key', LLM_JSON_MODE: 'off' },
    baseUrl: 'https://x.openai.azure.com/openai/v1',
    sleep: async () => {},
    fetch: async (url, init) => { seen.push(init); return new Response(`bad key ${KEY}`, { status: 401 }); },
  });
  await assert.rejects(
    provider.complete({ messages: [] }),
    (err) => err instanceof ProviderError && err.fatal && !err.message.includes(KEY) && err.message.includes('[redacted]'),
  );
  assert.equal(seen[0].headers['api-key'], KEY);
  assert.equal(seen[0].headers.authorization, undefined);
  assert.equal(JSON.parse(seen[0].body).response_format, undefined);
});

test('openai provider gives up on persistent 429 and retries 5xx', async () => {
  let calls = 0;
  const provider = createProvider('openai', {
    env: { LLM_API_KEY: KEY, LLM_MODEL: 'm' },
    sleep: async () => {},
    fetch: async () => { calls++; return new Response('', { status: 429 }); },
  });
  await assert.rejects(provider.complete({ messages: [] }), /429/);
  assert.equal(calls, 6);

  const flaky = [new Response('', { status: 503 }), okResponse('fine')];
  const p2 = createProvider('openai', { env: { LLM_API_KEY: KEY, LLM_MODEL: 'm' }, sleep: async () => {}, fetch: async () => flaky.shift() });
  assert.equal(await p2.complete({ messages: [] }), 'fine');
});

test('openai provider config errors are explicit and fatal', () => {
  const base = { LLM_API_KEY: KEY, LLM_MODEL: 'm' };
  assert.throws(() => createProvider('openai', { env: { LLM_MODEL: 'm' } }), /API key/);
  assert.throws(() => createProvider('openai', { env: { LLM_API_KEY: KEY } }), /model/i);
  assert.throws(() => createProvider('openai', { env: base, baseUrl: 'http://example.com/v1' }), /https/);
  assert.doesNotThrow(() => createProvider('openai', { env: base, baseUrl: 'http://localhost:11434/v1' }));
  assert.throws(() => createProvider('nope'), /Unknown provider/);
});

test('copilot provider sends the prompt on stdin with tools disabled and returns stdout', async () => {
  const calls = [];
  const spawn = (command, args) => {
    const child = new EventEmitter();
    child.stdout = new EventEmitter();
    child.stderr = new EventEmitter();
    child.kill = () => {};
    child.stdin = { on() {}, end(text) { calls.push({ command, args, stdin: text }); setImmediate(() => { child.stdout.emit('data', '{"translations":{}}\n'); child.emit('close', 0); }); } };
    return child;
  };
  const provider = createProvider('copilot', { env: {}, model: 'some-model', spawn });
  const out = await provider.complete({ messages: [{ role: 'system', content: 'SYS' }, { role: 'user', content: 'USER' }] });
  assert.equal(out, '{"translations":{}}');
  assert.equal(calls[0].command, 'copilot');
  assert.deepEqual(calls[0].args.slice(calls[0].args.indexOf('--available-tools'), calls[0].args.indexOf('--available-tools') + 2), ['--available-tools', '']);
  assert.ok(calls[0].args.includes('--no-custom-instructions'));
  assert.deepEqual(calls[0].args.slice(-2), ['--model', 'some-model']);
  assert.ok(calls[0].stdin.includes('[system]\nSYS') && calls[0].stdin.includes('[user]\nUSER'));
  assert.ok(!calls[0].args.some((a) => a.includes('USER')), 'prompt must not be on the command line');
});

test('copilot provider failures: missing CLI and sign-in problems are fatal, others are not', async () => {
  const failing = (emit) => () => {
    const child = new EventEmitter();
    child.stdout = new EventEmitter();
    child.stderr = new EventEmitter();
    child.kill = () => {};
    child.stdin = { on() {}, end() { setImmediate(() => emit(child)); } };
    return child;
  };
  const messages = [{ role: 'user', content: 'x' }];
  const missing = createProvider('copilot', { env: {}, spawn: failing((c) => c.emit('error', Object.assign(new Error('spawn copilot ENOENT'), { code: 'ENOENT' }))) });
  await assert.rejects(missing.complete({ messages }), (e) => e instanceof ProviderError && e.fatal && /not found/.test(e.message));

  const signedOut = createProvider('copilot', { env: {}, spawn: failing((c) => { c.stderr.emit('data', 'Please log in with copilot login'); c.emit('close', 1); }) });
  await assert.rejects(signedOut.complete({ messages }), (e) => e.fatal === true);

  const flaky = createProvider('copilot', { env: {}, spawn: failing((c) => { c.stderr.emit('data', 'something odd'); c.emit('close', 1); }) });
  await assert.rejects(flaky.complete({ messages }), (e) => e.fatal === false);
});

test('retry-after parsing handles seconds and HTTP dates', () => {
  assert.equal(parseRetryAfter('3'), 3000);
  assert.equal(parseRetryAfter('Wed, 21 Oct 2015 07:28:05 GMT', Date.parse('Wed, 21 Oct 2015 07:28:00 GMT')), 5000);
  assert.equal(parseRetryAfter(null), null);
  assert.equal(parseRetryAfter('soon'), null);
});

test('translateLanguage output is stable for a ready-made provider', async () => {
  const { file, stats } = await translateLanguage({ table: SOURCE_TABLE, existing: null, lang: 'de', provider: createProvider('mock') });
  assert.equal(stats.translated, SOURCE_TABLE.strings.length);
  assert.equal(file.strings['s2.sh4.r1c1.p1'].text, '[de: Tool]');
});

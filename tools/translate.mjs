#!/usr/bin/env node
// Translate a deck's string table with an LLM. See docs/string-table.md for the contract.
//
//   node tools/translate.mjs <slug> --to de,fr [--force] [--dry-run]
//                            [--provider copilot|openai|mock] [--model <name>] [--base-url <url>]
//
// Exit codes: 0 ok, 1 error (nothing partial written for the failing language),
// 2 finished but some strings were skipped (they fall back to source in the viewer).

import { readFile, writeFile, rename, mkdir } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createProvider, PROVIDERS } from './llm-provider.mjs';
import { planLanguage, makeBatches, serializeTranslation, translateLanguage, validateSourceTable } from './translate-core.mjs';

const USAGE = `Usage: node tools/translate.mjs <slug> --to <lang>[,<lang>...] [options]

Options:
  --to <langs>       Comma-separated BCP 47 tags, e.g. de,pt-BR,ar (required)
  --force            Retranslate every string, ignoring existing translations
  --dry-run          Show what would be translated; no LLM calls, nothing written
  --provider <name>  ${PROVIDERS.join(' | ')} (default: copilot)
  --model <name>     Model name (or LLM_MODEL); copilot defaults to the CLI's own model
  --base-url <url>   OpenAI-compatible endpoint (or LLM_BASE_URL); openai provider only
  --batch-size <n>   Max strings per request (default 30)
  --decks-dir <dir>  Decks directory (default: <repo>/decks)
  -h, --help         Show this help

Provider "copilot" runs the GitHub Copilot CLI (sign in once with \`copilot\`); it uses the
models in your Copilot subscription and needs no key. Optional: LLM_REASONING_EFFORT, COPILOT_BIN.
Provider "openai": LLM_API_KEY or OPENAI_API_KEY, LLM_MODEL, LLM_BASE_URL,
LLM_AUTH_HEADER=bearer|api-key, LLM_JSON_MODE=off.`;

const LANG_RE = /^[A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,8})*$/;
const SLUG_RE = /^[A-Za-z0-9][A-Za-z0-9._-]*$/;
const REPO_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

class UsageError extends Error {}

export function parseArgs(argv) {
  const opts = { force: false, dryRun: false, provider: 'copilot', langs: [] };
  const positional = [];
  const takeValue = (i, flag) => {
    if (i + 1 >= argv.length || argv[i + 1].startsWith('--')) throw new UsageError(`${flag} needs a value`);
    return argv[i + 1];
  };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    switch (arg) {
      case '-h':
      case '--help':
        opts.help = true;
        break;
      case '--force':
        opts.force = true;
        break;
      case '--dry-run':
        opts.dryRun = true;
        break;
      case '--to':
        opts.langs.push(...takeValue(i++, arg).split(',').map((s) => s.trim()).filter(Boolean));
        break;
      case '--provider':
        opts.provider = takeValue(i++, arg);
        break;
      case '--model':
        opts.model = takeValue(i++, arg);
        break;
      case '--base-url':
        opts.baseUrl = takeValue(i++, arg);
        break;
      case '--batch-size':
        opts.batchSize = Number(takeValue(i++, arg));
        break;
      case '--decks-dir':
        opts.decksDir = takeValue(i++, arg);
        break;
      default:
        if (arg.startsWith('-')) throw new UsageError(`Unknown option ${arg}`);
        positional.push(arg);
    }
  }
  if (opts.help) return opts;
  if (positional.length !== 1) throw new UsageError('Expected exactly one <slug>');
  opts.slug = positional[0];
  if (!SLUG_RE.test(opts.slug)) throw new UsageError(`Invalid slug "${opts.slug}"`);
  if (!opts.langs.length) throw new UsageError('--to is required');
  for (const lang of opts.langs) {
    if (!LANG_RE.test(lang)) throw new UsageError(`Invalid language tag "${lang}" (expected BCP 47, e.g. de, pt-BR)`);
  }
  opts.langs = [...new Set(opts.langs)];
  if (opts.batchSize !== undefined && !(Number.isInteger(opts.batchSize) && opts.batchSize > 0)) {
    throw new UsageError('--batch-size must be a positive integer');
  }
  return opts;
}

async function readJson(file, what) {
  let raw;
  try {
    raw = await readFile(file, 'utf8');
  } catch (err) {
    if (err.code === 'ENOENT') return { missing: true };
    throw err;
  }
  try {
    return { raw, data: JSON.parse(raw) };
  } catch {
    throw new Error(`${what} is not valid JSON: ${file}`);
  }
}

async function writeAtomic(file, content) {
  const tmp = `${file}.${process.pid}.tmp`;
  await writeFile(tmp, content, 'utf8');
  await rename(tmp, file);
}

// `deps` lets tests inject env, streams, a decks directory and a ready-made provider.
export async function run(argv, deps = {}) {
  const out = deps.stdout ?? ((line) => console.log(line));
  const err = deps.stderr ?? ((line) => console.error(line));

  let opts;
  try {
    opts = parseArgs(argv);
  } catch (e) {
    if (!(e instanceof UsageError)) throw e;
    err(`${e.message}\n\n${USAGE}`);
    return 1;
  }
  if (opts.help) {
    out(USAGE);
    return 0;
  }

  try {
    const decksDir = path.resolve(opts.decksDir ?? deps.decksDir ?? path.join(REPO_ROOT, 'decks'));
    const sourceFile = path.join(decksDir, opts.slug, 'strings.json');
    const source = await readJson(sourceFile, 'strings.json');
    if (source.missing) throw new Error(`No source table at ${sourceFile}. Run the PPTX converter first.`);
    const table = source.data;
    validateSourceTable(table);

    const langs = opts.langs.filter((lang) => {
      const same = lang.toLowerCase() === table.sourceLang.toLowerCase();
      if (same) err(`Skipping ${lang}: same as sourceLang`);
      return !same;
    });

    let provider = null;
    if (!opts.dryRun && langs.length) {
      provider = deps.provider ?? createProvider(opts.provider, { env: deps.env, model: opts.model, baseUrl: opts.baseUrl, fetch: deps.fetch, sleep: deps.sleep });
    }
    const batchOptions = opts.batchSize ? { maxItems: opts.batchSize } : undefined;

    let skipped = 0;
    for (const lang of langs) {
      const target = path.join(decksDir, opts.slug, `strings.${lang}.json`);
      const prior = await readJson(target, path.basename(target)).catch((e) => {
        if (!opts.force) throw e;
        return { missing: false, raw: null, data: null };
      });
      if (!prior.missing && !opts.force) {
        const d = prior.data;
        if (d?.version !== 1 || typeof d.strings !== 'object' || d.strings === null) {
          throw new Error(`${target} is not a v1 translation file; fix it or re-run with --force`);
        }
      }
      // With --force the old file is ignored entirely.
      const existing = prior.missing || opts.force ? null : prior.data;
      const sameSource = existing && existing.sourceLang === table.sourceLang ? existing : null;

      if (opts.dryRun) {
        const plan = planLanguage(table, sameSource, { force: opts.force });
        const batches = makeBatches(plan.todo, batchOptions);
        out(`${lang}: ${plan.todo.length} to translate in ${batches.length} batch(es), ${plan.kept.size} up to date, ${plan.dropped.length} orphaned`);
        for (const item of plan.todo) out(`  ${item.id} [${item.kind}]`);
        continue;
      }

      out(`${lang}: translating with ${provider.name}...`);
      const { file, stats } = await translateLanguage({
        table,
        existing: sameSource,
        lang,
        provider,
        force: opts.force,
        batchOptions,
        log: out,
      });

      const serialized = serializeTranslation(file);
      if (prior.missing || prior.raw !== serialized) {
        await mkdir(path.dirname(target), { recursive: true });
        await writeAtomic(target, serialized);
      }
      const shown = path.relative(process.cwd(), target);
      out(`${lang}: ${stats.translated} translated, ${stats.kept} kept, ${stats.failed.size} skipped -> ${shown.startsWith('..') ? target : shown}`);
      for (const [id, errors] of stats.failed) {
        err(`  skipped ${id}: ${errors.join('; ')}`);
        skipped++;
      }
    }
    if (skipped) err(`${skipped} string(s) skipped; they will fall back to the source text. Re-run to retry them.`);
    return skipped ? 2 : 0;
  } catch (e) {
    err(`Error: ${e.message}`);
    return 1;
  }
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
  process.exitCode = await run(process.argv.slice(2));
}

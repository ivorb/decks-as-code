// LLM provider interface for tools/translate.mjs.
//
// A provider is any object shaped like:
//   { name: string, complete({ messages, signal? }) => Promise<string> }
// where messages are OpenAI-style { role: 'system' | 'user' | 'assistant', content }.
// Providers throw ProviderError; `fatal` means retrying or continuing is pointless
// (bad credentials, unknown model, ...), so the CLI aborts instead of skipping strings.

import { spawn } from 'node:child_process';

export class ProviderError extends Error {
  constructor(message, { status, fatal = false } = {}) {
    super(message);
    this.name = 'ProviderError';
    this.status = status;
    this.fatal = fatal;
  }
}

export const PROVIDERS = ['copilot', 'openai', 'mock'];

const DEFAULT_BASE_URL = 'https://api.openai.com/v1';
const REQUEST_TIMEOUT_MS = 120_000;
const COPILOT_TIMEOUT_MS = 300_000;
const MAX_RATE_LIMIT_RETRIES = 5;
const MAX_TRANSIENT_RETRIES = 3;
const MAX_WAIT_MS = 120_000;

const defaultSleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

export function createProvider(name, options = {}) {
  switch (name) {
    case 'copilot':
      return createCopilotProvider(options);
    case 'openai':
      return createOpenAICompatibleProvider(options);
    case 'mock':
      return createMockProvider();
    default:
      throw new Error(`Unknown provider "${name}". Available: ${PROVIDERS.join(', ')}`);
  }
}

// Uses the GitHub Copilot CLI in non-interactive mode, so translation runs on the
// models included with your Copilot subscription: no API key, nothing to configure
// beyond being signed in to `copilot`. Tools, MCP servers and custom instructions are
// all turned off, so the model can only answer; it cannot touch files or the shell.
// The prompt goes in on stdin (no command-line length limit, nothing in the process list).
//   --model <name>          optional; otherwise the CLI's default model
//   LLM_REASONING_EFFORT    low (default) | medium | high ...
//   COPILOT_BIN             path to the copilot executable if it is not on PATH
export function flattenMessages(messages) {
  const parts = messages.map(({ role, content }) => `[${role}]\n${content}`);
  parts.push('Reply with only the JSON object, no commentary.');
  return parts.join('\n\n');
}

function createCopilotProvider(options) {
  const env = options.env ?? process.env;
  const spawnImpl = options.spawn ?? spawn;
  const command = env.COPILOT_BIN || 'copilot';
  const model = options.model || env.LLM_MODEL || '';
  const effort = env.LLM_REASONING_EFFORT || 'low';

  const args = [
    '-s', '--no-ask-user', '--no-custom-instructions', '--disable-builtin-mcps', '--no-color',
    '--available-tools', '', '--reasoning-effort', effort,
  ];
  if (model) args.push('--model', model);

  return {
    name: 'copilot',
    complete({ messages, signal }) {
      return new Promise((resolve, reject) => {
        let child;
        try {
          child = spawnImpl(command, args, { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true });
        } catch (err) {
          reject(new ProviderError(`Could not start "${command}": ${err.message}`, { fatal: true }));
          return;
        }
        let stdout = '';
        let stderr = '';
        const timer = setTimeout(() => child.kill(), COPILOT_TIMEOUT_MS);
        signal?.addEventListener('abort', () => child.kill(), { once: true });

        child.stdout.on('data', (chunk) => { stdout += chunk; });
        child.stderr.on('data', (chunk) => { stderr += chunk; });
        child.on('error', (err) => {
          clearTimeout(timer);
          const missing = err.code === 'ENOENT';
          reject(new ProviderError(
            missing ? `GitHub Copilot CLI not found ("${command}"). Install it and run "copilot" once to sign in, or set COPILOT_BIN.` : `Copilot CLI failed: ${err.message}`,
            { fatal: missing },
          ));
        });
        child.on('close', (code) => {
          clearTimeout(timer);
          if (code === 0 && stdout.trim()) {
            resolve(stdout.trim());
            return;
          }
          const detail = (stderr || stdout).trim().slice(0, 500);
          const auth = /log ?in|auth|token|sign ?in|subscription|not (?:available|enabled)|unknown model|invalid model/i.test(detail);
          reject(new ProviderError(`Copilot CLI exited with code ${code}: ${detail || 'no output'}`, { fatal: auth }));
        });
        child.stdin.on('error', () => {});
        child.stdin.end(flattenMessages(messages));
      });
    },
  };
}

// Parses a Retry-After header (delta-seconds or HTTP-date) into milliseconds.
export function parseRetryAfter(value, now = Date.now()) {
  if (value == null || value === '') return null;
  const seconds = Number(value);
  if (Number.isFinite(seconds)) return Math.max(0, seconds * 1000);
  const date = Date.parse(value);
  return Number.isNaN(date) ? null : Math.max(0, date - now);
}

function isLocalhost(hostname) {
  return hostname === 'localhost' || hostname === '127.0.0.1' || hostname === '[::1]';
}

// Any OpenAI-compatible /chat/completions endpoint: OpenAI, Azure OpenAI / Foundry
// (…/openai/v1 with LLM_AUTH_HEADER=api-key), Ollama, vLLM, LiteLLM, ...
// Config comes from options first, then environment:
//   baseUrl  LLM_BASE_URL  | OPENAI_BASE_URL  (default https://api.openai.com/v1)
//   apiKey   LLM_API_KEY   | OPENAI_API_KEY   (environment only, never a CLI flag)
//   model    LLM_MODEL                         (required)
//   LLM_AUTH_HEADER   "bearer" (default) or "api-key" (Azure)
//   LLM_JSON_MODE     "off" to omit response_format (for hosts that reject it)
function createOpenAICompatibleProvider(options) {
  const env = options.env ?? process.env;
  const fetchImpl = options.fetch ?? globalThis.fetch;
  const sleep = options.sleep ?? defaultSleep;

  const baseUrl = (options.baseUrl || env.LLM_BASE_URL || env.OPENAI_BASE_URL || DEFAULT_BASE_URL).replace(/\/+$/, '');
  const apiKey = env.LLM_API_KEY || env.OPENAI_API_KEY || '';
  const model = options.model || env.LLM_MODEL || '';
  const authHeader = (env.LLM_AUTH_HEADER || 'bearer').toLowerCase();
  const jsonMode = (env.LLM_JSON_MODE || 'on').toLowerCase() !== 'off';

  if (!apiKey) {
    throw new ProviderError('No API key. Set LLM_API_KEY (or OPENAI_API_KEY) in the environment.', { fatal: true });
  }
  if (!model) {
    throw new ProviderError('No model. Pass --model <name> or set LLM_MODEL.', { fatal: true });
  }
  if (!['bearer', 'api-key'].includes(authHeader)) {
    throw new ProviderError('LLM_AUTH_HEADER must be "bearer" or "api-key".', { fatal: true });
  }
  let url;
  try {
    url = new URL(`${baseUrl}/chat/completions`);
  } catch {
    throw new ProviderError(`Invalid base URL: ${baseUrl}`, { fatal: true });
  }
  if (url.protocol !== 'https:' && !(url.protocol === 'http:' && isLocalhost(url.hostname))) {
    throw new ProviderError('Base URL must be https (http is only allowed for localhost); refusing to send the API key in clear text.', { fatal: true });
  }

  const redact = (text) => (apiKey ? String(text).split(apiKey).join('[redacted]') : String(text));

  const headers = { 'content-type': 'application/json' };
  if (authHeader === 'api-key') headers['api-key'] = apiKey;
  else headers.authorization = `Bearer ${apiKey}`;

  async function post(body, signal) {
    let rateLimited = 0;
    let transient = 0;
    for (;;) {
      let res;
      try {
        res = await fetchImpl(url, {
          method: 'POST',
          headers,
          body: JSON.stringify(body),
          signal: signal ?? AbortSignal.timeout(REQUEST_TIMEOUT_MS),
        });
      } catch (err) {
        if (signal?.aborted) throw err;
        if (++transient > MAX_TRANSIENT_RETRIES) {
          throw new ProviderError(`Network error: ${redact(err.message)}`);
        }
        await sleep(Math.min(1000 * 2 ** transient, MAX_WAIT_MS));
        continue;
      }

      if (res.ok) return res.json();

      const detail = redact((await res.text().catch(() => '')).slice(0, 500));
      if (res.status === 429) {
        if (++rateLimited > MAX_RATE_LIMIT_RETRIES) {
          throw new ProviderError(`Rate limited (429) after ${MAX_RATE_LIMIT_RETRIES} retries`, { status: 429 });
        }
        const wait = parseRetryAfter(res.headers.get('retry-after')) ?? 1000 * 2 ** rateLimited;
        await sleep(Math.min(wait, MAX_WAIT_MS));
        continue;
      }
      if (res.status >= 500) {
        if (++transient > MAX_TRANSIENT_RETRIES) {
          throw new ProviderError(`HTTP ${res.status}: ${detail}`, { status: res.status });
        }
        await sleep(Math.min(1000 * 2 ** transient, MAX_WAIT_MS));
        continue;
      }
      const fatal = [401, 403, 404].includes(res.status);
      throw new ProviderError(`HTTP ${res.status}: ${detail}`, { status: res.status, fatal });
    }
  }

  return {
    name: 'openai',
    async complete({ messages, signal }) {
      const body = { model, messages };
      if (jsonMode) body.response_format = { type: 'json_object' };
      const data = await post(body, signal);
      const choice = data?.choices?.[0];
      if (!choice || typeof choice.message?.content !== 'string') {
        throw new ProviderError('Unexpected response shape from provider (no choices[0].message.content)');
      }
      if (choice.finish_reason === 'length') {
        throw new ProviderError('Response was truncated (finish_reason=length); try a smaller --batch-size');
      }
      return choice.message.content;
    },
  };
}

// Deterministic offline fake: wraps each string as "[<lang>: <text>]", which keeps
// every tag and href intact. Understands the payload built by translate-core.mjs.
function createMockProvider() {
  return {
    name: 'mock',
    async complete({ messages }) {
      const last = messages[messages.length - 1];
      // The payload is always the last line of the last message, as compact JSON.
      const payload = JSON.parse(last.content.slice(last.content.lastIndexOf('\n') + 1));
      const translations = {};
      for (const item of payload.strings) {
        translations[item.id] = `[${payload.targetLang}: ${item.text}]`;
      }
      return JSON.stringify({ translations });
    },
  };
}

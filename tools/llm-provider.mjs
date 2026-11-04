// LLM provider interface for tools/translate.mjs.
//
// A provider is any object shaped like:
//   { name: string, complete({ messages, signal? }) => Promise<string> }
// where messages are OpenAI-style { role: 'system' | 'user' | 'assistant', content }.
// Providers throw ProviderError; `fatal` means retrying or continuing is pointless
// (bad credentials, unknown model, ...), so the CLI aborts instead of skipping strings.

export class ProviderError extends Error {
  constructor(message, { status, fatal = false } = {}) {
    super(message);
    this.name = 'ProviderError';
    this.status = status;
    this.fatal = fatal;
  }
}

export const PROVIDERS = ['openai', 'mock'];

const DEFAULT_BASE_URL = 'https://api.openai.com/v1';
const REQUEST_TIMEOUT_MS = 120_000;
const MAX_RATE_LIMIT_RETRIES = 5;
const MAX_TRANSIENT_RETRIES = 3;
const MAX_WAIT_MS = 120_000;

const defaultSleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

export function createProvider(name, options = {}) {
  switch (name) {
    case 'openai':
      return createOpenAICompatibleProvider(options);
    case 'mock':
      return createMockProvider();
    default:
      throw new Error(`Unknown provider "${name}". Available: ${PROVIDERS.join(', ')}`);
  }
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

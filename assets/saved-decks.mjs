/**
 * Decks converted in the browser and saved there, shared by the landing page and
 * the converter. Each is kept as its original .pptx in IndexedDB (too large for
 * sessionStorage) and converted again when opened. Nothing leaves the browser.
 */
const DB_NAME = 'decks-session';
const STORE = 'decks';

export const hasStorage = typeof indexedDB !== 'undefined';

const openDb = () => new Promise((resolve, reject) => {
  const request = indexedDB.open(DB_NAME, 1);
  request.onupgradeneeded = () => request.result.createObjectStore(STORE, { keyPath: 'id' });
  request.onsuccess = () => resolve(request.result);
  request.onerror = () => reject(request.error);
});

async function run(mode, action) {
  const db = await openDb();
  try {
    return await new Promise((resolve, reject) => {
      const transaction = db.transaction(STORE, mode);
      const request = action(transaction.objectStore(STORE));
      transaction.oncomplete = () => resolve(request?.result);
      transaction.onerror = () => reject(transaction.error);
      transaction.onabort = () => reject(transaction.error);
    });
  } finally {
    db.close();
  }
}

/** Newest first. Never throws: blocked storage just means an empty list. */
export async function listSaved() {
  if (!hasStorage) return [];
  try {
    return ((await run('readonly', (s) => s.getAll())) ?? []).sort((a, b) => b.savedAt - a.savedAt);
  } catch {
    return [];
  }
}

export const getSaved = (id) => run('readonly', (s) => s.get(id));
export const removeSaved = (id) => run('readwrite', (s) => s.delete(id));

/** `bytes` is the original .pptx as a Uint8Array. */
export async function saveDeck({ name, bytes, title, slug, author, slides }) {
  // Ask the browser not to evict saved decks under storage pressure; harmless if refused.
  navigator.storage?.persist?.().catch(() => {});
  const record = {
    id: `${name}:${bytes.byteLength}`,
    name,
    title,
    slug,
    author,
    slides,
    bytes: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    savedAt: Date.now(),
  };
  await run('readwrite', (s) => s.put(record));
  return record.id;
}

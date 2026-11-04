/**
 * ZIP helpers for the browser converter: read the .pptx (inflate every part) and
 * write the deck folder back out as a ZIP. No dependencies.
 */
import { listZipEntries } from './pptx-core.mjs';

export async function inflateRaw(bytes) {
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** Inflate every part up front, so the converter itself stays synchronous. */
export async function readPptx(bytes) {
  const files = new Map();
  for (const entry of listZipEntries(bytes)) {
    const raw = bytes.subarray(entry.dataStart, entry.dataStart + entry.compSize);
    if (entry.method === 0) files.set(entry.name, raw);
    else if (entry.method === 8) files.set(entry.name, await inflateRaw(raw));
    else throw new Error(`Unsupported compression in ${entry.name}`);
  }
  return files;
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[i] = c >>> 0;
  }
  return table;
})();

export const crc32 = (bytes) => {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
};

/** A ZIP with everything stored uncompressed; the payload is mostly PNG and SVG. */
export function makeZip(entries) {
  const encoder = new TextEncoder();
  const chunks = [];
  const central = [];
  let offset = 0;

  for (const { name, bytes } of entries) {
    const nameBytes = encoder.encode(name);
    const crc = crc32(bytes);

    const local = new Uint8Array(30 + nameBytes.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);
    lv.setUint16(6, 0x0800, true); // UTF-8 names
    lv.setUint32(14, crc, true);
    lv.setUint32(18, bytes.length, true);
    lv.setUint32(22, bytes.length, true);
    lv.setUint16(26, nameBytes.length, true);
    local.set(nameBytes, 30);
    chunks.push(local, bytes);

    const dir = new Uint8Array(46 + nameBytes.length);
    const dv = new DataView(dir.buffer);
    dv.setUint32(0, 0x02014b50, true);
    dv.setUint16(4, 20, true);
    dv.setUint16(6, 20, true);
    dv.setUint16(8, 0x0800, true);
    dv.setUint32(16, crc, true);
    dv.setUint32(20, bytes.length, true);
    dv.setUint32(24, bytes.length, true);
    dv.setUint16(28, nameBytes.length, true);
    dv.setUint32(42, offset, true);
    dir.set(nameBytes, 46);
    central.push(dir);

    offset += local.length + bytes.length;
  }

  const centralSize = central.reduce((n, c) => n + c.length, 0);
  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, offset, true);

  return new Blob([...chunks, ...central, end], { type: 'application/zip' });
}

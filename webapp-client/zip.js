// Minimal ZIP reader/writer for .docx manipulation, browser-only (uses
// CompressionStream/DecompressionStream 'deflate-raw'), no dependencies.

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(bytes) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) {
    c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
  }
  return (c ^ 0xFFFFFFFF) >>> 0;
}

async function deflateRaw(bytes) {
  const cs = new CompressionStream('deflate-raw');
  const writer = cs.writable.getWriter();
  writer.write(bytes);
  writer.close();
  const chunks = [];
  const reader = cs.readable.getReader();
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
  }
  const total = chunks.reduce((s, c) => s + c.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const c of chunks) { out.set(c, off); off += c.length; }
  return out;
}

async function inflateRaw(bytes) {
  const ds = new DecompressionStream('deflate-raw');
  const writer = ds.writable.getWriter();
  writer.write(bytes);
  writer.close();
  const chunks = [];
  const reader = ds.readable.getReader();
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
  }
  const total = chunks.reduce((s, c) => s + c.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const c of chunks) { out.set(c, off); off += c.length; }
  return out;
}

function readU16(dv, off) { return dv.getUint16(off, true); }
function readU32(dv, off) { return dv.getUint32(off, true); }

// Parses the End Of Central Directory + Central Directory of a ZIP buffer.
function parseZip(buffer) {
  const bytes = new Uint8Array(buffer);
  const dv = new DataView(buffer);
  // find EOCD signature 0x06054b50 scanning from the end
  let eocdOff = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 22 - 65536); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocdOff = i; break; }
  }
  if (eocdOff === -1) throw new Error('EOCD not found - not a valid zip');
  const cdCount = readU16(dv, eocdOff + 10);
  const cdOffset = readU32(dv, eocdOff + 16);

  const entries = [];
  let p = cdOffset;
  for (let i = 0; i < cdCount; i++) {
    const sig = readU32(dv, p);
    if (sig !== 0x02014b50) throw new Error('Bad central directory signature at ' + p);
    const method = readU16(dv, p + 10);
    const modTime = readU16(dv, p + 12);
    const modDate = readU16(dv, p + 14);
    const crc = readU32(dv, p + 16);
    const compSize = readU32(dv, p + 20);
    const uncompSize = readU32(dv, p + 24);
    const nameLen = readU16(dv, p + 28);
    const extraLen = readU16(dv, p + 30);
    const commentLen = readU16(dv, p + 32);
    const localHeaderOffset = readU32(dv, p + 42);
    const nameBytes = bytes.slice(p + 46, p + 46 + nameLen);
    const name = new TextDecoder('utf-8').decode(nameBytes);
    entries.push({ name, method, modTime, modDate, crc, compSize, uncompSize, localHeaderOffset });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return { entries, bytes };
}

// Returns the raw (compressed) data bytes for an entry, by reading its local
// file header to find the true data start (extra field length can differ
// from the central directory's).
function rawEntryData(bytes, dv, entry) {
  const off = entry.localHeaderOffset;
  const sig = readU32(dv, off);
  if (sig !== 0x04034b50) throw new Error('Bad local file header at ' + off);
  const nameLen = readU16(dv, off + 26);
  const extraLen = readU16(dv, off + 28);
  const dataStart = off + 30 + nameLen + extraLen;
  return bytes.slice(dataStart, dataStart + entry.compSize);
}

async function readEntryText(zip, name) {
  const entry = zip.entries.find(e => e.name === name);
  if (!entry) throw new Error('Entry not found: ' + name);
  const dv = new DataView(zip.bytes.buffer, zip.bytes.byteOffset, zip.bytes.byteLength);
  const raw = rawEntryData(zip.bytes, dv, entry);
  const data = entry.method === 0 ? raw : await inflateRaw(raw);
  return new TextDecoder('utf-8').decode(data);
}

function dosDateTime() {
  // fixed, arbitrary valid DOS date/time (2024-01-01 00:00:00) - good enough,
  // nothing in the pipeline depends on real timestamps
  return { time: 0, date: (44 << 9) | (1 << 5) | 1 };
}

// Rebuilds the zip, replacing the UNCOMPRESSED content of any entry named in
// `replacements` (Map<name, Uint8Array>); all other entries are copied
// byte-for-byte (local header + compressed data) from the original.
async function buildZip(zip, replacements) {
  const dv = new DataView(zip.bytes.buffer, zip.bytes.byteOffset, zip.bytes.byteLength);
  const parts = [];
  const newCentral = [];
  let offset = 0;

  for (const entry of zip.entries) {
    const nameBytes = new TextEncoder().encode(entry.name);
    if (replacements.has(entry.name)) {
      const uncompressed = replacements.get(entry.name);
      const compressed = await deflateRaw(uncompressed);
      const newCrc = crc32(uncompressed);
      const { time, date } = dosDateTime();

      const header = new Uint8Array(30 + nameBytes.length);
      const hdv = new DataView(header.buffer);
      hdv.setUint32(0, 0x04034b50, true);
      hdv.setUint16(4, 20, true);           // version needed
      hdv.setUint16(6, 0, true);            // flags
      hdv.setUint16(8, 8, true);            // method = deflate
      hdv.setUint16(10, time, true);
      hdv.setUint16(12, date, true);
      hdv.setUint32(14, newCrc, true);
      hdv.setUint32(18, compressed.length, true);
      hdv.setUint32(22, uncompressed.length, true);
      hdv.setUint16(26, nameBytes.length, true);
      hdv.setUint16(28, 0, true);           // extra len
      header.set(nameBytes, 30);

      parts.push(header, compressed);
      newCentral.push({
        name: entry.name, method: 8, modTime: time, modDate: date, crc: newCrc,
        compSize: compressed.length, uncompSize: uncompressed.length, localHeaderOffset: offset,
      });
      offset += header.length + compressed.length;
    } else {
      const off = entry.localHeaderOffset;
      const nameLen = readU16(dv, off + 26);
      const extraLen = readU16(dv, off + 28);
      const total = 30 + nameLen + extraLen + entry.compSize;
      const raw = zip.bytes.slice(off, off + total);
      parts.push(raw);
      newCentral.push({ ...entry, localHeaderOffset: offset });
      offset += raw.length;
    }
  }

  const cdStart = offset;
  for (const e of newCentral) {
    const nameBytes = new TextEncoder().encode(e.name);
    const rec = new Uint8Array(46 + nameBytes.length);
    const rdv = new DataView(rec.buffer);
    rdv.setUint32(0, 0x02014b50, true);
    rdv.setUint16(4, 20, true);   // version made by
    rdv.setUint16(6, 20, true);   // version needed
    rdv.setUint16(8, 0, true);    // flags
    rdv.setUint16(10, e.method, true);
    rdv.setUint16(12, e.modTime, true);
    rdv.setUint16(14, e.modDate, true);
    rdv.setUint32(16, e.crc, true);
    rdv.setUint32(20, e.compSize, true);
    rdv.setUint32(24, e.uncompSize, true);
    rdv.setUint16(28, nameBytes.length, true);
    rdv.setUint16(30, 0, true);   // extra len
    rdv.setUint16(32, 0, true);   // comment len
    rdv.setUint16(34, 0, true);   // disk number
    rdv.setUint16(36, 0, true);   // internal attrs
    rdv.setUint32(38, 0, true);   // external attrs
    rdv.setUint32(42, e.localHeaderOffset, true);
    rec.set(nameBytes, 46);
    parts.push(rec);
    offset += rec.length;
  }
  const cdSize = offset - cdStart;

  const eocd = new Uint8Array(22);
  const edv = new DataView(eocd.buffer);
  edv.setUint32(0, 0x06054b50, true);
  edv.setUint16(4, 0, true);
  edv.setUint16(6, 0, true);
  edv.setUint16(8, newCentral.length, true);
  edv.setUint16(10, newCentral.length, true);
  edv.setUint32(12, cdSize, true);
  edv.setUint32(16, cdStart, true);
  edv.setUint16(20, 0, true);
  parts.push(eocd);

  const totalLen = parts.reduce((s, p) => s + p.length, 0);
  const out = new Uint8Array(totalLen);
  let o = 0;
  for (const part of parts) { out.set(part, o); o += part.length; }
  return out;
}

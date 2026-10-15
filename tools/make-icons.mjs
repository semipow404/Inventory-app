// One-off: writes icon-192.png and icon-512.png (navy square with a white 2×2 stillage). Run: node tools/make-icons.mjs
import { deflateSync } from 'node:zlib';
import { writeFileSync } from 'node:fs';

const table = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
const crc = buf => { let c = 0xffffffff; for (const b of buf) c = table[(c ^ b) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
const chunk = (type, data) => {
  const body = Buffer.concat([Buffer.from(type), data]);
  const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
  const sum = Buffer.alloc(4); sum.writeUInt32BE(crc(body));
  return Buffer.concat([len, body, sum]);
};

function png(size) {
  const m = size * 0.2, s = size * 0.6, t = Math.max(2, size / 40);
  const raw = Buffer.alloc(size * (size * 3 + 1)); // each scanline: filter byte 0, then RGB pixels
  for (let y = 0; y < size; y++) for (let x = 0; x < size; x++) {
    const inBox = x >= m - t && x <= m + s + t && y >= m - t && y <= m + s + t;
    const line = inBox && [m, m + s / 2, m + s].some(v => Math.abs(x - v) < t || Math.abs(y - v) < t);
    raw.set(line ? [255, 255, 255] : [0, 5, 41], y * (size * 3 + 1) + 1 + x * 3);
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0); ihdr.writeUInt32BE(size, 4); ihdr[8] = 8; ihdr[9] = 2; // 8-bit RGB
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

for (const size of [192, 512]) writeFileSync(`icon-${size}.png`, png(size));

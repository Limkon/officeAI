import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    c ^= buf[i];
    for (let k = 0; k < 8; k++) {
      c = (c >>> 1) ^ (c & 1 ? 0xedb88320 : 0);
    }
  }
  return (c ^ 0xffffffff) >>> 0;
}

function pngChunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const typeBuf = Buffer.from(type, 'ascii');
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(crc32(Buffer.concat([typeBuf, data])), 0);
  return Buffer.concat([len, typeBuf, data, crcBuf]);
}

function makeIconPng(size) {
  const raw = Buffer.alloc((size * 4 + 1) * size);
  const cx = size / 2;
  const cy = size / 2;
  const radius = size * 0.44;

  for (let y = 0; y < size; y++) {
    const rowStart = y * (size * 4 + 1);
    raw[rowStart] = 0; // filter None
    for (let x = 0; x < size; x++) {
      const idx = rowStart + 1 + x * 4;
      const dx = Math.abs(x + 0.5 - cx);
      const dy = Math.abs(y + 0.5 - cy);
      const cornerR = size * 0.22;
      const innerHalf = size * 0.46 - cornerR;

      let inside = true;
      if (dx > innerHalf && dy > innerHalf) {
        const dist = Math.hypot(dx - innerHalf, dy - innerHalf);
        inside = dist <= cornerR;
      } else if (dx > size * 0.46 || dy > size * 0.46) {
        inside = false;
      }

      if (!inside) {
        raw[idx] = 0;
        raw[idx + 1] = 0;
        raw[idx + 2] = 0;
        raw[idx + 3] = 0;
        continue;
      }

      // Indigo -> Blue gradient background
      const t = (x + y) / (2 * size);
      let r = Math.round(59 + (99 - 59) * t);
      let g = Math.round(130 + (102 - 130) * t);
      let b = Math.round(246 + (241 - 246) * t);

      // Draw a crisp white sparkle/AI star in the center
      const nx = (x + 0.5 - cx) / radius;
      const ny = (y + 0.5 - cy) / radius;
      const star = Math.pow(Math.abs(nx), 0.58) + Math.pow(Math.abs(ny), 0.58);
      if (star < 0.72) {
        r = 255;
        g = 255;
        b = 255;
      }

      raw[idx] = r;
      raw[idx + 1] = g;
      raw[idx + 2] = b;
      raw[idx + 3] = 255;
    }
  }

  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // RGBA
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const idat = zlib.deflateSync(raw);
  return Buffer.concat([
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', idat),
    pngChunk('IEND', Buffer.alloc(0))
  ]);
}

const targets = [
  'packages/office-addin/assets',
  'packages/wps-addin/assets',
  'packages/desktop-app/assets'
];

for (const dir of targets) {
  fs.mkdirSync(dir, { recursive: true });
  for (const size of [16, 32, 80, 256]) {
    fs.writeFileSync(path.join(dir, `icon-${size}.png`), makeIconPng(size));
  }
}

console.log('Generated PNG icons for office-addin, wps-addin, and desktop-app.');

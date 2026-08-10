'use strict';

const zlib = require('zlib');

/**
 * Minimal dependency-free PNG decoder — just enough to perceptually hash
 * screenshots so the crawler can tell "same screen" from "new screen" without
 * pulling in a native image library.
 *
 * Supports 8-bit non-interlaced greyscale / RGB / RGBA, which is what both
 * `adb exec-out screencap -p` and Unity's ScreenCapture produce.
 */

const CHANNELS = { 0: 1, 2: 3, 4: 2, 6: 4 };

function decode(buffer) {
  if (buffer.length < 8 || buffer[0] !== 0x89 || buffer.toString('ascii', 1, 4) !== 'PNG') {
    throw new Error('not a PNG');
  }
  let offset = 8;
  let width = 0;
  let height = 0;
  let bitDepth = 0;
  let colorType = 0;
  let interlace = 0;
  const idat = [];

  while (offset + 8 <= buffer.length) {
    const len = buffer.readUInt32BE(offset);
    const type = buffer.toString('ascii', offset + 4, offset + 8);
    const start = offset + 8;
    if (type === 'IHDR') {
      width = buffer.readUInt32BE(start);
      height = buffer.readUInt32BE(start + 4);
      bitDepth = buffer[start + 8];
      colorType = buffer[start + 9];
      interlace = buffer[start + 12];
    } else if (type === 'IDAT') {
      idat.push(buffer.subarray(start, start + len));
    } else if (type === 'IEND') {
      break;
    }
    offset = start + len + 4; // + CRC
  }

  if (bitDepth !== 8) throw new Error(`unsupported PNG bit depth ${bitDepth}`);
  if (interlace !== 0) throw new Error('interlaced PNG not supported');
  const channels = CHANNELS[colorType];
  if (!channels) throw new Error(`unsupported PNG color type ${colorType}`);

  const raw = zlib.inflateSync(Buffer.concat(idat));
  const stride = width * channels;
  const out = Buffer.alloc(height * stride);

  let pos = 0;
  for (let y = 0; y < height; y++) {
    const filter = raw[pos++];
    const line = raw.subarray(pos, pos + stride);
    pos += stride;
    const cur = out.subarray(y * stride, (y + 1) * stride);
    const prev = y > 0 ? out.subarray((y - 1) * stride, y * stride) : null;

    for (let x = 0; x < stride; x++) {
      const a = x >= channels ? cur[x - channels] : 0;
      const b = prev ? prev[x] : 0;
      const c = prev && x >= channels ? prev[x - channels] : 0;
      const v = line[x];
      let val;
      switch (filter) {
        case 0: val = v; break;
        case 1: val = v + a; break;
        case 2: val = v + b; break;
        case 3: val = v + ((a + b) >> 1); break;
        case 4: {
          const p = a + b - c;
          const pa = Math.abs(p - a);
          const pb = Math.abs(p - b);
          const pc = Math.abs(p - c);
          val = v + (pa <= pb && pa <= pc ? a : pb <= pc ? b : c);
          break;
        }
        default: throw new Error(`unknown PNG filter ${filter}`);
      }
      cur[x] = val & 0xff;
    }
  }

  return { width, height, channels, data: out };
}

/** Greyscale value at (x, y). */
function grey(img, x, y) {
  const i = (y * img.width + x) * img.channels;
  if (img.channels <= 2) return img.data[i];
  return (img.data[i] * 299 + img.data[i + 1] * 587 + img.data[i + 2] * 114) / 1000;
}

/**
 * 16x16 difference hash → 256-bit fingerprint as a hex string.
 * Tolerant of animation and antialiasing, sensitive to layout changes.
 */
function perceptualHash(pngBuffer, size = 16) {
  const img = decode(pngBuffer);
  const w = size + 1;
  const h = size;
  const cells = new Float64Array(w * h);

  // box-average downsample
  for (let cy = 0; cy < h; cy++) {
    const y0 = Math.floor((cy * img.height) / h);
    const y1 = Math.max(y0 + 1, Math.floor(((cy + 1) * img.height) / h));
    for (let cx = 0; cx < w; cx++) {
      const x0 = Math.floor((cx * img.width) / w);
      const x1 = Math.max(x0 + 1, Math.floor(((cx + 1) * img.width) / w));
      let sum = 0;
      let n = 0;
      const stepY = Math.max(1, Math.floor((y1 - y0) / 8));
      const stepX = Math.max(1, Math.floor((x1 - x0) / 8));
      for (let y = y0; y < y1; y += stepY) {
        for (let x = x0; x < x1; x += stepX) {
          sum += grey(img, x, y);
          n++;
        }
      }
      cells[cy * w + cx] = n ? sum / n : 0;
    }
  }

  let bits = '';
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < size; x++) {
      bits += cells[y * w + x] > cells[y * w + x + 1] ? '1' : '0';
    }
  }
  let hex = '';
  for (let i = 0; i < bits.length; i += 4) {
    hex += parseInt(bits.slice(i, i + 4), 2).toString(16);
  }
  return hex;
}

function hammingDistance(a, b) {
  if (!a || !b || a.length !== b.length) return Number.MAX_SAFE_INTEGER;
  let d = 0;
  for (let i = 0; i < a.length; i++) {
    let x = parseInt(a[i], 16) ^ parseInt(b[i], 16);
    while (x) {
      d += x & 1;
      x >>= 1;
    }
  }
  return d;
}

module.exports = { decode, perceptualHash, hammingDistance };

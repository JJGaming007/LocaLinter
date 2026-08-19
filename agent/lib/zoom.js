'use strict';

/**
 * Cutting a string out of a screenshot and blowing it up.
 *
 * The route map has told the model to "crop before calling" a defect since the
 * Thai pass, because at full-screen size a Thai tone mark or a Vietnamese hook
 * is a handful of pixels and it is the easiest thing in the build to imagine.
 * That instruction was unfollowable: the model only ever received the whole
 * 2400x1080 frame. Driving it by hand on 2026-08-19 the rule killed two
 * candidate findings outright and confirmed a third, which is the whole
 * difference between a report a translator trusts and one they stop reading.
 *
 * So the crop happens here, and the model is shown the string the size a person
 * would zoom it to before deciding.
 *
 * Pure JS on purpose. The agent ships as a single executable and inside an
 * Electron app, and a native image library would have to be rebuilt for both.
 */

const { PNG } = require('pngjs');

/** Nearest-neighbour, deliberately: smoothing invents pixels, and a diacritic is a few pixels. */
function upscale(src, factor) {
  const out = new PNG({ width: src.width * factor, height: src.height * factor });
  for (let y = 0; y < out.height; y++) {
    const sy = (y / factor) | 0;
    for (let x = 0; x < out.width; x++) {
      const sx = (x / factor) | 0;
      const s = (sy * src.width + sx) << 2;
      const d = (y * out.width + x) << 2;
      out.data[d] = src.data[s];
      out.data[d + 1] = src.data[s + 1];
      out.data[d + 2] = src.data[s + 2];
      out.data[d + 3] = src.data[s + 3];
    }
  }
  return out;
}

/**
 * Crop a normalised rect out of a PNG and magnify it.
 *
 * The rect comes from the model and is therefore approximate, so it is padded
 * generously: a mark clipped off the top of the crop is exactly the mistake
 * this exists to prevent, and context above and below is what makes a "does it
 * fit inside its button" question answerable at all.
 *
 * @param {Buffer} png            the full screenshot
 * @param {{x,y,w,h}} rect        normalised 0-1, centre-less (x,y = top-left)
 * @param {object} [opts]
 * @returns {{buffer: Buffer, width: number, height: number, factor: number}|null}
 */
// Target width for the magnified crop. A full 2400x1080 frame is downscaled
// before the model ever sees it, so a diacritic in it survives as a couple of
// pixels; the crop is sent close to the largest size that is still processed
// efficiently, which is where the mark becomes unmistakable.
const TARGET_WIDTH = 900;
const MAX_SIDE = 1500;

function cropAndZoom(png, rect, { pad = 0.6, minWidth = TARGET_WIDTH, maxFactor = 8, maxPixels = MAX_SIDE } = {}) {
  if (!png || !rect) return null;
  let img;
  try {
    img = PNG.sync.read(png);
  } catch {
    return null;                      // not a PNG we can read; skip rather than throw
  }

  const nx = Number(rect.x), ny = Number(rect.y);
  const nw = Number(rect.w), nh = Number(rect.h);
  if (![nx, ny, nw, nh].every(Number.isFinite) || nw <= 0 || nh <= 0) return null;

  // Pad by a share of the rect's own size, with a floor so a very small rect
  // still gets usable surroundings.
  const padX = Math.max(nw * pad, 0.02);
  const padY = Math.max(nh * pad, 0.02);

  const x0 = Math.max(0, Math.round((nx - padX) * img.width));
  const y0 = Math.max(0, Math.round((ny - padY) * img.height));
  const x1 = Math.min(img.width, Math.round((nx + nw + padX) * img.width));
  const y1 = Math.min(img.height, Math.round((ny + nh + padY) * img.height));
  const w = x1 - x0, h = y1 - y0;
  if (w < 8 || h < 8) return null;

  const cut = new PNG({ width: w, height: h });
  PNG.bitblt(img, cut, x0, y0, w, h, 0, 0);

  // Enough magnification that a tone mark is unmistakable, without sending a
  // needlessly huge image.
  let factor = Math.max(1, Math.min(maxFactor, Math.ceil(minWidth / w)));
  while (factor > 1 && (w * factor > maxPixels || h * factor > maxPixels)) factor -= 1;

  const big = factor > 1 ? upscale(cut, factor) : cut;
  return { buffer: PNG.sync.write(big), width: big.width, height: big.height, factor };
}

/** The size of a PNG without decoding all of it. */
function pngSize(buf) {
  if (!buf || buf.length < 24) return null;
  return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) };
}

module.exports = { cropAndZoom, pngSize };

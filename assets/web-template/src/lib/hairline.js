// White-background detector for deck images — the web counterpart of
// render_pptx._is_white_backgrounded: the same eight border samples (corners
// + edge midpoints), the same near-white floor (min channel >= 245). Any
// failure (tainted canvas, decode error) reads as "not white": no hairline
// rather than a wrong one.
export function isWhiteBackgrounded(img) {
  try {
    const canvas = document.createElement('canvas')
    canvas.width = img.naturalWidth
    canvas.height = img.naturalHeight
    const ctx = canvas.getContext('2d', { willReadFrequently: true })
    ctx.drawImage(img, 0, 0)
    const { width: w, height: h } = canvas
    const points = [
      [0, 0], [w - 1, 0], [0, h - 1], [w - 1, h - 1],
      [w >> 1, 0], [w >> 1, h - 1], [0, h >> 1], [w - 1, h >> 1],
    ]
    return points.every(([x, y]) => {
      const [r, g, b] = ctx.getImageData(x, y, 1, 1).data
      return Math.min(r, g, b) >= 245
    })
  } catch {
    return false
  }
}

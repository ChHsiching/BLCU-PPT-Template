// Geometry and typography helpers for the stepped web renderer.
//
// The manifest (src/manifest.json, overwritten by scripts/scaffold_web.py from
// templates/<id>/manifest.json) is the single source of truth for regions and
// typography, exactly as it is for renderer-pptx. This module only translates
// manifest inches/points into stage pixels — it invents no layout of its own.
import manifest from '../manifest.json'

export const SLIDE_W_IN = manifest.slide_size.w
export const SLIDE_H_IN = manifest.slide_size.h

// Fixed 16:9 stage; 1280/13.33 ≈ 7.5/720 ≈ 96 px per inch, so 1 pt = 4/3 px.
export const STAGE_W = 1280
export const STAGE_H = 720

const PX_PER_IN = STAGE_W / SLIDE_W_IN

export function inToPx(v) {
  return v * PX_PER_IN
}

export function ptToPx(pt) {
  return pt * (96 / 72)
}

// Absolute-positioned region box, mirroring a manifest region {x, y, w, h}.
export function regionStyle(region) {
  return {
    position: 'absolute',
    left: inToPx(region.x),
    top: inToPx(region.y),
    width: inToPx(region.w),
    height: inToPx(region.h),
  }
}

// CJK-aware width in the spirit of validate_deck.text_width (East Asian W/F
// chars count 1, the rest 0.5). The ranges cover the CJK blocks realistic deck
// text uses; exotica (emoji, rare Hangul extensions) may count 0.5 here, which
// can only shift the long-title size switch, never a budget check — budgets
// are enforced by the Python validator alone.
const WIDE_RANGES = [
  [0x1100, 0x115f], // Hangul Jamo
  [0x2e80, 0x303e], // CJK radicals, Kangxi, CJK symbols and punctuation
  [0x3041, 0x33ff], // Kana, CJK compatibility
  [0x3400, 0x4dbf], // CJK Extension A
  [0x4e00, 0x9fff], // CJK Unified Ideographs
  [0xa960, 0xa97c], // Hangul Jamo Extended-A
  [0xac00, 0xd7a3], // Hangul syllables
  [0xd7b0, 0xd7fb], // Hangul Jamo Extended-B
  [0xf900, 0xfaff], // CJK compatibility ideographs
  [0xfe30, 0xfe4f], // CJK compatibility forms
  [0xff00, 0xff60], // fullwidth forms
  [0xffe0, 0xffe6],
  [0x20000, 0x2fffd], // CJK Extension B+
  [0x30000, 0x3fffd],
]

function isWide(ch) {
  const cp = ch.codePointAt(0)
  return WIDE_RANGES.some(([lo, hi]) => cp >= lo && cp <= hi)
}

export function textWidth(s) {
  let units = 0
  for (const ch of s) units += isWide(ch) ? 2 : 1
  return units / 2
}

// Long-title size switch, mirroring render_pptx._title_size_pt.
export function titleSizePt(text) {
  const t = manifest.typography.title
  return textWidth(text) > t.long_title_over_chars ? t.long_title_size_pt : t.size_pt
}

// CSS font stack from the manifest's web_fallbacks: latin face first (so latin
// glyphs use it, like the pptx latin/ea split), then the CJK face and its
// declared fallbacks, then a generic family. Local fonts win; the bundled
// Noto families (imported in main.jsx) only serve machines without them.
export function fontStack(face, latin) {
  const fallbacks = manifest.typography.web_fallbacks[face] ?? []
  const generic = fallbacks.some((f) => f.includes('Serif')) ? 'serif' : 'sans-serif'
  return [latin, face, ...fallbacks, generic].filter(Boolean).join(', ')
}

// ---- deck accessors (blocks arrive in reading order) ----

export function blockOf(page, type) {
  return page.blocks.find((b) => b.type === type) ?? null
}

export function blocksOf(page, type) {
  return page.blocks.filter((b) => b.type === type)
}

export function textsOf(page) {
  return blocksOf(page, 'text').map((b) => b.text)
}

export function titleOf(page) {
  return blockOf(page, 'title').text
}

// ---- step model (步进式) ----
//
// A page contributes 1 + (number of list items) steps: the page itself, then
// one step per list item revealed. advance()/back() walk this flat list; the
// player derives {page, revealed} from the global step index.
export function buildSteps(pages) {
  const steps = []
  pages.forEach((page, p) => {
    const reveals = blocksOf(page, 'list').reduce((n, b) => n + b.items.length, 0)
    for (let k = 0; k <= reveals; k++) steps.push({ page: p, revealed: k })
  })
  return steps
}

export function firstStepOfPage(steps, pageIndex) {
  return steps.findIndex((s) => s.page === pageIndex)
}
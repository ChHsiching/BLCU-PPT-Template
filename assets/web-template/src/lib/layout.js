// Geometry and typography helpers for the stepped web renderer.
//
// The manifest (src/manifest.json, overwritten by scripts/scaffold_web.py from
// templates/<id>/manifest.json) is the single source of truth for regions and
// the typography.tokens role system, exactly as it is for renderer-pptx. This
// module only translates manifest inches/points into stage pixels — it
// invents no layout or style of its own (the mirrors of render_pptx helpers
// are noted per export).
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

// The token tree (typography.tokens) — the style single source of truth.
export const TK = manifest.typography.tokens

// ---- role resolution (mirrors render_pptx._role_style / _region_role) ----

// CSS font stack from the manifest's web_fallbacks: the token face first
// (@fontsource bundles it, so every machine renders the same face), the
// declared local fallbacks behind it, generic family last. An explicit latin
// face (only the master's page number carries one) goes in front so latin
// glyphs use it, like the pptx latin/ea split.
export function fontStack(face = TK.face, latin) {
  const fallbacks = manifest.typography.web_fallbacks[face] ?? []
  const generic = fallbacks.some((f) => f.includes('Serif')) ? 'serif' : 'sans-serif'
  return [latin ?? face, face, ...fallbacks, generic].filter(Boolean).join(', ')
}

// A named tokens role -> CSS style: face/size/weight/color all from tokens
// (latin inherits the token family default like render_pptx._role_style).
export function roleStyle(roleName) {
  const role = TK.roles[roleName]
  return {
    fontFamily: fontStack(role.face ?? TK.face, role.face ?? TK.latin_face),
    fontSize: ptToPx(role.size_pt),
    fontWeight: TK.weights[role.weight] ?? TK.weights.regular,
    color: role.color,
  }
}

// Region -> role name via typography.tokens.role_bindings.
export function regionRoleName(archetype, region) {
  return TK.role_bindings[archetype][region]
}

export function regionRole(archetype, region) {
  return roleStyle(regionRoleName(archetype, region))
}

// Title role with the long-title downgrade (over title_long.over_chars in
// CJK-width units -> 28pt; weight/color unchanged), mirroring
// render_pptx._title_style.
export function titleRoleName(text, archetype) {
  const name = regionRoleName(archetype, 'title')
  const long = TK.roles.title_long
  if (name === 'title' && long && textWidth(text) > (long.over_chars ?? Infinity)) {
    return 'title_long'
  }
  return name
}

export function titleRoleStyle(text, archetype) {
  return roleStyle(titleRoleName(text, archetype))
}

// ---- spacing (mirrors render_pptx._styled_paragraph rhythm / _textbox_shape) ----

// Body-flow text carries the rhythm; PowerPoint renders spcPct 150% as the
// measured line_pitch_em and applies space_before to every paragraph, the
// first included (inside the box) — the web mirrors both so line pitch and
// paragraph gaps match the pptx. Ceremonial single-line roles use
// single_pitch_em (spcPct 100%) instead.
export function rhythmLineHeight() {
  return { lineHeight: TK.spacing.line_pitch_em }
}

export function singleLineHeight() {
  return { lineHeight: TK.spacing.single_pitch_em }
}

export function paraBeforePx() {
  return ptToPx(TK.spacing.space_before_pt)
}

// Token textbox insets (0.1in sides / 0.05in ends), like every pptx box.
export function textboxPadding() {
  const s = TK.spacing
  return { padding: `${inToPx(s.textbox_inset_v_in)}px ${inToPx(s.textbox_inset_in)}px` }
}

// ---- emphasis (mirrors render_pptx._split_emphasis) ----

// Split into {text, emph} segments: paired **marker** segments emphasize;
// markers around empty or unpaired content never match and stay literal.
export function splitEmphasis(text) {
  const marker = TK.emphasis?.marker ?? ''
  if (!marker || typeof text !== 'string' || !text.includes(marker)) {
    return [{ text, emph: false }]
  }
  const pattern = new RegExp(
    `${marker.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(.+?)${marker.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}`,
    'gs',
  )
  const out = []
  let last = 0
  for (const m of text.matchAll(pattern)) {
    if (m.index > last) out.push({ text: text.slice(last, m.index), emph: false })
    out.push({ text: m[1], emph: true })
    last = m.index + m[0].length
  }
  if (last < text.length) out.push({ text: text.slice(last), emph: false })
  return out.length ? out : [{ text, emph: false }]
}

export function emphasisStyle() {
  const e = TK.emphasis ?? {}
  return { color: e.color, fontWeight: TK.weights[e.weight] ?? TK.weights.bold }
}

// ---- caption scrim + hairline ----

// One caption line at single-line pitch plus both v-insets — exactly the
// pptx strip height (render_pptx._caption_strip_height_in), resolved through
// the owning archetype's bound caption role like everything else.
export function captionStripPx(roleName = 'caption') {
  const role = TK.roles[roleName]
  return ptToPx(role.size_pt * TK.spacing.single_pitch_em)
    + 2 * inToPx(TK.spacing.textbox_inset_v_in)
}

export function rgba(hex, alphaPct) {
  const h = hex.replace('#', '')
  const [r, g, b] = [0, 2, 4].map((k) => parseInt(h.slice(k, k + 2), 16))
  return `rgba(${r}, ${g}, ${b}, ${alphaPct / 100})`
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

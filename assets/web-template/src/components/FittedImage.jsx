import { useState } from 'react'
import manifest from '../manifest.json'
import {
  regionStyle, inToPx, ptToPx, roleStyle, regionRoleName,
  captionStripPx, rgba, singleLineHeight,
} from '../lib/layout'
import { isWhiteBackgrounded } from '../lib/hairline'
import { assetUrl } from '../lib/assets'

const TK = manifest.typography.tokens
const HAIRLINE = TK.image?.hairline

// Contain-fitted deck image in a manifest slot, mirroring render_pptx.
// _add_fitted_picture: aspect kept, centered, upscaled when the slot is
// larger than the natural size (the pptx pic frame never sits at natural
// size) — the fitted box is computed in px from the natural size at load and
// the img then fills it exactly. White-backgrounded images get the manifest
// hairline (outline: no layout shift, drawn at the edge like the pptx line)
// so they do not dissolve into the white slide; photos get nothing.
function FittedImg({ src, alt, slot, children }) {
  const [fit, setFit] = useState(null)
  const [hairline, setHairline] = useState(false)

  const onLoad = (e) => {
    const img = e.currentTarget
    const scale = Math.min(
      inToPx(slot.w) / img.naturalWidth,
      inToPx(slot.h) / img.naturalHeight,
    )
    setFit({ width: img.naturalWidth * scale, height: img.naturalHeight * scale })
    setHairline(Boolean(HAIRLINE) && isWhiteBackgrounded(img))
  }

  const imgStyle = {
    ...(fit ? { width: '100%', height: '100%' } : {}),
    ...(hairline && HAIRLINE
      ? { outline: `${ptToPx(HAIRLINE.width_pt)}px solid ${HAIRLINE.color}` }
      : {}),
  }

  return (
    <div className="image-fit" style={fit ?? undefined}>
      <img src={src} alt={alt} draggable={false} onLoad={onLoad} style={imgStyle} />
      {children}
    </div>
  )
}

// White-on-black-scrim strip over the fitted image's bottom edge — it rides
// inside the fitted wrapper so it hugs the picture, not the slot
// (render_pptx._add_caption); text style and scrim both resolve through the
// owning archetype's role_binding. The v-inset padding puts the caption line
// where the pptx strip's top inset + centered anchor put it.
function CaptionStrip({ text, archetype }) {
  const roleName = regionRoleName(archetype, 'caption')
  const scrim = TK.roles[roleName].scrim
  return (
    <div
      className="image-caption"
      data-role="caption"
      style={{
        ...roleStyle(roleName),
        ...singleLineHeight(), // the strip height is one 1.2em caption line
        padding: `${inToPx(TK.spacing.textbox_inset_v_in)}px 0`,
        height: captionStripPx(roleName),
        background: rgba(scrim.color, scrim.alpha_pct),
      }}
    >
      <span>{text}</span>
    </div>
  )
}

export function FittedImage({ slot, block, archetype }) {
  const captionable = block.caption && regionRoleName(archetype, 'caption')
  return (
    <div className="image-slot" style={regionStyle(slot)}>
      <FittedImg src={assetUrl(block.path)} alt={block.caption ?? ''} slot={slot}>
        {captionable && <CaptionStrip text={block.caption} archetype={archetype} />}
      </FittedImg>
    </div>
  )
}

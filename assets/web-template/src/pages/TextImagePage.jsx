import manifest from '../manifest.json'
import { TitleBar } from '../components/TitleBar'
import { regionStyle, ptToPx, fontStack, blocksOf, titleOf } from '../lib/layout'
import { assetUrl } from '../lib/assets'

const T = manifest.typography

// FittedSlot: image contain-fitted into a manifest slot (aspect kept,
// centered), mirroring render_pptx._add_fitted_picture; an optional caption
// overlays the slot's internal bottom edge (render_pptx._add_caption).
function FittedSlot({ slot, block }) {
  return (
    <div className="image-slot" style={regionStyle(slot)}>
      <img src={assetUrl(block.path)} alt={block.caption ?? ''} draggable={false} />
      {block.caption && (
        <div
          className="image-caption"
          style={{
            fontFamily: fontStack(T.caption.face),
            fontSize: ptToPx(T.caption.size_pt),
          }}
        >
          {block.caption}
        </div>
      )}
    </div>
  )
}

// text-image: subhead + text/list band + 1-4 images in the manifest slots.
// Text blocks are always visible; list items ("- " prefixed, like the pptx
// text box) reveal one step at a time.
export function TextImagePage({ page, arch, revealed }) {
  const body = T.body
  const subheadBlock = page.blocks.find((b) => b.type === 'subhead')
  const slots = arch.regions.image_slots
  const images = blocksOf(page, 'image')

  const paragraphs = []
  let revealCursor = 0
  for (const block of page.blocks) {
    if (block.type === 'text') {
      paragraphs.push({ text: block.text, visible: true })
    } else if (block.type === 'list') {
      for (const item of block.items) {
        const myReveal = revealCursor++
        paragraphs.push({ text: `- ${item}`, visible: myReveal < revealed })
      }
    }
  }

  // 单图变体：左文右图整版（与 render_pptx.fill_text_image 的单图分支一致）；
  // 多图仍走模板原 slots。
  const single = images.length === 1 && arch.regions.image_primary

  return (
    <>
      <TitleBar text={titleOf(page)} arch={arch} />
      {subheadBlock && (
        <div
          className="region-text subhead"
          style={{
            ...regionStyle(arch.regions.subhead),
            fontFamily: fontStack(body.face),
            fontSize: ptToPx(T.subhead_size_pt),
          }}
        >
          {subheadBlock.text}
        </div>
      )}
      {paragraphs.length > 0 && (
        <div
          className="region-text"
          style={{
            ...regionStyle(single ? arch.regions.text_column : arch.regions.text),
            fontFamily: fontStack(body.face),
            fontSize: ptToPx(body.secondary_size_pt),
          }}
        >
          {paragraphs.map((p, i) => (
            <p key={i} className={p.visible ? 'is-visible' : 'is-hidden'}>
              {p.text}
            </p>
          ))}
        </div>
      )}
      {single ? (
        <FittedSlot slot={arch.regions.image_primary} block={images[0]} />
      ) : (
        images.map((block, i) => <FittedSlot key={i} slot={slots[i]} block={block} />)
      )}
    </>
  )
}

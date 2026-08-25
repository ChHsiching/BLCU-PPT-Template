import manifest from '../manifest.json'
import { TitleBar } from '../components/TitleBar'
import { regionStyle, ptToPx, fontStack, blocksOf, titleOf } from '../lib/layout'

const T = manifest.typography

// Image URL: deck paths are scaffold-relative ("material/images/x.png",
// rewritten by scripts/scaffold_web.py); BASE_URL keeps them correct under
// any Vite base and in the relocatable production build. Each segment is
// percent-encoded so filenames with URL metacharacters (#, ?, spaces, CJK)
// resolve to the right file instead of being parsed as fragment/query.
export function assetUrl(path) {
  const encoded = path.split('/').map(encodeURIComponent).join('/')
  const base = import.meta.env.BASE_URL
  return base.endsWith('/') ? base + encoded : base + '/' + encoded
}

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
            ...regionStyle(arch.regions.text),
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
      {images.map((block, i) => (
        <FittedSlot key={i} slot={slots[i]} block={block} />
      ))}
    </>
  )
}

import { TitleBar } from '../components/TitleBar'
import { FittedImage } from '../components/FittedImage'
import { EmphasisText } from '../components/EmphasisText'
import {
  regionStyle, textboxPadding, singleLineHeight, rhythmLineHeight,
  paraBeforePx, regionRole, regionRoleName, blocksOf, titleOf,
} from '../lib/layout'

// text-image: subhead + text/list band + 1-4 images in the manifest slots.
// Text blocks are always visible; list items ("- " prefixed, like the pptx
// text box) reveal one step at a time. role_bindings map text and text_column
// to the secondary role; the band flows with the rhythm and the emphasis
// convention.
export function TextImagePage({ page, arch, revealed }) {
  const archetype = page.archetype
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
      <TitleBar text={titleOf(page)} arch={arch} archetype={archetype} />
      {subheadBlock && (
        <div
          className="region-text subhead"
          data-role={regionRoleName(archetype, 'subhead')}
          style={{
            ...regionStyle(arch.regions.subhead),
            ...textboxPadding(),
            ...singleLineHeight(),
            ...regionRole(archetype, 'subhead'),
          }}
        >
          <span>{subheadBlock.text}</span>
        </div>
      )}
      {paragraphs.length > 0 && (
        <div
          className="region-text"
          data-role={regionRoleName(archetype, single ? 'text_column' : 'text')}
          data-rhythm="1"
          style={{
            ...regionStyle(single ? arch.regions.text_column : arch.regions.text),
            ...textboxPadding(),
            ...rhythmLineHeight(),
            ...regionRole(archetype, single ? 'text_column' : 'text'),
          }}
        >
          {paragraphs.map((p, i) => (
            <p key={i} className={p.visible ? 'is-visible' : 'is-hidden'}
               style={{ marginTop: paraBeforePx() }}>
              <EmphasisText text={p.text} />
            </p>
          ))}
        </div>
      )}
      {single ? (
        <FittedImage slot={arch.regions.image_primary} block={images[0]} archetype={archetype} />
      ) : (
        images.map((block, i) => (
          <FittedImage key={i} slot={slots[i]} block={block} archetype={archetype} />
        ))
      )}
    </>
  )
}

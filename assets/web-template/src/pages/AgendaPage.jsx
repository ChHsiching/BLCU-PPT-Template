import {
  regionStyle, textboxPadding, singleLineHeight, rhythmLineHeight,
  paraBeforePx, regionRole, regionRoleName, titleOf, blocksOf,
} from '../lib/layout'
import { EmphasisText } from '../components/EmphasisText'

// agenda: centered topic label (the page's title block) + section list, each
// item revealed one step at a time. Hidden items keep their line (no reflow),
// the freshly revealed one fades in. Both boxes carry the roles from
// role_bindings; the list flows with the body rhythm (line pitch + 12pt
// before every item, first included — like the pptx AgendaList box).
export function AgendaPage({ page, arch, revealed }) {
  const items = blocksOf(page, 'list').flatMap((b) => b.items)
  return (
    <>
      <div
        className="region-text agenda-label"
        data-role={regionRoleName(page.archetype, 'label')}
        style={{
          ...regionStyle(arch.regions.label),
          ...textboxPadding(),
          ...singleLineHeight(),
          ...regionRole(page.archetype, 'label'),
        }}
      >
        {titleOf(page)}
      </div>
      {items.length > 0 && (
        <div
          className="region-text agenda-list"
          data-role={regionRoleName(page.archetype, 'list')}
          data-rhythm="1"
          style={{
            ...regionStyle(arch.regions.list),
            // vertically center a sparse list, like the pptx anchor="ctr":
            // bottom-anchored emptiness reads as imbalance
            display: 'flex', flexDirection: 'column', justifyContent: 'center',
            ...textboxPadding(),
            ...rhythmLineHeight(),
            ...regionRole(page.archetype, 'list'),
          }}
        >
          {items.map((item, i) => (
            <div
              key={i}
              className={i < revealed ? 'agenda-item is-visible' : 'agenda-item is-hidden'}
              style={{ marginTop: paraBeforePx() }}
            >
              <EmphasisText text={item} />
            </div>
          ))}
        </div>
      )}
    </>
  )
}

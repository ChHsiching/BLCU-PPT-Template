import manifest from '../manifest.json'
import { regionStyle, ptToPx, fontStack, titleOf, blocksOf } from '../lib/layout'

const T = manifest.typography

// agenda: centered topic label (the page's title block) + section list, each
// item revealed one step at a time. Hidden items keep their line (no reflow),
// the freshly revealed one fades in.
export function AgendaPage({ page, arch, revealed }) {
  const label = T.agenda_label
  const list = T.agenda_list
  const items = blocksOf(page, 'list').flatMap((b) => b.items)
  return (
    <>
      <div
        className="region-text agenda-label"
        style={{
          ...regionStyle(arch.regions.label),
          fontFamily: fontStack(label.face),
          fontSize: ptToPx(label.size_pt),
        }}
      >
        {titleOf(page)}
      </div>
      {items.length > 0 && (
        <div
          className="agenda-list"
          style={{
            ...regionStyle(arch.regions.list),
            fontFamily: fontStack(list.face),
            fontSize: ptToPx(list.size_pt),
          }}
        >
          {items.map((item, i) => (
            <div
              key={i}
              className={i < revealed ? 'agenda-item is-visible' : 'agenda-item is-hidden'}
            >
              {item}
            </div>
          ))}
        </div>
      )}
    </>
  )
}

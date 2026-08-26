import manifest from '../manifest.json'
import {
  regionStyle, inToPx, singleLineHeight, titleRoleName, titleRoleStyle,
} from '../lib/layout'

const INSET_H_IN = manifest.typography.tokens.spacing.textbox_inset_in

// Content-page title bar: the page's title block in the archetype's title
// region, vertically centered like the template's title placeholder; the
// resolved role (title / long downgrade) lands in data-role for the QA gate.
export function TitleBar({ text, arch, archetype }) {
  return (
    <div
      className="title-bar"
      data-role={titleRoleName(text, archetype)}
      style={{
        ...regionStyle(arch.regions.title),
        padding: `0 ${inToPx(INSET_H_IN)}px`,
        ...singleLineHeight(),
        ...titleRoleStyle(text, archetype),
      }}
    >
      {text}
    </div>
  )
}

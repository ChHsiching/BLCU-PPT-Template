import {
  regionStyle, textboxPadding, singleLineHeight, titleRoleName,
  titleRoleStyle, regionRole, regionRoleName, titleOf, textsOf,
} from '../lib/layout'

// cover + closing share one layout (closing is a content variant of cover in
// the manifest): centered title in the upper band, joined presenter line in
// the lower band. Roles resolve through role_bindings (cover/closing both
// bind title -> title with the long downgrade, subtitle -> cover_subtitle).
export function CoverPage({ page, arch }) {
  const title = titleOf(page)
  const subtitle = textsOf(page).join(' ')
  return (
    <>
      <div
        className="region-text cover-title"
        data-role={titleRoleName(title, page.archetype)}
        style={{
          ...regionStyle(arch.regions.title),
          ...textboxPadding(),
          ...singleLineHeight(),
          ...titleRoleStyle(title, page.archetype),
        }}
      >
        {title}
      </div>
      {subtitle && (
        <div
          className="region-text cover-subtitle"
          data-role={regionRoleName(page.archetype, 'subtitle')}
          style={{
            ...regionStyle(arch.regions.subtitle),
            ...textboxPadding(),
            ...singleLineHeight(),
            ...regionRole(page.archetype, 'subtitle'),
          }}
        >
          {subtitle}
        </div>
      )}
    </>
  )
}

import manifest from '../manifest.json'
import { regionStyle } from '../lib/layout'
import { assetUrl } from '../lib/assets'

// 母版品牌层复刻 (#17): bands and logos measured on the slide masters, all
// geometry/media from manifest.brand_layer — this component invents nothing.
// Rendered before the page content (both absolutely positioned), so content
// text draws on top of the bands exactly like the pptx master z-order.
const BL = manifest.brand_layer
const BAND_COLOR = manifest.typography.tokens.colors.band

function BrandBand({ band }) {
  return (
    <div
      className="brand-band"
      aria-hidden="true"
      style={{ ...regionStyle(band), background: BAND_COLOR }}
    />
  )
}

function BrandLogo({ element }) {
  return (
    <img
      className="brand-logo"
      aria-hidden="true"
      draggable={false}
      alt=""
      src={assetUrl(`brand/${element.media}`)}
      style={regionStyle(element)}
    />
  )
}

// master2「标题和内容」: top band under the title + bottom-left corner logo.
export function ContentBrand() {
  return (
    <>
      <BrandBand band={BL.content.top_band} />
      <BrandLogo element={BL.content.corner_logo} />
    </>
  )
}

// master1「三logo标题页」(cover + closing): three top logos, mid band under
// the title, bottom-right logo bar.
export function CoverBrand() {
  return (
    <>
      {BL.cover.logos.map((element, i) => (
        <BrandLogo key={i} element={element} />
      ))}
      <BrandBand band={BL.cover.mid_band} />
      <BrandLogo element={BL.cover.corner_logo_bar} />
    </>
  )
}

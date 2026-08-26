// Asset URL helper for everything under public/ (deck images, brand media).
//
// Deck paths are scaffold-relative ("material/images/x.png", rewritten by
// scripts/scaffold_web.py); manifest brand media is "brand/<name>" (shipped by
// scripts/export_brand_assets.py). BASE_URL keeps them correct under any Vite
// base and in the relocatable production build. Each segment is
// percent-encoded so filenames with URL metacharacters (#, ?, spaces, CJK)
// resolve to the right file instead of being parsed as fragment/query.
export function assetUrl(path) {
  const encoded = path.split('/').map(encodeURIComponent).join('/')
  const base = import.meta.env.BASE_URL
  return base.endsWith('/') ? base + encoded : base + '/' + encoded
}

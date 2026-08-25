import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App'

// Noto fallbacks are bundled (not CDN): the local font stack (SimHei /
// STZhongsong / KaiTi from the manifest) wins when present; these serve
// machines without the Chinese fonts — no missing-glyph boxes either way.
import '@fontsource/noto-sans-sc/400.css'
import '@fontsource/noto-sans-sc/700.css'
import '@fontsource/noto-serif-sc/400.css'
import '@fontsource/noto-serif-sc/700.css'
import './styles.css'

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)

import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App'

// Noto Sans SC is the primary face (typography.tokens.face; @fontsource
// bundles 400/700, woff2 subsets load on demand via unicode-range), so every
// machine renders the same face and weights. The local fallbacks in
// typography.web_fallbacks only serve a blocked-font failure. Noto Serif SC
// backs the master's page-number face (华文中宋), the one legacy face left.
import '@fontsource/noto-sans-sc/400.css'
import '@fontsource/noto-sans-sc/700.css'
import '@fontsource/noto-serif-sc/400.css'
import './styles.css'

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)

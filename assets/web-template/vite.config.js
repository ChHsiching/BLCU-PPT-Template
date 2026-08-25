import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// base './' keeps the production build relocatable: dist/ opens from any
// subpath (or file://) without rewriting asset URLs.
export default defineConfig({
  base: './',
  plugins: [react()],
})

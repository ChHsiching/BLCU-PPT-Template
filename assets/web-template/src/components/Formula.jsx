import { useMemo } from 'react'
import katex from 'katex'
import 'katex/dist/katex.min.css'

// Display formula via KaTeX (the web counterpart of the pptx OMML chain).
// throwOnError: false renders a red error inline instead of crashing a slide;
// the try/catch is a backstop for inputs KaTeX rejects outright.
export function Formula({ latex }) {
  const html = useMemo(() => {
    try {
      return katex.renderToString(latex, { displayMode: true, throwOnError: false })
    } catch {
      return `<span style="color:#c00">${latex.replace(/[<>&]/g, '')}</span>`
    }
  }, [latex])
  return <div className="formula" dangerouslySetInnerHTML={{ __html: html }} />
}

import type { DocMetadata } from '../types/docMetadata.js'

/**
 * Tiptap HTML → Word-friendly HTML enhancement for export.
 * Wraps HTML in a full document, injects @page rules, font declarations, etc.
 */
export function enhanceHtmlForExport(html: string, metadata?: Partial<DocMetadata>): string {
  const pageWidth = metadata?.paperSize?.width ?? 210
  const pageHeight = metadata?.paperSize?.height ?? 297
  const margins = metadata?.margins ?? { top: 25.4, bottom: 25.4, left: 31.8, right: 31.8 }
  const defaultFont = metadata?.defaultFont ?? '宋体'
  const defaultFontSize = metadata?.defaultFontSize ?? 12

  const pageStyle = `
    @page {
      size: ${pageWidth}mm ${pageHeight}mm;
      margin: ${margins.top}mm ${margins.right}mm ${margins.bottom}mm ${margins.left}mm;
    }
  `

  const bodyStyle = `
    body {
      font-family: "${defaultFont}", "SimSun", serif;
      font-size: ${defaultFontSize}pt;
      line-height: 1.5;
      color: #000000;
    }
  `

  const tableStyle = `
    table {
      border-collapse: collapse;
      width: 100%;
    }
    td, th {
      border: 1px solid #000000;
      padding: 4px 8px;
    }
  `

  const imgStyle = `
    img {
      max-width: 100%;
      height: auto;
    }
  `

  let enhancedBody = enhanceTableBorders(html)
  enhancedBody = ensureImageAbsoluteSizes(enhancedBody)
  enhancedBody = normalizeParagraphSpacing(enhancedBody)

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
${pageStyle}
${bodyStyle}
${tableStyle}
${imgStyle}
</style>
</head>
<body>
${enhancedBody}
</body>
</html>`
}

function enhanceTableBorders(html: string): string {
  return html.replace(
    /<table([^>]*)>/gi,
    (match, attrs: string) => {
      if (!attrs.includes('border')) {
        return `<table${attrs} border="1" cellpadding="4">`
      }
      return match
    }
  )
}

function ensureImageAbsoluteSizes(html: string): string {
  return html.replace(/<img([^>]*)>/gi, (match, attrs: string) => {
    if (attrs.includes('blob:')) {
      console.warn('Export HTML contains blob: URL image, which will not render in LO')
    }
    return match
  })
}

function normalizeParagraphSpacing(html: string): string {
  return html
}

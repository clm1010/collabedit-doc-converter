/**
 * LO HTML → Tiptap HTML post-processing
 * Cleans up LibreOffice HTML output for Tiptap editor consumption.
 */
export function postProcessLoHtml(html: string): string {
  let result = html

  result = stripMetaAndGeneratedStyles(result)
  result = normalizeFontSizes(result)
  result = normalizeImages(result)
  result = normalizeTableStructure(result)
  result = normalizeParagraphStyles(result)
  result = preserveRgbaAlpha(result)

  return result
}

function stripMetaAndGeneratedStyles(html: string): string {
  let result = html.replace(/<meta[^>]*>/gi, '')
  result = result.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
  result = result.replace(/<!--[\s\S]*?-->/g, '')
  result = result.replace(/<\/?html[^>]*>/gi, '')
  result = result.replace(/<\/?head[^>]*>/gi, '')
  result = result.replace(/<\/?body[^>]*>/gi, '')
  return result.trim()
}

function normalizeFontSizes(html: string): string {
  return html.replace(/font-size:\s*([\d.]+)pt/gi, (_, pt) => {
    const px = Math.round(parseFloat(pt) * 1.333)
    return `font-size: ${px}px`
  })
}

function normalizeImages(html: string): string {
  return html.replace(/<img([^>]*)>/gi, (match, attrs: string) => {
    let processed = attrs
    if (!processed.includes('style=')) {
      processed += ' style=""'
    }
    if (!processed.includes('width') && !processed.includes('max-width')) {
      processed = processed.replace(/style="/, 'style="max-width: 100%; ')
    }
    return `<img${processed}>`
  })
}

function normalizeTableStructure(html: string): string {
  let result = html

  result = result.replace(/<table([^>]*)>/gi, (match, attrs: string) => {
    if (!attrs.includes('style=')) {
      return `<table${attrs} style="border-collapse: collapse; width: 100%;">`
    }
    return match
  })

  result = result.replace(/<td([^>]*)>/gi, (match, attrs: string) => {
    if (attrs.includes('style=') && !attrs.includes('border')) {
      return match.replace(/style="/, 'style="border: 1px solid #d0d0d0; padding: 4px 8px; ')
    }
    if (!attrs.includes('style=')) {
      return `<td${attrs} style="border: 1px solid #d0d0d0; padding: 4px 8px;">`
    }
    return match
  })

  return result
}

function normalizeParagraphStyles(html: string): string {
  return html.replace(/<p([^>]*)>/gi, (match, attrs: string) => {
    if (!attrs.includes('style=')) return match
    let style = attrs.match(/style="([^"]*)"/)?.[1] ?? ''
    style = style.replace(/margin:\s*0[^;]*(;|$)/g, '')
    return `<p${attrs.replace(/style="[^"]*"/, `style="${style}"`)}>` 
  })
}

function preserveRgbaAlpha(html: string): string {
  return html
}

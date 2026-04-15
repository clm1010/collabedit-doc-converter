import juice from 'juice'

/**
 * LO HTML → Tiptap HTML post-processing
 * Cleans up LibreOffice HTML output for Tiptap editor consumption.
 */
export function postProcessLoHtml(html: string): string {
  let result = html

  const hadStyleBlock = /<style[^>]*>/i.test(result)
  const styleMatch = result.match(/<style[^>]*>([\s\S]*?)<\/style>/i)
  if (styleMatch) {
    console.log(`[postprocess] style block sample (first 800 chars):\n${styleMatch[1].substring(0, 800)}`)
  }

  result = cleanLoDefaultCss(result)

  const beforeLen = result.length
  result = inlineCssWithJuice(result)
  const afterLen = result.length
  console.log(`[postprocess] juice: hadStyleBlock=${hadStyleBlock}, len ${beforeLen} → ${afterLen} (${afterLen > beforeLen ? '+' : ''}${afterLen - beforeLen})`)

  const styledSpans = result.match(/<span[^>]*style="[^"]{5,}"[^>]*>/gi)
  if (styledSpans) {
    console.log(`[postprocess] sample styled spans after juice (first 5):`)
    for (const el of styledSpans.slice(0, 5)) {
      console.log(`  ${el.substring(0, 250)}`)
    }
  }

  result = stripMetaAndGeneratedStyles(result)
  result = normalizeFontSizes(result)
  result = normalizeImages(result)
  result = normalizeTableStructure(result)
  result = normalizeParagraphStyles(result)

  return result
}

/**
 * Clean LO's default CSS before juice inlining.
 *
 * LO generates two kinds of CSS rules:
 * 1. Tag-level: p.western, h1.western etc. — LO defaults (Arial font, theme colors)
 *    These have higher specificity (0,1,1) than document classes.
 * 2. Class-level: .T1, .T2, .P1 etc. — actual document formatting (correct fonts/colors)
 *    These have lower specificity (0,1,0).
 *
 * Because of this specificity mismatch, juice inlines the WRONG values.
 * Fix: strip font-family, font-size, color from tag-level selectors before juice,
 * so only document-specific class-level styles get inlined.
 */
function cleanLoDefaultCss(html: string): string {
  return html.replace(/<style[^>]*>([\s\S]*?)<\/style>/gi, (fullMatch, css: string) => {
    const cleaned = css.replace(
      /([^{}]+)\{([^}]+)\}/g,
      (_rule, selector: string, decls: string) => {
        const sel = selector.trim()
        if (/^[a-z]/i.test(sel)) {
          const filtered = decls
            .split(';')
            .filter(d => {
              const prop = d.trim().toLowerCase().split(':')[0]?.trim()
              return prop !== 'font-family' && prop !== 'font-size' && prop !== 'color'
            })
            .join(';')
          return `${selector}{${filtered}}`
        }
        return `${selector}{${decls}}`
      }
    )
    return fullMatch.replace(css, cleaned)
  })
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

function inlineCssWithJuice(html: string): string {
  try {
    return juice(html, {
      removeStyleTags: false,
      preserveMediaQueries: false,
      preserveFontFaces: false,
      applyStyleTags: true,
      applyAttributeStyles: true
    })
  } catch (err) {
    console.warn('[postprocess] juice inlining failed, skipping:', err)
    return html
  }
}

/**
 * Validates and processes images in HTML for export.
 * Ensures all images are base64-encoded since LO container can't access external URLs.
 */
export interface ImageValidationResult {
  html: string
  warnings: string[]
}

export function validateAndProcessImages(html: string): ImageValidationResult {
  const warnings: string[] = []
  let processed = html

  processed = processed.replace(/<img([^>]*)src="([^"]*)"([^>]*)>/gi, (match, pre, src: string, post) => {
    if (src.startsWith('data:image/')) {
      return match
    }

    if (src.startsWith('blob:')) {
      warnings.push(`Image with blob: URL detected — should be converted to base64 before export`)
      return match
    }

    if (src.startsWith('http://') || src.startsWith('https://')) {
      warnings.push(`External image URL detected: ${src.substring(0, 80)}... — may not render in exported document`)
      return match
    }

    warnings.push(`Unknown image source format: ${src.substring(0, 50)}`)
    return match
  })

  return { html: processed, warnings }
}

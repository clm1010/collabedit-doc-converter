import type { DocxArchive } from './zipExtractor.js'
import { parseXml, ensureArray, getAttr, getWVal } from './xmlParser.js'

/**
 * 检测文档是否为红头文档（altChunk 特征或红线/红字特征）
 *
 * 红头文档特征：
 * 1. 包含 w:altChunk 引用（嵌入的子文档）
 * 2. 开头段落包含红色大字标题
 * 3. 包含红色底边框横线
 */
export function detectRedHead(archive: DocxArchive): boolean {
  const documentXml = archive.getText('word/document.xml')
  if (!documentXml) return false

  // 检测 altChunk
  if (documentXml.includes('w:altChunk') || documentXml.includes('r:id="altChunk')) {
    return true
  }

  const parsed = parseXml(documentXml)
  const doc = parsed['w:document'] as Record<string, unknown> | undefined
  const body = doc?.['w:body'] as Record<string, unknown> | undefined
  if (!body) return false

  const paragraphs = ensureArray(body['w:p'] as Record<string, unknown>[])

  // 检查前 10 个段落
  const checkCount = Math.min(paragraphs.length, 10)
  let hasRedColor = false
  let hasRedBorder = false
  let hasLargeFont = false

  for (let i = 0; i < checkCount; i++) {
    const p = paragraphs[i]
    const pPr = p['w:pPr'] as Record<string, unknown> | undefined

    // 检查红色边框
    if (pPr) {
      const pBdr = pPr['w:pBdr'] as Record<string, unknown> | undefined
      if (pBdr) {
        const bottom = pBdr['w:bottom'] as Record<string, unknown> | undefined
        if (bottom) {
          const color = getAttr(bottom, 'w:color')?.toLowerCase()
          if (color === 'ff0000' || color === 'red' || color === '#ff0000') {
            hasRedBorder = true
          }
        }
      }
    }

    // 检查 run 中的红色和大字
    const runs = ensureArray(p['w:r'] as Record<string, unknown>[])
    for (const r of runs) {
      const rPr = r['w:rPr'] as Record<string, unknown> | undefined
      if (!rPr) continue

      const color = rPr['w:color'] as Record<string, unknown> | undefined
      if (color) {
        const val = getWVal(color)?.toLowerCase()
        if (val === 'ff0000' || val === 'red') hasRedColor = true
      }

      const sz = rPr['w:sz'] as Record<string, unknown> | undefined
      if (sz) {
        const val = Number(getWVal(sz) ?? 0)
        if (val >= 44) hasLargeFont = true // 22pt = 44 half-points
      }
    }
  }

  return (hasRedColor && hasLargeFont) || (hasRedBorder && hasRedColor)
}

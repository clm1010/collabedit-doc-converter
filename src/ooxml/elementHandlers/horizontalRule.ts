import type { TiptapNode } from '../../types/tiptapJson.js'
import type { ParseContext, ParagraphProperties } from '../../types/ooxml.js'
import { createNode } from '../../types/tiptapJson.js'
import { getAttr } from '../xmlParser.js'

/**
 * 检测红头红线（特征：段落边框为红色底边线，通常无文字或仅包含空格）
 * 返回 ColoredHorizontalRule 节点或 null
 */
export function detectHorizontalRule(
  pPr: Record<string, unknown>,
  hasContent: boolean,
  ctx: ParseContext,
): TiptapNode | null {
  const pBdr = pPr['w:pBdr'] as Record<string, unknown> | undefined
  if (!pBdr) return null

  const bottom = pBdr['w:bottom'] as Record<string, unknown> | undefined
  if (!bottom) return null

  const color = getAttr(bottom, 'w:color')
  const val = getAttr(bottom, 'w:val')
  const sz = getAttr(bottom, 'w:sz')

  // 如果有底边框且颜色为红色系，且段落无实质内容
  if (color && isRedishColor(color) && !hasContent) {
    return createNode('horizontalRule', { 'data-line-color': 'red' })
  }

  // 通用水平线检测：仅有边框无内容的段落
  if (val && val !== 'none' && !hasContent) {
    const attrs: Record<string, unknown> = {}
    if (color && color !== 'auto') {
      attrs['data-line-color'] = color.startsWith('#') ? color : `#${color}`
    }
    return createNode('horizontalRule', Object.keys(attrs).length > 0 ? attrs : undefined)
  }

  return null
}

function isRedishColor(color: string): boolean {
  const normalized = color.toLowerCase().replace('#', '')
  // 常见的红色值
  return normalized === 'ff0000' ||
    normalized === 'red' ||
    normalized === 'cc0000' ||
    normalized === 'e60000' ||
    (normalized.length === 6 &&
      parseInt(normalized.substring(0, 2), 16) > 180 &&
      parseInt(normalized.substring(2, 4), 16) < 80 &&
      parseInt(normalized.substring(4, 6), 16) < 80)
}

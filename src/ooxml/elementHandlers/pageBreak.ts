import type { TiptapNode } from '../../types/tiptapJson.js'
import { createNode } from '../../types/tiptapJson.js'
import { getAttr } from '../xmlParser.js'

/** 检查段落是否包含分页符，返回 pageBreak 节点或 null */
export function checkPageBreak(run: Record<string, unknown>): TiptapNode | null {
  const br = run['w:br'] as Record<string, unknown> | undefined
  if (!br) return null

  const type = getAttr(br, 'w:type')
  if (type === 'page') {
    return createNode('pageBreak')
  }

  return null
}

/** 检查段落属性中的分页符 */
export function checkParagraphPageBreak(pPr: Record<string, unknown>): boolean {
  const pageBreakBefore = pPr['w:pageBreakBefore'] as Record<string, unknown> | undefined
  if (pageBreakBefore !== undefined) {
    // 存在即生效（除非 val="0"）
    const val = getAttr(pageBreakBefore, 'w:val')
    return val !== '0' && val !== 'false'
  }
  return false
}

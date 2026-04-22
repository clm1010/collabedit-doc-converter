import type { TiptapNode, TiptapMark } from '../../types/tiptapJson.js'
import type { ParseContext, RunProperties } from '../../types/ooxml.js'
import { ensureArray, getAttr } from '../xmlParser.js'
import { handleRun } from './run.js'
import { resolveRelTarget } from '../relationships.js'

/**
 * 处理 w:hyperlink。
 *
 * 重要：TOC 的每一条通常是
 *   <w:hyperlink>
 *     <w:r><w:t>标题</w:t></w:r>
 *     <w:r><w:tab/></w:r>
 *     <w:r><w:fldChar begin/></w:r>
 *     <w:r><w:instrText>PAGEREF _Toc... \h</w:instrText></w:r>
 *     <w:r><w:fldChar separate/></w:r>
 *     <w:r><w:t>3</w:t></w:r>
 *     <w:r><w:fldChar end/></w:r>
 *   </w:hyperlink>
 *
 * 因此这里也需要一套与 paragraph.processInlineContent 等价的 fldChar 状态机，
 * 否则 PAGEREF 的 instrText 会被 run.handleRun 误判成需要警告的未知域指令文本。
 */
export function handleHyperlink(
  hl: Record<string, unknown>,
  ctx: ParseContext,
  parentStyleRPr: RunProperties,
): TiptapNode[] {
  const rId = getAttr(hl, 'r:id')
  const anchor = getAttr(hl, 'w:anchor')

  let href: string | undefined
  if (rId) {
    href = resolveRelTarget(ctx.relationships, rId)
  } else if (anchor) {
    href = `#${anchor}`
  }

  const nodes: TiptapNode[] = []
  const runs = ensureArray(hl['w:r'] as Record<string, unknown>[])

  // fldChar 状态机：'none' → 正常输出；'instr' → 吞掉 instrText；'display' → 正常输出
  let fldState: 'none' | 'instr' | 'display' = 'none'

  for (const run of runs) {
    const fldChar = run['w:fldChar'] as Record<string, unknown> | undefined
    if (fldChar) {
      const type = (fldChar as Record<string, unknown>)['@_w:fldCharType'] as string | undefined
      if (type === 'begin') { fldState = 'instr'; continue }
      if (type === 'separate') { fldState = 'display'; continue }
      if (type === 'end') { fldState = 'none'; continue }
    }

    if (fldState === 'instr') {
      // 静默吞掉指令文本所在 run，不走 handleRun，避免重复告警。
      continue
    }

    const runNodes = handleRun(run, ctx, parentStyleRPr)
    for (const node of runNodes) {
      if (href && node.type === 'text') {
        const linkMark: TiptapMark = { type: 'link', attrs: { href, target: '_blank' } }
        node.marks = [...(node.marks ?? []), linkMark]
      }
      nodes.push(node)
    }
  }

  return nodes
}

import type { TiptapNode, TiptapMark } from '../../types/tiptapJson.js'
import type { ParseContext, RunProperties } from '../../types/ooxml.js'
import { ensureArray, getAttr } from '../xmlParser.js'
import { handleRun } from './run.js'
import { resolveRelTarget } from '../relationships.js'

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
  for (const run of runs) {
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

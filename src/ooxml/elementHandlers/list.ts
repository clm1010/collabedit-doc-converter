import type { TiptapNode } from '../../types/tiptapJson.js'
import type { ParseContext } from '../../types/ooxml.js'
import { createNode } from '../../types/tiptapJson.js'
import { getNumberingLevel, isBulletFormat } from '../numberingResolver.js'

interface ListParagraph {
  node: TiptapNode
  numId: number
  ilvl: number
}

/**
 * 将带有 __numPr 标记的段落包裹到 bulletList/orderedList 结构中。
 * 此函数在所有段落处理完成后调用，对文档级节点数组做后处理。
 */
export function wrapListItems(nodes: TiptapNode[], ctx: ParseContext): TiptapNode[] {
  const result: TiptapNode[] = []
  let i = 0

  while (i < nodes.length) {
    const node = nodes[i]
    const numPr = (node as unknown as Record<string, unknown>).__numPr as { numId: number; ilvl: number } | undefined

    if (!numPr) {
      result.push(node)
      i++
      continue
    }

    // 收集连续的同 numId 列表段落
    const listItems: ListParagraph[] = []
    while (i < nodes.length) {
      const current = nodes[i]
      const currentNumPr = (current as unknown as Record<string, unknown>).__numPr as { numId: number; ilvl: number } | undefined
      if (!currentNumPr) break
      listItems.push({ node: current, numId: currentNumPr.numId, ilvl: currentNumPr.ilvl })
      // 清理临时标记
      delete (current as unknown as Record<string, unknown>).__numPr
      i++
    }

    // 构建列表结构
    const listTree = buildListTree(listItems, 0, ctx)
    result.push(...listTree)
  }

  return result
}

function buildListTree(
  items: ListParagraph[],
  baseLevel: number,
  ctx: ParseContext,
): TiptapNode[] {
  const result: TiptapNode[] = []
  let i = 0

  while (i < items.length) {
    const item = items[i]
    const level = item.ilvl

    if (level < baseLevel) break

    if (level === baseLevel) {
      const numLvl = getNumberingLevel(ctx.numbering, item.numId, item.ilvl)
      const isBullet = numLvl ? isBulletFormat(numLvl.numFmt) : true

      const listItemNodes: TiptapNode[] = []
      listItemNodes.push(createNode('listItem', undefined, [item.node]))
      i++

      while (i < items.length && items[i].ilvl >= baseLevel) {
        if (items[i].ilvl === baseLevel) {
          listItemNodes.push(createNode('listItem', undefined, [items[i].node]))
          i++
        } else {
          // 子级项：收集到当前子级结束
          const subStart = i
          while (i < items.length && items[i].ilvl > baseLevel) {
            i++
          }
          const subItems = items.slice(subStart, i)
          const subTree = buildListTree(subItems, baseLevel + 1, ctx)
          const lastListItem = listItemNodes[listItemNodes.length - 1]
          if (lastListItem) {
            lastListItem.content = [...(lastListItem.content ?? []), ...subTree]
          }
        }
      }

      const listType = isBullet ? 'bulletList' : 'orderedList'
      result.push(createNode(listType, undefined, listItemNodes))
    } else {
      // level > baseLevel, 跳过（由递归处理）
      i++
    }
  }

  return result
}

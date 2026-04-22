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

    // 构建列表结构：baseLevel 取该段内最小 ilvl（见 buildListTree 的 skip 分支说明）。
    const listTree = buildListTree(listItems, minIlvl(listItems), ctx)
    result.push(...listTree)
  }

  return result
}

function minIlvl(items: ListParagraph[]): number {
  let m = items[0].ilvl
  for (let j = 1; j < items.length; j++) {
    if (items[j].ilvl < m) m = items[j].ilvl
  }
  return m
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
          // 递归时 baseLevel 取 subItems 实际最小 ilvl，
          // 避免跳级（如父 ilvl=3 直接嵌套 ilvl=5）时
          // buildListTree 的 else 分支把跳级 item 丢弃。
          const subTree = buildListTree(subItems, minIlvl(subItems), ctx)
          const lastListItem = listItemNodes[listItemNodes.length - 1]
          if (lastListItem) {
            lastListItem.content = [...(lastListItem.content ?? []), ...subTree]
          }
        }
      }

      const listType = isBullet ? 'bulletList' : 'orderedList'
      result.push(createNode(listType, undefined, listItemNodes))
    } else {
      // level > baseLevel 且此时 outer while 中还没任何 level==baseLevel 的 listItem
      // 可以挂子树（否则就会被 inner while 的 else 分支吃掉）。
      // 这种"段首 leading 高级别 items"直接递归成独立列表，避免 items 被整段丢掉。
      const subStart = i
      while (i < items.length && items[i].ilvl > baseLevel) i++
      const subItems = items.slice(subStart, i)
      if (subItems.length > 0) {
        const subTree = buildListTree(subItems, minIlvl(subItems), ctx)
        result.push(...subTree)
      }
    }
  }

  return result
}

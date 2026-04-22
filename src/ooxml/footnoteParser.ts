import type { TiptapNode } from '../types/tiptapJson.js'
import type { ParseContext, RelationshipMap } from '../types/ooxml.js'
import type { DocxArchive } from './zipExtractor.js'
import { parseXml, parseOrdered, findOrderedByPath, ensureArray, getAttr } from './xmlParser.js'
import { parseRelationships } from './relationships.js'
import { handleParagraph, handleTable } from './elementHandlers/index.js'
import { wrapListItems } from './elementHandlers/list.js'

export interface FootnoteData {
  id: number
  noteType: 'normal' | 'separator' | 'continuationSeparator' | string
  content: TiptapNode[]
}

/**
 * 解析 word/footnotes.xml 或 word/endnotes.xml
 * 返回 id → FootnoteData 的 Map
 */
export function parseFootnotes(
  archive: DocxArchive,
  partPath: 'word/footnotes.xml' | 'word/endnotes.xml',
  rootCtx: ParseContext,
): Map<number, FootnoteData> {
  const result = new Map<number, FootnoteData>()
  const xml = archive.getText(partPath)
  if (!xml) return result

  const relsPath = partPath.replace(/^word\//, 'word/_rels/') + '.rels'
  const localRels: RelationshipMap = parseRelationships(archive, relsPath)
  const parsed = parseXml(xml)
  const ordered = parseOrdered(xml)

  const rootKey = partPath.endsWith('footnotes.xml') ? 'w:footnotes' : 'w:endnotes'
  const itemTag = partPath.endsWith('footnotes.xml') ? 'w:footnote' : 'w:endnote'
  const root = parsed[rootKey] as Record<string, unknown> | undefined
  if (!root) return result

  const ctx: ParseContext = {
    ...rootCtx,
    relationships: localRels,
    orderedRoot: ordered,
    partPath,
    // 脚注/尾注部件里即便出现 heading，也不应消费主文档的章节编号计数器，
    // 置 undefined 让 paragraph 处理器跳过 counter.advance。
    headingNumberingCounter: undefined,
  }

  const items = ensureArray(root[itemTag] as Record<string, unknown>[])
  for (const item of items) {
    const idStr = getAttr(item, 'w:id')
    const typeStr = getAttr(item, 'w:type') ?? 'normal'
    if (idStr == null) continue
    const id = Number(idStr)
    // 跳过 -1/0 之类的分隔符占位？保留，上层可按 noteType 判定。
    const content: TiptapNode[] = []
    const paragraphs = ensureArray(item['w:p'] as Record<string, unknown>[])
    for (const p of paragraphs) {
      const r = handleParagraph(p, ctx)
      if (r) {
        if (Array.isArray(r)) content.push(...r)
        else content.push(r)
      }
    }
    const tables = ensureArray(item['w:tbl'] as Record<string, unknown>[])
    for (const t of tables) {
      content.push(handleTable(t, ctx))
    }
    const wrapped = wrapListItems(content, ctx)
    result.set(id, { id, noteType: typeStr as FootnoteData['noteType'], content: wrapped })
  }

  // 抑制未使用告警
  void findOrderedByPath
  return result
}

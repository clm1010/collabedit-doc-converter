import type { TiptapNode } from '../types/tiptapJson.js'
import type { ParseContext, RelationshipMap } from '../types/ooxml.js'
import type { DocxArchive } from './zipExtractor.js'
import { parseXml, parseOrdered, findOrderedByPath, ensureArray, getAttr } from './xmlParser.js'
import { parseRelationships } from './relationships.js'
import { handleParagraph, handleTable } from './elementHandlers/index.js'
import { wrapListItems } from './elementHandlers/list.js'

/**
 * 解析页眉/页脚 XML 为 Tiptap 节点数组。
 * 使用独立的 relationships（word/_rels/headerN.xml.rels）以解析内部图片等引用。
 */
export function parseHeaderFooterXml(
  archive: DocxArchive,
  partPath: string,
  rootCtx: ParseContext,
): TiptapNode[] {
  const xml = archive.getText(partPath)
  if (!xml) return []

  // 独立的 rels
  const relsPath = partPath.replace(/^word\//, 'word/_rels/') + '.rels'
  const localRels: RelationshipMap = parseRelationships(archive, relsPath)

  const parsed = parseXml(xml)
  const ordered = parseOrdered(xml)

  // header/footer root tag: w:hdr / w:ftr
  const rootKey = Object.keys(parsed).find(k => k === 'w:hdr' || k === 'w:ftr')
  if (!rootKey) return []

  const root = parsed[rootKey] as Record<string, unknown> | undefined
  const orderedRoot = findOrderedByPath(ordered, [rootKey])
  if (!root) return []

  // 基于 rootCtx 构建局部 ctx（仅替换 relationships）。
  // 同时标记 partPath + 禁用 headingNumberingCounter：页眉/页脚里若出现 heading 样式，
  // 不应推进主文档的章节编号（会打乱正文的 1 / 1.2 / 1.2.1 序列）。
  const ctx: ParseContext = {
    ...rootCtx,
    relationships: localRels,
    orderedRoot: ordered,
    partPath,
    headingNumberingCounter: undefined,
  }

  const nodes: TiptapNode[] = []

  if (orderedRoot) {
    const pIdx = { v: 0 }
    const tblIdx = { v: 0 }
    const pArr = ensureArray(root['w:p'] as Record<string, unknown>[])
    const tblArr = ensureArray(root['w:tbl'] as Record<string, unknown>[])
    for (const child of orderedRoot.children) {
      if (child.tag === 'w:p') {
        const p = pArr[pIdx.v++]
        if (!p) continue
        const result = handleParagraph(p, ctx)
        if (result) {
          if (Array.isArray(result)) nodes.push(...result)
          else nodes.push(result)
        }
      } else if (child.tag === 'w:tbl') {
        const t = tblArr[tblIdx.v++]
        if (!t) continue
        nodes.push(handleTable(t, ctx))
      }
    }
  } else {
    // 回退：无序遍历
    const pArr = ensureArray(root['w:p'] as Record<string, unknown>[])
    for (const p of pArr) {
      const r = handleParagraph(p, ctx)
      if (r) {
        if (Array.isArray(r)) nodes.push(...r)
        else nodes.push(r)
      }
    }
    const tblArr = ensureArray(root['w:tbl'] as Record<string, unknown>[])
    for (const t of tblArr) nodes.push(handleTable(t, ctx))
  }

  return wrapListItems(nodes, ctx)
}

/**
 * 基于 document.xml 的 sectPr 中的 w:headerReference / w:footerReference 建立精确映射：
 * Returns { headers: { [type]: nodes }, footers: { [type]: nodes } }
 * type 为 'default' | 'first' | 'even'
 */
export function parseHeadersFooters(
  archive: DocxArchive,
  documentXml: string,
  documentRels: RelationshipMap,
  rootCtx: ParseContext,
): {
  headers: Record<string, TiptapNode[]>
  footers: Record<string, TiptapNode[]>
} {
  const headers: Record<string, TiptapNode[]> = {}
  const footers: Record<string, TiptapNode[]> = {}

  const parsed = parseXml(documentXml)
  const doc = parsed['w:document'] as Record<string, unknown> | undefined
  const body = doc?.['w:body'] as Record<string, unknown> | undefined
  if (!body) return { headers, footers }

  // 收集所有 sectPr（body 直接 sectPr + 段落内 sectPr）
  const sectPrs: Record<string, unknown>[] = []
  const bodySectPr = body['w:sectPr'] as Record<string, unknown> | undefined
  if (bodySectPr) sectPrs.push(bodySectPr)

  const paragraphs = ensureArray(body['w:p'] as Record<string, unknown>[])
  for (const p of paragraphs) {
    const pPr = p['w:pPr'] as Record<string, unknown> | undefined
    if (!pPr) continue
    const inner = pPr['w:sectPr'] as Record<string, unknown> | undefined
    if (inner) sectPrs.push(inner)
  }

  for (const sectPr of sectPrs) {
    const headerRefs = ensureArray(sectPr['w:headerReference'] as Record<string, unknown>[])
    for (const ref of headerRefs) {
      const type = (getAttr(ref, 'w:type') ?? 'default') as 'default' | 'first' | 'even'
      const rId = getAttr(ref, 'r:id')
      if (!rId) continue
      const target = documentRels[rId]?.target
      if (!target) continue
      const path = target.startsWith('word/') ? target : `word/${target}`
      if (!headers[type]) {
        headers[type] = parseHeaderFooterXml(archive, path, rootCtx)
      }
    }
    const footerRefs = ensureArray(sectPr['w:footerReference'] as Record<string, unknown>[])
    for (const ref of footerRefs) {
      const type = (getAttr(ref, 'w:type') ?? 'default') as 'default' | 'first' | 'even'
      const rId = getAttr(ref, 'r:id')
      if (!rId) continue
      const target = documentRels[rId]?.target
      if (!target) continue
      const path = target.startsWith('word/') ? target : `word/${target}`
      if (!footers[type]) {
        footers[type] = parseHeaderFooterXml(archive, path, rootCtx)
      }
    }
  }

  return { headers, footers }
}

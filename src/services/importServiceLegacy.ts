/**
 * Legacy 导入路径（选择性保存引入前的旧实现）。
 *
 * 保留作为新管线异常时的兜底：当 engine/importPipeline.ts 抛异常时，
 * importService.ts 会自动回退到此处，保证最低可用性不低于改造前。
 *
 * 该文件是 Git 历史里 importService.ts 的直译拷贝，除更名为 `importDocxLegacy`
 * 外不引入任何新逻辑；如需维护 legacy 行为，直接修改此文件。
 */
import type { Buffer as NodeBuffer } from 'node:buffer'
import { extractDocx } from '../ooxml/zipExtractor.js'
import { parseXml, parseOrdered, findOrderedByPath, ensureArray } from '../ooxml/xmlParser.js'
import type { OrderedNode } from '../ooxml/xmlParser.js'
import { parseDocumentRelationships } from '../ooxml/relationships.js'
import { resolveStyles } from '../ooxml/styleResolver.js'
import { resolveTheme } from '../ooxml/themeResolver.js'
import { resolveNumbering } from '../ooxml/numberingResolver.js'
import { extractImages } from '../ooxml/imageExtractor.js'
import { detectRedHead } from '../ooxml/redheadDetector.js'
import { extractMetadata } from '../utils/metadataExtractor.js'
import {
  handleParagraph,
  handleTable,
  handleSdt,
  parseTocParagraph,
  isTocStyledParagraph,
} from '../ooxml/elementHandlers/index.js'
import { wrapListItems } from '../ooxml/elementHandlers/list.js'
import { checkParagraphPageBreak } from '../ooxml/elementHandlers/pageBreak.js'
import { detectHorizontalRule } from '../ooxml/elementHandlers/horizontalRule.js'
import { parseHeadersFooters } from '../ooxml/headerFooterParser.js'
import { parseFootnotes } from '../ooxml/footnoteParser.js'
import { parseSectionProperties } from '../ooxml/sectionParser.js'
import type { ParseContext } from '../types/ooxml.js'
import type { TiptapNode, TiptapDoc, ImportResponse, ImportLogs, FootnoteData } from '../types/tiptapJson.js'
import { createDoc, createNode } from '../types/tiptapJson.js'
import type { SectionDefinition } from '../types/docMetadata.js'

export async function importDocxLegacy(fileBuffer: NodeBuffer): Promise<ImportResponse> {
  const logs: ImportLogs = { info: [], warn: [], error: [] }
  const startTime = Date.now()

  const archive = extractDocx(fileBuffer)
  logs.info.push(`ZIP extracted, ${archive.listFiles().length} files`)

  const relationships = parseDocumentRelationships(archive)
  const { styles, docDefaults } = resolveStyles(archive)
  const theme = resolveTheme(archive)
  const numbering = resolveNumbering(archive)
  const images = extractImages(archive)
  const metadata = extractMetadata(archive, logs)
  const isRedHead = detectRedHead(archive)

  metadata.isRedHead = isRedHead
  logs.info.push(`Styles: ${styles.size}, Images: ${images.size}, IsRedHead: ${isRedHead}`)

  const ctx: ParseContext = {
    styles,
    numbering,
    relationships,
    images,
    theme,
    docDefaults,
    logs,
  }

  const documentXml = archive.getText('word/document.xml')
  if (!documentXml) {
    logs.error.push('word/document.xml not found')
    return {
      data: { content: createDoc([createNode('paragraph')]) },
      metadata,
      logs,
    }
  }

  const parsed = parseXml(documentXml)
  const orderedDoc = parseOrdered(documentXml)
  ctx.orderedRoot = orderedDoc

  const doc = parsed['w:document'] as Record<string, unknown> | undefined
  const body = doc?.['w:body'] as Record<string, unknown> | undefined

  if (!body) {
    logs.error.push('w:body not found in document.xml')
    return {
      data: { content: createDoc([createNode('paragraph')]) },
      metadata,
      logs,
    }
  }

  const orderedBody = findOrderedByPath(orderedDoc, ['w:document', 'w:body'])

  const pArr = ensureArray(body['w:p'] as Record<string, unknown>[])
  const tblArr = ensureArray(body['w:tbl'] as Record<string, unknown>[])
  const sdtArr = ensureArray(body['w:sdt'] as Record<string, unknown>[])

  const nodes: TiptapNode[] = []

  let footnotesData: FootnoteData[] | undefined
  let endnotesData: FootnoteData[] | undefined
  if (metadata.hasFootnotes) {
    const m = parseFootnotes(archive, 'word/footnotes.xml', ctx)
    footnotesData = Array.from(m.values())
    ctx.footnotes = new Map(Array.from(m.entries()).map(([k, v]) => [k, v.content]))
  }
  if (metadata.hasEndnotes) {
    const m = parseFootnotes(archive, 'word/endnotes.xml', ctx)
    endnotesData = Array.from(m.values())
    ctx.endnotes = new Map(Array.from(m.entries()).map(([k, v]) => [k, v.content]))
  }

  if (orderedBody) {
    let pIdx = 0, tblIdx = 0, sdtIdx = 0
    for (const child of orderedBody.children) {
      if (child.tag === 'w:p') {
        const p = pArr[pIdx++]
        if (!p) continue
        processParagraph(p, nodes, ctx)
      } else if (child.tag === 'w:tbl') {
        const t = tblArr[tblIdx++]
        if (!t) continue
        nodes.push(handleTable(t, ctx))
      } else if (child.tag === 'w:sdt') {
        const s = sdtArr[sdtIdx++]
        if (!s) continue
        const sdtNodes = handleSdt(s, child, ctx, parseBlockContentForSdt)
        nodes.push(...sdtNodes)
      }
    }
  } else {
    for (const p of pArr) processParagraph(p, nodes, ctx)
    for (const t of tblArr) nodes.push(handleTable(t, ctx))
    for (const s of sdtArr) {
      const sdtNodes = handleSdt(s, null, ctx, parseBlockContentForSdt)
      nodes.push(...sdtNodes)
    }
  }

  const finalNodes = wrapListItems(nodes, ctx)

  const content: TiptapDoc = createDoc(
    finalNodes.length > 0 ? finalNodes : [createNode('paragraph')],
  )

  try {
    const hf = parseHeadersFooters(archive, documentXml, relationships, ctx)
    for (const [k, v] of Object.entries(hf.headers)) {
      if (v.length > 0) {
        (metadata.headers as Record<string, unknown>)[k] = v
      }
    }
    for (const [k, v] of Object.entries(hf.footers)) {
      if (v.length > 0) {
        (metadata.footers as Record<string, unknown>)[k] = v
      }
    }
  } catch (err) {
    logs.warn.push(`Header/footer parsing failed: ${err instanceof Error ? err.message : String(err)}`)
  }

  try {
    const sections: SectionDefinition[] = []
    const bodySectPr = body['w:sectPr'] as Record<string, unknown> | undefined
    const collectFromSectPr = (sectPr: Record<string, unknown>) => {
      const section = parseSectionProperties(sectPr)
      sections.push({
        pageSetup: section.pageSetup,
        headerRefs: section.headerRefs,
        footerRefs: section.footerRefs,
        type: section.type,
        titlePg: section.titlePg,
      })
    }
    for (const p of pArr) {
      const pPr = p['w:pPr'] as Record<string, unknown> | undefined
      if (!pPr) continue
      const inner = pPr['w:sectPr'] as Record<string, unknown> | undefined
      if (inner) collectFromSectPr(inner)
    }
    if (bodySectPr) collectFromSectPr(bodySectPr)
    metadata.sections = sections
  } catch (err) {
    logs.warn.push(`Section parsing failed: ${err instanceof Error ? err.message : String(err)}`)
  }

  const elapsed = Date.now() - startTime
  logs.info.push(`Import completed in ${elapsed}ms, ${finalNodes.length} top-level nodes`)

  return {
    data: { content },
    metadata,
    logs,
    ...(footnotesData ? { footnotes: footnotesData } : {}),
    ...(endnotesData ? { endnotes: endnotesData } : {}),
  }
}

function processParagraph(
  p: Record<string, unknown>,
  nodes: TiptapNode[],
  ctx: ParseContext,
) {
  const pPr = p['w:pPr'] as Record<string, unknown> | undefined
  if (pPr && checkParagraphPageBreak(pPr)) {
    nodes.push(createNode('pageBreak'))
  }

  if (pPr) {
    const hasContent = hasRealContent(p)
    const hr = detectHorizontalRule(pPr, hasContent, ctx)
    if (hr) {
      nodes.push(hr)
      return
    }
  }

  if (isTocStyledParagraph(p, ctx)) {
    const entry = parseTocParagraph(p, ctx)
    if (entry) {
      nodes.push(entry)
      return
    }
  }

  const result = handleParagraph(p, ctx)
  if (result) {
    if (Array.isArray(result)) nodes.push(...result)
    else nodes.push(result)
  }
}

function parseBlockContentForSdt(
  parent: Record<string, unknown>,
  orderedParent: OrderedNode | null,
  ctx: ParseContext,
): TiptapNode[] {
  const nodes: TiptapNode[] = []
  const pArr = ensureArray(parent['w:p'] as Record<string, unknown>[])
  const tblArr = ensureArray(parent['w:tbl'] as Record<string, unknown>[])
  const sdtArr = ensureArray(parent['w:sdt'] as Record<string, unknown>[])

  if (orderedParent) {
    let pIdx = 0, tblIdx = 0, sdtIdx = 0
    for (const child of orderedParent.children) {
      if (child.tag === 'w:p') {
        const p = pArr[pIdx++]
        if (!p) continue
        processParagraph(p, nodes, ctx)
      } else if (child.tag === 'w:tbl') {
        const t = tblArr[tblIdx++]
        if (!t) continue
        nodes.push(handleTable(t, ctx))
      } else if (child.tag === 'w:sdt') {
        const s = sdtArr[sdtIdx++]
        if (!s) continue
        nodes.push(...handleSdt(s, child, ctx, parseBlockContentForSdt))
      }
    }
  } else {
    for (const p of pArr) processParagraph(p, nodes, ctx)
    for (const t of tblArr) nodes.push(handleTable(t, ctx))
    for (const s of sdtArr) nodes.push(...handleSdt(s, null, ctx, parseBlockContentForSdt))
  }

  return wrapListItems(nodes, ctx)
}

function hasRealContent(p: Record<string, unknown>): boolean {
  const runs = ensureArray(p['w:r'] as Record<string, unknown>[])
  for (const r of runs) {
    const texts = ensureArray(r['w:t'] as unknown[])
    for (const t of texts) {
      const text = typeof t === 'string' ? t : (typeof t === 'object' && t !== null ? (t as Record<string, unknown>)['#text'] as string : '')
      if (text && text.trim().length > 0) return true
    }
    if (r['w:drawing'] || r['w:pict']) return true
  }
  return false
}

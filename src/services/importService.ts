import { extractDocx } from '../ooxml/zipExtractor.js'
import { parseXml, ensureArray } from '../ooxml/xmlParser.js'
import { parseDocumentRelationships } from '../ooxml/relationships.js'
import { resolveStyles } from '../ooxml/styleResolver.js'
import { resolveTheme } from '../ooxml/themeResolver.js'
import { resolveNumbering } from '../ooxml/numberingResolver.js'
import { extractImages } from '../ooxml/imageExtractor.js'
import { detectRedHead } from '../ooxml/redheadDetector.js'
import { extractMetadata } from '../utils/metadataExtractor.js'
import { handleParagraph, handleTable } from '../ooxml/elementHandlers/index.js'
import { wrapListItems } from '../ooxml/elementHandlers/list.js'
import { checkParagraphPageBreak } from '../ooxml/elementHandlers/pageBreak.js'
import { detectHorizontalRule } from '../ooxml/elementHandlers/horizontalRule.js'
import type { ParseContext } from '../types/ooxml.js'
import type { TiptapNode, TiptapDoc, ImportResponse, ImportLogs } from '../types/tiptapJson.js'
import { createDoc, createNode } from '../types/tiptapJson.js'

export async function importDocx(fileBuffer: Buffer): Promise<ImportResponse> {
  const logs: ImportLogs = { info: [], warn: [], error: [] }
  const startTime = Date.now()

  // 1. ZIP 解压
  const archive = extractDocx(fileBuffer)
  logs.info.push(`ZIP extracted, ${archive.listFiles().length} files`)

  // 2. 解析基础资源
  const relationships = parseDocumentRelationships(archive)
  const { styles, docDefaults } = resolveStyles(archive)
  const theme = resolveTheme(archive)
  const numbering = resolveNumbering(archive)
  const images = extractImages(archive)
  const metadata = extractMetadata(archive)
  const isRedHead = detectRedHead(archive)

  metadata.isRedHead = isRedHead
  logs.info.push(`Styles: ${styles.size}, Images: ${images.size}, IsRedHead: ${isRedHead}`)

  // 3. 构建解析上下文
  const ctx: ParseContext = {
    styles,
    numbering,
    relationships,
    images,
    theme,
    docDefaults,
    logs,
  }

  // 4. 解析 document.xml 主体
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

  // 5. 遍历 body 子元素
  // 注意：fast-xml-parser 将同标签合并为数组，段落和表格分开处理。
  // 对于段落和表格交错的文档（如表格后还有段落），顺序可能不精确，
  // 但大多数文档以段落为主，表格穿插，实际影响不大。
  const nodes: TiptapNode[] = []

  // 先处理段落
  const paragraphs = ensureArray(body['w:p'] as Record<string, unknown>[])
  for (const p of paragraphs) {
    const pPr = p['w:pPr'] as Record<string, unknown> | undefined
    if (pPr && checkParagraphPageBreak(pPr)) {
      nodes.push(createNode('pageBreak'))
    }

    if (pPr) {
      const hasContent = hasRealContent(p)
      const hr = detectHorizontalRule(pPr, hasContent, ctx)
      if (hr) {
        nodes.push(hr)
        continue
      }
    }

    const result = handleParagraph(p, ctx)
    if (result) {
      if (Array.isArray(result)) nodes.push(...result)
      else nodes.push(result)
    }
  }

  // 再处理表格（插入到段落之后）
  const tables = ensureArray(body['w:tbl'] as Record<string, unknown>[])
  for (const tbl of tables) {
    nodes.push(handleTable(tbl, ctx))
  }

  // 6. 列表后处理
  const finalNodes = wrapListItems(nodes, ctx)

  const content: TiptapDoc = createDoc(
    finalNodes.length > 0 ? finalNodes : [createNode('paragraph')],
  )

  const elapsed = Date.now() - startTime
  logs.info.push(`Import completed in ${elapsed}ms, ${finalNodes.length} top-level nodes`)

  return { data: { content }, metadata, logs }
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

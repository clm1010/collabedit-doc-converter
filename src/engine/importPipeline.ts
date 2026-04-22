/**
 * 选择性保存高保真方案：导入管线（Import Pipeline）。
 *
 * 本模块把原 `services/importService.ts` 的串行流程包装到一起，
 * 并在每个顶层块生成的 Tiptap 节点上回填 `__origRange / __origHash / __origPart`。
 *
 * 四步：
 *   Step 1 - extract : fflate.unzipSync 保留原始字节到 RawDocxArchive
 *   Step 2 - index   : xmlRangeIndexer 扫描 word/document.xml，得到顶层元素字节范围
 *   Step 3 - parse   : 复用现有 handleParagraph / handleTable / handleSdt 产出节点
 *   Step 4 - meta    : 页眉/页脚/脚注/尾注；本阶段走现有 parser（range 索引留待 phase 3）
 *
 * 设计约束：
 *   - 不改动现有 handler 签名。orchestrator 拿到 TiptapNode 后，按范围索引中的
 *     同序位置回填 origAttrs。"按序位置对齐"成立的前提：fast-xml-parser 解析出
 *     `body['w:p']` / `body['w:tbl']` / `body['w:sdt']` 的顺序与 XML 出现顺序一致，
 *     且 xmlRangeIndexer 也按字节出现顺序产出；两侧按 `tag` + 递增计数器对齐。
 *   - 未能识别 range 时（历史文档 / 解析失败 / 范围索引失败），节点以 null 值
 *     保留字段，导出侧自然降级到全量重序列化（等价 legacy 行为）。
 */

import type { Buffer as NodeBuffer } from 'node:buffer'
import { extractRawDocx, type RawDocxArchive } from './zipExtractor.js'
import { indexTopLevelRanges, sliceRange, type TopLevelRange } from './xmlRangeIndexer.js'
import { hashXmlRange } from './hasher.js'
import { computeContentFingerprint } from './contentFingerprint.js'
import { parseXml, parseOrdered, findOrderedByPath, ensureArray } from '../ooxml/xmlParser.js'
import type { OrderedNode } from '../ooxml/xmlParser.js'
import { parseDocumentRelationships } from '../ooxml/relationships.js'
import { resolveStyles } from '../ooxml/styleResolver.js'
import { resolveTheme } from '../ooxml/themeResolver.js'
import { resolveNumbering } from '../ooxml/numberingResolver.js'
import { HeadingNumberingCounter } from '../ooxml/headingNumberingCounter.js'
import { extractImages } from '../ooxml/imageExtractor.js'
import { detectRedHead } from '../ooxml/redheadDetector.js'
import { extractMetadata } from '../utils/metadataExtractor.js'
import { extractDocx } from '../ooxml/zipExtractor.js'
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
import type {
  TiptapNode,
  TiptapDoc,
  ImportResponse,
  ImportLogs,
  FootnoteData,
} from '../types/tiptapJson.js'
import { createDoc, createNode } from '../types/tiptapJson.js'
import type { SectionDefinition } from '../types/docMetadata.js'

const DOCUMENT_XML_PATH = 'word/document.xml'
const TEXT_DECODER = new TextDecoder('utf-8')

export async function importDocxPipeline(fileBuffer: NodeBuffer): Promise<ImportResponse> {
  const logs: ImportLogs = { info: [], warn: [], error: [] }
  const startTime = Date.now()

  // ---- Step 1: extract raw bytes ----
  const rawArchive = extractRawDocx(fileBuffer)
  // 保留旧 DocxArchive 形态，供仍按字符串解析的老 handler（relationships / styles / 等）复用。
  const legacyArchive = extractDocx(fileBuffer)
  logs.info.push(`ZIP extracted, ${rawArchive.parts.length} files`)

  const relationships = parseDocumentRelationships(legacyArchive)
  const { styles, docDefaults } = resolveStyles(legacyArchive)
  const theme = resolveTheme(legacyArchive)
  const numbering = resolveNumbering(legacyArchive)
  const images = extractImages(legacyArchive)
  const metadata = extractMetadata(legacyArchive, logs)
  const isRedHead = detectRedHead(legacyArchive)

  metadata.isRedHead = isRedHead
  logs.info.push(
    `Styles: ${styles.size}, Images: ${images.size}, IsRedHead: ${isRedHead}`,
  )

  const ctx: ParseContext = {
    styles,
    numbering,
    relationships,
    images,
    theme,
    docDefaults,
    logs,
    rawArchive,
    partPath: DOCUMENT_XML_PATH,
    sdtXmlMap: new Map(),
    headingNumberingCounter: new HeadingNumberingCounter({ numbering }),
  }

  const documentBytes = rawArchive.partsByPath.get(DOCUMENT_XML_PATH)
  if (!documentBytes) {
    logs.error.push('word/document.xml not found')
    return {
      data: { content: createDoc([createNode('paragraph')]) },
      metadata,
      logs,
    }
  }
  const documentXml = TEXT_DECODER.decode(documentBytes)

  // ---- Step 2: index top-level ranges ----
  const rangeIndex = indexTopLevelRanges(documentBytes, 'w:body')
  if (rangeIndex) {
    ctx.topLevelRanges = rangeIndex.ranges
    logs.info.push(`Top-level ranges indexed: ${rangeIndex.ranges.length}`)
  } else {
    logs.warn.push('Range index not available; selective save disabled for this document')
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

  // 为脚注 / 尾注 handler 准备引用
  let footnotesData: FootnoteData[] | undefined
  let endnotesData: FootnoteData[] | undefined
  if (metadata.hasFootnotes) {
    const m = parseFootnotes(legacyArchive, 'word/footnotes.xml', ctx)
    footnotesData = Array.from(m.values())
    ctx.footnotes = new Map(Array.from(m.entries()).map(([k, v]) => [k, v.content]))
  }
  if (metadata.hasEndnotes) {
    const m = parseFootnotes(legacyArchive, 'word/endnotes.xml', ctx)
    endnotesData = Array.from(m.values())
    ctx.endnotes = new Map(Array.from(m.entries()).map(([k, v]) => [k, v.content]))
  }

  // ---- Step 3: parse body with orig-attrs backfill ----
  const nodes: TiptapNode[] = []
  walkBodyWithRanges({
    orderedBody,
    pArr,
    tblArr,
    sdtArr,
    ctx,
    documentBytes,
    rangeIndex,
    out: nodes,
    partPath: DOCUMENT_XML_PATH,
  })

  const finalNodes = wrapListItems(nodes, ctx)
  const content: TiptapDoc = createDoc(
    finalNodes.length > 0 ? finalNodes : [createNode('paragraph')],
  )

  // ---- Step 4: meta (headers/footers/sections) ----
  try {
    const hf = parseHeadersFooters(legacyArchive, documentXml, relationships, ctx)
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
    logs.warn.push(
      `Header/footer parsing failed: ${err instanceof Error ? err.message : String(err)}`,
    )
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
    logs.warn.push(
      `Section parsing failed: ${err instanceof Error ? err.message : String(err)}`,
    )
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

// -------------------------------------------------------------------
// orchestrator 层的 body 遍历 + origAttrs 回填
// -------------------------------------------------------------------

interface WalkBodyArgs {
  orderedBody: OrderedNode | null | undefined
  pArr: Record<string, unknown>[]
  tblArr: Record<string, unknown>[]
  sdtArr: Record<string, unknown>[]
  ctx: ParseContext
  documentBytes: Uint8Array
  rangeIndex: ReturnType<typeof indexTopLevelRanges>
  out: TiptapNode[]
  partPath: string
}

function walkBodyWithRanges(args: WalkBodyArgs): void {
  const { orderedBody, pArr, tblArr, sdtArr, ctx, documentBytes, rangeIndex, out, partPath } = args

  // 按 tag 维度的"范围列表"，用计数器对齐每个 handler 调用到对应字节范围。
  const rangesByTag = groupRangesByTag(rangeIndex?.ranges ?? [])
  const rangeCursor: Record<string, number> = {}
  const nextRange = (tag: string): TopLevelRange | null => {
    const arr = rangesByTag.get(tag)
    if (!arr) return null
    const idx = rangeCursor[tag] ?? 0
    if (idx >= arr.length) return null
    rangeCursor[tag] = idx + 1
    return arr[idx]
  }

  const appendWithOrig = (node: TiptapNode | TiptapNode[] | null, range: TopLevelRange | null) => {
    if (!node) return
    const arr = Array.isArray(node) ? node : [node]
    for (const n of arr) attachOrigAttrs(n, range, documentBytes, partPath)
    out.push(...arr)
  }

  if (orderedBody) {
    let pIdx = 0
    let tblIdx = 0
    let sdtIdx = 0
    for (const child of orderedBody.children) {
      if (child.tag === 'w:p') {
        const p = pArr[pIdx++]
        if (!p) continue
        const range = nextRange('w:p')
        processParagraph(p, ctx, (node) => appendWithOrig(node, range))
      } else if (child.tag === 'w:tbl') {
        const t = tblArr[tblIdx++]
        if (!t) continue
        const range = nextRange('w:tbl')
        appendWithOrig(handleTable(t, ctx), range)
      } else if (child.tag === 'w:sdt') {
        const s = sdtArr[sdtIdx++]
        if (!s) continue
        const range = nextRange('w:sdt')
        const sdtNodes = handleSdt(s, child, ctx, parseBlockContentForSdt)
        attachSdtOrigAttrs(sdtNodes, range, documentBytes, partPath, ctx)
        out.push(...sdtNodes)
      }
    }
  } else {
    // 无有序树（xmlParser 降级），按 tag 分组顺序处理。
    for (const p of pArr) {
      const range = nextRange('w:p')
      processParagraph(p, ctx, (node) => appendWithOrig(node, range))
    }
    for (const t of tblArr) {
      const range = nextRange('w:tbl')
      appendWithOrig(handleTable(t, ctx), range)
    }
    for (const s of sdtArr) {
      const range = nextRange('w:sdt')
      const sdtNodes = handleSdt(s, null, ctx, parseBlockContentForSdt)
      attachSdtOrigAttrs(sdtNodes, range, documentBytes, partPath, ctx)
      out.push(...sdtNodes)
    }
  }
}

function groupRangesByTag(ranges: TopLevelRange[]): Map<string, TopLevelRange[]> {
  const map = new Map<string, TopLevelRange[]>()
  for (const r of ranges) {
    const list = map.get(r.tag)
    if (list) list.push(r)
    else map.set(r.tag, [r])
  }
  return map
}

/**
 * 把 range/hash/partPath 写入节点 attrs。
 *   - 仅当 range 可用时回填，否则维持 null，导出侧自动走 legacy。
 *   - attrs 已存在时做浅合并，不覆盖已有字段。
 */
function attachOrigAttrs(
  node: TiptapNode,
  range: TopLevelRange | null,
  documentBytes: Uint8Array,
  partPath: string,
): void {
  if (!range) return
  // 排除 tiptap 包装类节点（bulletList / orderedList 之类）：它们是由后续 wrapListItems
  // 合成的，这里遇到的都是 handler 直接产出的块级叶子节点。
  const attrs = (node.attrs ?? {}) as Record<string, unknown>
  if (!('__origRange' in attrs)) {
    attrs.__origRange = [range.start, range.end]
  }
  if (!('__origHash' in attrs)) {
    attrs.__origHash = hashXmlRange(sliceRange(documentBytes, range))
  }
  if (!('__origPart' in attrs)) {
    attrs.__origPart = partPath
  }
  // contentFp 依赖当前 node 的内容，所以要在 attrs 补齐后（排除 __orig*）再算一次。
  // computeContentFingerprint 会自动忽略 __orig* / data-origin-* 等元数据 attrs。
  node.attrs = attrs
  if (!('__origContentFp' in attrs)) {
    attrs.__origContentFp = computeContentFingerprint(node)
  }
}

/**
 * SDT 的 TOC 特殊规则（plan 第 2 节）：
 *   - 整块 w:sdt 视为原子单元。sdtXmlMap 里保存该 SDT 的原始字节。
 *   - handleSdt 产出的节点序列可能是若干 tocEntry；把它们视为"共享同一 SDT"：
 *     · 第一个节点写入 __origSdtXml 字符串形式（供导出 localSerializer 直接输出）；
 *     · 其余节点仅写 __origSdtId，通过 ctx.sdtXmlMap 查找共享字节。
 *   - 非 tocEntry 的 SDT（如富文本 SDT）当前仍走普通 origRange 回填。
 */
function attachSdtOrigAttrs(
  nodes: TiptapNode[],
  range: TopLevelRange | null,
  documentBytes: Uint8Array,
  partPath: string,
  ctx: ParseContext,
): void {
  if (!range || nodes.length === 0) return
  const sdtBytes = sliceRange(documentBytes, range)
  const isTocGroup = nodes.every((n) => n.type === 'tocEntry')

  if (isTocGroup) {
    const sdtId = `sdt_${partPath}_${range.start}_${range.end}`
    ctx.sdtXmlMap?.set(sdtId, sdtBytes)
    const sdtXmlString = TEXT_DECODER.decode(sdtBytes)
    for (let i = 0; i < nodes.length; i++) {
      const n = nodes[i]
      const attrs = (n.attrs ?? {}) as Record<string, unknown>
      attrs.__origSdtId = sdtId
      attrs.__origPart = partPath
      if (i === 0) {
        attrs.__origSdtXml = sdtXmlString
      }
      // 首条 tocEntry 额外附带 range，供 nodeClassifier 做范围定位。
      if (i === 0) {
        attrs.__origRange = [range.start, range.end]
        attrs.__origHash = hashXmlRange(sdtBytes)
      }
      n.attrs = attrs
      attrs.__origContentFp = computeContentFingerprint(n)
    }
    return
  }

  // 非 TOC 形态的 SDT：整块作为一个节点单位处理（通常 handleSdt 返回 1 个节点）。
  for (const n of nodes) {
    attachOrigAttrs(n, range, documentBytes, partPath)
  }
}

// -------------------------------------------------------------------
// 复刻 importService 中的 processParagraph / parseBlockContentForSdt
// -------------------------------------------------------------------

function processParagraph(
  p: Record<string, unknown>,
  ctx: ParseContext,
  emit: (node: TiptapNode | TiptapNode[] | null) => void,
): void {
  const pPr = p['w:pPr'] as Record<string, unknown> | undefined
  if (pPr && checkParagraphPageBreak(pPr)) {
    // 分页符节点没有独立 range（它是段落级属性的副产物），暂不回填 origAttrs
    emit(createNode('pageBreak'))
  }

  if (pPr) {
    const hasContent = hasRealContent(p)
    const hr = detectHorizontalRule(pPr, hasContent, ctx)
    if (hr) {
      emit(hr)
      return
    }
  }

  if (isTocStyledParagraph(p, ctx)) {
    const entry = parseTocParagraph(p, ctx)
    if (entry) {
      emit(entry)
      return
    }
  }

  const result = handleParagraph(p, ctx)
  if (result) emit(result as TiptapNode | TiptapNode[])
}

function parseBlockContentForSdt(
  parent: Record<string, unknown>,
  orderedParent: OrderedNode | null,
  ctx: ParseContext,
): TiptapNode[] {
  // SDT 内容层级不做选择性保存的 range 回填；整个 w:sdt 作为原子块，
  // 由 attachSdtOrigAttrs 统一打在外层 TiptapNode 上。
  const nodes: TiptapNode[] = []
  const pArr = ensureArray(parent['w:p'] as Record<string, unknown>[])
  const tblArr = ensureArray(parent['w:tbl'] as Record<string, unknown>[])
  const sdtArr = ensureArray(parent['w:sdt'] as Record<string, unknown>[])

  if (orderedParent) {
    let pIdx = 0
    let tblIdx = 0
    let sdtIdx = 0
    for (const child of orderedParent.children) {
      if (child.tag === 'w:p') {
        const p = pArr[pIdx++]
        if (!p) continue
        processParagraph(p, ctx, (n) => {
          if (!n) return
          if (Array.isArray(n)) nodes.push(...n)
          else nodes.push(n)
        })
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
    for (const p of pArr) {
      processParagraph(p, ctx, (n) => {
        if (!n) return
        if (Array.isArray(n)) nodes.push(...n)
        else nodes.push(n)
      })
    }
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
      const text =
        typeof t === 'string'
          ? t
          : typeof t === 'object' && t !== null
          ? ((t as Record<string, unknown>)['#text'] as string)
          : ''
      if (text && text.trim().length > 0) return true
    }
    if (r['w:drawing'] || r['w:pict']) return true
  }
  return false
}

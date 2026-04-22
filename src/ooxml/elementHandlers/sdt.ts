import type { TiptapNode } from '../../types/tiptapJson.js'
import type { ParseContext } from '../../types/ooxml.js'
import type { OrderedNode } from '../xmlParser.js'
import { createNode } from '../../types/tiptapJson.js'
import { ensureArray, getOrderedAttr } from '../xmlParser.js'

/**
 * 判断 SDT 是否为 TOC（Table of Contents）
 * 特征：w:sdtPr > w:docPartObj > w:docPartGallery[@w:val="Table of Contents"]
 */
function isTocSdt(sdt: Record<string, unknown>): boolean {
  const sdtPr = sdt['w:sdtPr'] as Record<string, unknown> | undefined
  if (!sdtPr) return false
  const docPartObj = sdtPr['w:docPartObj'] as Record<string, unknown> | undefined
  if (!docPartObj) return false
  const gallery = docPartObj['w:docPartGallery'] as Record<string, unknown> | undefined
  if (!gallery) return false
  const val = (gallery as Record<string, unknown>)['@_w:val']
  return typeof val === 'string' && val === 'Table of Contents'
}

/**
 * 判断 OrderedNode 形式的 SDT 是否为 TOC
 */
function isTocSdtOrdered(sdtOrdered: OrderedNode): boolean {
  const sdtPr = sdtOrdered.children.find(c => c.tag === 'w:sdtPr')
  if (!sdtPr) return false
  const docPartObj = sdtPr.children.find(c => c.tag === 'w:docPartObj')
  if (!docPartObj) return false
  const gallery = docPartObj.children.find(c => c.tag === 'w:docPartGallery')
  if (!gallery) return false
  return getOrderedAttr(gallery, 'w:val') === 'Table of Contents'
}

/**
 * 判断段落是否带有 TOC 条目样式（TOC1/TOC2/…/TOC9、toc 1 等形态）。
 *
 * 老版 Word 以及许多国产办公软件生成的目录**不会**把 TOC 包进 `w:sdt`，
 * 而是以裸的 TOC 域 + 多段 "TOC n" 样式段落的形式出现：
 *
 *   <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:fldChar begin/></w:r>...</w:p>
 *   <w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:hyperlink>...</w:hyperlink></w:p>
 *   ...
 *
 * 因此除了 SDT 分支外，importService 还会按样式名在正文层拦截这些段落并转成 tocEntry。
 */
export function isTocStyledParagraph(
  p: Record<string, unknown>,
  ctx: ParseContext,
): boolean {
  const pPr = p['w:pPr'] as Record<string, unknown> | undefined
  if (!pPr) return false
  const pStyle = pPr['w:pStyle'] as Record<string, unknown> | undefined
  const styleId = pStyle ? (pStyle['@_w:val'] as string | undefined) : undefined
  if (!styleId) return false
  if (/^toc\s*\d$/i.test(styleId)) return true
  const styleEntry = ctx.styles.get(styleId)
  const name = styleEntry?.name
  if (name && /^toc\s*\d$/i.test(name)) return true
  return false
}

/**
 * 从 TOC 段落中提取一个 tocEntry 节点
 * - text: 标题文本
 * - pageNumber: 页码（右侧末尾数字文本）
 * - level: 依据段落样式名（TOC1/TOC2/...）推断
 * - href: 首个 hyperlink 的 anchor
 */
export function parseTocParagraph(
  p: Record<string, unknown>,
  ctx: ParseContext,
): TiptapNode | null {
  // 收集所有文本片段
  const textSegments: string[] = []
  let href: string | undefined

  const hyperlinks = ensureArray(p['w:hyperlink'] as Record<string, unknown>[])
  for (const hl of hyperlinks) {
    const anchor = (hl['@_w:anchor'] as string | undefined)
    if (anchor && !href) href = `#${anchor}`
    const hlRuns = ensureArray(hl['w:r'] as Record<string, unknown>[])
    for (const r of hlRuns) {
      collectRunText(r, textSegments)
    }
  }

  const runs = ensureArray(p['w:r'] as Record<string, unknown>[])
  for (const r of runs) {
    collectRunText(r, textSegments)
  }

  const fullText = textSegments.join('').replace(/\u0013|\u0014|\u0015/g, '').trim()
  if (!fullText) return null

  // 尝试拆分「文本 + 页码」：通常末尾是一串数字
  let text = fullText
  let pageNumber = ''
  const pageMatch = fullText.match(/^(.*?)[\s\t.]+([0-9ivxlcIVXLC]+)\s*$/)
  if (pageMatch) {
    text = pageMatch[1].trim()
    pageNumber = pageMatch[2].trim()
  }

  // 样式推断 level
  let level = 1
  const pPr = p['w:pPr'] as Record<string, unknown> | undefined
  if (pPr) {
    const pStyle = pPr['w:pStyle'] as Record<string, unknown> | undefined
    const styleId = pStyle ? (pStyle['@_w:val'] as string | undefined) : undefined
    if (styleId) {
      const match = styleId.match(/^toc\s*(\d)/i) || styleId.match(/^TOC(\d)$/)
      if (match) level = Math.max(1, Math.min(9, Number(match[1])))
      else {
        const style = ctx.styles.get(styleId)
        const name = style?.name ?? styleId
        const m2 = name.match(/toc\s*(\d)/i)
        if (m2) level = Math.max(1, Math.min(9, Number(m2[1])))
      }
    }
  }

  const attrs: Record<string, unknown> = { text, pageNumber, level }
  if (href) attrs.href = href

  return createNode('tocEntry', attrs)
}

function collectRunText(run: Record<string, unknown>, out: string[]): void {
  const ts = ensureArray(run['w:t'] as unknown[])
  for (const t of ts) {
    if (typeof t === 'string') out.push(t)
    else if (t && typeof t === 'object') {
      const inner = (t as Record<string, unknown>)['#text']
      if (typeof inner === 'string') out.push(inner)
    }
  }
  const tabs = ensureArray(run['w:tab'] as Record<string, unknown>[])
  for (let i = 0; i < tabs.length; i++) out.push('\t')
}

/**
 * 处理 w:sdt 元素
 * - TOC：为每个段落生成 tocEntry 节点
 * - 其他：递归处理 sdtContent 内部的段落/表格（使用上层 parseBlockContent 回调）
 */
export function handleSdt(
  sdt: Record<string, unknown>,
  sdtOrdered: OrderedNode | null,
  ctx: ParseContext,
  parseBlockContent: (
    parent: Record<string, unknown>,
    orderedParent: OrderedNode | null,
    ctx: ParseContext,
  ) => TiptapNode[],
): TiptapNode[] {
  const sdtContent = sdt['w:sdtContent'] as Record<string, unknown> | undefined
  if (!sdtContent) return []

  const orderedContent = sdtOrdered
    ? (sdtOrdered.children.find(c => c.tag === 'w:sdtContent') ?? null)
    : null

  if (isTocSdt(sdt) || (sdtOrdered && isTocSdtOrdered(sdtOrdered))) {
    // 序列化整块 SDT 的 XML 作为 rawXml 以供精确还原（此处简化：仅保留 entries）
    const entries: TiptapNode[] = []
    const paragraphs = ensureArray(sdtContent['w:p'] as Record<string, unknown>[])
    for (const p of paragraphs) {
      const entry = parseTocParagraph(p, ctx)
      if (entry) entries.push(entry)
    }
    if (entries.length === 0) {
      ctx.logs.warn.push('TOC SDT parsed but no entries extracted')
      return parseBlockContent(sdtContent, orderedContent, ctx)
    }
    return entries
  }

  // 普通 SDT：透传内部内容
  return parseBlockContent(sdtContent, orderedContent, ctx)
}

/**
 * 从段落节点中提取 w:bookmarkStart 的 name 列表（用于段落 attrs.bookmarks）
 */
export function extractBookmarkNames(p: Record<string, unknown>): string[] {
  const names: string[] = []
  const bookmarks = ensureArray(p['w:bookmarkStart'] as Record<string, unknown>[])
  for (const bk of bookmarks) {
    const name = bk['@_w:name'] as string | undefined
    if (name && !name.startsWith('_GoBack')) names.push(name)
  }
  return names
}

import type { TiptapNode } from '../../types/tiptapJson.js'
import type { ParseContext, ParagraphProperties, RunProperties } from '../../types/ooxml.js'
import { createNode } from '../../types/tiptapJson.js'
import { ensureArray, getWVal } from '../xmlParser.js'
import { parseParagraphProperties, parseRunProperties } from '../styleResolver.js'
import { handleRun } from './run.js'
import { handleHyperlink } from './hyperlink.js'

const HEADING_STYLE_RE = /^(?:heading|Heading|标题)\s*(\d)$/i
const HEADING_STYLE_PREFIXES = [
  'Heading', 'heading', '标题',
  'Title', 'Subtitle',
]

const QUOTE_STYLES = new Set([
  'quote', 'Quote', '引用',
  'blockquote', 'BlockQuote', 'Block Quote',
  'IntenseQuote', 'Intense Quote', '明显引用',
])

const CODE_STYLES = new Set([
  'code', 'Code', '代码',
  'htmlcode', 'HTMLCode', 'HTML Code',
  'PlainText', 'Plain Text', '纯文本',
])

const BLOCK_TYPES_IN_INLINE = new Set(['pageBreak'])

/** 将 OOXML 的 jc 值映射为 Tiptap textAlign */
function mapTextAlign(jc: string | undefined): string | undefined {
  if (!jc) return undefined
  const map: Record<string, string> = {
    left: 'left', start: 'left',
    center: 'center',
    right: 'right', end: 'right',
    both: 'justify', distribute: 'justify',
  }
  return map[jc]
}

/** 判断段落样式是否为 heading，返回 level 1-6，否则 null */
function detectHeadingLevel(
  pStyleId: string | undefined,
  pPr: ParagraphProperties,
  ctx: ParseContext,
): number | null {
  // 1. outlineLvl 直接指定
  if (pPr.outlineLvl != null && pPr.outlineLvl >= 0 && pPr.outlineLvl <= 5) {
    return pPr.outlineLvl + 1
  }

  if (!pStyleId) return null

  // 2. 样式名匹配
  const style = ctx.styles.get(pStyleId)
  const styleName = style?.name ?? pStyleId

  const match = styleName.match(HEADING_STYLE_RE)
  if (match) {
    const lvl = Number(match[1])
    if (lvl >= 1 && lvl <= 6) return lvl
  }

  // 3. Title → h1, Subtitle → h2
  if (/^(Title|标题)$/i.test(styleName)) return 1
  if (/^(Subtitle|副标题)$/i.test(styleName)) return 2

  // 4. 样式链上检查 outlineLvl
  if (style?.pPr.outlineLvl != null && style.pPr.outlineLvl >= 0 && style.pPr.outlineLvl <= 5) {
    return style.pPr.outlineLvl + 1
  }

  return null
}

export function handleParagraph(
  p: Record<string, unknown>,
  ctx: ParseContext,
): TiptapNode | TiptapNode[] | null {
  // 解析段落属性
  const pPrNode = p['w:pPr'] as Record<string, unknown> | undefined
  const directPPr = pPrNode ? parseParagraphProperties(pPrNode) : {}

  // 获取段落样式
  let pStyleId: string | undefined
  let stylePPr: ParagraphProperties = {}
  let styleRPr: RunProperties = {}

  if (pPrNode) {
    const pStyleNode = pPrNode['w:pStyle'] as Record<string, unknown> | undefined
    if (pStyleNode) {
      pStyleId = getWVal(pStyleNode)
      if (pStyleId) {
        const style = ctx.styles.get(pStyleId)
        if (style) {
          stylePPr = style.pPr
          styleRPr = style.rPr
        }
      }
    }
  }

  const mergedPPr: ParagraphProperties = { ...stylePPr, ...directPPr }
  const textAlign = mapTextAlign(mergedPPr.jc)

  // 解析段落内的 rPr（段落默认格式）
  const pRPrNode = pPrNode?.['w:rPr'] as Record<string, unknown> | undefined
  const paragraphRPr = pRPrNode ? parseRunProperties(pRPrNode) : {}
  const mergedStyleRPr: RunProperties = { ...styleRPr, ...paragraphRPr }

  // 检查是否为编号列表（交给 list handler 处理）
  const numPr = mergedPPr.numPr ?? directPPr.numPr
  if (numPr && numPr.numId > 0) {
    return buildListParagraph(p, numPr, mergedPPr, mergedStyleRPr, textAlign, ctx)
  }

  // 处理所有子元素
  const content = processInlineContent(p, mergedStyleRPr, ctx)

  // 判断 heading
  const headingLevel = detectHeadingLevel(pStyleId, mergedPPr, ctx)

  // 判断 blockquote / codeBlock（基于样式名称）
  const styleName = pStyleId ? (ctx.styles.get(pStyleId)?.name ?? pStyleId) : undefined
  const isBlockquote = !!(styleName && QUOTE_STYLES.has(styleName))
  const isCodeBlock = !!(styleName && CODE_STYLES.has(styleName))

  // 检测行内内容中混入的块级节点（如 w:br type="page" 产生的 pageBreak）
  const hasBlockInInline = content.some(n => BLOCK_TYPES_IN_INLINE.has(n.type))
  if (hasBlockInInline) {
    return splitAtBlockNodes(content, headingLevel, textAlign, isBlockquote, isCodeBlock)
  }

  if (headingLevel) {
    const attrs: Record<string, unknown> = { level: headingLevel }
    if (textAlign) attrs.textAlign = textAlign
    return createNode('heading', attrs, content.length > 0 ? content : undefined)
  }

  if (isBlockquote) {
    const paraAttrs: Record<string, unknown> = {}
    if (textAlign) paraAttrs.textAlign = textAlign
    const innerPara = createNode(
      'paragraph',
      Object.keys(paraAttrs).length > 0 ? paraAttrs : undefined,
      content.length > 0 ? content : undefined,
    )
    return createNode('blockquote', undefined, [innerPara])
  }

  if (isCodeBlock) {
    const textContent = content
      .filter(n => n.type === 'text')
      .map(n => n.text || '')
      .join('')
    return createNode('codeBlock', { language: '' }, textContent
      ? [{ type: 'text', text: textContent }]
      : undefined)
  }

  // 普通段落
  const attrs: Record<string, unknown> = {}
  if (textAlign) attrs.textAlign = textAlign

  return createNode(
    'paragraph',
    Object.keys(attrs).length > 0 ? attrs : undefined,
    content.length > 0 ? content : undefined,
  )
}

/**
 * 当行内内容中混入块级节点时，按块级节点边界拆分，
 * 返回 [paragraphOrHeading, pageBreak, paragraphOrHeading, ...] 的数组。
 */
function splitAtBlockNodes(
  inlineContent: TiptapNode[],
  headingLevel: number | null,
  textAlign: string | undefined,
  isBlockquote: boolean,
  isCodeBlock: boolean,
): TiptapNode[] {
  const result: TiptapNode[] = []
  let buffer: TiptapNode[] = []

  const flushBuffer = () => {
    const nodes = buffer
    buffer = []
    if (isCodeBlock) {
      const textContent = nodes.filter(n => n.type === 'text').map(n => n.text || '').join('')
      result.push(createNode('codeBlock', { language: '' }, textContent ? [{ type: 'text', text: textContent }] : undefined))
      return
    }
    const wrapperType = headingLevel ? 'heading' : 'paragraph'
    const attrs: Record<string, unknown> = {}
    if (headingLevel) attrs.level = headingLevel
    if (textAlign) attrs.textAlign = textAlign
    const wrapped = createNode(
      wrapperType,
      Object.keys(attrs).length > 0 ? attrs : undefined,
      nodes.length > 0 ? nodes : undefined,
    )
    if (isBlockquote) {
      result.push(createNode('blockquote', undefined, [wrapped]))
    } else {
      result.push(wrapped)
    }
  }

  for (const node of inlineContent) {
    if (BLOCK_TYPES_IN_INLINE.has(node.type)) {
      flushBuffer()
      result.push(node)
    } else {
      buffer.push(node)
    }
  }
  if (buffer.length > 0) {
    flushBuffer()
  }

  return result.length > 0 ? result : [createNode('paragraph')]
}

/** 处理段落内所有内联子元素 */
export function processInlineContent(
  parent: Record<string, unknown>,
  parentStyleRPr: RunProperties,
  ctx: ParseContext,
): TiptapNode[] {
  const content: TiptapNode[] = []

  // OOXML 段落子元素顺序：w:pPr, w:r, w:hyperlink, w:bookmarkStart/End, w:fldSimple...
  // fast-xml-parser 不保留混合顺序，需要根据属性逐个处理
  // 按照 fast-xml-parser 的结构，同类标签合并为数组

  const runs = ensureArray(parent['w:r'] as Record<string, unknown>[])
  for (const run of runs) {
    content.push(...handleRun(run, ctx, parentStyleRPr))
  }

  const hyperlinks = ensureArray(parent['w:hyperlink'] as Record<string, unknown>[])
  for (const hl of hyperlinks) {
    content.push(...handleHyperlink(hl, ctx, parentStyleRPr))
  }

  return content
}

/** 构建列表段落（返回带有 numPr 信息的 paragraph，由 list handler 后续包裹） */
function buildListParagraph(
  p: Record<string, unknown>,
  numPr: { numId: number; ilvl: number },
  _pPr: ParagraphProperties,
  styleRPr: RunProperties,
  textAlign: string | undefined,
  ctx: ParseContext,
): TiptapNode {
  const content = processInlineContent(p, styleRPr, ctx)

  const attrs: Record<string, unknown> = {}
  if (textAlign) attrs.textAlign = textAlign

  const node = createNode(
    'paragraph',
    Object.keys(attrs).length > 0 ? attrs : undefined,
    content.length > 0 ? content : undefined,
  )

  // 附加列表信息（供 list handler 使用）
  ;(node as unknown as Record<string, unknown>).__numPr = numPr
  return node
}

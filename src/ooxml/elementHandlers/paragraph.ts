import type { TiptapNode, TiptapMark } from '../../types/tiptapJson.js'
import type { ParseContext, ParagraphProperties, RunProperties } from '../../types/ooxml.js'
import { createNode, createTextNode } from '../../types/tiptapJson.js'
import { ensureArray, getWVal } from '../xmlParser.js'
import { parseParagraphProperties, parseRunProperties } from '../styleResolver.js'
import { handleRun } from './run.js'
import { handleHyperlink } from './hyperlink.js'
import { extractBookmarkNames } from './sdt.js'

const HEADING_STYLE_RE = /^(?:heading|Heading|标题)\s*(\d)$/i

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

function detectHeadingLevel(
  pStyleId: string | undefined,
  pPr: ParagraphProperties,
  ctx: ParseContext,
): number | null {
  if (pPr.outlineLvl != null && pPr.outlineLvl >= 0 && pPr.outlineLvl <= 5) {
    return pPr.outlineLvl + 1
  }

  if (!pStyleId) return null

  const style = ctx.styles.get(pStyleId)
  const styleName = style?.name ?? pStyleId

  const match = styleName.match(HEADING_STYLE_RE)
  if (match) {
    const lvl = Number(match[1])
    if (lvl >= 1 && lvl <= 6) return lvl
  }

  if (/^(Title|标题)$/i.test(styleName)) return 1
  if (/^(Subtitle|副标题)$/i.test(styleName)) return 2

  if (style?.pPr.outlineLvl != null && style.pPr.outlineLvl >= 0 && style.pPr.outlineLvl <= 5) {
    return style.pPr.outlineLvl + 1
  }

  return null
}

export function handleParagraph(
  p: Record<string, unknown>,
  ctx: ParseContext,
): TiptapNode | TiptapNode[] | null {
  const pPrNode = p['w:pPr'] as Record<string, unknown> | undefined
  const directPPr = pPrNode ? parseParagraphProperties(pPrNode) : {}

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

  const pRPrNode = pPrNode?.['w:rPr'] as Record<string, unknown> | undefined
  const paragraphRPr = pRPrNode ? parseRunProperties(pRPrNode) : {}
  const mergedStyleRPr: RunProperties = { ...styleRPr, ...paragraphRPr }

  // numPr 来源区分：
  //   - directPPr.numPr：段落自身 pPr 直接写的 numPr，意味着这是一个真列表项。
  //   - stylePPr.numPr：从段落样式（常见于 Heading 1/2/3）继承来的 numPr，
  //     通常是"章节自动编号"的装饰，不应把段落分类为列表项。
  //
  // 分流策略（对齐 docx-js-editor 的理念：heading/list 不互斥，style 上的 numPr
  // 视为编号装饰）：
  //   1. 先判定 headingLevel。只要段落语义是 heading（样式名 heading N / outlineLvl
  //      有效），就走 heading 分支，即使 style 层挂了 numPr。
  //   2. 否则若 directPPr.numPr 存在，视为真实列表项，走 list 流水线。
  //   3. 否则若仅 stylePPr.numPr 存在，也走 list 流水线（保持旧行为的向后兼容，
  //      避免影响已正确处理的非 heading 样式列表）。
  const headingLevel = detectHeadingLevel(pStyleId, mergedPPr, ctx)
  const styleNumPr = stylePPr.numPr
  const directNumPr = directPPr.numPr

  if (!headingLevel) {
    const numPr = directNumPr ?? styleNumPr
    if (numPr && numPr.numId > 0) {
      return buildListParagraph(p, numPr, mergedPPr, mergedStyleRPr, textAlign, ctx, buildIndentAttrs(mergedPPr.ind))
    }
  }

  const rawContent = processInlineContent(p, mergedStyleRPr, ctx)
  const content = consolidateRuns(rawContent)

  const styleName = pStyleId ? (ctx.styles.get(pStyleId)?.name ?? pStyleId) : undefined
  const isBlockquote = !!(styleName && QUOTE_STYLES.has(styleName))
  const isCodeBlock = !!(styleName && CODE_STYLES.has(styleName))

  // 书签名附加到段落 attrs
  const bookmarks = extractBookmarkNames(p)

  const hasBlockInInline = content.some(n => BLOCK_TYPES_IN_INLINE.has(n.type))
  const indentAttrs = buildIndentAttrs(mergedPPr.ind)

  // 标题章节编号：只要段落是 heading 且样式层挂着 numPr，就让全局计数器推进一次，
  // 产出 "1"、"1.2"、"1.2.1" 之类前缀文本，作为 **真实 text 节点** 拼到 heading
  // content 的最前面。这样用户可以像编辑普通文字一样选中、修改、删除它，符合
  // "导入后编号就是文档的一部分" 的直觉。
  //
  // 额外仍然把完整前缀保存到 attrs.numberingText，以及把原始 w:numPr 引用记到
  // attrs.__origNumPr：
  //   - numberingText 用于导出阶段判断"用户是否改过编号"（如果 heading 开头的
  //     文字仍等于 numberingText + ' '，说明没动过，导出时可以脱掉让 Word
  //     自动重编号；否则整体保留）。
  //   - __origNumPr 让 localSerializer 有机会回写 w:numPr，保留原 numbering.xml
  //     的样式绑定。
  //
  // 计数器只对主文档正文（word/document.xml）生效，避免脚注 / 页眉页脚等
  // 其它部件的 heading（极罕见）把正文编号拨乱。
  let headingNumberingText: string | undefined
  let headingOrigNumPr: { numId: number; ilvl: number } | undefined
  if (headingLevel) {
    const headingNumPr = directNumPr ?? styleNumPr
    if (headingNumPr && headingNumPr.numId > 0) {
      headingOrigNumPr = { numId: headingNumPr.numId, ilvl: headingNumPr.ilvl ?? 0 }
      if (ctx.partPath === 'word/document.xml' && ctx.headingNumberingCounter) {
        const prefix = ctx.headingNumberingCounter.advance(
          headingOrigNumPr.numId,
          headingOrigNumPr.ilvl,
        )
        if (prefix) headingNumberingText = prefix
      }
    }
  }

  // 把编号作为真实 text 节点 prepend 到 heading 内容里。
  // 仅在当前段落是 heading 且有编号时生效；非 heading 段落不受影响。
  const contentWithNumbering =
    headingLevel && headingNumberingText
      ? prependNumberingText(content, headingNumberingText)
      : content

  if (hasBlockInInline) {
    return splitAtBlockNodes(
      contentWithNumbering,
      headingLevel,
      textAlign,
      isBlockquote,
      isCodeBlock,
      bookmarks,
      indentAttrs,
      headingNumberingText,
      headingOrigNumPr,
    )
  }

  const baseAttrs: Record<string, unknown> = {}
  if (textAlign) baseAttrs.textAlign = textAlign
  if (bookmarks.length > 0) baseAttrs.bookmarks = bookmarks
  Object.assign(baseAttrs, indentAttrs)

  if (headingLevel) {
    const attrs: Record<string, unknown> = { level: headingLevel, ...baseAttrs }
    if (headingNumberingText) attrs.numberingText = headingNumberingText
    if (headingOrigNumPr) attrs.__origNumPr = headingOrigNumPr
    return createNode(
      'heading',
      attrs,
      contentWithNumbering.length > 0 ? contentWithNumbering : undefined,
    )
  }

  if (isBlockquote) {
    const paraAttrs: Record<string, unknown> = { ...baseAttrs }
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

  return createNode(
    'paragraph',
    Object.keys(baseAttrs).length > 0 ? baseAttrs : undefined,
    content.length > 0 ? content : undefined,
  )
}

/**
 * 把章节编号文本作为 text 节点 prepend 到 heading 的 inline content 前。
 *
 *   - 编号和正文之间统一插入一个空格，贴近 Word 渲染的视觉间距。
 *   - 若首个 inline 节点也是 text，直接把前缀并入同一个 text 节点，
 *     避免产生不必要的节点分裂；否则新建 text 节点。
 *   - 前缀 text 不附加任何 marks（编号在视觉上按 heading 本体样式呈现即可，
 *     不继承首个 run 的 bold/color/size，避免"编号特别抢眼"）。
 */
function prependNumberingText(
  content: TiptapNode[],
  numberingText: string,
): TiptapNode[] {
  const prefix = `${numberingText} `
  if (content.length === 0) {
    return [{ type: 'text', text: prefix }]
  }
  const first = content[0]
  if (first.type === 'text' && !first.marks && typeof first.text === 'string') {
    return [{ type: 'text', text: prefix + first.text }, ...content.slice(1)]
  }
  return [{ type: 'text', text: prefix }, ...content]
}

/**
 * 将 OOXML 的 w:ind 换算为 CSS：
 *   w:firstLine / w:hanging / w:left / w:right 单位是 1/20 磅（twip），
 *   1 twip = 1/1440 英寸，1 英寸 = 96px，所以 px = twip / 15。
 *   *Chars 单位是 1/100 字符，直接用 em 保持与字号联动。
 * 优先使用 Chars 版（和字号同步更符合中文排版"首行缩进 2 字符"意图）。
 * hanging 需要转为负的 textIndent。
 */
function twipsToPx(twip: number): number {
  // 保留两位小数避免整数化造成的漂移
  return Math.round((twip / 15) * 100) / 100
}

export function buildIndentAttrs(
  ind: ParagraphProperties['ind'],
): Record<string, string> {
  if (!ind) return {}
  const out: Record<string, string> = {}

  // 首行缩进 / 悬挂缩进
  if (ind.firstLineChars != null && ind.firstLineChars > 0) {
    out.textIndent = `${ind.firstLineChars / 100}em`
  } else if (ind.firstLine != null && ind.firstLine > 0) {
    out.textIndent = `${twipsToPx(ind.firstLine)}px`
  } else if (ind.hangingChars != null && ind.hangingChars > 0) {
    out.textIndent = `-${ind.hangingChars / 100}em`
  } else if (ind.hanging != null && ind.hanging > 0) {
    out.textIndent = `-${twipsToPx(ind.hanging)}px`
  }

  // 左缩进
  if (ind.leftChars != null && ind.leftChars > 0) {
    out.indent = `${ind.leftChars / 100}em`
  } else if (ind.left != null && ind.left > 0) {
    out.indent = `${twipsToPx(ind.left)}px`
  }

  // 右缩进
  if (ind.rightChars != null && ind.rightChars > 0) {
    out.indentRight = `${ind.rightChars / 100}em`
  } else if (ind.right != null && ind.right > 0) {
    out.indentRight = `${twipsToPx(ind.right)}px`
  }

  return out
}

function splitAtBlockNodes(
  inlineContent: TiptapNode[],
  headingLevel: number | null,
  textAlign: string | undefined,
  isBlockquote: boolean,
  isCodeBlock: boolean,
  bookmarks: string[],
  indentAttrs: Record<string, string>,
  headingNumberingText?: string,
  headingOrigNumPr?: { numId: number; ilvl: number },
): TiptapNode[] {
  const result: TiptapNode[] = []
  let buffer: TiptapNode[] = []
  let firstSegment = true

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
    if (firstSegment && bookmarks.length > 0) attrs.bookmarks = bookmarks
    // 只在首段应用 textIndent，避免被拆分段落的每一段都缩进；margin-left 保留所有段
    if (firstSegment && indentAttrs.textIndent) attrs.textIndent = indentAttrs.textIndent
    if (indentAttrs.indent) attrs.indent = indentAttrs.indent
    if (indentAttrs.indentRight) attrs.indentRight = indentAttrs.indentRight
    // 章节编号 / 原始 numPr：仅首段保留，避免被拆分后重复显示
    if (firstSegment && headingLevel && headingNumberingText) attrs.numberingText = headingNumberingText
    if (firstSegment && headingLevel && headingOrigNumPr) attrs.__origNumPr = headingOrigNumPr
    firstSegment = false
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

/**
 * 处理段落内所有内联子元素
 * 使用有序树保持 w:r / w:hyperlink / w:fldSimple 等的交错顺序
 */
export function processInlineContent(
  parent: Record<string, unknown>,
  parentStyleRPr: RunProperties,
  ctx: ParseContext,
): TiptapNode[] {
  const content: TiptapNode[] = []

  // 试图使用 orderedRoot 定位当前段落并按原始顺序遍历
  // 由于单个段落的定位开销较大，这里退而采用轻量近似：先 runs，再 hyperlinks，再 fldSimple
  // 但我们对 runs 使用 fldChar 状态机以处理复合域
  const runs = ensureArray(parent['w:r'] as Record<string, unknown>[])
  const hyperlinks = ensureArray(parent['w:hyperlink'] as Record<string, unknown>[])
  const fldSimples = ensureArray(parent['w:fldSimple'] as Record<string, unknown>[])

  // 复合域状态机：遍历所有 runs，跟踪 fldChar begin/separate/end
  let fldState: 'none' | 'instr' | 'display' = 'none'
  let instrBuf: string[] = []

  for (const run of runs) {
    // 判断该 run 内是否有 fldChar
    const fldChar = run['w:fldChar'] as Record<string, unknown> | undefined
    if (fldChar) {
      const type = (fldChar as Record<string, unknown>)['@_w:fldCharType'] as string | undefined
      if (type === 'begin') {
        fldState = 'instr'
        instrBuf = []
        continue
      } else if (type === 'separate') {
        fldState = 'display'
        continue
      } else if (type === 'end') {
        fldState = 'none'
        instrBuf = []
        continue
      }
    }

    if (fldState === 'instr') {
      // 收集 w:instrText（此 run 的文本算指令）
      const instrs = ensureArray(run['w:instrText'] as unknown[])
      for (const it of instrs) {
        const text = typeof it === 'string' ? it : (typeof it === 'object' && it !== null ? (it as Record<string, unknown>)['#text'] as string : undefined)
        if (text) instrBuf.push(text)
      }
      continue
    }

    // fldState === 'none' 或 'display'：两种情况都正常输出 run 内的文本
    content.push(...handleRun(run, ctx, parentStyleRPr))
  }

  for (const hl of hyperlinks) {
    content.push(...handleHyperlink(hl, ctx, parentStyleRPr))
  }

  // w:fldSimple → 保留其内部 run 的显示文本；
  // w:instr 属性本身是域指令（如 "PAGE \* MERGEFORMAT"），不作为可见文本输出，也不再冗余告警。
  for (const fs of fldSimples) {
    const innerRuns = ensureArray(fs['w:r'] as Record<string, unknown>[])
    for (const r of innerRuns) {
      content.push(...handleRun(r, ctx, parentStyleRPr))
    }
  }

  return content
}

/**
 * 合并相邻、格式完全相同的 text 节点
 */
function consolidateRuns(nodes: TiptapNode[]): TiptapNode[] {
  if (nodes.length < 2) return nodes
  const out: TiptapNode[] = []
  for (const node of nodes) {
    const last = out[out.length - 1]
    if (
      node.type === 'text' &&
      last &&
      last.type === 'text' &&
      marksEqual(last.marks, node.marks)
    ) {
      last.text = (last.text ?? '') + (node.text ?? '')
    } else {
      out.push(node)
    }
  }
  return out
}

function marksEqual(a?: TiptapMark[], b?: TiptapMark[]): boolean {
  const aa = a ?? []
  const bb = b ?? []
  if (aa.length !== bb.length) return false
  const stableKey = (m: TiptapMark) => `${m.type}:${JSON.stringify(m.attrs ?? {})}`
  const aKeys = aa.map(stableKey).sort()
  const bKeys = bb.map(stableKey).sort()
  for (let i = 0; i < aKeys.length; i++) {
    if (aKeys[i] !== bKeys[i]) return false
  }
  return true
}

function buildListParagraph(
  p: Record<string, unknown>,
  numPr: { numId: number; ilvl: number },
  _pPr: ParagraphProperties,
  styleRPr: RunProperties,
  textAlign: string | undefined,
  ctx: ParseContext,
  indentAttrs: Record<string, string>,
): TiptapNode {
  const rawContent = processInlineContent(p, styleRPr, ctx)
  const content = consolidateRuns(rawContent)

  const bookmarks = extractBookmarkNames(p)

  const attrs: Record<string, unknown> = {}
  if (textAlign) attrs.textAlign = textAlign
  if (bookmarks.length > 0) attrs.bookmarks = bookmarks
  // 列表段落：不应用 indent（列表层级自己处理左缩进），但保留 textIndent（首行缩进显式声明）
  if (indentAttrs.textIndent) attrs.textIndent = indentAttrs.textIndent

  const node = createNode(
    'paragraph',
    Object.keys(attrs).length > 0 ? attrs : undefined,
    content.length > 0 ? content : undefined,
  )

  ;(node as unknown as Record<string, unknown>).__numPr = numPr
  return node
}

void createTextNode

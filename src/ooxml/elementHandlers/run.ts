import type { TiptapMark, TiptapNode } from '../../types/tiptapJson.js'
import type { ParseContext, RunProperties } from '../../types/ooxml.js'
import { createTextNode, createNode } from '../../types/tiptapJson.js'
import { ensureArray, getWVal, getAttr } from '../xmlParser.js'
import { parseRunProperties } from '../styleResolver.js'
import { resolveThemeColor } from '../themeResolver.js'
import { extractImagesFromRun } from './image.js'

/** 半点(half-point) → px，Word 中 1pt = 2 half-points */
function halfPointsToPx(halfPts: number): string {
  const px = Math.round(halfPts / 2 * 1.333)
  return `${px}px`
}

/** 将 RunProperties 转换为 Tiptap marks 数组 */
export function buildMarks(rPr: RunProperties, ctx: ParseContext): TiptapMark[] {
  const marks: TiptapMark[] = []
  const styleAttrs: Record<string, unknown> = {}

  if (rPr.bold) marks.push({ type: 'bold' })
  if (rPr.italic) marks.push({ type: 'italic' })
  if (rPr.underline) marks.push({ type: 'underline' })
  if (rPr.strike) marks.push({ type: 'strike' })
  if (rPr.superscript) marks.push({ type: 'superscript' })
  if (rPr.subscript) marks.push({ type: 'subscript' })

  const color = rPr.color
  if (color) styleAttrs.color = color

  if (rPr.fontSize) styleAttrs.fontSize = halfPointsToPx(rPr.fontSize)

  if (rPr.fontFamily) {
    styleAttrs.fontFamily = rPr.fontFamily
  }

  if (Object.keys(styleAttrs).length > 0) {
    marks.push({ type: 'textStyle', attrs: styleAttrs })
  }

  const highlightColor = rPr.highlight || rPr.shading
  if (highlightColor) {
    marks.push({
      type: 'highlight',
      attrs: { color: resolveHighlightColor(highlightColor) },
    })
  }

  void ctx
  return marks
}

const HIGHLIGHT_COLOR_MAP: Record<string, string> = {
  yellow: '#FFFF00', green: '#00FF00', cyan: '#00FFFF',
  magenta: '#FF00FF', blue: '#0000FF', red: '#FF0000',
  darkBlue: '#00008B', darkCyan: '#008B8B', darkGreen: '#006400',
  darkMagenta: '#8B008B', darkRed: '#8B0000', darkYellow: '#808000',
  darkGray: '#A9A9A9', lightGray: '#D3D3D3', black: '#000000',
  white: '#FFFFFF',
}

function resolveHighlightColor(val: string): string {
  if (val.startsWith('#')) return val
  if (/^[0-9a-fA-F]{6}$/.test(val)) return `#${val}`
  return HIGHLIGHT_COLOR_MAP[val] ?? val
}

/** 合并样式中的 rPr 和直接格式化的 rPr */
export function mergeRunProperties(
  docDefaults: RunProperties,
  stylePr: RunProperties,
  directPr: RunProperties,
): RunProperties {
  return { ...docDefaults, ...stylePr, ...directPr }
}

// ════════ Symbol font unicode mapping (minimal subset) ════════
// Symbol font (MS Symbol) Basic Latin → Greek / math mapping
const SYMBOL_FONT_MAP: Record<string, string> = {
  F020: ' ', F021: '!', F022: '∀', F023: '#', F024: '∃', F025: '%',
  F026: '&', F027: '∋', F028: '(', F029: ')', F02A: '∗', F02B: '+',
  F02C: ',', F02D: '−', F02E: '.', F02F: '/',
  F0B7: '\u2022', // bullet
  F0A7: '\u2022', // sometimes bullet
  F0F0: '→',
}

function mapSymbolChar(font: string | undefined, char: string | undefined): string | undefined {
  if (!char) return undefined
  const key = char.toUpperCase().replace(/^0X/, '').padStart(4, '0')
  if (!font) return SYMBOL_FONT_MAP[key]
  const fontLc = font.toLowerCase()
  if (fontLc.includes('symbol') || fontLc.includes('wingdings')) {
    // 简易回退：把 0xF0XX 私用区映射到 0x00XX 基本拉丁
    const mapped = SYMBOL_FONT_MAP[key]
    if (mapped) return mapped
    const num = parseInt(key, 16)
    if (!isNaN(num) && num >= 0xF020 && num <= 0xF07E) {
      return String.fromCharCode(num - 0xF000)
    }
  }
  return SYMBOL_FONT_MAP[key]
}

/**
 * 解包 mc:AlternateContent 容器（优先 mc:Choice，回退 mc:Fallback）
 * 将解包后的内容合并到当前 run 级的处理上下文
 * 返回虚拟 drawing 数组（展平）
 */
function unpackAlternateContent(ac: Record<string, unknown>): Record<string, unknown>[] {
  const result: Record<string, unknown>[] = []
  const choices = ensureArray(ac['mc:Choice'] as Record<string, unknown>[])
  const fallback = ensureArray(ac['mc:Fallback'] as Record<string, unknown>[])

  const candidates = choices.length > 0 ? choices : fallback
  for (const candidate of candidates) {
    result.push(candidate)
  }
  return result
}

/** 处理 w:r 元素，返回 TiptapNode[] */
export function handleRun(
  run: Record<string, unknown>,
  ctx: ParseContext,
  parentStyleRPr: RunProperties,
): TiptapNode[] {
  const nodes: TiptapNode[] = []

  const rPrNode = run['w:rPr'] as Record<string, unknown> | undefined
  const directRPr = rPrNode ? parseRunProperties(rPrNode) : {}

  let styleRPr: RunProperties = {}
  if (rPrNode) {
    const rStyleNode = rPrNode['w:rStyle'] as Record<string, unknown> | undefined
    if (rStyleNode) {
      const styleId = getWVal(rStyleNode)
      if (styleId) {
        const style = ctx.styles.get(styleId)
        if (style) styleRPr = style.rPr
      }
    }
  }

  if (rPrNode) {
    const colorNode = rPrNode['w:color'] as Record<string, unknown> | undefined
    if (colorNode && !directRPr.color) {
      const themeColor = getAttr(colorNode, 'w:themeColor')
      if (themeColor) {
        const tint = getAttr(colorNode, 'w:themeTint')
        const shade = getAttr(colorNode, 'w:themeShade')
        const resolved = resolveThemeColor(ctx.theme.colors, themeColor, tint, shade)
        if (resolved) directRPr.color = resolved
      }
    }

    const rFonts = rPrNode['w:rFonts'] as Record<string, unknown> | undefined
    if (rFonts && !directRPr.fontFamily) {
      const themeRef = getAttr(rFonts, 'w:asciiTheme') ?? getAttr(rFonts, 'w:eastAsiaTheme')
      if (themeRef) {
        if (themeRef.includes('major')) {
          directRPr.fontFamily = ctx.theme.fonts.majorEastAsia || ctx.theme.fonts.majorLatin
        } else if (themeRef.includes('minor')) {
          directRPr.fontFamily = ctx.theme.fonts.minorEastAsia || ctx.theme.fonts.minorLatin
        }
      }
    }
  }

  const mergedRPr = mergeRunProperties(ctx.docDefaults.rPr, { ...parentStyleRPr, ...styleRPr }, directRPr)
  const marks = buildMarks(mergedRPr, ctx)
  const marksOrUndef = marks.length > 0 ? marks : undefined

  // w:t 文本
  const texts = ensureArray(run['w:t'] as unknown[])
  for (const t of texts) {
    const text = typeof t === 'string' ? t : (typeof t === 'object' && t !== null ? (t as Record<string, unknown>)['#text'] as string : undefined)
    if (text) {
      nodes.push(createTextNode(text, marksOrUndef))
    }
  }

  // w:instrText → 域指令文本。
  // 正常情况下 instrText 应由上层（paragraph.processInlineContent / hyperlink.handleHyperlink）
  // 的 fldChar 状态机在 'instr' 阶段消费掉；能流到这里说明上层状态机没对上（例如字段跨段落、
  // 或 begin/separate/end 缺失）。为避免将 "PAGEREF _Toc... \h"、"PAGE \* MERGEFORMAT" 等
  // 污染成可见文本，这里统一静默忽略，不再逐条告警（之前会把导入日志刷爆 50+ 条）。
  // 如需排查可以把下面的注释打开，按需改成 ctx.logs.info 记录。
  // const instrs = ensureArray(run['w:instrText'] as unknown[])
  // for (const it of instrs) { ... }

  // w:tab 数组化（TOC 常见多个 tab）
  const tabs = ensureArray(run['w:tab'] as Record<string, unknown>[])
  for (let i = 0; i < tabs.length; i++) {
    nodes.push(createTextNode('\t', marksOrUndef))
  }

  // w:br (inline break, not page break)
  const brs = ensureArray(run['w:br'] as Record<string, unknown>[])
  for (const br of brs) {
    const brType = getAttr(br, 'w:type')
    if (brType === 'page') {
      nodes.push({ type: 'pageBreak' })
    } else {
      nodes.push({ type: 'hardBreak' })
    }
  }

  // w:cr → hardBreak
  if (run['w:cr'] !== undefined) {
    const crs = ensureArray(run['w:cr'] as unknown[])
    for (let i = 0; i < crs.length; i++) nodes.push({ type: 'hardBreak' })
  }

  // w:noBreakHyphen → \u2011
  if (run['w:noBreakHyphen'] !== undefined) {
    nodes.push(createTextNode('\u2011', marksOrUndef))
  }

  // w:softHyphen → \u00AD
  if (run['w:softHyphen'] !== undefined) {
    nodes.push(createTextNode('\u00AD', marksOrUndef))
  }

  // w:sym → 符号字符
  const syms = ensureArray(run['w:sym'] as Record<string, unknown>[])
  for (const sym of syms) {
    const font = getAttr(sym, 'w:font')
    const char = getAttr(sym, 'w:char')
    const mapped = mapSymbolChar(font, char)
    if (mapped) {
      nodes.push(createTextNode(mapped, marksOrUndef))
    } else if (char) {
      const num = parseInt(char, 16)
      if (!isNaN(num)) {
        nodes.push(createTextNode(String.fromCharCode(num), marksOrUndef))
      }
    }
  }

  // w:fldChar 复合域：begin/separate/end 状态由段落级处理器管理，此处暂不处理具体状态
  // 但如果单 run 内就包含 fldChar，记录日志
  if (run['w:fldChar'] !== undefined) {
    // 无法单 run 还原域值。保留 run 内其他文本（上面已处理）
  }

  // w:footnoteReference / w:endnoteReference → 上标数字
  const fnRefs = ensureArray(run['w:footnoteReference'] as Record<string, unknown>[])
  for (const ref of fnRefs) {
    const idStr = getAttr(ref, 'w:id')
    if (idStr != null) {
      // 上标样式
      const supMarks: TiptapMark[] = [...(marksOrUndef ?? []), { type: 'superscript' }]
      nodes.push(createTextNode(`[${idStr}]`, supMarks))
    }
  }
  const enRefs = ensureArray(run['w:endnoteReference'] as Record<string, unknown>[])
  for (const ref of enRefs) {
    const idStr = getAttr(ref, 'w:id')
    if (idStr != null) {
      const supMarks: TiptapMark[] = [...(marksOrUndef ?? []), { type: 'superscript' }]
      nodes.push(createTextNode(`[${idStr}]`, supMarks))
    }
  }

  // mc:AlternateContent 解包（Word 把 drawing 包在降级容器下）
  if (run['mc:AlternateContent'] !== undefined) {
    const acs = ensureArray(run['mc:AlternateContent'] as Record<string, unknown>[])
    for (const ac of acs) {
      const containers = unpackAlternateContent(ac)
      for (const c of containers) {
        // 展平：把 drawing / pict 等放到虚拟 run 上重新处理
        const fakeRun: Record<string, unknown> = {}
        if (c['w:drawing']) fakeRun['w:drawing'] = c['w:drawing']
        if (c['w:pict']) fakeRun['w:pict'] = c['w:pict']
        const imgs = extractImagesFromRun(fakeRun, ctx)
        nodes.push(...imgs)
      }
    }
  }

  // w:drawing / w:pict → images
  const images = extractImagesFromRun(run, ctx)
  nodes.push(...images)

  // 未支持元素的警告
  if (run['w:object'] !== undefined) {
    ctx.logs.warn.push('Unsupported: w:object (embedded OLE)')
  }
  if (run['w:ruby'] !== undefined) {
    ctx.logs.warn.push('Unsupported: w:ruby (East Asian annotation)')
  }
  // cached layout hint, intentionally skipped
  void run['w:lastRenderedPageBreak']

  // 提取文本框内容（wps:wsp > wps:txbx > w:txbxContent）中的段落并展平为文本节点
  // 复杂文本框通过 image.handleDrawing 处理；这里处理 AlternateContent 内未被覆盖的情况由 image.ts 负责
  void createNode

  return nodes
}

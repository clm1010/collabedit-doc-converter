import type { TiptapMark, TiptapNode } from '../../types/tiptapJson.js'
import type { ParseContext, RunProperties } from '../../types/ooxml.js'
import { createTextNode } from '../../types/tiptapJson.js'
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

  // textStyle attrs
  let color = rPr.color
  if (!color && rPr.color === undefined) {
    // 检查主题颜色
  }
  if (color) styleAttrs.color = color

  if (rPr.fontSize) styleAttrs.fontSize = halfPointsToPx(rPr.fontSize)

  if (rPr.fontFamily) {
    styleAttrs.fontFamily = rPr.fontFamily
  }

  if (Object.keys(styleAttrs).length > 0) {
    marks.push({ type: 'textStyle', attrs: styleAttrs })
  }

  // highlight
  const highlightColor = rPr.highlight || rPr.shading
  if (highlightColor) {
    marks.push({
      type: 'highlight',
      attrs: { color: resolveHighlightColor(highlightColor) },
    })
  }

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

/** 处理 w:r 元素，返回 TiptapNode[] */
export function handleRun(
  run: Record<string, unknown>,
  ctx: ParseContext,
  parentStyleRPr: RunProperties,
): TiptapNode[] {
  const nodes: TiptapNode[] = []

  // 解析直接格式化
  const rPrNode = run['w:rPr'] as Record<string, unknown> | undefined
  const directRPr = rPrNode ? parseRunProperties(rPrNode) : {}

  // 检查 rPr 中引用的样式
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

  // 处理主题颜色引用
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

    // 主题字体引用
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

  // 处理文本内容
  const texts = ensureArray(run['w:t'] as unknown[])
  for (const t of texts) {
    const text = typeof t === 'string' ? t : (typeof t === 'object' && t !== null ? (t as Record<string, unknown>)['#text'] as string : undefined)
    if (text) {
      nodes.push(createTextNode(text, marks.length > 0 ? marks : undefined))
    }
  }

  // w:tab → tab character
  if (run['w:tab'] !== undefined) {
    nodes.push(createTextNode('\t', marks.length > 0 ? marks : undefined))
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

  // w:drawing / w:pict → images
  const images = extractImagesFromRun(run, ctx)
  nodes.push(...images)

  return nodes
}

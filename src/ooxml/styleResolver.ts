import type { ResolvedStyle, ParagraphProperties, RunProperties } from '../types/ooxml.js'
import { parseXml, ensureArray, getAttr, getWVal } from './xmlParser.js'
import type { DocxArchive } from './zipExtractor.js'

export interface StyleResolverResult {
  styles: Map<string, ResolvedStyle>
  docDefaults: { rPr: RunProperties; pPr: ParagraphProperties }
}

export function resolveStyles(archive: DocxArchive): StyleResolverResult {
  const xml = archive.getText('word/styles.xml')
  const styles = new Map<string, ResolvedStyle>()
  const docDefaults: { rPr: RunProperties; pPr: ParagraphProperties } = {
    rPr: {},
    pPr: {},
  }

  if (!xml) return { styles, docDefaults }

  const parsed = parseXml(xml)
  const root = parsed['w:styles'] as Record<string, unknown> | undefined
  if (!root) return { styles, docDefaults }

  // 1. 解析 docDefaults
  const dd = root['w:docDefaults'] as Record<string, unknown> | undefined
  if (dd) {
    const rPrDefault = dd['w:rPrDefault'] as Record<string, unknown> | undefined
    if (rPrDefault) {
      const rPr = rPrDefault['w:rPr'] as Record<string, unknown> | undefined
      if (rPr) docDefaults.rPr = parseRunProperties(rPr)
    }
    const pPrDefault = dd['w:pPrDefault'] as Record<string, unknown> | undefined
    if (pPrDefault) {
      const pPr = pPrDefault['w:pPr'] as Record<string, unknown> | undefined
      if (pPr) docDefaults.pPr = parseParagraphProperties(pPr)
    }
  }

  // 2. 第一轮：解析所有样式（不解析继承）
  const rawStyles = ensureArray(root['w:style'] as Record<string, unknown>[])
  for (const style of rawStyles) {
    const styleId = getAttr(style, 'w:styleId')
    if (!styleId) continue

    const type = (getAttr(style, 'w:type') ?? 'paragraph') as ResolvedStyle['type']
    const name = getWVal(style['w:name'] as Record<string, unknown>)
    const basedOn = getWVal(style['w:basedOn'] as Record<string, unknown>)

    const pPr = style['w:pPr'] as Record<string, unknown> | undefined
    const rPr = style['w:rPr'] as Record<string, unknown> | undefined

    styles.set(styleId, {
      styleId,
      name,
      type,
      basedOn,
      pPr: pPr ? parseParagraphProperties(pPr) : {},
      rPr: rPr ? parseRunProperties(rPr) : {},
    })
  }

  // 3. 第二轮：递归解析继承链（子覆盖父）
  const resolved = new Map<string, ResolvedStyle>()
  for (const [id] of styles) {
    resolveInheritance(id, styles, resolved, new Set())
  }

  return { styles: resolved, docDefaults }
}

function resolveInheritance(
  styleId: string,
  raw: Map<string, ResolvedStyle>,
  resolved: Map<string, ResolvedStyle>,
  visited: Set<string>,
): ResolvedStyle {
  if (resolved.has(styleId)) return resolved.get(styleId)!

  const style = raw.get(styleId)
  if (!style) {
    const empty: ResolvedStyle = { styleId, type: 'paragraph', pPr: {}, rPr: {} }
    resolved.set(styleId, empty)
    return empty
  }

  if (visited.has(styleId)) {
    resolved.set(styleId, style)
    return style
  }
  visited.add(styleId)

  if (style.basedOn && raw.has(style.basedOn)) {
    const parent = resolveInheritance(style.basedOn, raw, resolved, visited)
    const merged: ResolvedStyle = {
      ...style,
      pPr: { ...parent.pPr, ...style.pPr },
      rPr: { ...parent.rPr, ...style.rPr },
    }
    resolved.set(styleId, merged)
    return merged
  }

  resolved.set(styleId, style)
  return style
}

export function parseParagraphProperties(pPr: Record<string, unknown>): ParagraphProperties {
  const result: ParagraphProperties = {}

  const jc = pPr['w:jc'] as Record<string, unknown> | undefined
  if (jc) result.jc = getWVal(jc)

  const outlineLvl = pPr['w:outlineLvl'] as Record<string, unknown> | undefined
  if (outlineLvl) {
    const val = getWVal(outlineLvl)
    if (val != null) result.outlineLvl = Number(val)
  }

  const numPr = pPr['w:numPr'] as Record<string, unknown> | undefined
  if (numPr) {
    const numId = getWVal(numPr['w:numId'] as Record<string, unknown>)
    const ilvl = getWVal(numPr['w:ilvl'] as Record<string, unknown>)
    if (numId != null) {
      result.numPr = { numId: Number(numId), ilvl: Number(ilvl ?? 0) }
    }
  }

  const spacing = pPr['w:spacing'] as Record<string, unknown> | undefined
  if (spacing) {
    result.spacing = {
      before: numOrUndef(getAttr(spacing, 'w:before')),
      after: numOrUndef(getAttr(spacing, 'w:after')),
      line: numOrUndef(getAttr(spacing, 'w:line')),
      lineRule: getAttr(spacing, 'w:lineRule'),
    }
  }

  const ind = pPr['w:ind'] as Record<string, unknown> | undefined
  if (ind) {
    result.ind = {
      left: numOrUndef(getAttr(ind, 'w:left')),
      right: numOrUndef(getAttr(ind, 'w:right')),
      firstLine: numOrUndef(getAttr(ind, 'w:firstLine')),
      hanging: numOrUndef(getAttr(ind, 'w:hanging')),
    }
  }

  return result
}

export function parseRunProperties(rPr: Record<string, unknown>): RunProperties {
  const result: RunProperties = {}

  if (rPr['w:b'] !== undefined) result.bold = !isFalseVal(rPr['w:b'])
  if (rPr['w:bCs'] !== undefined && result.bold === undefined) result.bold = !isFalseVal(rPr['w:bCs'])
  if (rPr['w:i'] !== undefined) result.italic = !isFalseVal(rPr['w:i'])
  if (rPr['w:iCs'] !== undefined && result.italic === undefined) result.italic = !isFalseVal(rPr['w:iCs'])
  if (rPr['w:u'] !== undefined) {
    const uVal = getWVal(rPr['w:u'] as Record<string, unknown>)
    result.underline = uVal !== 'none' && uVal !== undefined
  }
  if (rPr['w:strike'] !== undefined) result.strike = !isFalseVal(rPr['w:strike'])

  const vertAlign = rPr['w:vertAlign'] as Record<string, unknown> | undefined
  if (vertAlign) {
    const val = getWVal(vertAlign)
    if (val === 'superscript') { result.superscript = true; result.vertAlign = 'superscript' }
    if (val === 'subscript') { result.subscript = true; result.vertAlign = 'subscript' }
  }

  const color = rPr['w:color'] as Record<string, unknown> | undefined
  if (color) {
    const val = getWVal(color)
    if (val && val !== 'auto') result.color = normalizeColor(val)
  }

  const sz = rPr['w:sz'] as Record<string, unknown> | undefined
  if (sz) {
    const val = getWVal(sz)
    if (val) result.fontSize = Number(val)
  }

  const rFonts = rPr['w:rFonts'] as Record<string, unknown> | undefined
  if (rFonts) {
    result.fontFamily =
      getAttr(rFonts, 'w:eastAsia') ??
      getAttr(rFonts, 'w:ascii') ??
      getAttr(rFonts, 'w:hAnsi') ??
      getAttr(rFonts, 'w:cs')
  }

  const highlight = rPr['w:highlight'] as Record<string, unknown> | undefined
  if (highlight) result.highlight = getWVal(highlight)

  const shd = rPr['w:shd'] as Record<string, unknown> | undefined
  if (shd) {
    const fill = getAttr(shd, 'w:fill')
    if (fill && fill !== 'auto') result.shading = normalizeColor(fill)
  }

  return result
}

function normalizeColor(val: string): string {
  if (val.startsWith('#')) return val
  if (/^[0-9a-fA-F]{6}$/.test(val)) return `#${val}`
  return val
}

function isFalseVal(node: unknown): boolean {
  if (node == null) return false
  if (typeof node === 'object') {
    const val = getWVal(node as Record<string, unknown>)
    return val === '0' || val === 'false'
  }
  return false
}

function numOrUndef(val: string | undefined): number | undefined {
  if (val == null) return undefined
  const n = Number(val)
  return isNaN(n) ? undefined : n
}

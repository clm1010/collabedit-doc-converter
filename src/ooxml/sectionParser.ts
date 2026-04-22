import { ensureArray, getAttr } from './xmlParser.js'

export interface SectionPageSetup {
  width?: number // mm
  height?: number
  margins?: { top?: number; bottom?: number; left?: number; right?: number }
  orientation?: 'portrait' | 'landscape'
}

export interface SectionHeaderFooterRef {
  default?: string // rId
  first?: string
  even?: string
}

export interface ParsedSection {
  pageSetup?: SectionPageSetup
  headerRefs?: SectionHeaderFooterRef
  footerRefs?: SectionHeaderFooterRef
  titlePg?: boolean
  type?: string // continuous, nextPage...
}

function twipsToMm(twips: number): number {
  return Math.round((twips / 1440) * 25.4 * 10) / 10
}

export function parseSectionProperties(
  sectPr: Record<string, unknown>,
): ParsedSection {
  const section: ParsedSection = {}

  const pgSz = sectPr['w:pgSz'] as Record<string, unknown> | undefined
  if (pgSz) {
    const w = getAttr(pgSz, 'w:w')
    const h = getAttr(pgSz, 'w:h')
    const orient = getAttr(pgSz, 'w:orient')
    section.pageSetup = section.pageSetup ?? {}
    if (w) section.pageSetup.width = twipsToMm(Number(w))
    if (h) section.pageSetup.height = twipsToMm(Number(h))
    if (orient === 'landscape') section.pageSetup.orientation = 'landscape'
    else if (orient === 'portrait') section.pageSetup.orientation = 'portrait'
  }

  const pgMar = sectPr['w:pgMar'] as Record<string, unknown> | undefined
  if (pgMar) {
    section.pageSetup = section.pageSetup ?? {}
    const top = getAttr(pgMar, 'w:top')
    const bottom = getAttr(pgMar, 'w:bottom')
    const left = getAttr(pgMar, 'w:left')
    const right = getAttr(pgMar, 'w:right')
    const margins: NonNullable<SectionPageSetup['margins']> = {}
    if (top) margins.top = twipsToMm(Number(top))
    if (bottom) margins.bottom = twipsToMm(Number(bottom))
    if (left) margins.left = twipsToMm(Number(left))
    if (right) margins.right = twipsToMm(Number(right))
    section.pageSetup.margins = margins
  }

  const typeNode = sectPr['w:type'] as Record<string, unknown> | undefined
  if (typeNode) {
    const v = getAttr(typeNode, 'w:val')
    if (v) section.type = v
  }

  if (sectPr['w:titlePg'] !== undefined) section.titlePg = true

  const headerRefs: SectionHeaderFooterRef = {}
  const headerRefArr = ensureArray(sectPr['w:headerReference'] as Record<string, unknown>[])
  for (const ref of headerRefArr) {
    const type = getAttr(ref, 'w:type') ?? 'default'
    const rId = getAttr(ref, 'r:id')
    if (rId && (type === 'default' || type === 'first' || type === 'even')) {
      headerRefs[type] = rId
    }
  }
  if (Object.keys(headerRefs).length > 0) section.headerRefs = headerRefs

  const footerRefs: SectionHeaderFooterRef = {}
  const footerRefArr = ensureArray(sectPr['w:footerReference'] as Record<string, unknown>[])
  for (const ref of footerRefArr) {
    const type = getAttr(ref, 'w:type') ?? 'default'
    const rId = getAttr(ref, 'r:id')
    if (rId && (type === 'default' || type === 'first' || type === 'even')) {
      footerRefs[type] = rId
    }
  }
  if (Object.keys(footerRefs).length > 0) section.footerRefs = footerRefs

  return section
}

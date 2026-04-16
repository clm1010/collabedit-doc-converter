import type { AbstractNum, NumInstance, NumberingLevel, ParagraphProperties, RunProperties } from '../types/ooxml.js'
import { parseXml, ensureArray, getAttr, getWVal } from './xmlParser.js'
import { parseParagraphProperties, parseRunProperties } from './styleResolver.js'
import type { DocxArchive } from './zipExtractor.js'

export interface NumberingResult {
  abstracts: Map<number, AbstractNum>
  instances: Map<number, NumInstance>
}

export function resolveNumbering(archive: DocxArchive): NumberingResult {
  const xml = archive.getText('word/numbering.xml')
  const abstracts = new Map<number, AbstractNum>()
  const instances = new Map<number, NumInstance>()

  if (!xml) return { abstracts, instances }

  const parsed = parseXml(xml)
  const root = parsed['w:numbering'] as Record<string, unknown> | undefined
  if (!root) return { abstracts, instances }

  // 解析 abstractNum
  const abstractNums = ensureArray(root['w:abstractNum'] as Record<string, unknown>[])
  for (const an of abstractNums) {
    const id = Number(getAttr(an, 'w:abstractNumId') ?? -1)
    if (id < 0) continue

    const levels: NumberingLevel[] = []
    const lvls = ensureArray(an['w:lvl'] as Record<string, unknown>[])
    for (const lvl of lvls) {
      const ilvl = Number(getAttr(lvl, 'w:ilvl') ?? 0)
      const numFmt = getWVal(lvl['w:numFmt'] as Record<string, unknown>) ?? 'decimal'
      const lvlText = getWVal(lvl['w:lvlText'] as Record<string, unknown>) ?? ''
      const start = Number(getWVal(lvl['w:start'] as Record<string, unknown>) ?? 1)

      const pPrNode = lvl['w:pPr'] as Record<string, unknown> | undefined
      const rPrNode = lvl['w:rPr'] as Record<string, unknown> | undefined

      levels.push({
        level: ilvl,
        numFmt,
        lvlText,
        start,
        pPr: pPrNode ? parseParagraphProperties(pPrNode) : {},
        rPr: rPrNode ? parseRunProperties(rPrNode) : {},
      })
    }

    abstracts.set(id, { abstractNumId: id, levels })
  }

  // 解析 num → abstractNum 映射及 lvlOverride
  const nums = ensureArray(root['w:num'] as Record<string, unknown>[])
  for (const num of nums) {
    const numId = Number(getAttr(num, 'w:numId') ?? -1)
    if (numId < 0) continue

    const abstractNumIdRef = num['w:abstractNumId'] as Record<string, unknown> | undefined
    const abstractNumId = Number(getWVal(abstractNumIdRef) ?? -1)

    const overrides = new Map<number, Partial<NumberingLevel>>()
    const lvlOverrides = ensureArray(num['w:lvlOverride'] as Record<string, unknown>[])
    for (const ovr of lvlOverrides) {
      const ilvl = Number(getAttr(ovr, 'w:ilvl') ?? 0)
      const startOverride = ovr['w:startOverride'] as Record<string, unknown> | undefined
      if (startOverride) {
        overrides.set(ilvl, { start: Number(getWVal(startOverride) ?? 1) })
      }
    }

    instances.set(numId, { numId, abstractNumId, overrides })
  }

  return { abstracts, instances }
}

/** 获取指定 numId + ilvl 的编号级别信息 */
export function getNumberingLevel(
  numbering: NumberingResult,
  numId: number,
  ilvl: number,
): NumberingLevel | null {
  const instance = numbering.instances.get(numId)
  if (!instance) return null

  const abstract = numbering.abstracts.get(instance.abstractNumId)
  if (!abstract) return null

  const baseLvl = abstract.levels.find((l) => l.level === ilvl)
  if (!baseLvl) return null

  const override = instance.overrides.get(ilvl)
  if (override) return { ...baseLvl, ...override }
  return baseLvl
}

/** 判断编号格式是 bullet 还是 ordered */
export function isBulletFormat(numFmt: string): boolean {
  return numFmt === 'bullet' || numFmt === 'none'
}

/** 判断是否为中文/CJK 编号格式 */
export function isChineseNumberingFormat(numFmt: string): boolean {
  const chineseFmts = new Set([
    'chineseCounting', 'chineseCountingThousand', 'chineseLegalSimplified',
    'ideographTraditional', 'ideographLegalTraditional', 'ideographDigital',
    'ideographEnclosedCircle', 'ideographZodiac', 'ideographZodiacTraditional',
    'japaneseCounting', 'japaneseLegal', 'japaneseDigitalTenThousand',
    'decimalEnclosedCircle', 'decimalEnclosedCircleChinese',
  ])
  return chineseFmts.has(numFmt)
}

/** fast-xml-parser 解析后的通用节点结构 */
export type XmlNode = Record<string, unknown>

/** .rels 关系映射 rId → target */
export interface RelationshipMap {
  [rId: string]: {
    target: string
    type: string
  }
}

/** 样式继承链中的解析结果 */
export interface ResolvedStyle {
  styleId: string
  name?: string
  type: 'paragraph' | 'character' | 'table' | 'numbering'
  basedOn?: string
  pPr: ParagraphProperties
  rPr: RunProperties
}

export interface ParagraphProperties {
  jc?: string // justification: left | center | right | both
  outlineLvl?: number // heading level (0-based)
  numPr?: { numId: number; ilvl: number }
  pBdr?: Record<string, BorderDef>
  spacing?: { before?: number; after?: number; line?: number; lineRule?: string }
  ind?: { left?: number; right?: number; firstLine?: number; hanging?: number }
}

export interface RunProperties {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  superscript?: boolean
  subscript?: boolean
  color?: string
  fontSize?: number // half-points
  fontFamily?: string
  highlight?: string
  shading?: string
  vertAlign?: 'superscript' | 'subscript' | 'baseline'
}

export interface BorderDef {
  val?: string
  color?: string
  sz?: number
  space?: number
}

export interface NumberingLevel {
  level: number
  numFmt: string
  lvlText: string
  start: number
  pPr: ParagraphProperties
  rPr: RunProperties
}

export interface AbstractNum {
  abstractNumId: number
  levels: NumberingLevel[]
}

export interface NumInstance {
  numId: number
  abstractNumId: number
  overrides: Map<number, Partial<NumberingLevel>>
}

export interface ThemeColors {
  dk1: string
  lt1: string
  dk2: string
  lt2: string
  accent1: string
  accent2: string
  accent3: string
  accent4: string
  accent5: string
  accent6: string
  hlink: string
  folHlink: string
  [key: string]: string
}

export interface ThemeFonts {
  majorLatin: string
  majorEastAsia: string
  minorLatin: string
  minorEastAsia: string
}

export interface ExtractedImage {
  base64: string
  mime: string
  fileName: string
  relPath: string
}

/** 解析上下文 - 在整个文档解析过程中传递 */
export interface ParseContext {
  styles: Map<string, ResolvedStyle>
  numbering: { abstracts: Map<number, AbstractNum>; instances: Map<number, NumInstance> }
  relationships: RelationshipMap
  images: Map<string, ExtractedImage>
  theme: { colors: ThemeColors; fonts: ThemeFonts }
  docDefaults: { rPr: RunProperties; pPr: ParagraphProperties }
  logs: { info: string[]; warn: string[]; error: string[] }
}

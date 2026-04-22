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
  ind?: {
    left?: number
    right?: number
    firstLine?: number
    hanging?: number
    // 字符单位（百分之一字符），中文模板常用来表达 "首行缩进 2 字符"
    leftChars?: number
    rightChars?: number
    firstLineChars?: number
    hangingChars?: number
  }
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
  /** 有序解析树根节点（document.xml），用于在 cell/inline 层获取子元素交错顺序 */
  orderedRoot?: import('../ooxml/xmlParser.js').OrderedNode[]
  /** 脚注上下文：id → 脚注内容 */
  footnotes?: Map<number, import('./tiptapJson.js').TiptapNode[]>
  endnotes?: Map<number, import('./tiptapJson.js').TiptapNode[]>

  // ---- 选择性保存高保真方案：原始字节档案 + 当前部件信息（可选） ----
  /**
   * DOCX 原始字节档案。引入后 handler 可按需读取原始 XML 字节做字节级对齐或
   * 图片 rels 追溯；未设置时回退到历史行为。
   */
  rawArchive?: import('../engine/zipExtractor.js').RawDocxArchive
  /**
   * 当前正在处理的部件路径（如 "word/document.xml" / "word/header2.xml"）。
   * 由 importPipeline / headerFooterParser / footnoteParser 切换上下文时设置。
   * image handler 借此把 `data-origin-part` 写入节点 attrs。
   */
  partPath?: string
  /**
   * 当前部件对应的顶层元素字节范围列表（按文档顺序）。
   * orchestrator 层消费它来为每个顶层 TiptapNode 回填 __origRange/__origHash/__origPart。
   */
  topLevelRanges?: import('../engine/xmlRangeIndexer.js').TopLevelRange[]
  /**
   * SDT id → 原始 w:sdt 字节子串。一个 SDT 若含多条 tocEntry，所有 tocEntry
   * 共享同一条记录（第一条 tocEntry 节点的 __origSdtXml 会写入该字节的 string 形式；
   * 其余 tocEntry 的 __origSdtId 引用该条目）。
   */
  sdtXmlMap?: Map<string, Uint8Array>
  /**
   * 章节标题编号计数器。由 importPipeline 创建，按文档顺序逐个 heading
   * 调用 `advance(numId, ilvl)` 生成 "1"、"1.2"、"1.2.1" 等前缀文本，
   * 并写入 heading attrs.numberingText，前端编辑器据此以 ::before 显示。
   * 仅在正文遍历时消费；页眉/页脚/脚注通常不挂章节编号，可不使用。
   */
  headingNumberingCounter?: import('../ooxml/headingNumberingCounter.js').HeadingNumberingCounter
}

import {
  Document, Packer, Paragraph, TextRun, ImageRun,
  Table, TableRow, TableCell,
  HeadingLevel, AlignmentType,
  Header, Footer,
  PageBreak, BorderStyle, WidthType,
  ExternalHyperlink, ShadingType, UnderlineType,
  convertMillimetersToTwip, LevelFormat,
  VerticalAlign, HeightRule,
  BookmarkStart, BookmarkEnd,
  TabStopType, TabStopPosition,
} from 'docx'
import type { TiptapDoc, TiptapNode, TiptapMark } from '../types/tiptapJson.js'
import type { DocMetadata } from '../types/docMetadata.js'

// ═══════════════════════════════════════════════
// Font size conversion helpers
// ═══════════════════════════════════════════════

function pxToHalfPoints(raw: string | number): number | undefined {
  const n = typeof raw === 'number' ? raw : parseFloat(String(raw))
  if (isNaN(n) || n <= 0) return undefined
  if (typeof raw === 'string' && raw.endsWith('pt')) return Math.round(n * 2)
  return Math.round(n * 1.5)
}

function parsePixel(val: unknown): number | undefined {
  if (typeof val === 'number') return val
  if (typeof val === 'string') {
    const n = parseFloat(val)
    return isNaN(n) ? undefined : n
  }
  return undefined
}

function pxToTwips(px: number): number {
  return Math.round(px * 15)
}

// ═══════════════════════════════════════════════
// Mark → TextRun options
// ═══════════════════════════════════════════════

function marksToRunOptions(marks?: TiptapMark[]): Record<string, unknown> {
  const opts: Record<string, unknown> = {}
  if (!marks) return opts

  for (const mark of marks) {
    switch (mark.type) {
      case 'bold':
        opts.bold = true
        break
      case 'italic':
        opts.italics = true
        break
      case 'underline':
        opts.underline = { type: UnderlineType.SINGLE }
        break
      case 'strike':
        opts.strike = true
        break
      case 'superscript':
        opts.superScript = true
        break
      case 'subscript':
        opts.subScript = true
        break
      case 'textStyle': {
        const a = mark.attrs || {}
        if (a.color) opts.color = String(a.color).replace(/^#/, '')
        if (a.fontSize) {
          const s = pxToHalfPoints(a.fontSize as string | number)
          if (s) opts.size = s
        }
        if (a.fontFamily) opts.font = String(a.fontFamily)
        break
      }
      case 'highlight': {
        const c = String(mark.attrs?.color || '#ffff00').replace(/^#/, '')
        opts.shading = { type: ShadingType.SOLID, color: 'auto', fill: c }
        break
      }
    }
  }
  return opts
}

function toAlignment(val?: unknown): (typeof AlignmentType)[keyof typeof AlignmentType] | undefined {
  switch (val) {
    case 'left': return AlignmentType.LEFT
    case 'center': return AlignmentType.CENTER
    case 'right': return AlignmentType.RIGHT
    case 'justify': return AlignmentType.JUSTIFIED
    default: return undefined
  }
}

// ═══════════════════════════════════════════════
// Inline content conversion
// ═══════════════════════════════════════════════

type InlineChild = TextRun | ImageRun | ExternalHyperlink

function convertInlineContent(
  nodes: TiptapNode[] | undefined,
  warnings: string[],
): InlineChild[] {
  if (!nodes) return []
  const result: InlineChild[] = []

  for (const node of nodes) {
    switch (node.type) {
      case 'text': {
        const linkMark = node.marks?.find(m => m.type === 'link')
        const otherMarks = node.marks?.filter(m => m.type !== 'link')
        const runOpts = marksToRunOptions(otherMarks)
        runOpts.text = node.text || ''

        if (linkMark) {
          result.push(new ExternalHyperlink({
            children: [new TextRun(runOpts as any)],
            link: String(linkMark.attrs?.href || '#'),
          }))
        } else {
          result.push(new TextRun(runOpts as any))
        }
        break
      }
      case 'image': {
        const imgRun = makeImageRun(node, warnings)
        if (imgRun) result.push(imgRun)
        break
      }
      case 'hardBreak':
        result.push(new TextRun({ break: 1 }))
        break
      case 'pageBreak':
        // 行内场景 pageBreak → 页分隔
        result.push(new TextRun({ children: [new PageBreak()] } as any))
        break
      case 'tocEntry': {
        // 内联形式的 tocEntry（极少见，走文字 + 制表位 + 页码布局）
        const text = String(node.attrs?.text || '')
        const page = String(node.attrs?.pageNumber || '')
        result.push(new TextRun({ text }))
        if (page) result.push(new TextRun({ text: `\t${page}` }))
        break
      }
      default:
        warnings.push(`Unknown inline node: ${node.type}`)
    }
  }
  return result
}

// ═══════════════════════════════════════════════
// Image
// ═══════════════════════════════════════════════

function detectImageType(src: string): 'jpg' | 'png' | 'gif' | 'bmp' {
  if (src.includes('image/png')) return 'png'
  if (src.includes('image/gif')) return 'gif'
  if (src.includes('image/bmp')) return 'bmp'
  return 'jpg'
}

function makeImageRun(node: TiptapNode, warnings: string[]): ImageRun | null {
  const src = String(node.attrs?.src || '')
  if (!src) { warnings.push('Image node has no src'); return null }

  const m = src.match(/^data:image\/[^;]+;base64,(.+)$/)
  if (!m) {
    warnings.push(`Non-base64 image skipped: ${src.slice(0, 60)}...`)
    return null
  }

  const data = Buffer.from(m[1], 'base64')
  const w = parsePixel(node.attrs?.width) || 200
  const h = parsePixel(node.attrs?.height) || 200
  const type = detectImageType(src)

  return new ImageRun({ type, data, transformation: { width: w, height: h } })
}

// ═══════════════════════════════════════════════
// Paragraph & Heading
// ═══════════════════════════════════════════════

let bookmarkIdSeq = 0
function nextBookmarkId(): number {
  return ++bookmarkIdSeq
}

function convertParagraph(
  node: TiptapNode,
  extraOpts: Record<string, unknown> | undefined,
  warnings: string[],
): Paragraph {
  const align = toAlignment(node.attrs?.textAlign)
  const inlineChildren = convertInlineContent(node.content, warnings)

  // bookmarks 支持：将 bookmarkStart/End 包裹 children
  const bookmarks = node.attrs?.bookmarks
  let children: unknown[] = inlineChildren
  if (Array.isArray(bookmarks) && bookmarks.length > 0) {
    const wrapped: unknown[] = []
    for (const name of bookmarks) {
      if (typeof name === 'string') {
        wrapped.push(new BookmarkStart(name, nextBookmarkId()))
      }
    }
    wrapped.push(...inlineChildren)
    for (const name of bookmarks) {
      if (typeof name === 'string') {
        wrapped.push(new BookmarkEnd(nextBookmarkId()))
      }
    }
    children = wrapped
  }

  const opts: Record<string, unknown> = { children }
  if (align) opts.alignment = align
  if (extraOpts) Object.assign(opts, extraOpts)

  return new Paragraph(opts as any)
}

const HEADING_MAP: Record<number, (typeof HeadingLevel)[keyof typeof HeadingLevel]> = {
  1: HeadingLevel.HEADING_1,
  2: HeadingLevel.HEADING_2,
  3: HeadingLevel.HEADING_3,
  4: HeadingLevel.HEADING_4,
  5: HeadingLevel.HEADING_5,
  6: HeadingLevel.HEADING_6,
}

function convertHeading(node: TiptapNode, warnings: string[]): Paragraph {
  const level = Number(node.attrs?.level || 1)
  return convertParagraph(
    node,
    { heading: HEADING_MAP[level] || HeadingLevel.HEADING_1 },
    warnings,
  )
}

// ═══════════════════════════════════════════════
// TocEntry
// ═══════════════════════════════════════════════

/**
 * tocEntry 节点导出
 * 方案：使用制表位对齐生成 "文字....页码" 样式
 * 如果有 rawXml 则优先还原（当前简化：总是重新生成）
 */
function convertTocEntry(node: TiptapNode): Paragraph {
  const text = String(node.attrs?.text || '')
  const page = String(node.attrs?.pageNumber || '')
  const level = Math.max(1, Math.min(9, Number(node.attrs?.level || 1)))
  const href = node.attrs?.href ? String(node.attrs.href) : undefined

  const indent = 400 * (level - 1)
  const runs: TextRun[] = []
  runs.push(new TextRun({ text }))
  if (page) {
    runs.push(new TextRun({ text: `\t${page}` }))
  }

  const children: unknown[] = runs
  if (href && href.startsWith('#')) {
    const anchor = href.slice(1)
    // 简化：忽略 internal hyperlink，正文直接输出文本
    void anchor
  }

  return new Paragraph({
    children,
    indent: indent ? { left: indent } : undefined,
    tabStops: [
      { type: TabStopType.RIGHT, position: TabStopPosition.MAX, leader: 'dot' as any },
    ],
  } as any)
}

// ═══════════════════════════════════════════════
// Table
// ═══════════════════════════════════════════════

function convertTable(node: TiptapNode, warnings: string[]): Table {
  const rows: TableRow[] = []

  for (const rowNode of node.content || []) {
    if (rowNode.type !== 'tableRow') continue

    const cells: TableCell[] = []
    for (const cellNode of rowNode.content || []) {
      if (cellNode.type !== 'tableCell' && cellNode.type !== 'tableHeader') continue

      const cellChildren = convertBlockNodes(cellNode.content || [], 0, warnings)
      if (cellChildren.length === 0) cellChildren.push(new Paragraph({}))

      const cellOpts: Record<string, unknown> = { children: cellChildren }

      const colspan = Number(cellNode.attrs?.colspan || 1)
      const rowspan = Number(cellNode.attrs?.rowspan || 1)
      if (colspan > 1) cellOpts.columnSpan = colspan
      if (rowspan > 1) cellOpts.rowSpan = rowspan

      if (cellNode.attrs?.backgroundColor) {
        cellOpts.shading = {
          type: ShadingType.SOLID, color: 'auto',
          fill: String(cellNode.attrs.backgroundColor).replace(/^#/, ''),
        }
      }

      if (cellNode.attrs?.verticalAlign) {
        const vaMap: Record<string, (typeof VerticalAlign)[keyof typeof VerticalAlign]> = {
          top: VerticalAlign.TOP, center: VerticalAlign.CENTER, bottom: VerticalAlign.BOTTOM,
        }
        const va = vaMap[String(cellNode.attrs.verticalAlign)]
        if (va) cellOpts.verticalAlign = va
      }

      if (cellNode.attrs?.colwidth) {
        const cw = Array.isArray(cellNode.attrs.colwidth)
          ? cellNode.attrs.colwidth[0]
          : cellNode.attrs.colwidth
        if (cw) cellOpts.width = { size: pxToTwips(Number(cw)), type: WidthType.DXA }
      }

      cells.push(new TableCell(cellOpts as any))
    }

    const rowOpts: Record<string, unknown> = { children: cells }
    if (rowNode.attrs?.height) {
      rowOpts.height = { value: pxToTwips(Number(rowNode.attrs.height)), rule: HeightRule.ATLEAST }
    }
    rows.push(new TableRow(rowOpts as any))
  }

  const tableOpts: Record<string, unknown> = { rows }
  if (node.attrs?.tableWidth) {
    const tw = String(node.attrs.tableWidth)
    if (tw.endsWith('%')) {
      tableOpts.width = { size: parseFloat(tw) * 50, type: WidthType.PERCENTAGE }
    } else {
      tableOpts.width = { size: pxToTwips(parseFloat(tw)), type: WidthType.DXA }
    }
  }

  return new Table(tableOpts as any)
}

// ═══════════════════════════════════════════════
// Lists
// ═══════════════════════════════════════════════

function convertList(
  node: TiptapNode,
  level: number,
  warnings: string[],
): Paragraph[] {
  const isBullet = node.type === 'bulletList'
  const reference = isBullet ? 'bullet-list' : 'ordered-list'
  const results: Paragraph[] = []

  for (const listItem of node.content || []) {
    if (listItem.type !== 'listItem') continue

    let isFirst = true
    for (const child of listItem.content || []) {
      if (child.type === 'paragraph') {
        const numOpts = isFirst ? { numbering: { reference, level } } : undefined
        results.push(convertParagraph(child, numOpts, warnings))
        isFirst = false
      } else if (child.type === 'bulletList' || child.type === 'orderedList') {
        results.push(...convertList(child, level + 1, warnings))
      } else {
        const blocks = convertBlockNodes([child], level, warnings)
        for (const b of blocks) {
          if (b instanceof Paragraph) results.push(b)
        }
      }
    }
  }

  return results
}

// ═══════════════════════════════════════════════
// Horizontal Rule
// ═══════════════════════════════════════════════

function convertHorizontalRule(node: TiptapNode): Paragraph {
  const color = String(node.attrs?.['data-line-color'] || '000000').replace(/^#/, '')
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color } },
    spacing: { after: 120 },
  })
}

// ═══════════════════════════════════════════════
// Blockquote
// ═══════════════════════════════════════════════

function convertBlockquote(
  node: TiptapNode,
  level: number,
  warnings: string[],
): (Paragraph | Table)[] {
  const results: (Paragraph | Table)[] = []
  for (const child of node.content || []) {
    if (child.type === 'paragraph') {
      results.push(convertParagraph(child, {
        indent: { left: convertMillimetersToTwip(12.7) },
        border: { left: { style: BorderStyle.SINGLE, size: 6, color: '999999' } },
      }, warnings))
    } else {
      results.push(...convertBlockNodes([child], level, warnings))
    }
  }
  return results
}

// ═══════════════════════════════════════════════
// Code Block
// ═══════════════════════════════════════════════

function convertCodeBlock(node: TiptapNode, _warnings: string[]): Paragraph {
  const text = (node.content || [])
    .filter(c => c.type === 'text')
    .map(c => c.text || '')
    .join('')

  return new Paragraph({
    children: [
      new TextRun({
        text,
        font: 'Courier New',
        size: 20,
        shading: { type: ShadingType.SOLID, color: 'auto', fill: 'f5f5f5' },
      } as any),
    ],
  })
}

// ═══════════════════════════════════════════════
// Block-level dispatch
// ═══════════════════════════════════════════════

function convertBlockNodes(
  nodes: TiptapNode[],
  listLevel: number,
  warnings: string[],
): (Paragraph | Table)[] {
  const result: (Paragraph | Table)[] = []

  for (const node of nodes) {
    switch (node.type) {
      case 'paragraph':
        result.push(convertParagraph(node, undefined, warnings))
        break
      case 'heading':
        result.push(convertHeading(node, warnings))
        break
      case 'table':
        result.push(convertTable(node, warnings))
        break
      case 'bulletList':
      case 'orderedList':
        result.push(...convertList(node, listLevel, warnings))
        break
      case 'pageBreak':
        result.push(new Paragraph({ children: [new PageBreak()] }))
        break
      case 'horizontalRule':
        result.push(convertHorizontalRule(node))
        break
      case 'blockquote':
        result.push(...convertBlockquote(node, listLevel, warnings))
        break
      case 'codeBlock':
        result.push(convertCodeBlock(node, warnings))
        break
      case 'tocEntry':
        result.push(convertTocEntry(node))
        break
      case 'image': {
        const imgRun = makeImageRun(node, warnings)
        if (imgRun) {
          const align = toAlignment(node.attrs?.align)
          result.push(new Paragraph({
            children: [imgRun],
            alignment: align || AlignmentType.CENTER,
          }))
        }
        break
      }
      default:
        warnings.push(`Unknown block node: ${node.type}`)
    }
  }

  return result
}

// ═══════════════════════════════════════════════
// Numbering definitions
// ═══════════════════════════════════════════════

const BULLET_CHARS = ['\u2022', '\u25CB', '\u25AA']

const ORDERED_LEVEL_DEFS: { format: string; textFn: (i: number) => string }[] = [
  { format: LevelFormat.DECIMAL, textFn: (i) => `%${i + 1}.` },
  { format: LevelFormat.LOWER_LETTER, textFn: (i) => `%${i + 1}.` },
  { format: LevelFormat.LOWER_ROMAN, textFn: (i) => `%${i + 1}.` },
  { format: 'chineseCountingThousand' as any, textFn: (i) => `%${i + 1}\u3001` },
  { format: LevelFormat.DECIMAL, textFn: (i) => `(%${i + 1})` },
  { format: 'chineseCountingThousand' as any, textFn: (i) => `(%${i + 1})` },
  { format: LevelFormat.DECIMAL, textFn: (i) => `%${i + 1}.` },
  { format: LevelFormat.LOWER_LETTER, textFn: (i) => `%${i + 1})` },
  { format: LevelFormat.LOWER_ROMAN, textFn: (i) => `%${i + 1})` },
]

function buildNumberingConfig() {
  return {
    config: [
      {
        reference: 'bullet-list',
        levels: Array.from({ length: 9 }, (_, i) => ({
          level: i,
          format: LevelFormat.BULLET,
          text: BULLET_CHARS[i % 3],
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720 * (i + 1), hanging: 360 } } },
        })),
      },
      {
        reference: 'ordered-list',
        levels: Array.from({ length: 9 }, (_, i) => {
          const def = ORDERED_LEVEL_DEFS[i % ORDERED_LEVEL_DEFS.length]
          return {
            level: i,
            format: def.format,
            text: def.textFn(i),
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720 * (i + 1), hanging: 360 } } },
          }
        }),
      },
    ],
  }
}

// ═══════════════════════════════════════════════
// Header / Footer helpers
// ═══════════════════════════════════════════════

/**
 * 将 header/footer 值转换为 Paragraph[]（支持新旧两种格式）
 * - 新：TiptapNode[]（富内容） → convertBlockNodes
 * - 旧：string（HTML 字符串） → 简单去标签兜底
 */
function headerFooterToChildren(
  val: TiptapNode[] | string | undefined,
  warnings: string[],
): Paragraph[] {
  if (!val) return []
  if (typeof val === 'string') {
    const text = val.replace(/<[^>]*>/g, '').trim()
    if (!text) return []
    return [new Paragraph({ children: [new TextRun({ text })] })]
  }
  const blocks = convertBlockNodes(val, 0, warnings)
  return blocks.filter((b): b is Paragraph => b instanceof Paragraph)
}

// ═══════════════════════════════════════════════
// Main entry point
// ═══════════════════════════════════════════════

export async function jsonToDocx(
  doc: TiptapDoc,
  metadata?: Partial<DocMetadata>,
): Promise<{ buffer: Buffer; warnings: string[] }> {
  bookmarkIdSeq = 0
  const warnings: string[] = []

  // 收集 footnotes/endnotes（从 TiptapDoc 本身未携带，仅来源于 metadata 附带）
  // 此处本版不对 footnotes 做完整还原，仅保留字段
  const footnotes: Record<number, { children: Paragraph[] }> = {}

  const children = convertBlockNodes(doc.content || [], 0, warnings)
  if (children.length === 0) children.push(new Paragraph({}))

  const sectionProps: Record<string, unknown> = {}
  if (metadata) {
    const page: Record<string, unknown> = {}
    if (metadata.paperSize) {
      page.size = {
        width: convertMillimetersToTwip(metadata.paperSize.width),
        height: convertMillimetersToTwip(metadata.paperSize.height),
      }
    }
    if (metadata.margins) {
      page.margin = {
        top: convertMillimetersToTwip(metadata.margins.top),
        bottom: convertMillimetersToTwip(metadata.margins.bottom),
        left: convertMillimetersToTwip(metadata.margins.left),
        right: convertMillimetersToTwip(metadata.margins.right),
      }
    }
    if (Object.keys(page).length > 0) sectionProps.page = page
  }

  const headers: Record<string, Header> = {}
  const footers: Record<string, Footer> = {}

  const h = metadata?.headers
  const f = metadata?.footers

  if (h?.default) {
    headers.default = new Header({ children: headerFooterToChildren(h.default, warnings) })
  }
  if (h?.first) {
    headers.first = new Header({ children: headerFooterToChildren(h.first, warnings) })
  }
  if (h?.even) {
    headers.even = new Header({ children: headerFooterToChildren(h.even, warnings) })
  }
  if (f?.default) {
    footers.default = new Footer({ children: headerFooterToChildren(f.default, warnings) })
  }
  if (f?.first) {
    footers.first = new Footer({ children: headerFooterToChildren(f.first, warnings) })
  }
  if (f?.even) {
    footers.even = new Footer({ children: headerFooterToChildren(f.even, warnings) })
  }

  // 多节属性：此版仅应用全局 section 的 page 设置 + 首节 headerRefs/footerRefs
  // 若后续要支持真正多节，需要在 convertBlockNodes 中插入 section 分隔
  if (metadata?.sections && metadata.sections.length > 0) {
    const first = metadata.sections[0]
    if (first.pageSetup) {
      const ps = first.pageSetup
      const page: Record<string, unknown> = (sectionProps.page as Record<string, unknown>) ?? {}
      if (ps.width && ps.height) {
        page.size = {
          width: convertMillimetersToTwip(ps.width),
          height: convertMillimetersToTwip(ps.height),
          orientation: ps.orientation,
        }
      }
      if (ps.margins) {
        page.margin = {
          top: ps.margins.top != null ? convertMillimetersToTwip(ps.margins.top) : undefined,
          bottom: ps.margins.bottom != null ? convertMillimetersToTwip(ps.margins.bottom) : undefined,
          left: ps.margins.left != null ? convertMillimetersToTwip(ps.margins.left) : undefined,
          right: ps.margins.right != null ? convertMillimetersToTwip(ps.margins.right) : undefined,
        }
      }
      sectionProps.page = page
    }
    if (first.titlePg) {
      sectionProps.titlePage = true
    }
  }

  const section: Record<string, unknown> = {
    properties: sectionProps,
    children,
  }
  if (Object.keys(headers).length > 0) section.headers = headers
  if (Object.keys(footers).length > 0) section.footers = footers

  const documentOpts: Record<string, unknown> = {
    numbering: buildNumberingConfig() as any,
    sections: [section as any],
  }
  if (Object.keys(footnotes).length > 0) {
    documentOpts.footnotes = footnotes
  }

  const document = new Document(documentOpts as any)

  const arrayBuffer = await Packer.toBuffer(document)
  return { buffer: Buffer.from(arrayBuffer), warnings }
}

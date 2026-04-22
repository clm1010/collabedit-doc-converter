import { XMLParser } from 'fast-xml-parser'

const alwaysArrayTagsSet = new Set<string>([
  // Core block elements
  'w:p', 'w:r', 'w:t', 'w:tbl', 'w:tr', 'w:tc',
  'w:hyperlink', 'w:bookmarkStart', 'w:bookmarkEnd',
  'w:drawing', 'w:pict', 'w:br',
  'w:gridCol',
  'w:style', 'w:abstractNum', 'w:num', 'w:lvl', 'w:lvlOverride',
  'Relationship',
  // P2: SDT / TOC
  'w:sdt', 'w:sdtContent', 'w:sdtPr',
  // P4: Footnotes / endnotes
  'w:footnote', 'w:endnote', 'w:footnoteReference', 'w:endnoteReference',
  // P6: Text boxes / AlternateContent
  'mc:Choice', 'mc:Fallback', 'wps:wsp', 'wps:txbx', 'w:txbxContent',
  // Group shapes / 绘图画布 / VML 旧版文本框（slogan、水印常用）
  'wpg:wgp', 'wpg:grpSp', 'wps:grpSp',
  'wpc:wpc',
  'v:shape', 'v:group', 'v:rect', 'v:oval', 'v:roundrect', 'v:textbox',
  // P7: Fields / special characters
  'w:fldSimple', 'w:fldChar', 'w:instrText', 'w:sym', 'w:tab',
])

const defaultParserOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  isArray: (tagName: string, _jpath: unknown, _isLeafNode: boolean, isAttribute: boolean) => {
    if (isAttribute) return false
    return alwaysArrayTagsSet.has(tagName)
  },
  preserveOrder: false,
  trimValues: true,
  parseTagValue: false,
  parseAttributeValue: false,
}

const orderedParserOptions = {
  ...defaultParserOptions,
  preserveOrder: true,
}

let _parser: XMLParser | null = null
let _orderedParser: XMLParser | null = null

function getParser(): XMLParser {
  if (!_parser) _parser = new XMLParser(defaultParserOptions)
  return _parser
}

function getOrderedParser(): XMLParser {
  if (!_orderedParser) _orderedParser = new XMLParser(orderedParserOptions)
  return _orderedParser
}

export function parseXml(xmlString: string): Record<string, unknown> {
  return getParser().parse(xmlString) as Record<string, unknown>
}

/**
 * 有序节点树结构
 * fast-xml-parser 的 preserveOrder 模式返回形如：
 *   [{ "w:p": [ { "w:r": [...] }, ... ], ":@": { "@_attr": "val" } }, ...]
 * 我们规范化为 { tag, attrs, children } 的形式便于遍历。
 */
export interface OrderedNode {
  tag: string // 标签名，例如 'w:p'；纯文本节点 tag === '#text'
  attrs: Record<string, string>
  children: OrderedNode[]
  text?: string // 仅 #text 节点使用
}

function toOrderedNodes(raw: unknown[]): OrderedNode[] {
  const result: OrderedNode[] = []
  if (!Array.isArray(raw)) return result
  for (const item of raw) {
    if (!item || typeof item !== 'object') continue
    const obj = item as Record<string, unknown>
    const attrs = (obj[':@'] as Record<string, unknown>) ?? {}
    // 找出标签键（非 ':@'）
    let tag: string | undefined
    let childArr: unknown[] | undefined
    for (const k of Object.keys(obj)) {
      if (k === ':@') continue
      tag = k
      const v = obj[k]
      if (Array.isArray(v)) childArr = v
      break
    }
    if (!tag) continue

    if (tag === '#text') {
      const textVal = (obj['#text'] as unknown)
      result.push({
        tag: '#text',
        attrs: {},
        children: [],
        text: typeof textVal === 'string' ? textVal : String(textVal ?? ''),
      })
      continue
    }

    const normalizedAttrs: Record<string, string> = {}
    for (const [ak, av] of Object.entries(attrs)) {
      normalizedAttrs[ak] = String(av)
    }

    result.push({
      tag,
      attrs: normalizedAttrs,
      children: childArr ? toOrderedNodes(childArr) : [],
    })
  }
  return result
}

export function parseOrdered(xml: string): OrderedNode[] {
  const raw = getOrderedParser().parse(xml) as unknown[]
  return toOrderedNodes(raw)
}

/** 按路径定位有序节点（路径为标签名数组，例如 ['w:document', 'w:body']） */
export function findOrderedByPath(
  roots: OrderedNode[],
  path: string[],
): OrderedNode | null {
  if (path.length === 0) return null
  let nodes = roots
  let found: OrderedNode | null = null
  for (const tag of path) {
    found = nodes.find(n => n.tag === tag) ?? null
    if (!found) return null
    nodes = found.children
  }
  return found
}

/**
 * 安全地获取嵌套属性值
 * 支持点分路径如 'w:body.w:p'
 */
export function getXmlVal(obj: unknown, path: string): unknown {
  if (obj == null) return undefined
  const parts = path.split('.')
  let current: unknown = obj
  for (const part of parts) {
    if (current == null || typeof current !== 'object') return undefined
    current = (current as Record<string, unknown>)[part]
  }
  return current
}

/** 确保值是数组 */
export function ensureArray<T>(val: T | T[] | undefined | null): T[] {
  if (val == null) return []
  return Array.isArray(val) ? val : [val]
}

/** 获取属性值 */
export function getAttr(node: unknown, attrName: string): string | undefined {
  if (node == null || typeof node !== 'object') return undefined
  const key = `@_${attrName}`
  return (node as Record<string, unknown>)[key] as string | undefined
}

/** 获取有序节点的属性值 */
export function getOrderedAttr(node: OrderedNode, attrName: string): string | undefined {
  return node.attrs[`@_${attrName}`]
}

/** 获取 w:val 属性（OOXML 最常见模式） */
export function getWVal(node: unknown): string | undefined {
  return getAttr(node, 'w:val')
}

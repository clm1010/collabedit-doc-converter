import { XMLParser } from 'fast-xml-parser'

const defaultParserOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  isArray: (tagName: string, _jpath: unknown, _isLeafNode: boolean, isAttribute: boolean) => {
    if (isAttribute) return false
    const alwaysArrayTags = new Set([
      'w:p', 'w:r', 'w:t', 'w:tbl', 'w:tr', 'w:tc',
      'w:hyperlink', 'w:bookmarkStart', 'w:bookmarkEnd',
      'w:drawing', 'w:pict', 'w:br',
      'w:gridCol',
      'w:style', 'w:abstractNum', 'w:num', 'w:lvl', 'w:lvlOverride',
      'Relationship',
    ])
    return alwaysArrayTags.has(tagName)
  },
  preserveOrder: false,
  trimValues: true,
  parseTagValue: false,
  parseAttributeValue: false,
}

let _parser: XMLParser | null = null

function getParser(): XMLParser {
  if (!_parser) {
    _parser = new XMLParser(defaultParserOptions)
  }
  return _parser
}

export function parseXml(xmlString: string): Record<string, unknown> {
  return getParser().parse(xmlString) as Record<string, unknown>
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

/** 获取 w:val 属性（OOXML 最常见模式） */
export function getWVal(node: unknown): string | undefined {
  return getAttr(node, 'w:val')
}

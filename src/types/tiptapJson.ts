export interface TiptapMark {
  type: string
  attrs?: Record<string, unknown>
}

export interface TiptapNode {
  type: string
  attrs?: Record<string, unknown>
  content?: TiptapNode[]
  marks?: TiptapMark[]
  text?: string
}

export interface TiptapDoc {
  type: 'doc'
  content: TiptapNode[]
}

/**
 * 脚注/尾注数据
 */
export interface FootnoteData {
  id: number
  noteType: string // 'normal' | 'separator' | 'continuationSeparator' | ...
  content: TiptapNode[]
}

// 命名注释：前端对应类型名为 ImportResult (see collabedit-fe/src/api/converter/index.ts)
export interface ImportResponse {
  data: { content: TiptapDoc }
  metadata: import('./docMetadata.js').DocMetadata
  logs: ImportLogs
  footnotes?: FootnoteData[]
  endnotes?: FootnoteData[]
}

export interface ImportLogs {
  info: string[]
  warn: string[]
  error: string[]
}

export function createDoc(content: TiptapNode[]): TiptapDoc {
  return { type: 'doc', content }
}

export function createNode(
  type: string,
  attrs?: Record<string, unknown>,
  content?: TiptapNode[],
  marks?: TiptapMark[],
  text?: string,
): TiptapNode {
  const node: TiptapNode = { type }
  if (attrs && Object.keys(attrs).length > 0) node.attrs = attrs
  if (content && content.length > 0) node.content = content
  if (marks && marks.length > 0) node.marks = marks
  if (text !== undefined) node.text = text
  return node
}

export function createTextNode(text: string, marks?: TiptapMark[]): TiptapNode {
  return createNode('text', undefined, undefined, marks, text)
}

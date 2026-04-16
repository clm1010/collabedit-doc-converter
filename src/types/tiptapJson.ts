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

export interface ImportResponse {
  data: { content: TiptapDoc }
  metadata: import('./docMetadata.js').DocMetadata
  logs: ImportLogs
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

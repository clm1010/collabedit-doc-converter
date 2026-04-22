import type { TiptapNode } from './tiptapJson.js'

export type HeaderFooterType = 'default' | 'first' | 'even'

/**
 * 页眉/页脚富内容 Map：type → Tiptap 节点数组
 * 兼容旧版字符串格式（旧数据库数据）仍通过 string 透传。
 */
export type HeaderFooterMap = Partial<Record<HeaderFooterType, TiptapNode[] | string>>

export interface SectionHeaderFooterRefs {
  default?: string
  first?: string
  even?: string
}

export interface SectionDefinition {
  pageSetup?: {
    width?: number
    height?: number
    margins?: { top?: number; bottom?: number; left?: number; right?: number }
    orientation?: 'portrait' | 'landscape'
  }
  /** 该节引用的 header/footer rId */
  headerRefs?: SectionHeaderFooterRefs
  footerRefs?: SectionHeaderFooterRefs
  /** 节类型：continuous / nextPage / oddPage / evenPage */
  type?: string
  /** 首页不同 */
  titlePg?: boolean
  /** （兼容）旧字段：直接存 HTML 字符串的 header/footer */
  headerFooter?: {
    header?: string
    footer?: string
  }
}

export interface DocMetadata {
  paperSize: { width: number; height: number }
  margins: { top: number; bottom: number; left: number; right: number }
  defaultFont: string
  defaultFontSize: number
  headers: HeaderFooterMap
  footers: HeaderFooterMap
  sections: SectionDefinition[]
  hasFootnotes: boolean
  hasEndnotes: boolean
  hasComments: boolean
  numberingDefinitions: object[]
  // TODO: implement style extraction in later phase
  customStyles: object[]
  isRedHead?: boolean
}

export interface ExportRequest {
  content: import('./tiptapJson.js').TiptapDoc
  metadata?: Partial<DocMetadata>
}

export interface HealthStatus {
  status: 'ok' | 'degraded'
  unoserver: boolean
}

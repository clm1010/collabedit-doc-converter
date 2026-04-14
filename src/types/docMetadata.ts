export interface DocMetadata {
  paperSize: { width: number; height: number }
  margins: { top: number; bottom: number; left: number; right: number }
  defaultFont: string
  defaultFontSize: number
  headers: { default?: string; first?: string; even?: string }
  footers: { default?: string; first?: string; even?: string }
  sections: Array<{
    pageSetup?: {
      width?: number
      height?: number
      margins?: { top?: number; bottom?: number; left?: number; right?: number }
      orientation?: 'portrait' | 'landscape'
    }
    headerFooter?: {
      header?: string
      footer?: string
    }
  }>
  hasFootnotes: boolean
  hasEndnotes: boolean
  numberingDefinitions: object[]
  customStyles: object[]
}

export interface ImportResult {
  html: string
  metadata: DocMetadata
}

export interface ExportRequest {
  html: string
  metadata?: Partial<DocMetadata>
  format: 'docx' | 'pdf'
}

export interface HealthStatus {
  status: 'ok' | 'degraded'
  unoserver: boolean
}

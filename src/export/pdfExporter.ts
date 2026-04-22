/**
 * HTML → PDF via Puppeteer
 *
 * 三个关键点缺一不可：
 *   1. emulateMediaType('print') 让前端 @media print CSS 生效
 *   2. waitUntil: 'load' 已由前端把图片内联成 data URL，不需要 networkidle0
 *   3. printBackground: true 让公文红头 / 表格底纹 / 彩色边框出现在 PDF 中
 */
import { acquirePage, releasePage } from '../services/pagePool.js'
import { env } from '../config/env.js'

export interface PdfMargin {
  top?: string
  bottom?: string
  left?: string
  right?: string
}

export interface PdfOptions {
  format?: 'A4' | 'A3' | 'Letter' | 'Legal' | 'Tabloid'
  margin?: PdfMargin
  displayHeaderFooter?: boolean
  headerTemplate?: string
  footerTemplate?: string
  landscape?: boolean
  printBackground?: boolean
}

const DEFAULT_MARGIN: Required<PdfMargin> = {
  top: '20mm',
  bottom: '20mm',
  left: '20mm',
  right: '20mm',
}

export async function htmlToPdf(
  html: string,
  options?: PdfOptions,
): Promise<{ buffer: Buffer; warnings: string[] }> {
  if (!html || typeof html !== 'string') {
    throw new Error('htmlToPdf: html must be a non-empty string')
  }

  const page = await acquirePage()
  try {
    await page.emulateMediaType('print')

    await page.setContent(html, {
      waitUntil: 'load',
      timeout: env.pdfRenderTimeout,
    })

    const pdf = await page.pdf({
      format: options?.format ?? 'A4',
      margin: {
        top: options?.margin?.top ?? DEFAULT_MARGIN.top,
        bottom: options?.margin?.bottom ?? DEFAULT_MARGIN.bottom,
        left: options?.margin?.left ?? DEFAULT_MARGIN.left,
        right: options?.margin?.right ?? DEFAULT_MARGIN.right,
      },
      printBackground: options?.printBackground ?? true,
      displayHeaderFooter: options?.displayHeaderFooter ?? false,
      headerTemplate: options?.headerTemplate,
      footerTemplate: options?.footerTemplate,
      landscape: options?.landscape ?? false,
      timeout: env.pdfRenderTimeout,
    })

    return {
      buffer: Buffer.from(pdf),
      warnings: [],
    }
  } finally {
    await releasePage(page)
  }
}

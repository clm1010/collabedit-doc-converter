import { Router, Request, Response, NextFunction } from 'express'
import multer from 'multer'
import { exportToDocx, exportToPdf } from '../services/exportService.js'

const router = Router()

// 选择性保存高保真方案：/export/docx 升级为同时支持 multipart/form-data 与 JSON：
//   multipart 字段：
//     - content      : JSON 字符串（必填，Tiptap Doc）
//     - metadata     : JSON 字符串（可选，Partial<DocMetadata>）
//     - originalDocx : File（可选，原始 DOCX 字节，触发选择性保存路径；
//                      未提供时等价 legacy JSON 请求）
//   JSON 请求：保留 legacy 行为，content / metadata 从 req.body 读取。
// upload 上限与导入一致：50MB。
// multer 对普通 field 的默认上限：fieldSize=1MB / fields=1000 / parts=1000。
// content 字段承载整份 Tiptap JSON，长文档容易突破 1MB，这里放宽到 64MB 并放宽 parts 上限，
// 避免触发 "Field value too long"（LIMIT_FIELD_VALUE）/ "Too many parts" 错误。
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024,
    fieldSize: 64 * 1024 * 1024,
    fieldNameSize: 1024,
    fields: 32,
    parts: 64,
  },
})

// 解析 content / metadata / originalDocx，兼容两种 Content-Type。
const parseExportBody = (
  req: Request
): {
  content: any
  metadata: any
  originalDocx: Buffer | null
  originalFileName: string | null
  mode: string | null
} => {
  const contentType = req.headers['content-type'] || ''
  const mode = (req.headers['x-export-mode'] as string) || null

  // multipart/form-data：字段值是字符串，需 JSON.parse；originalDocx 在 req.file
  if (contentType.includes('multipart/form-data')) {
    const body = (req.body || {}) as Record<string, string>
    const content = body.content ? JSON.parse(body.content) : null
    const metadata = body.metadata ? JSON.parse(body.metadata) : undefined
    const file = (req as any).file as Express.Multer.File | undefined
    return {
      content,
      metadata,
      originalDocx: file?.buffer ?? null,
      originalFileName: file?.originalname ?? null,
      mode,
    }
  }

  // legacy：JSON body
  const { content, metadata } = req.body || {}
  return {
    content,
    metadata,
    originalDocx: null,
    originalFileName: null,
    mode,
  }
}

// multer 的错误发生在流式解析阶段，无法被 route handler 的 try/catch 捕获，
// 因此用一个本路由作用域的中间件统一转成 JSON 400，避免 500 裸堆栈。
const handleMulterErrors = (
  err: unknown,
  _req: Request,
  res: Response,
  next: NextFunction,
) => {
  if (err instanceof multer.MulterError) {
    console.warn('[export/docx] multer rejected request:', err.code, err.field ?? '')
    res.status(413).json({
      error: 'multipart upload limit exceeded',
      code: err.code,
      field: err.field,
    })
    return
  }
  next(err)
}

router.post('/docx', upload.single('originalDocx'), handleMulterErrors, async (req: Request, res: Response) => {
  try {
    const parsed = parseExportBody(req)
    const { content, metadata, originalDocx, mode } = parsed

    if (!content || typeof content !== 'object') {
      res.status(400).json({ error: 'Missing or invalid field: content (must be a TiptapDoc JSON object)' })
      return
    }
    if (content.type !== 'doc' || !Array.isArray(content.content)) {
      res.status(400).json({ error: 'Invalid TiptapDoc: must have type="doc" and content array' })
      return
    }

    const start = Date.now()
    const requestedMode: 'selective' | 'legacy' | undefined =
      mode === 'selective' || mode === 'legacy' ? mode : undefined

    console.log(
      `[export/docx] Converting Tiptap JSON to DOCX (${content.content.length} nodes,` +
        ` requestedMode=${requestedMode ?? 'auto'},` +
        ` originalDocx=${originalDocx ? originalDocx.length + 'B' : 'none'})`,
    )

    const result = await exportToDocx(content, metadata, {
      originalDocx,
      mode: requestedMode,
    })
    const elapsed = Date.now() - start

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )
    res.setHeader('Content-Disposition', 'attachment; filename="export.docx"')
    res.setHeader('X-Export-Mode', result.mode)
    if (result.fallbackReason) {
      // 规范：仅"selective 失败回退"设 X-Export-Fallback: true；
      // 客户端主动选择 legacy 或缺 originalDocx 不算 fallback。
      if (result.telemetry.fallback) {
        res.setHeader('X-Export-Fallback', 'true')
      }
      res.setHeader('X-Export-Fallback-Reason', result.fallbackReason)
    }
    const t = result.telemetry
    res.setHeader(
      'X-Export-Stats',
      [
        `nodes=${t.totalNodes}`,
        `modified=${t.modifiedCount}`,
        `new=${t.newCount}`,
        `unchangedRatio=${t.unchangedRatio.toFixed(4)}`,
        `originalSize=${t.originalSize}`,
        `outputSize=${t.outputSize}`,
        `durationMs=${t.durationMs}`,
        `relsAppended=${t.relsAppended}`,
        `newMedia=${t.newMediaFiles}`,
      ].join(','),
    )
    res.send(result.buffer)
    // 结构化单行 JSON 日志，方便下游 ELK/Loki 直接采集
    console.log(
      '[export/docx.telemetry] ' +
        JSON.stringify({
          ...t,
          warnings: result.warnings.length,
          elapsedOuterMs: elapsed,
        }),
    )
  } catch (err: any) {
    console.error('[export/docx] Error:', err)
    res.status(500).json({ error: err.message || 'DOCX export failed' })
  }
})

// /export/pdf（Puppeteer 版）：接收完整 HTML + 可选 PdfOptions，返回 PDF 字节流
// body: { html: string, options?: PdfOptions }
router.post('/pdf', async (req, res) => {
  try {
    const { html, options } = req.body ?? {}

    if (typeof html !== 'string' || html.length === 0) {
      res.status(400).json({
        error: 'Missing or invalid field: html (must be a non-empty string)',
      })
      return
    }

    const start = Date.now()
    console.log(`[export/pdf] Rendering HTML to PDF (${html.length} bytes)`)
    const result = await exportToPdf(html, options)
    const elapsed = Date.now() - start

    res.setHeader('Content-Type', 'application/pdf')
    res.setHeader('Content-Disposition', 'attachment; filename="export.pdf"')
    res.send(result.buffer)
    console.log(
      `[export/pdf] Done in ${elapsed}ms. File size: ${result.buffer.length} bytes` +
        `${result.warnings.length ? ` (${result.warnings.length} warnings)` : ''}`,
    )
  } catch (err: any) {
    console.error('[export/pdf] Error:', err)
    res.status(500).json({ error: err.message || 'PDF export failed' })
  }
})

export default router

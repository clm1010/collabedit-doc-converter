import { Router } from 'express'
import multer from 'multer'
import { exportToDocx, exportToPdf } from '../services/exportService.js'
import { mergeExport } from '../services/mergeExportService.js'

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
})

const router = Router()

router.post('/', async (req, res) => {
  try {
    const { html, metadata, format } = req.body

    if (!html) {
      res.status(400).json({ error: 'Missing required field: html' })
      return
    }

    if (!format || !['docx', 'pdf'].includes(format)) {
      res.status(400).json({ error: 'Invalid format. Must be "docx" or "pdf".' })
      return
    }

    console.log(`[export] Converting HTML (${html.length} chars) to ${format}`)

    let result: { buffer: Buffer; warnings: string[] }

    if (format === 'pdf') {
      result = await exportToPdf(html, metadata)
    } else {
      result = await exportToDocx(html, metadata)
    }

    if (result.warnings.length > 0) {
      console.warn('[export] Warnings:', result.warnings)
    }

    const mimeType =
      format === 'pdf'
        ? 'application/pdf'
        : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

    const ext = format === 'pdf' ? 'pdf' : 'docx'

    res.setHeader('Content-Type', mimeType)
    res.setHeader('Content-Disposition', `attachment; filename="export.${ext}"`)

    if (result.warnings.length > 0) {
      res.setHeader('X-Export-Warnings', JSON.stringify(result.warnings))
    }

    res.send(result.buffer)
    console.log(`[export] Done. File size: ${result.buffer.length} bytes`)
  } catch (err: any) {
    console.error('[export] Error:', err)
    res.status(500).json({ error: err.message || 'Export conversion failed' })
  }
})

router.post('/merge', upload.single('originalFile'), async (req, res) => {
  try {
    if (!req.file) {
      res.status(400).json({ error: 'No original file uploaded.' })
      return
    }

    const { editedHtml, metadata } = req.body
    if (!editedHtml) {
      res.status(400).json({ error: 'Missing required field: editedHtml' })
      return
    }

    const parsedMetadata = metadata ? JSON.parse(metadata) : undefined
    const result = await mergeExport(req.file.buffer, editedHtml, parsedMetadata)

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
    res.setHeader('Content-Disposition', 'attachment; filename="merged.docx"')
    res.send(result)
  } catch (err: any) {
    console.error('[merge-export] Error:', err)
    res.status(500).json({ error: err.message || 'Merge export failed' })
  }
})

export default router

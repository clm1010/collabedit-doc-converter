import { Router } from 'express'
import { exportToDocx, exportToPdf } from '../services/exportService.js'

const router = Router()

router.post('/docx', async (req, res) => {
  try {
    const { content, metadata } = req.body

    if (!content || typeof content !== 'object') {
      res.status(400).json({ error: 'Missing or invalid field: content (must be a TiptapDoc JSON object)' })
      return
    }
    if (content.type !== 'doc' || !Array.isArray(content.content)) {
      res.status(400).json({ error: 'Invalid TiptapDoc: must have type="doc" and content array' })
      return
    }

    const start = Date.now()
    console.log(`[export/docx] Converting Tiptap JSON to DOCX (${content.content.length} nodes)`)
    const result = await exportToDocx(content, metadata)
    const elapsed = Date.now() - start

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )
    res.setHeader('Content-Disposition', 'attachment; filename="export.docx"')
    res.send(result.buffer)
    console.log(`[export/docx] Done in ${elapsed}ms. File size: ${result.buffer.length} bytes${result.warnings.length ? ` (${result.warnings.length} warnings)` : ''}`)
  } catch (err: any) {
    console.error('[export/docx] Error:', err)
    res.status(500).json({ error: err.message || 'DOCX export failed' })
  }
})

router.post('/pdf', async (req, res) => {
  try {
    const { content, metadata } = req.body

    if (!content || typeof content !== 'object') {
      res.status(400).json({ error: 'Missing or invalid field: content (must be a TiptapDoc JSON object)' })
      return
    }
    if (content.type !== 'doc' || !Array.isArray(content.content)) {
      res.status(400).json({ error: 'Invalid TiptapDoc: must have type="doc" and content array' })
      return
    }

    const start = Date.now()
    console.log(`[export/pdf] Converting Tiptap JSON to PDF (${content.content.length} nodes)`)
    const result = await exportToPdf(content, metadata)
    const elapsed = Date.now() - start

    res.setHeader('Content-Type', 'application/pdf')
    res.setHeader('Content-Disposition', 'attachment; filename="export.pdf"')
    res.send(result.buffer)
    console.log(`[export/pdf] Done in ${elapsed}ms. File size: ${result.buffer.length} bytes${result.warnings.length ? ` (${result.warnings.length} warnings)` : ''}`)
  } catch (err: any) {
    console.error('[export/pdf] Error:', err)
    res.status(500).json({ error: err.message || 'PDF export failed' })
  }
})

export default router

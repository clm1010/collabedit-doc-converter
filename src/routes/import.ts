import { Router } from 'express'
import multer from 'multer'
import { importDocx } from '../services/importService.js'

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
})

const router = Router()

router.post('/', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      res.status(400).json({ error: 'No file uploaded. Field name must be "file".' })
      return
    }

    const allowedMimes = [
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'application/octet-stream',
    ]
    if (!allowedMimes.includes(req.file.mimetype) && !req.file.originalname.endsWith('.docx')) {
      res.status(400).json({ error: 'Only DOCX files are supported.' })
      return
    }

    const start = Date.now()
    console.log(`[import] Processing ${req.file.originalname} (${req.file.size} bytes)`)
    const result = await importDocx(req.file.buffer)
    const elapsed = Date.now() - start
    console.log(`[import] Done in ${elapsed}ms. ${result.data.content.content.length} nodes, ${result.logs.warn.length} warnings`)

    res.json(result)
  } catch (err: any) {
    console.error('[import] Error:', err)
    res.status(500).json({ error: err.message || 'Import conversion failed' })
  }
})

export default router

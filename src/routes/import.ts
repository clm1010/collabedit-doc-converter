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

    console.log(`[import] Processing ${req.file.originalname} (${req.file.size} bytes)`)
    const result = await importDocx(req.file.buffer)
    console.log(`[import] Done. HTML length: ${result.html.length}`)

    res.json(result)
  } catch (err: any) {
    console.error('[import] Error:', err)
    res.status(500).json({ error: err.message || 'Import conversion failed' })
  }
})

export default router

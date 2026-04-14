import { Router } from 'express'
import { checkHealth } from '../services/unoClient.js'

const router = Router()

router.get('/', async (_req, res) => {
  const unoserverOk = await checkHealth()

  res.json({
    status: unoserverOk ? 'ok' : 'degraded',
    unoserver: unoserverOk,
    timestamp: new Date().toISOString(),
  })
})

export default router

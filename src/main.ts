import express from 'express'
import cors from 'cors'
import morgan from 'morgan'
import { env } from './config/env.js'
import { concurrencyMiddleware } from './middleware/concurrency.js'
import importRouter from './routes/import.js'
import exportRouter from './routes/export.js'
import healthRouter from './routes/health.js'

const app = express()

app.use(cors())
app.use(morgan('dev'))
app.use(express.json({ limit: env.maxFileSize }))

app.use('/import', concurrencyMiddleware, importRouter)
app.use('/export', concurrencyMiddleware, exportRouter)
app.use('/health', healthRouter)

app.get('/', (_req, res) => {
  res.json({ service: 'collabedit-doc-converter', version: '0.1.0' })
})

app.listen(env.port, () => {
  console.log(`[converter] Service started on port ${env.port}`)
  console.log(`[converter] unoserver URL: ${env.unoserverUrl}`)
  console.log(`[converter] Max concurrent: ${env.maxConcurrent}`)
})

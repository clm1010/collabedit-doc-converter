import express from 'express'
import cors from 'cors'
import morgan from 'morgan'
import { env } from './config/env.js'
import { concurrencyMiddleware } from './middleware/concurrency.js'
import importRouter from './routes/import.js'
import exportRouter from './routes/export.js'
import healthRouter from './routes/health.js'
import { shutdown as shutdownPagePool } from './services/pagePool.js'

const app = express()

app.use(cors())
app.use(morgan('dev'))
app.use(express.json({ limit: env.maxFileSize }))

app.use('/import', concurrencyMiddleware, importRouter)
app.use('/export', concurrencyMiddleware, exportRouter)
app.use('/health', healthRouter)

app.get('/', (_req, res) => {
  res.json({ service: 'collabedit-doc-converter', version: '0.3.0' })
})

const server = app.listen(env.port, () => {
  console.log(`[converter] Service started on port ${env.port}`)
  console.log(`[converter] Max concurrent: ${env.maxConcurrent}`)
  console.log(`[converter] Max body size: ${env.maxFileSize}`)
  console.log(
    `[converter] Puppeteer mode (Chromium at ${env.chromiumExecutablePath || 'bundled'})`,
  )
})

let isShuttingDown = false
async function gracefulShutdown(signal: string, exitCode = 0) {
  if (isShuttingDown) return
  isShuttingDown = true
  console.log(`[converter] Received ${signal}, shutting down...`)

  server.close((err) => {
    if (err) console.warn('[converter] HTTP server close error:', err)
  })

  try {
    await shutdownPagePool()
  } catch (err) {
    console.warn('[converter] Page pool shutdown error:', err)
  }

  setTimeout(() => process.exit(exitCode), 500).unref()
}

process.on('SIGTERM', () => {
  void gracefulShutdown('SIGTERM')
})
process.on('SIGINT', () => {
  void gracefulShutdown('SIGINT')
})
process.on('uncaughtException', (err) => {
  console.error('[converter] uncaughtException:', err)
  void gracefulShutdown('uncaughtException', 1)
})
process.on('unhandledRejection', (reason) => {
  console.error('[converter] unhandledRejection:', reason)
})

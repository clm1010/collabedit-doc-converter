import dotenv from 'dotenv'

dotenv.config()

export const env = {
  port: Number(process.env.PORT ?? 3002),
  unoserverUrl: process.env.UNOSERVER_URL ?? 'http://localhost:2004',
  maxFileSize: process.env.MAX_FILE_SIZE ?? '50mb',
  convertTimeout: Number(process.env.CONVERT_TIMEOUT ?? 60000),
  maxConcurrent: Number(process.env.MAX_CONCURRENT ?? 5),
  importTimeout: Number(process.env.IMPORT_TIMEOUT ?? 30000),
  exportTimeout: Number(process.env.EXPORT_TIMEOUT ?? 120000),
  logLevel: process.env.LOG_LEVEL ?? 'info',
}

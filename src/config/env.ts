import dotenv from 'dotenv'

dotenv.config()

export const env = {
  port: Number(process.env.PORT ?? 3002),
  maxFileSize: process.env.MAX_FILE_SIZE ?? '200mb',
  maxConcurrent: Number(process.env.MAX_CONCURRENT ?? 5),
  importTimeout: Number(process.env.IMPORT_TIMEOUT ?? 30000),
  exportTimeout: Number(process.env.EXPORT_TIMEOUT ?? 120000),
  logLevel: process.env.LOG_LEVEL ?? 'info',
  // Puppeteer 相关
  // CHROMIUM_EXECUTABLE_PATH：
  //   - 未设置 / 空值：使用 puppeteer 自带下载的 Chrome for Testing（Docker 场景的默认路径）。
  //   - 显式设置：用宿主机安装的浏览器（本地开发场景，例如 Windows 下的 Chrome.exe）。
  chromiumExecutablePath: process.env.CHROMIUM_EXECUTABLE_PATH ?? '',
  pdfRenderTimeout: Number(process.env.PDF_RENDER_TIMEOUT ?? 60000),
}

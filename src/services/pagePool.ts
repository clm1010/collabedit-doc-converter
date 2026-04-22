/**
 * Puppeteer Browser 单例 + Page 简单池
 *
 * 设计目标：
 * - Lazy launch：第一次请求到来才启动 Chromium，避免服务启动时抢占资源
 * - Browser 单例常驻：每个 PDF 请求 acquirePage() 取一个新 Page，用完 close()
 * - 崩溃保护：browser.on('disconnected') 标记状态，下次请求自动重新 launch
 * - 健康探测：isHealthy() 只返回内存 bool，不触发真实 launch，避免 launch 风暴
 * - 优雅关闭：shutdown() 在 SIGTERM/SIGINT/uncaughtException 时被调
 *
 * 并发上限依赖上层 concurrencyMiddleware（p-limit(MAX_CONCURRENT)）兜底，
 * 本模块只做单例与连通性管理，不重复实现并发控制。
 */
import puppeteer, {
  type Browser,
  type Page,
  type LaunchOptions,
} from 'puppeteer'
import { env } from '../config/env.js'

let browserPromise: Promise<Browser> | null = null
let currentBrowser: Browser | null = null

const LAUNCH_ARGS = [
  '--no-sandbox',
  '--disable-setuid-sandbox',
  '--disable-dev-shm-usage',
  '--disable-gpu',
  '--font-render-hinting=none',
  '--disable-features=IsolateOrigins,site-per-process',
]

async function launchBrowser(): Promise<Browser> {
  const options: LaunchOptions = {
    headless: true,
    args: LAUNCH_ARGS,
  }
  // 若显式指定了 executablePath 则使用（兼容宿主机 Chrome 开发模式）；
  // 否则 puppeteer 默认使用自带下载的 Chrome for Testing（Docker 镜像中的主要路径）。
  if (env.chromiumExecutablePath) {
    options.executablePath = env.chromiumExecutablePath
  }

  console.log(
    `[pagePool] Launching Chromium: executablePath=${options.executablePath ?? '(bundled)'}`,
  )
  const browser = await puppeteer.launch(options)

  browser.on('disconnected', () => {
    console.warn('[pagePool] Browser disconnected, will re-launch on next request')
    if (currentBrowser === browser) {
      currentBrowser = null
      browserPromise = null
    }
  })

  currentBrowser = browser
  return browser
}

async function getBrowser(): Promise<Browser> {
  if (currentBrowser && currentBrowser.connected) {
    return currentBrowser
  }
  if (!browserPromise) {
    browserPromise = launchBrowser().catch((err) => {
      browserPromise = null
      throw err
    })
  }
  return browserPromise
}

export async function acquirePage(): Promise<Page> {
  const browser = await getBrowser()
  const page = await browser.newPage()
  return page
}

export async function releasePage(page: Page): Promise<void> {
  try {
    if (!page.isClosed()) {
      await page.close({ runBeforeUnload: false })
    }
  } catch (err) {
    console.warn('[pagePool] releasePage error (ignored):', err)
  }
}

/**
 * 健康状态：仅返回内存中的 browser 连通 bool，不触发真实 launch。
 * 专供 /health 路由使用，避免频繁探活引发 Chromium launch 风暴。
 */
export function isHealthy(): boolean {
  return !!(currentBrowser && currentBrowser.connected)
}

export async function shutdown(): Promise<void> {
  const browser = currentBrowser
  currentBrowser = null
  browserPromise = null
  if (browser) {
    try {
      await browser.close()
      console.log('[pagePool] Browser closed')
    } catch (err) {
      console.warn('[pagePool] shutdown error (ignored):', err)
    }
  }
}

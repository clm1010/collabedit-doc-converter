import { Router } from 'express'
import { isHealthy } from '../services/pagePool.js'

const router = Router()

router.get('/', (_req, res) => {
  const chromiumOk = isHealthy()

  // 说明：
  //   - chromium: 来自 PagePool 的内存 bool（不触发真实 launch）；lazy launch 策略下，
  //     服务刚启动、尚未处理任何 PDF 请求时会是 false，这是正常的。
  //   - unoserver: 保留字段且恒为 false，兼容旧前端缓存 (_healthCache) 与可能的监控脚本。
  //   - status: 始终为 'ok'（Puppeteer 模式下不存在"外部依赖失败"的降级态）。
  res.json({
    status: 'ok',
    chromium: chromiumOk,
    unoserver: false,
    timestamp: new Date().toISOString(),
  })
})

export default router

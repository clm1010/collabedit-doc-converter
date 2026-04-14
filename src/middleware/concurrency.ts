import type { Request, Response, NextFunction } from 'express'
import pLimit from 'p-limit'
import { env } from '../config/env.js'

const limit = pLimit(env.maxConcurrent)

export function concurrencyMiddleware(req: Request, res: Response, next: NextFunction) {
  limit(
    () =>
      new Promise<void>((resolve) => {
        res.on('finish', resolve)
        res.on('close', resolve)
        next()
      })
  ).catch(next)
}

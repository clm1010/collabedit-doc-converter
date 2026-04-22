/**
 * 导入服务入口。
 *
 * 选择性保存高保真方案（phase 2）：默认走 `engine/importPipeline.ts`，
 * 它保留 DOCX 原始字节档案并为每个顶层块回填 __origRange/__origHash/__origPart。
 *
 * 降级策略：
 *   - 环境变量 `ENGINE_SELECTIVE_SAVE_ENABLED=false` 可强制走 legacy 实现（保留在
 *     `./importServiceLegacy.ts`）。
 *   - 新管线抛异常时自动 catch → 回退到 legacy，保证导入可用性不低于改造前。
 */

import type { Buffer as NodeBuffer } from 'node:buffer'
import { importDocxPipeline } from '../engine/importPipeline.js'
import { importDocxLegacy } from './importServiceLegacy.js'
import type { ImportResponse } from '../types/tiptapJson.js'

export async function importDocx(fileBuffer: NodeBuffer): Promise<ImportResponse> {
  const selectiveSaveEnabled = process.env.ENGINE_SELECTIVE_SAVE_ENABLED !== 'false'
  if (!selectiveSaveEnabled) {
    return importDocxLegacy(fileBuffer)
  }
  try {
    return await importDocxPipeline(fileBuffer)
  } catch (err) {
    console.warn(
      '[importService] selective-save pipeline failed, falling back to legacy:',
      err instanceof Error ? err.message : String(err),
    )
    return importDocxLegacy(fileBuffer)
  }
}

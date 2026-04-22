/**
 * 选择性保存高保真方案：DOCX 重打包（Rezipper）。
 *
 * 输入：原 RawDocxArchive + 一组 overrides（路径 → 新字节）。
 * 输出：新 DOCX 的 Uint8Array。
 *
 * 行为：
 *   - 未在 overrides 中出现的部件：原样字节回写（关键不变量：
 *     严格保留原字节，压缩级别/时间戳不会使 DOCX 变坏）。
 *   - 出现在 overrides 中的部件：使用新字节。
 *   - overrides 可以新增原 archive 中没有的路径（供阶段 4 追加 media 文件）。
 *   - overrides 中值为 null 的路径表示删除（阶段 4 可能用到）。
 *
 * 实现选择：
 *   - 使用 fflate.zipSync，开销可控；对较大 DOCX 也能在百毫秒级完成。
 *   - 压缩级别默认 6（fflate 默认）。Word 对压缩级别不敏感，任意合法 deflate 都接受。
 *   - 按 archive.parts 原顺序打包 → 保证 [Content_Types].xml 在首。
 *     对于 overrides 中新增路径，追加到原有序列尾部。
 *
 * 注意：
 *   - fflate.zipSync 的值格式是 `Uint8Array | [Uint8Array, ZipOptions]`；
 *     统一用前者最简。
 *   - ZIP 对目录项不做显式处理（DOCX 不包含独立目录项）。
 */

import { zipSync } from 'fflate'
import type { RawDocxArchive } from './zipExtractor.js'

export type PartOverrides = Map<string, Uint8Array | null>

export interface RezipOptions {
  /** overrides：路径 → 新字节；null 表示删除；未出现视为原样 */
  overrides?: PartOverrides
  /** 压缩级别 0~9；默认 6 */
  level?: number
}

export interface RezipResult {
  bytes: Uint8Array
  stats: {
    totalParts: number
    overriddenParts: number
    addedParts: number
    deletedParts: number
    unchangedParts: number
  }
}

export function rezipDocx(
  archive: RawDocxArchive,
  options?: RezipOptions,
): RezipResult {
  const overrides = options?.overrides ?? new Map<string, Uint8Array | null>()
  const level = options?.level ?? 6

  const zipInput: Record<string, Uint8Array> = {}
  const seen = new Set<string>()

  let overriddenParts = 0
  let deletedParts = 0
  let unchangedParts = 0

  // 1. 按原 parts 顺序处理
  for (const part of archive.parts) {
    seen.add(part.path)
    if (!overrides.has(part.path)) {
      zipInput[part.path] = part.bytes
      unchangedParts++
      continue
    }
    const override = overrides.get(part.path)
    if (override === null) {
      deletedParts++
      continue // 跳过，不打进 zip
    }
    zipInput[part.path] = override as Uint8Array
    overriddenParts++
  }

  // 2. 处理 overrides 中 archive 未有的新路径（phase 4: 新增 media/rels）
  let addedParts = 0
  for (const [path, bytes] of overrides) {
    if (seen.has(path)) continue
    if (bytes === null) continue
    zipInput[path] = bytes
    addedParts++
  }

  const bytes = zipSync(zipInput, { level: level as any })
  return {
    bytes,
    stats: {
      totalParts: Object.keys(zipInput).length,
      overriddenParts,
      addedParts,
      deletedParts,
      unchangedParts,
    },
  }
}

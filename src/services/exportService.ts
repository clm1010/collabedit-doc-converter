import type { TiptapDoc } from '../types/tiptapJson.js'
import type { DocMetadata } from '../types/docMetadata.js'
import { jsonToDocx } from '../export/jsonToDocx.js'
import { htmlToPdf, type PdfOptions } from '../export/pdfExporter.js'
import {
  exportDocxPipeline,
  type ExportPipelineResult,
} from '../engine/exportPipeline.js'

export interface ExportDocxResult {
  buffer: Buffer
  warnings: string[]
  /**
   * 导出模式：
   *   - 'selective' : 走选择性保存引擎（需要 originalDocx）
   *   - 'legacy'    : 走全量 jsonToDocx
   */
  mode: 'selective' | 'legacy'
  /** selective 模式下产出的命中率 / 字节统计，legacy 时为 undefined */
  stats?: ExportPipelineResult['stats']
  /** fallback 原因（selective 失败回退 legacy 时会填） */
  fallbackReason?: string
  /**
   * 结构化埋点（plan 第 7.17 节要求）。
   * 无论 selective / legacy / fallback 都会写入，便于统一采集。
   */
  telemetry: ExportTelemetry
}

export interface ExportTelemetry {
  mode: 'selective' | 'legacy'
  /** 是否是 selective 失败降级 */
  fallback: boolean
  fallbackReason?: string
  /** 原始 DOCX 字节数，legacy 或无 originalDocx 时为 0 */
  originalSize: number
  /** 最终输出字节数 */
  outputSize: number
  durationMs: number
  /** reuse 字节数 / 原 document.xml 字节数；仅 selective 有意义，否则 0 */
  unchangedRatio: number
  /** 被改动的顶层节点数；selective 下等价 classifier.regenerateNodes */
  modifiedCount: number
  /** 新增的顶层节点数（无 __origRange）；legacy / fallback 时不可区分，填 0 */
  newCount: number
  /** 节点总数；selective 下为 classifier.totalNodes，否则 content.content.length */
  totalNodes: number
  /** 追加 rels 条目数（阶段 4） */
  relsAppended: number
  /** 追加的 Content_Types Default 条目数（阶段 4） */
  contentTypesAppended: number
  /** 新增 media 文件数（阶段 4） */
  newMediaFiles: number
}

export interface ExportDocxOptions {
  /** 原始 DOCX 字节；无则不走 selective */
  originalDocx?: Buffer | Uint8Array | null
  /**
   * 请求端期望的导出模式：
   *   - 'selective' ：显式走选择性保存（缺 originalDocx 会回退 legacy 并打 fallback）
   *   - 'legacy'    : 强制跳过选择性保存
   *   - undefined   ：按环境开关 + originalDocx 存在性决定
   */
  mode?: 'selective' | 'legacy'
}

/**
 * 导出 DOCX 的统一入口。
 *
 * 决策逻辑：
 *   1. ENGINE_SELECTIVE_SAVE_ENABLED 环境变量（默认 enabled）；禁用则强制 legacy。
 *   2. options.mode === 'legacy' 则强制 legacy。
 *   3. 没有 originalDocx 则强制 legacy（加 fallbackReason）。
 *   4. 以上条件都不触发 → 走 selective；若 selective 抛错，回退 legacy 并记录原因。
 */
export async function exportToDocx(
  content: TiptapDoc,
  metadata?: Partial<DocMetadata>,
  options?: ExportDocxOptions,
): Promise<ExportDocxResult> {
  const overallStart = Date.now()
  const selectiveEnabled = process.env.ENGINE_SELECTIVE_SAVE_ENABLED !== 'false'
  const explicitLegacy = options?.mode === 'legacy'
  const hasOriginal =
    !!options?.originalDocx && options.originalDocx.byteLength > 0
  const originalSize = hasOriginal
    ? (options!.originalDocx as Buffer | Uint8Array).byteLength
    : 0

  const topLevelCount = content.content?.length ?? 0

  const buildLegacy = async (
    reason: string,
    isFallback: boolean,
  ): Promise<ExportDocxResult> => {
    const legacy = await jsonToDocx(content, metadata)
    const duration = Date.now() - overallStart
    return {
      buffer: legacy.buffer,
      warnings: isFallback
        ? [...legacy.warnings, `selective fallback: ${reason}`]
        : legacy.warnings,
      mode: 'legacy',
      fallbackReason: reason,
      telemetry: {
        mode: 'legacy',
        fallback: isFallback,
        fallbackReason: reason,
        originalSize,
        outputSize: legacy.buffer.length,
        durationMs: duration,
        unchangedRatio: 0,
        modifiedCount: topLevelCount,
        newCount: 0,
        totalNodes: topLevelCount,
        relsAppended: 0,
        contentTypesAppended: 0,
        newMediaFiles: 0,
      },
    }
  }

  if (!selectiveEnabled) return buildLegacy('ENGINE_SELECTIVE_SAVE_ENABLED=false', false)
  if (explicitLegacy) return buildLegacy('X-Export-Mode=legacy', false)
  if (!hasOriginal) return buildLegacy('no-original-docx', false)

  try {
    const result = await exportDocxPipeline({
      content,
      metadata,
      originalDocx: options!.originalDocx as Buffer | Uint8Array,
      imageMapperOptions: {
        cleanOrphanMedia: process.env.ENGINE_CLEAN_ORPHAN_MEDIA === 'true',
      },
    })
    const duration = Date.now() - overallStart
    const cs = result.stats.classifier
    const ps = result.stats.patcher
    const origDocXml = ps.originalDocumentXmlBytes || 1
    const unchangedRatio = ps.reusedBytes / origDocXml
    // newCount：统计顶层节点里没有 __origRange 的节点（前端新增块）
    let newCount = 0
    for (const n of content.content ?? []) {
      const r = (n.attrs as Record<string, unknown> | undefined)?.__origRange
      if (!Array.isArray(r)) newCount += 1
    }
    return {
      buffer: result.buffer,
      warnings: result.warnings,
      mode: 'selective',
      stats: result.stats,
      telemetry: {
        mode: 'selective',
        fallback: false,
        originalSize,
        outputSize: result.buffer.length,
        durationMs: duration,
        unchangedRatio,
        modifiedCount: cs.regenerateNodes,
        newCount,
        totalNodes: cs.totalNodes,
        relsAppended: result.stats.rels.relsAppended,
        contentTypesAppended: result.stats.contentTypes.defaultsAppended,
        newMediaFiles: result.stats.media.newFiles,
      },
    }
  } catch (err) {
    const reason = err instanceof Error ? err.message : String(err)
    console.warn('[exportService] selective-save failed, fallback to legacy:', reason)
    return buildLegacy(reason, true)
  }
}

export async function exportToPdf(
  html: string,
  options?: PdfOptions,
): Promise<{ buffer: Buffer; warnings: string[] }> {
  return htmlToPdf(html, options)
}

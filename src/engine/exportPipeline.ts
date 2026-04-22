/**
 * 选择性保存高保真方案：导出管线（Export Pipeline）总入口。
 *
 * 编排步骤：
 *   1. extractRawDocx(originalDocx)                ← 获得 RawDocxArchive（保留字节）
 *   2. indexTopLevelRanges(word/document.xml)      ← 顶层范围索引
 *   3. classifyNodes(tiptap.content)               ← 分类 segments
 *   4. serializeSegments(regenerate segments)      ← 本地序列化 regenerate 片段
 *      （含图片 rId 复用 / 新增 → ImageRefMapper）
 *   5. patchDocumentXml(...)                       ← 按 range 拼接新 document.xml
 *   6. patchRelsXml + patchContentTypesXml         ← 追加新 rels / Content_Types
 *   7. rezipDocx(archive, { overrides })           ← 重打包（含新 media 字节）
 *
 * 输出：{ buffer, stats, mode }，供路由层设置响应头 / 日志。
 *
 * 降级与 fallback：
 *   - 任何步骤失败，抛出错误；上层 `exportService` 捕获后回退到 `jsonToDocx` 全量路径。
 *   - 如 rangeIndex 不可用（XML 异常），抛错即可，触发 fallback。
 *   - 如 classifier 得到的 reuseNodes=0（没有节点能复用），仍然可以跑完流程
 *     （全部 regenerate），只是效果等价 legacy；上层可选择性直接走 legacy 避免浪费。
 *
 * 局限（由后续阶段补齐）：
 *   - SDT / TOC 原子处理、numbering 匹配 在阶段 5 加。
 */

import type { Buffer as NodeBuffer } from 'node:buffer'
import { extractRawDocx } from './zipExtractor.js'
import { indexTopLevelRanges } from './xmlRangeIndexer.js'
import { classifyNodes, type ClassifierStats } from './nodeClassifier.js'
import { serializeBlockNodesToXml } from './localSerializer.js'
import { patchDocumentXml } from './xmlPatcher.js'
import { rezipDocx, type PartOverrides } from './rezipper.js'
import {
  ImageRefMapper,
  relsPathFor,
  type ImageRefMapperOptions,
} from './imageRefMapper.js'
import { patchRelsXml } from './relsPatcher.js'
import { patchContentTypesXml } from './contentTypesPatcher.js'
import { NumberingMapper } from './numberingMapper.js'
import type { TiptapDoc } from '../types/tiptapJson.js'
import type { DocMetadata } from '../types/docMetadata.js'

const DOCUMENT_XML_PATH = 'word/document.xml'
const NUMBERING_XML_PATH = 'word/numbering.xml'
const CONTENT_TYPES_PATH = '[Content_Types].xml'

export interface ExportPipelineResult {
  buffer: Buffer
  warnings: string[]
  stats: {
    classifier: ClassifierStats
    patcher: {
      insertedFragments: number
      reusedBytes: number
      generatedBytes: number
      droppedBytes: number
      deletedRangeCount: number
      originalDocumentXmlBytes: number
      newDocumentXmlBytes: number
    }
    rels: {
      partsPatched: number
      relsAppended: number
    }
    contentTypes: {
      defaultsAppended: number
    }
    media: {
      newFiles: number
    }
    rezip: {
      totalParts: number
      overriddenParts: number
      addedParts: number
      deletedParts: number
      unchangedParts: number
    }
    elapsedMs: number
  }
}

export interface ExportPipelineInput {
  content: TiptapDoc
  originalDocx: Buffer | Uint8Array
  metadata?: Partial<DocMetadata>
  /** 阶段 4：图片 / 孤立 media 控制 */
  imageMapperOptions?: ImageRefMapperOptions
}

export async function exportDocxPipeline(
  input: ExportPipelineInput,
): Promise<ExportPipelineResult> {
  const start = Date.now()
  const warnings: string[] = []
  void input.metadata

  // Step 1: extract
  const archive = extractRawDocx(input.originalDocx)
  const documentBytes = archive.partsByPath.get(DOCUMENT_XML_PATH)
  if (!documentBytes) {
    throw new Error('exportPipeline: original DOCX has no word/document.xml')
  }

  // Step 2: index
  const rangeIndex = indexTopLevelRanges(documentBytes, 'w:body')
  if (!rangeIndex) {
    throw new Error(
      'exportPipeline: failed to index body of original document.xml',
    )
  }

  // Step 3: classify
  const topLevel = input.content.content ?? []
  const { segments, stats: classifierStats } = classifyNodes(topLevel, {
    partPath: DOCUMENT_XML_PATH,
  })

  // Step 4: serialize regenerate segments
  const mapper = new ImageRefMapper(archive, input.imageMapperOptions)
  const numberingMapper = new NumberingMapper(
    archive.partsByPath.get(NUMBERING_XML_PATH),
  )
  const regenSegments = segments.filter((s) => s.kind === 'regenerate')
  const regenFragments: Uint8Array[] = []
  for (const seg of regenSegments) {
    try {
      const xml = await serializeBlockNodesToXml(seg.nodes, {
        partPath: DOCUMENT_XML_PATH,
        imageRefMapper: mapper,
        numberingMapper,
      })
      regenFragments.push(xml)
    } catch (err) {
      throw new Error(
        `exportPipeline: localSerializer failed for segment with ${seg.nodes.length} node(s): ` +
          (err instanceof Error ? err.message : String(err)),
      )
    }
  }

  // Step 5: patch document.xml
  const patchResult = patchDocumentXml({
    originalBytes: documentBytes,
    rangeIndex,
    segments,
    regeneratedFragments: regenFragments,
  })

  // Step 6: patch rels (per part) + Content_Types
  const overrides: PartOverrides = new Map()
  overrides.set(DOCUMENT_XML_PATH, patchResult.bytes)

  let relsPartsPatched = 0
  let relsAppendedTotal = 0
  for (const [partPath, newRels] of mapper.getPendingRelsByPart()) {
    if (newRels.length === 0) continue
    const relsPath = relsPathFor(partPath)
    const originalRelsBytes = archive.partsByPath.get(relsPath)
    const out = patchRelsXml({ originalBytes: originalRelsBytes, newRels })
    overrides.set(relsPath, out.bytes)
    relsPartsPatched += 1
    relsAppendedTotal += out.appendedCount
  }

  let defaultsAppended = 0
  const pendingTypes = mapper.getPendingContentTypes()
  if (pendingTypes.length > 0) {
    const ctBytes = archive.partsByPath.get(CONTENT_TYPES_PATH)
    if (ctBytes) {
      const out = patchContentTypesXml({
        originalBytes: ctBytes,
        newDefaults: pendingTypes,
      })
      if (out.appendedCount > 0) {
        overrides.set(CONTENT_TYPES_PATH, out.bytes)
        defaultsAppended = out.appendedCount
      }
    } else {
      warnings.push(
        'exportPipeline: [Content_Types].xml not found in original archive; new image extensions may not be recognized by Word',
      )
    }
  }

  // 新增 media 文件
  for (const m of mapper.pendingMedia) {
    overrides.set(m.zipPath, m.bytes)
  }

  // Step 7: rezip
  const rezip = rezipDocx(archive, { overrides })

  const elapsedMs = Date.now() - start
  return {
    buffer: Buffer.from(rezip.bytes),
    warnings,
    stats: {
      classifier: classifierStats,
      patcher: {
        insertedFragments: patchResult.insertedFragments,
        reusedBytes: patchResult.reusedBytes,
        generatedBytes: patchResult.generatedBytes,
        droppedBytes: patchResult.droppedBytes,
        deletedRangeCount: patchResult.deletedRangeCount,
        originalDocumentXmlBytes: documentBytes.length,
        newDocumentXmlBytes: patchResult.bytes.length,
      },
      rels: {
        partsPatched: relsPartsPatched,
        relsAppended: relsAppendedTotal,
      },
      contentTypes: { defaultsAppended },
      media: { newFiles: mapper.pendingMedia.length },
      rezip: rezip.stats,
      elapsedMs,
    },
  }
}

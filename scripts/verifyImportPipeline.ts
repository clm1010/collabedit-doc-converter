/**
 * 阶段 1+2 验证脚本：对样本 DOCX 跑 importDocxPipeline，统计关键指标。
 *
 * 用法：
 *   npx tsx scripts/verifyImportPipeline.ts <docxPath> [<docxPath>...]
 *
 * 每个文件输出：
 *   - ZIP 部件数 / document.xml 字节数
 *   - 顶层范围索引数（按 tag 分组）
 *   - 顶层节点数 / 其中挂 __origRange 的个数 / __origHash 个数 / 不同 hash 个数
 *   - SDT 原子块数（sdtXmlMap 尺寸）
 *   - 图片节点数 / 其中挂 data-origin-rid 的个数
 *   - hash 自洽检查：取前 N 个挂了 origHash 的节点，用 hasher 重新算一次比对
 *   - 是否走了 legacy（importService 的 catch 分支不会被触发，这里直接调 pipeline）
 *   - 采样前 3 条 range 的字节片段预览（去换行）
 */

import { readFileSync } from 'node:fs'
import { basename } from 'node:path'
import { importDocxPipeline } from '../src/engine/importPipeline.js'
import { extractRawDocx } from '../src/engine/zipExtractor.js'
import { indexTopLevelRanges, sliceRange } from '../src/engine/xmlRangeIndexer.js'
import { hashXmlRange } from '../src/engine/hasher.js'
import { computeContentFingerprint } from '../src/engine/contentFingerprint.js'
import type { TiptapNode } from '../src/types/tiptapJson.js'

const PREVIEW_BYTES = 80
const HASH_SAMPLES = 5

function flatten(nodes: TiptapNode[]): TiptapNode[] {
  const out: TiptapNode[] = []
  const walk = (n: TiptapNode) => {
    out.push(n)
    if (n.content) for (const c of n.content) walk(c)
  }
  for (const n of nodes) walk(n)
  return out
}

function bytesPreview(bytes: Uint8Array, max: number): string {
  const slice = bytes.subarray(0, max)
  return new TextDecoder('utf-8', { fatal: false })
    .decode(slice)
    .replace(/\s+/g, ' ')
    .trim()
}

async function verifyOne(docxPath: string): Promise<boolean> {
  const label = basename(docxPath)
  console.log('\n' + '='.repeat(80))
  console.log(`[${label}]`)
  console.log('='.repeat(80))

  const buffer = readFileSync(docxPath)
  console.log(`File size: ${buffer.length} bytes`)

  // 先直接验证 extractRawDocx + indexTopLevelRanges 这一层
  const raw = extractRawDocx(buffer)
  console.log(`ZIP parts: ${raw.parts.length}`)

  const documentBytes = raw.partsByPath.get('word/document.xml')
  if (!documentBytes) {
    console.log('  [FAIL] word/document.xml not found')
    return false
  }
  console.log(`document.xml bytes: ${documentBytes.length}`)

  const rangeIndex = indexTopLevelRanges(documentBytes, 'w:body')
  if (!rangeIndex) {
    console.log('  [FAIL] range index not produced')
    return false
  }
  const tagHist: Record<string, number> = {}
  for (const r of rangeIndex.ranges) {
    tagHist[r.tag] = (tagHist[r.tag] ?? 0) + 1
  }
  console.log(
    `Top-level ranges: ${rangeIndex.ranges.length} ` +
      Object.entries(tagHist)
        .map(([k, v]) => `${k}=${v}`)
        .join(', '),
  )
  console.log(
    `Container content: [${rangeIndex.containerContentStart}, ${rangeIndex.containerContentEnd})`,
  )

  // 跑完整 pipeline
  let result
  let pipelineErr: unknown = null
  try {
    result = await importDocxPipeline(buffer)
  } catch (err) {
    pipelineErr = err
  }
  if (pipelineErr) {
    console.log(
      `  [FAIL] pipeline threw: ${pipelineErr instanceof Error ? pipelineErr.message : String(pipelineErr)}`,
    )
    if (pipelineErr instanceof Error && pipelineErr.stack) {
      console.log(pipelineErr.stack.split('\n').slice(0, 6).join('\n'))
    }
    return false
  }

  const topLevel = result!.data.content.content
  const all = flatten(topLevel)

  const withRange = topLevel.filter(
    (n) => Array.isArray((n.attrs as any)?.__origRange),
  )
  const withHash = topLevel.filter(
    (n) => typeof (n.attrs as any)?.__origHash === 'string',
  )
  const withPart = topLevel.filter(
    (n) => typeof (n.attrs as any)?.__origPart === 'string',
  )
  const uniqueHashes = new Set(
    withHash.map((n) => (n.attrs as any).__origHash as string),
  )
  const withFp = topLevel.filter(
    (n) => typeof (n.attrs as any)?.__origContentFp === 'string',
  )
  // contentFp 自洽：对前 N 个节点重新算一次 fp，应当等于 attrs 里存的
  let fpOk = 0
  let fpBad = 0
  for (const n of withFp.slice(0, HASH_SAMPLES)) {
    const stored = (n.attrs as any).__origContentFp as string
    const recomputed = computeContentFingerprint(n)
    if (recomputed === stored) fpOk++
    else fpBad++
  }

  const imageNodes = all.filter((n) => n.type === 'image')
  const imagesWithRid = imageNodes.filter(
    (n) => typeof (n.attrs as any)?.['data-origin-rid'] === 'string',
  )
  const imagesWithTarget = imageNodes.filter(
    (n) => typeof (n.attrs as any)?.['data-origin-target'] === 'string',
  )
  const imagesWithPart = imageNodes.filter(
    (n) => typeof (n.attrs as any)?.['data-origin-part'] === 'string',
  )

  const tocEntries = all.filter((n) => n.type === 'tocEntry')
  const tocWithSdtId = tocEntries.filter(
    (n) => typeof (n.attrs as any)?.__origSdtId === 'string',
  )
  const tocWithSdtXml = tocEntries.filter(
    (n) => typeof (n.attrs as any)?.__origSdtXml === 'string',
  )

  console.log(
    `Top-level nodes: ${topLevel.length}` +
      ` | origRange: ${withRange.length}` +
      ` | origHash: ${withHash.length}` +
      ` | origPart: ${withPart.length}` +
      ` | uniqueHash: ${uniqueHashes.size}` +
      ` | origContentFp: ${withFp.length}`,
  )
  console.log(
    `ContentFp self-check (first ${Math.min(HASH_SAMPLES, withFp.length)}): ok=${fpOk}, mismatch=${fpBad}`,
  )
  console.log(
    `Images: ${imageNodes.length}` +
      ` | with rid: ${imagesWithRid.length}` +
      ` | with target: ${imagesWithTarget.length}` +
      ` | with part: ${imagesWithPart.length}`,
  )
  if (tocEntries.length > 0) {
    console.log(
      `TocEntries: ${tocEntries.length}` +
        ` | with sdtId: ${tocWithSdtId.length}` +
        ` | with sdtXml (leader): ${tocWithSdtXml.length}`,
    )
  }

  // 覆盖率检查：范围索引里有 N 个 w:p + w:tbl，顶层节点里应当大致等于 origRange 数
  // （pageBreak 节点无 range、bulletList 包装会吞掉一层 paragraph，所以允许偏差）
  const rangeCountTracked =
    (tagHist['w:p'] ?? 0) + (tagHist['w:tbl'] ?? 0) + (tagHist['w:sdt'] ?? 0)
  if (withRange.length === 0 && rangeCountTracked > 0) {
    console.log(
      `  [WARN] ranges=${rangeCountTracked} but 0 top-level nodes got __origRange — orchestrator misalignment?`,
    )
  }

  // hash 自洽：随机抽前 N 个挂 range 的节点，重新计算 hash 对齐
  const samples = withRange.slice(0, HASH_SAMPLES)
  let hashOk = 0
  let hashBad = 0
  for (const n of samples) {
    const [start, end] = (n.attrs as any).__origRange as [number, number]
    const storedHash = (n.attrs as any).__origHash as string
    const recomputed = hashXmlRange(documentBytes.subarray(start, end))
    if (recomputed === storedHash) hashOk++
    else hashBad++
  }
  console.log(
    `Hash self-check (first ${samples.length}): ok=${hashOk}, mismatch=${hashBad}`,
  )

  // 前 3 条 range 字节预览
  console.log('First 3 ranges preview:')
  for (let i = 0; i < Math.min(3, rangeIndex.ranges.length); i++) {
    const r = rangeIndex.ranges[i]
    const slice = sliceRange(documentBytes, r)
    console.log(
      `  [${i}] ${r.tag} [${r.start}, ${r.end}) len=${slice.length}  ${bytesPreview(
        slice,
        PREVIEW_BYTES,
      )}`,
    )
  }

  // 日志摘要
  console.log(`Logs: info=${result!.logs.info.length} warn=${result!.logs.warn.length} error=${result!.logs.error.length}`)
  if (result!.logs.warn.length > 0) {
    console.log('  Warnings (first 3):')
    for (const w of result!.logs.warn.slice(0, 3)) console.log(`    - ${w}`)
  }
  if (result!.logs.error.length > 0) {
    console.log('  Errors:')
    for (const e of result!.logs.error) console.log(`    - ${e}`)
  }

  // 成功标准
  const passed =
    withRange.length > 0 &&
    withHash.length === withRange.length &&
    hashBad === 0 &&
    result!.logs.error.length === 0
  console.log(passed ? '  [PASS]' : '  [FAIL]')
  return passed
}

async function main() {
  const files = process.argv.slice(2)
  if (files.length === 0) {
    console.error('用法: npx tsx scripts/verifyImportPipeline.ts <docx> [<docx>...]')
    process.exit(2)
  }
  let passCount = 0
  let failCount = 0
  for (const f of files) {
    try {
      const ok = await verifyOne(f)
      if (ok) passCount++
      else failCount++
    } catch (err) {
      failCount++
      console.error(
        `\n[${basename(f)}] unhandled exception: ${err instanceof Error ? err.message : String(err)}`,
      )
      if (err instanceof Error && err.stack) console.error(err.stack)
    }
  }
  console.log('\n' + '='.repeat(80))
  console.log(`SUMMARY: pass=${passCount} fail=${failCount}`)
  console.log('='.repeat(80))
  process.exit(failCount > 0 ? 1 : 0)
}

main()

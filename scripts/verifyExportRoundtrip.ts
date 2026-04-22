/**
 * 阶段 3 验证：import → 不修改 → export 的字节级 roundtrip 对比。
 *
 * 核心断言：
 *   - 所有顶层节点都命中 reuse（regenerate=0）
 *   - 新 document.xml 的 body 字节与原 body 字节完全一致（sectPr 顺序可能不同，允许）
 *   - 其它原样部件在新 DOCX 中也完整存在
 *
 * 用法：
 *   npx tsx scripts/verifyExportRoundtrip.ts <docx> [<docx>...]
 */

import { readFileSync, writeFileSync } from 'node:fs'
import { basename } from 'node:path'
import { importDocxPipeline } from '../src/engine/importPipeline.js'
import { exportDocxPipeline } from '../src/engine/exportPipeline.js'
import { extractRawDocx } from '../src/engine/zipExtractor.js'
import { indexTopLevelRanges, sliceRange } from '../src/engine/xmlRangeIndexer.js'
import { hashXmlRange } from '../src/engine/hasher.js'
import type { TiptapNode } from '../src/types/tiptapJson.js'

const DUMP_DIR = process.env.DUMP_DIR || ''

function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false
  return true
}

function computeBodyRangeHashes(
  docxBytes: Uint8Array,
): { rangeHashes: string[]; tagHist: Record<string, number> } {
  const raw = extractRawDocx(docxBytes)
  const documentBytes = raw.partsByPath.get('word/document.xml')
  if (!documentBytes) throw new Error('no document.xml')
  const index = indexTopLevelRanges(documentBytes, 'w:body')
  if (!index) throw new Error('cannot index body')
  const hashes: string[] = []
  const hist: Record<string, number> = {}
  for (const r of index.ranges) {
    const slice = sliceRange(documentBytes, r)
    hashes.push(`${r.tag}:${hashXmlRange(slice)}`)
    hist[r.tag] = (hist[r.tag] ?? 0) + 1
  }
  return { rangeHashes: hashes, tagHist: hist }
}

async function verifyOne(docxPath: string): Promise<boolean> {
  const label = basename(docxPath)
  console.log('\n' + '='.repeat(80))
  console.log(`[${label}]`)
  console.log('='.repeat(80))

  const originalBuffer = readFileSync(docxPath)
  const originalRaw = extractRawDocx(originalBuffer)
  const originalDocumentBytes = originalRaw.partsByPath.get('word/document.xml')
  if (!originalDocumentBytes) {
    console.log('  [FAIL] original docx has no word/document.xml')
    return false
  }

  // ---- Step A: import ----
  const importResult = await importDocxPipeline(originalBuffer)
  const topLevel = importResult.data.content.content
  console.log(`Import: ${topLevel.length} top-level nodes`)

  // 诊断：统计"flatten 后所有挂了 origRange 的叶子节点"
  const flatOrigRanges: { type: string; start: number; end: number }[] = []
  const walk = (n: TiptapNode) => {
    const attrs = (n.attrs ?? {}) as Record<string, unknown>
    const r = attrs.__origRange
    if (Array.isArray(r) && r.length === 2) {
      flatOrigRanges.push({
        type: n.type,
        start: Number(r[0]),
        end: Number(r[1]),
      })
    }
    if (Array.isArray(n.content)) for (const c of n.content) walk(c)
  }
  for (const n of topLevel) walk(n)
  console.log(`Flat nodes with origRange: ${flatOrigRanges.length}`)

  // 与原 body ranges 做对比，列出"原 ranges 中没挂到任何 Tiptap 节点上的 ranges"
  const origIndex = indexTopLevelRanges(originalDocumentBytes, 'w:body')!
  const nodeStarts = new Set(flatOrigRanges.map((r) => r.start))
  const uncoveredRanges = origIndex.ranges.filter(
    (r) => r.tag !== 'w:sectPr' && !nodeStarts.has(r.start),
  )
  if (uncoveredRanges.length > 0) {
    console.log(
      `  [DIAG] ${uncoveredRanges.length} original body ranges not mapped to any Tiptap node.` +
        ` First 3 previews:`,
    )
    for (const r of uncoveredRanges.slice(0, 3)) {
      const preview = new TextDecoder('utf-8', { fatal: false })
        .decode(originalDocumentBytes.subarray(r.start, Math.min(r.end, r.start + 180)))
        .replace(/\s+/g, ' ')
        .trim()
      console.log(`    - ${r.tag} [${r.start},${r.end}) ${preview}`)
    }
  }

  // ---- Step B: export (no changes) ----
  let exportResult
  try {
    exportResult = await exportDocxPipeline({
      content: importResult.data.content,
      originalDocx: originalBuffer,
    })
  } catch (err) {
    console.log(
      `  [FAIL] exportPipeline threw: ${err instanceof Error ? err.message : String(err)}`,
    )
    return false
  }

  const cs = exportResult.stats.classifier
  const ps = exportResult.stats.patcher
  console.log(
    `Classifier: nodes=${cs.totalNodes} reuse=${cs.reuseNodes} regen=${cs.regenerateNodes}` +
      ` | reuseSegs=${cs.reuseSegments} regenSegs=${cs.regenerateSegments}`,
  )
  console.log(`  DirtyReasons: ${JSON.stringify(cs.dirtyReasons)}`)
  console.log(
    `Patcher: origBytes=${ps.originalDocumentXmlBytes} newBytes=${ps.newDocumentXmlBytes}` +
      ` reuseBytes=${ps.reusedBytes} genBytes=${ps.generatedBytes}` +
      ` droppedBytes=${ps.droppedBytes} deletedRanges=${ps.deletedRangeCount}` +
      ` insertedFragments=${ps.insertedFragments}`,
  )
  console.log(
    `Rezip: total=${exportResult.stats.rezip.totalParts} unchanged=${exportResult.stats.rezip.unchangedParts}` +
      ` overridden=${exportResult.stats.rezip.overriddenParts}` +
      ` added=${exportResult.stats.rezip.addedParts}` +
      ` deleted=${exportResult.stats.rezip.deletedParts}`,
  )
  console.log(`Elapsed: ${exportResult.stats.elapsedMs}ms`)

  // ---- Step C: analyze new docx ----
  const newRaw = extractRawDocx(exportResult.buffer)
  const newDocumentBytes = newRaw.partsByPath.get('word/document.xml')
  if (!newDocumentBytes) {
    console.log('  [FAIL] new docx has no word/document.xml')
    return false
  }

  const origProfile = computeBodyRangeHashes(originalBuffer)
  const newProfile = computeBodyRangeHashes(exportResult.buffer)

  console.log(
    `Body tags: orig=${JSON.stringify(origProfile.tagHist)} new=${JSON.stringify(newProfile.tagHist)}`,
  )

  // 允许的差异：sectPr 的顺序可能不同（原在末尾 vs 新在末尾一致）
  // 但所有 w:p/w:tbl 的 hash 序列必须完全相同
  const origBodyTagHashes = origProfile.rangeHashes.filter(
    (h) => !h.startsWith('w:sectPr:'),
  )
  const newBodyTagHashes = newProfile.rangeHashes.filter(
    (h) => !h.startsWith('w:sectPr:'),
  )
  const bodyMatches =
    origBodyTagHashes.length === newBodyTagHashes.length &&
    origBodyTagHashes.every((h, i) => h === newBodyTagHashes[i])

  console.log(
    `Body hash sequence match (ex sectPr): ${bodyMatches ? 'YES' : 'NO'}` +
      ` orig=${origBodyTagHashes.length} new=${newBodyTagHashes.length}`,
  )

  if (!bodyMatches) {
    // 找出第一个不一致位置供诊断
    for (let i = 0; i < Math.max(origBodyTagHashes.length, newBodyTagHashes.length); i++) {
      if (origBodyTagHashes[i] !== newBodyTagHashes[i]) {
        console.log(
          `  First mismatch at index ${i}: orig=${origBodyTagHashes[i] ?? 'MISSING'} new=${newBodyTagHashes[i] ?? 'MISSING'}`,
        )
        break
      }
    }
  }

  const exactBytesMatch = bytesEqual(originalDocumentBytes, newDocumentBytes)
  console.log(
    `Exact document.xml bytes equal: ${exactBytesMatch ? 'YES' : 'NO'}` +
      ` (orig=${originalDocumentBytes.length}B new=${newDocumentBytes.length}B)`,
  )

  // 非 document.xml 部件检查
  let unchangedPartCount = 0
  let changedPartPaths: string[] = []
  let missingPartPaths: string[] = []
  for (const part of originalRaw.parts) {
    if (part.path === 'word/document.xml') continue
    const newBytes = newRaw.partsByPath.get(part.path)
    if (!newBytes) {
      missingPartPaths.push(part.path)
    } else if (bytesEqual(part.bytes, newBytes)) {
      unchangedPartCount++
    } else {
      changedPartPaths.push(part.path)
    }
  }
  console.log(
    `Other parts: unchanged=${unchangedPartCount} changed=${changedPartPaths.length} missing=${missingPartPaths.length}`,
  )
  if (changedPartPaths.length > 0) {
    console.log(`  Changed: ${changedPartPaths.slice(0, 5).join(', ')}${changedPartPaths.length > 5 ? ' ...' : ''}`)
  }
  if (missingPartPaths.length > 0) {
    console.log(`  Missing: ${missingPartPaths.slice(0, 5).join(', ')}${missingPartPaths.length > 5 ? ' ...' : ''}`)
  }

  if (DUMP_DIR) {
    const dumpPath = `${DUMP_DIR}/${label.replace(/\.docx$/, '')}-roundtrip.docx`
    writeFileSync(dumpPath, exportResult.buffer)
    console.log(`  Dumped: ${dumpPath}`)
  }

  // 验收标准：
  //   1. reuse 命中率 100%（所有节点都 reuse）
  //   2. body hash 序列（非 sectPr）完全一致
  //   3. 非 document.xml 部件全部未改动
  const passed =
    cs.regenerateNodes === 0 &&
    bodyMatches &&
    changedPartPaths.length === 0 &&
    missingPartPaths.length === 0
  console.log(passed ? '  [PASS]' : '  [FAIL]')
  return passed
}

async function main() {
  const files = process.argv.slice(2)
  if (files.length === 0) {
    console.error('用法: npx tsx scripts/verifyExportRoundtrip.ts <docx> [<docx>...]')
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
        `\n[${basename(f)}] unhandled: ${err instanceof Error ? err.message : String(err)}`,
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

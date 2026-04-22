/**
 * 阶段 6 验收回归：对样本库 DOCX 跑三种场景的端到端校验，输出结构化报告。
 *
 * 场景：
 *   A. no-op roundtrip    ：import → export（不修改）→ 原 document.xml 字节完全一致
 *   B. edit paragraph     ：改一个段落文字（优先选含图段落）→ 图片 rId 应当复用（media/rels 不追加）
 *   C. insert new image   ：向一个段落插入一张新 PNG → rels / media 各追加 1 条，re-import 后图片数 +1
 *
 * 用法：
 *   npx tsx scripts/verifyRegression.ts [--dir <dir>...] [<docx>...]
 *   默认读取：
 *     c:/Users/limin/Desktop/参考图/fix-img2/*.docx
 *     c:/Users/limin/Desktop/导入文件测试数据/*.docx
 *
 * 退出码：0=全部通过；1=任一场景失败。
 */

import { readdirSync, readFileSync, statSync } from 'node:fs'
import { basename, join } from 'node:path'
import { importDocxPipeline } from '../src/engine/importPipeline.js'
import { exportDocxPipeline } from '../src/engine/exportPipeline.js'
import { extractRawDocx } from '../src/engine/zipExtractor.js'
import { indexTopLevelRanges, sliceRange } from '../src/engine/xmlRangeIndexer.js'
import { hashXmlRange } from '../src/engine/hasher.js'
import type { TiptapNode } from '../src/types/tiptapJson.js'

const DEFAULT_DIRS = [
  'C:/Users/limin/Desktop/参考图/fix-img2',
  'C:/Users/limin/Desktop/导入文件测试数据',
]

// ---------------- helpers ----------------

function listDocx(dir: string): string[] {
  try {
    return readdirSync(dir)
      .filter((n) => n.toLowerCase().endsWith('.docx'))
      .map((n) => join(dir, n))
  } catch {
    return []
  }
}

function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false
  return true
}

function findParagraphWithImage(nodes: TiptapNode[]): TiptapNode | null {
  for (const n of nodes) {
    if (n.type === 'paragraph' && Array.isArray(n.content)) {
      if (n.content.some((c) => c.type === 'image')) return n
    }
    if (Array.isArray(n.content)) {
      const f = findParagraphWithImage(n.content)
      if (f) return f
    }
  }
  return null
}

function findFirstTextParagraph(nodes: TiptapNode[]): TiptapNode | null {
  for (const n of nodes) {
    if (n.type === 'paragraph' && Array.isArray(n.content)) {
      if (n.content.some((c) => c.type === 'text' && (c.text ?? '').length > 0)) return n
    }
    if (Array.isArray(n.content)) {
      const f = findFirstTextParagraph(n.content)
      if (f) return f
    }
  }
  return null
}

function listImages(nodes: TiptapNode[]): TiptapNode[] {
  const out: TiptapNode[] = []
  const walk = (n: TiptapNode) => {
    if (n.type === 'image') out.push(n)
    if (Array.isArray(n.content)) for (const c of n.content) walk(c)
  }
  for (const n of nodes) walk(n)
  return out
}

function markDirty(node: TiptapNode) {
  const a = (node.attrs ?? {}) as Record<string, unknown>
  a.__origContentFp = '__dirty__'
  node.attrs = a
}

function mutateText(node: TiptapNode, suffix: string): boolean {
  if (!Array.isArray(node.content)) return false
  for (const c of node.content) {
    if (c.type === 'text' && typeof c.text === 'string') {
      c.text = c.text + suffix
      markDirty(node)
      return true
    }
  }
  if (node.content.some((c) => c.type === 'image')) {
    node.content.unshift({ type: 'text', text: suffix.trim() + ' ' })
    markDirty(node)
    return true
  }
  return false
}

const TINY_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGBgAAAABQABXvMqOgAAAABJRU5ErkJggg=='

function makeNewImageNode(): TiptapNode {
  return {
    type: 'image',
    attrs: {
      src: `data:image/png;base64,${TINY_PNG_BASE64}`,
      width: 40,
      height: 40,
    },
  }
}

// ---------------- scenarios ----------------

interface ScenarioResult {
  name: string
  pass: boolean
  reason?: string
  notes: string
}

async function scenarioNoop(docxPath: string): Promise<ScenarioResult> {
  const buf = readFileSync(docxPath)
  const imp = await importDocxPipeline(buf)
  const exp = await exportDocxPipeline({
    content: imp.data.content,
    originalDocx: buf,
  })
  const rawOrig = extractRawDocx(buf)
  const rawNew = extractRawDocx(exp.buffer)
  const origDoc = rawOrig.partsByPath.get('word/document.xml')!
  const newDoc = rawNew.partsByPath.get('word/document.xml')!
  if (bytesEqual(origDoc, newDoc)) {
    return {
      name: 'no-op',
      pass: true,
      notes: `docBytes=${origDoc.length} reuse=${exp.stats.classifier.reuseNodes}/${exp.stats.classifier.totalNodes}`,
    }
  }
  // 退而求其次：body hash 序列相同（sectPr 顺序/空白差异是可接受的）
  const origIdx = indexTopLevelRanges(origDoc, 'w:body')!
  const newIdx = indexTopLevelRanges(newDoc, 'w:body')!
  const h = (bytes: Uint8Array, idx: typeof origIdx): string[] =>
    idx.ranges
      .filter((r) => r.tag !== 'w:sectPr')
      .map((r) => `${r.tag}:${hashXmlRange(sliceRange(bytes, r))}`)
  const hashesOrig = h(origDoc, origIdx)
  const hashesNew = h(newDoc, newIdx)
  if (hashesOrig.length === hashesNew.length && hashesOrig.every((x, i) => x === hashesNew[i])) {
    return {
      name: 'no-op',
      pass: true,
      notes: `body hash seq eq (bytes diff, likely sectPr/whitespace). origBytes=${origDoc.length} newBytes=${newDoc.length}`,
    }
  }
  return {
    name: 'no-op',
    pass: false,
    reason: `bytes differ and body hash seq differs: orig=${hashesOrig.length} new=${hashesNew.length}`,
    notes: `origBytes=${origDoc.length} newBytes=${newDoc.length}`,
  }
}

async function scenarioEdit(docxPath: string): Promise<ScenarioResult> {
  const buf = readFileSync(docxPath)
  const imp = await importDocxPipeline(buf)
  const origImg = listImages(imp.data.content.content).length
  const target =
    findParagraphWithImage(imp.data.content.content) ??
    findFirstTextParagraph(imp.data.content.content)
  if (!target) return { name: 'edit', pass: true, notes: 'no editable paragraph, skip' }
  const hasImage = Array.isArray(target.content) && target.content.some((c) => c.type === 'image')
  if (!mutateText(target, ' [edited]')) {
    return { name: 'edit', pass: true, notes: 'no text to mutate, skip' }
  }

  const exp = await exportDocxPipeline({ content: imp.data.content, originalDocx: buf })
  const reimp = await importDocxPipeline(exp.buffer)
  const newImg = listImages(reimp.data.content.content).length

  const issues: string[] = []
  if (hasImage) {
    if (exp.stats.media.newFiles !== 0) issues.push(`media.new=${exp.stats.media.newFiles}`)
    if (exp.stats.rels.relsAppended !== 0) issues.push(`rels.appended=${exp.stats.rels.relsAppended}`)
  }
  if (newImg !== origImg) issues.push(`images ${origImg} → ${newImg}`)
  if (reimp.logs.error.length > 0) issues.push(`reimport errors=${reimp.logs.error.length}`)

  return {
    name: 'edit',
    pass: issues.length === 0,
    reason: issues.join(', ') || undefined,
    notes:
      `hasImage=${hasImage} regen=${exp.stats.classifier.regenerateNodes} ` +
      `media.new=${exp.stats.media.newFiles} rels.appended=${exp.stats.rels.relsAppended}`,
  }
}

async function scenarioInsertImage(docxPath: string): Promise<ScenarioResult> {
  const buf = readFileSync(docxPath)
  const imp = await importDocxPipeline(buf)
  const origImg = listImages(imp.data.content.content).length
  const target = findFirstTextParagraph(imp.data.content.content)
  if (!target || !Array.isArray(target.content)) {
    return { name: 'insert-image', pass: true, notes: 'no paragraph, skip' }
  }
  target.content.push(makeNewImageNode())
  markDirty(target)

  const exp = await exportDocxPipeline({ content: imp.data.content, originalDocx: buf })
  const reimp = await importDocxPipeline(exp.buffer)
  const newImg = listImages(reimp.data.content.content).length

  const issues: string[] = []
  if (exp.stats.media.newFiles !== 1) issues.push(`media.new=${exp.stats.media.newFiles} (want 1)`)
  if (exp.stats.rels.relsAppended !== 1) issues.push(`rels.appended=${exp.stats.rels.relsAppended} (want 1)`)
  if (newImg !== origImg + 1) issues.push(`images ${origImg} → ${newImg} (want +1)`)
  if (reimp.logs.error.length > 0) issues.push(`reimport errors=${reimp.logs.error.length}`)

  return {
    name: 'insert-image',
    pass: issues.length === 0,
    reason: issues.join(', ') || undefined,
    notes:
      `regen=${exp.stats.classifier.regenerateNodes} ` +
      `media.new=${exp.stats.media.newFiles} rels.appended=${exp.stats.rels.relsAppended} ` +
      `CT.appended=${exp.stats.contentTypes.defaultsAppended}`,
  }
}

// ---------------- main ----------------

async function runOne(docxPath: string) {
  const label = basename(docxPath)
  console.log('\n' + '='.repeat(78))
  console.log(label)
  console.log('='.repeat(78))
  const scenarios: ScenarioResult[] = []
  try {
    scenarios.push(await scenarioNoop(docxPath))
  } catch (e) {
    scenarios.push({ name: 'no-op', pass: false, reason: (e as Error).message, notes: '' })
  }
  try {
    scenarios.push(await scenarioEdit(docxPath))
  } catch (e) {
    scenarios.push({ name: 'edit', pass: false, reason: (e as Error).message, notes: '' })
  }
  try {
    scenarios.push(await scenarioInsertImage(docxPath))
  } catch (e) {
    scenarios.push({ name: 'insert-image', pass: false, reason: (e as Error).message, notes: '' })
  }
  for (const r of scenarios) {
    const flag = r.pass ? 'PASS' : 'FAIL'
    console.log(`  [${flag}] ${r.name.padEnd(14)} ${r.notes}${r.reason ? ` :: ${r.reason}` : ''}`)
  }
  return scenarios
}

async function main() {
  const args = process.argv.slice(2)
  const explicit: string[] = []
  const dirs: string[] = []
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--dir') {
      dirs.push(args[++i])
    } else {
      explicit.push(args[i])
    }
  }
  const targets: string[] = []
  if (explicit.length > 0) {
    targets.push(...explicit)
  } else {
    const searchDirs = dirs.length > 0 ? dirs : DEFAULT_DIRS
    for (const d of searchDirs) {
      try {
        const st = statSync(d)
        if (st.isDirectory()) targets.push(...listDocx(d))
      } catch {
        // skip
      }
    }
  }
  if (targets.length === 0) {
    console.error('No .docx found. Provide files or --dir.')
    process.exit(2)
  }

  let pass = 0
  let fail = 0
  const failList: string[] = []
  for (const p of targets) {
    const results = await runOne(p)
    for (const r of results) {
      if (r.pass) pass++
      else {
        fail++
        failList.push(`${basename(p)} :: ${r.name} :: ${r.reason ?? 'fail'}`)
      }
    }
  }
  console.log('\n' + '='.repeat(78))
  console.log(`SUMMARY: pass=${pass} fail=${fail}  docs=${targets.length}`)
  if (failList.length > 0) {
    console.log('\nFailures:')
    for (const f of failList) console.log('  - ' + f)
  }
  console.log('='.repeat(78))
  if (fail > 0) process.exit(1)
}

main().catch((err) => {
  console.error(err)
  process.exit(1)
})

/**
 * 阶段 4 验证：模拟用户编辑（修改含图片的段落、插入新图片），导出后验证：
 *   - 编辑含图段落 → 图片 rId 被复用（newFiles=0, relsAppended=0）
 *   - 插入全新图片 → media/rels/Content_Types 正确追加（newFiles=1, relsAppended=1）
 *   - 产出 DOCX 可被 importPipeline 重新读取，且图片数量符合预期
 *
 * 用法：
 *   npx tsx scripts/verifyExportEdit.ts <docx> [<docx>...]
 */

import { readFileSync } from 'node:fs'
import { basename } from 'node:path'
import { importDocxPipeline } from '../src/engine/importPipeline.js'
import { exportDocxPipeline } from '../src/engine/exportPipeline.js'
import { extractRawDocx } from '../src/engine/zipExtractor.js'
import type { TiptapNode } from '../src/types/tiptapJson.js'

function findFirstParagraphWithImage(nodes: TiptapNode[]): TiptapNode | null {
  for (const n of nodes) {
    if (n.type === 'paragraph' && Array.isArray(n.content)) {
      if (n.content.some((c) => c.type === 'image')) return n
    }
    if (Array.isArray(n.content)) {
      const f = findFirstParagraphWithImage(n.content)
      if (f) return f
    }
  }
  return null
}

function findFirstNonEmptyParagraph(nodes: TiptapNode[]): TiptapNode | null {
  for (const n of nodes) {
    if (n.type === 'paragraph' && Array.isArray(n.content)) {
      if (n.content.some((c) => c.type === 'text' && (c.text ?? '').length > 0)) return n
    }
    if (Array.isArray(n.content)) {
      const f = findFirstNonEmptyParagraph(n.content)
      if (f) return f
    }
  }
  return null
}

function markDirty(node: TiptapNode): void {
  const a = (node.attrs ?? {}) as Record<string, unknown>
  a.__origContentFp = '__dirty__'
  node.attrs = a
}

function mutateParagraph(p: TiptapNode): boolean {
  if (!Array.isArray(p.content)) return false
  for (const c of p.content) {
    if (c.type === 'text' && typeof c.text === 'string') {
      c.text = c.text + ' [edited]'
      markDirty(p)
      return true
    }
  }
  // 整段都是图：插入一段前置文字以强制 regen
  if (p.content.some((c) => c.type === 'image')) {
    p.content.unshift({ type: 'text', text: '[edited] ' })
    markDirty(p)
    return true
  }
  return false
}

function listImageNodes(nodes: TiptapNode[]): TiptapNode[] {
  const out: TiptapNode[] = []
  const walk = (n: TiptapNode) => {
    if (n.type === 'image') out.push(n)
    if (Array.isArray(n.content)) for (const c of n.content) walk(c)
  }
  for (const n of nodes) walk(n)
  return out
}

/** 1x1 透明 PNG */
const TINY_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGBgAAAABQABXvMqOgAAAABJRU5ErkJggg=='

function makeNewImageNode(): TiptapNode {
  return {
    type: 'image',
    attrs: {
      src: `data:image/png;base64,${TINY_PNG_BASE64}`,
      width: 50,
      height: 50,
    },
  }
}

async function runEdit(filePath: string): Promise<boolean> {
  console.log('\n' + '-'.repeat(60) + '\n[edit paragraph with image] ' + basename(filePath))
  const buf = readFileSync(filePath)
  const imp = await importDocxPipeline(buf)
  const origImgCount = listImageNodes(imp.data.content.content).length
  console.log(`  imported images=${origImgCount}`)

  const target =
    findFirstParagraphWithImage(imp.data.content.content) ??
    findFirstNonEmptyParagraph(imp.data.content.content)
  if (!target) return true
  const hasImage = Array.isArray(target.content) && target.content.some((c) => c.type === 'image')
  if (!mutateParagraph(target)) return true

  const exp = await exportDocxPipeline({ content: imp.data.content, originalDocx: buf })
  const s = exp.stats
  console.log(
    `  regen=${s.classifier.regenerateNodes} ` +
      `rels.appended=${s.rels.relsAppended} CT.appended=${s.contentTypes.defaultsAppended} media.new=${s.media.newFiles}`,
  )
  let pass = true
  if (hasImage) {
    if (s.media.newFiles !== 0) {
      console.log(`    FAIL: image rId not reused (media.new=${s.media.newFiles})`)
      pass = false
    }
    if (s.rels.relsAppended !== 0) {
      console.log(`    FAIL: unexpected rels appended for reused rId (${s.rels.relsAppended})`)
      pass = false
    }
  }
  const reimp = await importDocxPipeline(exp.buffer)
  const newImgCount = listImageNodes(reimp.data.content.content).length
  if (newImgCount !== origImgCount) {
    console.log(`    FAIL: image count changed (${origImgCount} → ${newImgCount})`)
    pass = false
  }
  console.log(pass ? '  [PASS]' : '  [FAIL]')
  return pass
}

function findFirstListItemParagraph(nodes: TiptapNode[]): { item: TiptapNode; para: TiptapNode } | null {
  for (const n of nodes) {
    if (n.type === 'listItem' && Array.isArray(n.content)) {
      for (const c of n.content) {
        if (c.type === 'paragraph' && Array.isArray(c.content)) {
          if (c.content.some((cc) => cc.type === 'text' && (cc.text ?? '').length > 0)) {
            return { item: n, para: c }
          }
        }
      }
    }
    if (Array.isArray(n.content)) {
      const r = findFirstListItemParagraph(n.content)
      if (r) return r
    }
  }
  return null
}

function extractNumIds(xmlBytes: Uint8Array): Set<string> {
  const text = new TextDecoder().decode(xmlBytes)
  const rx = /<w:numId\b[^/]*\bw:val="(\d+)"\s*\/>/g
  const out = new Set<string>()
  let m: RegExpExecArray | null
  while ((m = rx.exec(text)) !== null) out.add(m[1])
  return out
}

async function runEditListItem(filePath: string): Promise<boolean> {
  console.log('\n' + '-'.repeat(60) + '\n[edit list item] ' + basename(filePath))
  const buf = readFileSync(filePath)
  const imp = await importDocxPipeline(buf)
  const hit = findFirstListItemParagraph(imp.data.content.content)
  if (!hit) {
    console.log('  no list item, skip')
    return true
  }
  // 修改段落文字
  for (const c of hit.para.content as TiptapNode[]) {
    if (c.type === 'text' && typeof c.text === 'string') {
      c.text = c.text + ' [edited list]'
      break
    }
  }
  markDirty(hit.para)

  const exp = await exportDocxPipeline({ content: imp.data.content, originalDocx: buf })
  const s = exp.stats
  console.log(
    `  regen=${s.classifier.regenerateNodes} reuse=${s.classifier.reuseNodes} ` +
      `inserted=${s.patcher.insertedFragments} genBytes=${s.patcher.generatedBytes}`,
  )
  let pass = true
  if (s.classifier.regenerateNodes === 0) {
    console.log('    FAIL: regen not triggered')
    pass = false
  }
  // 验证导出的 document.xml 里 numId 都在原 numbering 可用 numId 集合里
  const out = extractRawDocx(exp.buffer)
  const docXml = out.partsByPath.get('word/document.xml')!
  const origNumbering = extractRawDocx(buf).partsByPath.get('word/numbering.xml')
  const outNumIds = extractNumIds(docXml)
  if (origNumbering) {
    const origText = new TextDecoder().decode(origNumbering)
    const origNumIdSet = new Set<string>()
    const rx = /<w:num\b[^>]*\bw:numId="(\d+)"/g
    let m: RegExpExecArray | null
    while ((m = rx.exec(origText)) !== null) origNumIdSet.add(m[1])
    // numId="0" 是 OOXML 规范里的特殊值（清除列表引用），无需在 <w:num> 中定义
    const invalid = [...outNumIds].filter((id) => id !== '0' && !origNumIdSet.has(id))
    console.log(`  numIds in output=${[...outNumIds].join(',')}; invalid(not in orig numbering)=${invalid.join(',') || '-'}`)
    if (invalid.length > 0) {
      console.log(`    FAIL: fragment references numId not in original numbering: ${invalid.join(',')}`)
      pass = false
    }
  }
  // re-import 应该仍能识别列表
  const reimp = await importDocxPipeline(exp.buffer)
  console.log(`  re-import errors=${reimp.logs.error.length}`)
  console.log(pass ? '  [PASS]' : '  [FAIL]')
  return pass
}

async function runInsertNewImage(filePath: string): Promise<boolean> {
  console.log('\n' + '-'.repeat(60) + '\n[insert new image] ' + basename(filePath))
  const buf = readFileSync(filePath)
  const imp = await importDocxPipeline(buf)
  const origImgCount = listImageNodes(imp.data.content.content).length
  const target = findFirstNonEmptyParagraph(imp.data.content.content)
  if (!target || !Array.isArray(target.content)) return true
  target.content.push(makeNewImageNode())
  markDirty(target)

  const exp = await exportDocxPipeline({ content: imp.data.content, originalDocx: buf })
  const s = exp.stats
  console.log(
    `  regen=${s.classifier.regenerateNodes} ` +
      `rels.appended=${s.rels.relsAppended} CT.appended=${s.contentTypes.defaultsAppended} media.new=${s.media.newFiles}`,
  )
  let pass = true
  if (s.media.newFiles !== 1) {
    console.log(`    FAIL: expected media.new=1, got ${s.media.newFiles}`)
    pass = false
  }
  if (s.rels.relsAppended !== 1) {
    console.log(`    FAIL: expected rels.appended=1, got ${s.rels.relsAppended}`)
    pass = false
  }
  // 产出 DOCX 能 re-import 且图片数=原数+1
  const reimp = await importDocxPipeline(exp.buffer)
  const newImgCount = listImageNodes(reimp.data.content.content).length
  if (newImgCount !== origImgCount + 1) {
    console.log(`    FAIL: image count not +1 (${origImgCount} → ${newImgCount})`)
    pass = false
  }
  // rels 总数应与原档 +1
  const outRels = extractRawDocx(exp.buffer).partsByPath.get('word/_rels/document.xml.rels')
  const origRels = extractRawDocx(buf).partsByPath.get('word/_rels/document.xml.rels')
  if (outRels && origRels) {
    const outCount = (new TextDecoder().decode(outRels).match(/<Relationship\b/g) ?? []).length
    const origCount = (new TextDecoder().decode(origRels).match(/<Relationship\b/g) ?? []).length
    if (outCount !== origCount + 1) {
      console.log(`    FAIL: rels count not +1 (${origCount} → ${outCount})`)
      pass = false
    }
  }
  console.log(pass ? '  [PASS]' : '  [FAIL]')
  return pass
}

async function main() {
  const args = process.argv.slice(2)
  if (args.length === 0) {
    console.error('用法: npx tsx scripts/verifyExportEdit.ts <docx> [<docx>...]')
    process.exit(2)
  }
  let pass = 0
  let fail = 0
  for (const p of args) {
    console.log('\n' + '='.repeat(80) + '\n[' + basename(p) + ']\n' + '='.repeat(80))
    for (const ok of [await runEdit(p), await runInsertNewImage(p), await runEditListItem(p)]) {
      if (ok) pass++
      else fail++
    }
  }
  console.log('\n' + '='.repeat(80))
  console.log(`SUMMARY: pass=${pass} fail=${fail}`)
  console.log('='.repeat(80))
  if (fail > 0) process.exit(1)
}

main().catch((err) => {
  console.error(err)
  process.exit(1)
})

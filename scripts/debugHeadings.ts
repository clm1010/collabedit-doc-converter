/**
 * 调试脚本：跑 importDocxPipeline，dump 所有 heading 节点的 level + 前 40 字文本。
 * 用法: npx tsx scripts/debugHeadings.ts <docxPath>
 */
import { readFileSync } from 'node:fs'
import { importDocxPipeline } from '../src/engine/importPipeline.js'
import type { TiptapNode } from '../src/types/tiptapJson.js'

function* walk(nodes: TiptapNode[]): Generator<TiptapNode> {
  for (const n of nodes) {
    yield n
    if (n.content) yield* walk(n.content)
  }
}

function textOf(node: TiptapNode): string {
  let out = ''
  const stack: TiptapNode[] = [node]
  while (stack.length) {
    const n = stack.shift()!
    if (n.type === 'text' && n.text) out += n.text
    if (n.content) stack.unshift(...n.content)
  }
  return out
}

async function main() {
  const path = process.argv[2]
  if (!path) {
    console.error('用法: npx tsx scripts/debugHeadings.ts <docx>')
    process.exit(2)
  }
  const buf = readFileSync(path)
  const res = await importDocxPipeline(buf)
  const top = res.data.content.content
  console.log(`顶层节点数: ${top.length}`)

  const typeHist: Record<string, number> = {}
  for (const n of walk(top)) {
    typeHist[n.type] = (typeHist[n.type] ?? 0) + 1
  }
  console.log('节点类型分布:', typeHist)

  // 查找所有段落文本包含目标关键词的节点，打印其类型和 level
  // 遍历所有节点查找 heading 并打印其 level / numberingText
  let idx = 0
  for (const n of walk(top)) {
    if (n.type !== 'heading') continue
    const a = (n.attrs ?? {}) as Record<string, unknown>
    const lvl = a.level
    const numText = a.numberingText ?? ''
    const origNumPr = a.__origNumPr
    const text = textOf(n).replace(/\s+/g, ' ').slice(0, 50)
    console.log(
      `#${idx++} H${lvl} numberingText="${numText}" origNumPr=${JSON.stringify(origNumPr ?? null)} | ${text}`,
    )
  }
  console.log(`heading 总数: ${idx}`)
}

main().catch((e) => {
  console.error(e)
  process.exit(1)
})

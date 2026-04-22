/**
 * Content Fingerprint：从 Tiptap 节点的"可编辑内容"算出稳定 64-bit 指纹。
 *
 * 用途：选择性保存导出时判定一个节点相对于原始 DOCX 是否被用户修改过。
 *
 *   - import 阶段：每个挂 __origRange 的顶层节点额外计算 `__origContentFp`
 *     并存入 attrs。该值代表"这个节点刚从 DOCX 解析出来时的内容形态"。
 *   - export 阶段：对当前 Tiptap JSON 的同一节点重新计算 contentFp，
 *     与 __origContentFp 比较：
 *       - 相同  → 节点未被改动，走字节级复用路径（`__origRange`）
 *       - 不同  → 节点被改动过，走 localSerializer 重生成路径
 *
 * 关键设计：
 *   - 必须是幂等、稳定的：同样 node 输入必须产生同样 hash，
 *     所以在序列化前做：对象键排序、移除 __orig / __dirty 辅助 attrs、
 *     剥离 undefined/null、空数组归一。
 *   - 和 OOXML 无关：不走 XML canonicalize，只对 JSON 本身指纹化。
 *     这样避开了 "docx 库重新打包产生的 OOXML 与 Word 原生 OOXML 不一致" 的陷阱。
 *   - 前端若要本地预判也可复用：后续 phase 可以把本文件的算法移植到 FE。
 */

import { fnv1a64 } from './hasher.js'
import type { TiptapNode } from '../types/tiptapJson.js'

/**
 * 不参与内容指纹的 attrs key 列表。
 *
 * - 所有以 `__orig` 开头的字段都是我们自己注入的元数据，不应参与 hash。
 * - `__dirty` 保留为可选的编辑脏标。
 * - `data-origin-*` 是图片的原始 rels 元数据，本身不影响"视觉内容"，排除。
 *   （它只影响导出时是否复用原 rId，不改动视觉不该触发 regenerate。）
 */
const EXCLUDED_ATTR_PATTERNS: RegExp[] = [
  /^__orig/,
  /^__dirty$/,
  /^data-origin-/,
]

function shouldExcludeAttr(key: string): boolean {
  return EXCLUDED_ATTR_PATTERNS.some((re) => re.test(key))
}

/**
 * 稳定序列化：对对象键按字典序排序，剥离 undefined，保留 null。
 * 数组顺序保留（顺序有语义）。
 */
function stableStringify(value: unknown): string {
  if (value === null) return 'null'
  if (value === undefined) return 'null'
  if (typeof value === 'number') {
    return Number.isFinite(value) ? String(value) : 'null'
  }
  if (typeof value === 'string' || typeof value === 'boolean') {
    return JSON.stringify(value)
  }
  if (Array.isArray(value)) {
    return '[' + value.map(stableStringify).join(',') + ']'
  }
  if (typeof value === 'object') {
    const obj = value as Record<string, unknown>
    const keys = Object.keys(obj).sort()
    const parts: string[] = []
    for (const k of keys) {
      const v = obj[k]
      if (v === undefined) continue
      parts.push(JSON.stringify(k) + ':' + stableStringify(v))
    }
    return '{' + parts.join(',') + '}'
  }
  // bigint / function / symbol 一律归 null
  return 'null'
}

/**
 * 清洗一个 Tiptap 节点：递归移除 __orig / data-origin- 等不参与指纹的字段。
 * 返回新的对象，不修改入参。
 */
export function cleanNodeForFingerprint(node: TiptapNode): unknown {
  const out: Record<string, unknown> = { type: node.type }

  if (node.attrs && typeof node.attrs === 'object') {
    const cleanedAttrs: Record<string, unknown> = {}
    for (const [k, v] of Object.entries(node.attrs)) {
      if (shouldExcludeAttr(k)) continue
      if (v === undefined) continue
      cleanedAttrs[k] = v
    }
    if (Object.keys(cleanedAttrs).length > 0) out.attrs = cleanedAttrs
  }

  if (Array.isArray(node.marks) && node.marks.length > 0) {
    out.marks = node.marks.map((m) => {
      const mo: Record<string, unknown> = { type: m.type }
      if (m.attrs && typeof m.attrs === 'object') {
        const ma: Record<string, unknown> = {}
        for (const [k, v] of Object.entries(m.attrs)) {
          if (shouldExcludeAttr(k)) continue
          if (v === undefined) continue
          ma[k] = v
        }
        if (Object.keys(ma).length > 0) mo.attrs = ma
      }
      return mo
    })
  }

  if (typeof node.text === 'string') out.text = node.text

  if (Array.isArray(node.content) && node.content.length > 0) {
    out.content = node.content.map(cleanNodeForFingerprint)
  }

  return out
}

/**
 * 计算节点的内容指纹（16 位小写 hex，用 FNV-1a 64-bit）。
 *
 * 相同的节点结构必然产生相同的 hash，跨进程/跨语言只要采用同样的
 * `cleanNodeForFingerprint` + `stableStringify` 就能复现。
 */
export function computeContentFingerprint(node: TiptapNode): string {
  const cleaned = cleanNodeForFingerprint(node)
  const json = stableStringify(cleaned)
  return fnv1a64(new TextEncoder().encode(json))
}

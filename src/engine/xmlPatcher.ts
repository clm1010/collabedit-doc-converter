/**
 * 选择性保存高保真方案：XML 部件字节级 Patcher。
 *
 * 通过"游标推进 + 跳过已删除 ranges"的方式拼出新 document.xml 字节：
 *
 *   [0, containerContentStart)   ← 原文档序言 + <w:body ...> 开标签
 *   body 区间 [containerContentStart, containerContentEnd)：
 *     按 segments 顺序：
 *       reuse 段 → 把原字节 [cursor, seg.end) 原样复制
 *                  （其中若跨越了被删除的顶层 ranges，则跳过那些 ranges 的字节；
 *                   ranges 之间的 whitespace / 非顶层元素字节照常保留。）
 *       regen 段 → 在当前 cursor 位置插入 localSerializer 产出的新 XML，cursor 不动
 *     末尾：把剩余字节 [cursor, containerContentEnd) 按同样规则复制
 *           （这会自然保留末尾 sectPr 和尾部 whitespace）
 *   [containerContentEnd, eof)   ← </w:body></w:document> 等尾部
 *
 * 这样的设计保证：
 *   - 用户未改动且未删除的节点：100% 字节保留（包括 w:p 属性里的 rsid/paraId 等）。
 *   - ranges 之间的 whitespace / bookmarkStart / proofErr / commentRangeStart 等
 *     非顶层 ranges 识别范围之外的字节也完整保留（只要它们没被删除的 range 吞掉）。
 *   - sectPr 保持在 body 末尾原位置（它本身也是顶层 range，不在 deletedRanges 中，
 *     通过末尾 sliceExcludingDeleted 自然写出）。
 *   - 被删除节点对应的原 range 字节被精确跳过。
 *
 * 不变量：
 *   - 所有 reuse segment 的 [start, end) 必须落在 [containerContentStart, containerContentEnd) 之内。
 *   - reuse segment 的 end 必须 > 所有之前已处理 segment 的 end（Tiptap 顺序 = body 字节顺序）。
 *     违反该不变量说明用户跨段移动了节点 —— 这种情况 phase 3 暂不支持跨度节点复用，
 *     由 classifier 降级为 regenerate 处理。
 */

import type { RangeIndex, TopLevelRange } from './xmlRangeIndexer.js'
import type { Segment } from './nodeClassifier.js'

export interface PatchDocumentXmlInput {
  /** 原 document.xml 字节 */
  originalBytes: Uint8Array
  /** 顶层范围索引（xmlRangeIndexer 产出） */
  rangeIndex: RangeIndex
  /** 分类器产出的段列表，顺序即新文档顺序 */
  segments: Segment[]
  /** 每个 regenerate 段对应的 XML 字节，顺序与 segments 中 regenerate 段出现顺序一致 */
  regeneratedFragments: Uint8Array[]
}

export interface PatchDocumentXmlOutput {
  /** 新 document.xml 字节 */
  bytes: Uint8Array
  /** 保留字节数（从 original body 原样复制过来的部分） */
  reusedBytes: number
  /** 新生成字节数（regenerate 片段） */
  generatedBytes: number
  /** 被跳过字节数（删除的原 ranges） */
  droppedBytes: number
  /** 已插入的 regenerate 片段数 */
  insertedFragments: number
  /** 被删除（未复用）的原 top-level ranges 数（不含 sectPr） */
  deletedRangeCount: number
}

export function patchDocumentXml(
  input: PatchDocumentXmlInput,
): PatchDocumentXmlOutput {
  const { originalBytes, rangeIndex, segments, regeneratedFragments } = input

  // 1. 识别"被复用"的原 ranges：凡是 start 落在任一 reuse segment [s.start, s.end) 内
  //    且 end <= s.end 的原 range，都算被复用（含 wrapper 合并情况）。
  const reusedStarts = new Set<number>()
  for (const seg of segments) {
    if (seg.kind !== 'reuse') continue
    validateReuseInBody(seg, rangeIndex)
    for (const r of rangeIndex.ranges) {
      if (r.start >= seg.start && r.end <= seg.end) {
        reusedStarts.add(r.start)
      }
    }
  }

  // 2. 被删除的 ranges = 原 ranges 中未被任何 reuse 段覆盖的（排除 sectPr，
  //    sectPr 总是保留在 body 末尾）。
  const deletedRanges = rangeIndex.ranges
    .filter((r) => r.tag !== 'w:sectPr' && !reusedStarts.has(r.start))
    .sort((a, b) => a.start - b.start)

  const pieces: Uint8Array[] = []
  let reusedBytes = 0
  let generatedBytes = 0
  let droppedBytes = 0
  let regenCursor = 0
  let cursor = rangeIndex.containerContentStart

  // 3. 头部：[0, containerContentStart)
  pieces.push(originalBytes.subarray(0, rangeIndex.containerContentStart))

  // 4. 按 segments 顺序推进 body
  for (const seg of segments) {
    if (seg.kind === 'reuse') {
      if (seg.end < cursor) {
        throw new Error(
          `xmlPatcher: reuse segment [${seg.start}, ${seg.end}) precedes cursor ${cursor}.` +
            ' selective save does not support cross-segment node reordering in phase 3',
        )
      }
      // 从 cursor 复制到 seg.end，跳过中间的 deletedRanges
      const { copied, dropped } = copyExcludingDeleted(
        originalBytes,
        cursor,
        seg.end,
        deletedRanges,
        pieces,
      )
      reusedBytes += copied
      droppedBytes += dropped
      cursor = seg.end
    } else {
      const frag = regeneratedFragments[regenCursor++]
      if (!frag) {
        throw new Error(
          `xmlPatcher: missing regenerated fragment at index ${regenCursor - 1}`,
        )
      }
      if (frag.length > 0) {
        pieces.push(frag)
        generatedBytes += frag.length
      }
    }
  }

  if (regenCursor !== regeneratedFragments.length) {
    throw new Error(
      `xmlPatcher: regeneratedFragments length ${regeneratedFragments.length}` +
        ` does not match regenerate segments ${regenCursor}`,
    )
  }

  // 5. 末尾：[cursor, containerContentEnd)，同样跳过 deletedRanges
  //    这里会自然保留 body 末尾的 w:sectPr 和其前后 whitespace。
  {
    const { copied, dropped } = copyExcludingDeleted(
      originalBytes,
      cursor,
      rangeIndex.containerContentEnd,
      deletedRanges,
      pieces,
    )
    reusedBytes += copied
    droppedBytes += dropped
  }

  // 6. 尾部：[containerContentEnd, eof)
  pieces.push(originalBytes.subarray(rangeIndex.containerContentEnd, originalBytes.length))

  // 合并
  const totalLen = pieces.reduce((sum, p) => sum + p.length, 0)
  const bytes = new Uint8Array(totalLen)
  let offset = 0
  for (const p of pieces) {
    bytes.set(p, offset)
    offset += p.length
  }

  return {
    bytes,
    reusedBytes,
    generatedBytes,
    droppedBytes,
    insertedFragments: regenCursor,
    deletedRangeCount: deletedRanges.length,
  }
}

/**
 * 把 bytes[from, to) 复制到 pieces 中，但跳过 deletedRanges 中与之相交的部分。
 * 返回实际复制 / 丢弃的字节数。
 *
 * deletedRanges 假定按 start 升序。
 */
function copyExcludingDeleted(
  bytes: Uint8Array,
  from: number,
  to: number,
  deletedRanges: TopLevelRange[],
  pieces: Uint8Array[],
): { copied: number; dropped: number } {
  if (from >= to) return { copied: 0, dropped: 0 }

  let c = from
  let copied = 0
  let dropped = 0

  for (const d of deletedRanges) {
    if (d.end <= c) continue
    if (d.start >= to) break
    // 相交
    if (d.start > c) {
      const slice = bytes.subarray(c, d.start)
      pieces.push(slice)
      copied += slice.length
    }
    const overlapStart = Math.max(c, d.start)
    const overlapEnd = Math.min(to, d.end)
    dropped += overlapEnd - overlapStart
    c = Math.max(c, d.end)
    if (c >= to) break
  }

  if (c < to) {
    const slice = bytes.subarray(c, to)
    pieces.push(slice)
    copied += slice.length
  }

  return { copied, dropped }
}

function validateReuseInBody(
  seg: { start: number; end: number },
  rangeIndex: RangeIndex,
): void {
  if (
    seg.start < rangeIndex.containerContentStart ||
    seg.end > rangeIndex.containerContentEnd
  ) {
    throw new Error(
      `xmlPatcher: reuse segment [${seg.start}, ${seg.end}) ` +
        `escapes body [${rangeIndex.containerContentStart}, ${rangeIndex.containerContentEnd})`,
    )
  }
  if (seg.end <= seg.start) {
    throw new Error(
      `xmlPatcher: invalid reuse range [${seg.start}, ${seg.end})`,
    )
  }
}

// 让 IDE 类型检查友好地引用 TopLevelRange
export type { TopLevelRange }

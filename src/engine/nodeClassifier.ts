/**
 * 选择性保存高保真方案：节点分类器（Node Classifier）。
 *
 * 作用：把导出请求中的 Tiptap 顶层节点数组分类为一串"段"（Segment），
 *      每段要么是"字节复用"（reuse，直接从原 document.xml 对应范围拷贝），
 *      要么是"重生成"（regenerate，由 localSerializer 把若干连续节点序列化成 OOXML）。
 *
 * 判定规则（针对顶层节点）：
 *   1. 节点必须具备三元组 __origRange + __origPart + __origContentFp 才有资格 reuse。
 *   2. __origPart 必须等于当前正在 patch 的部件（默认 'word/document.xml'）。
 *   3. 当前节点的 contentFp 必须等于 __origContentFp（未被用户改动）。
 *   4. 包装型节点（bulletList / orderedList / blockquote 等）本身没有 origRange，
 *      一律视为 dirty，进入 regenerate。
 *
 * 相邻合并：
 *   - 连续 reuse 段仅在"字节邻接"时合并成一个 reuse 段（start 连续）。
 *     否则保留多个独立 reuse 段，patcher 会按顺序 splice（隐含跳过删除字节）。
 *   - 连续 regenerate 段合并成一个 regenerate 段（nodes[] concat）。
 *
 * 删除的表达：
 *   - 当前 JSON 中缺失的节点，等于删除；只要它不在 segments 的任何 reuse 字节区间内，
 *     新 document.xml 自然不包含它。分类器不显式产出 "delete" 段。
 *
 * 返回的统计便于观察选择性保存的命中率（由 exportPipeline 打日志 / 上报）。
 */

import { computeContentFingerprint } from './contentFingerprint.js'
import type { TiptapNode } from '../types/tiptapJson.js'

export interface ReuseSegment {
  kind: 'reuse'
  /** 原 document.xml 的字节区间 [start, end) */
  start: number
  end: number
  /** 来源部件路径，如 'word/document.xml' */
  partPath: string
  /** reuse 段覆盖的节点，便于调试 / 日志 */
  nodes: TiptapNode[]
}

export interface RegenerateSegment {
  kind: 'regenerate'
  /** 需要重新序列化的连续节点（原始顺序） */
  nodes: TiptapNode[]
}

export type Segment = ReuseSegment | RegenerateSegment

export interface ClassifierStats {
  totalNodes: number
  reuseNodes: number
  regenerateNodes: number
  reuseBytes: number
  reuseSegments: number
  regenerateSegments: number
  /** 因 part mismatch / 缺少 fp / fp 对不上 / 包装节点 无法 reuse 的节点数 */
  dirtyReasons: {
    missingRange: number
    missingFp: number
    partMismatch: number
    fpMismatch: number
    wrapper: number
  }
}

export interface ClassifierOptions {
  /** 正在 patch 的部件，默认 'word/document.xml' */
  partPath?: string
}

export interface ClassifierResult {
  segments: Segment[]
  stats: ClassifierStats
}

const DEFAULT_PART = 'word/document.xml'

/**
 * 包装型节点：自身没有 __origRange，内容由后代块级叶子节点承载。
 *
 * listItem 也视作 wrapper：bulletList → listItem → paragraph/... 的结构里，
 * listItem 不挂 origRange，需要透明下穿到其子 paragraph。
 */
const WRAPPER_TYPES = new Set([
  'bulletList',
  'orderedList',
  'listItem',
  'blockquote',
])

interface ParsedRange {
  start: number
  end: number
}

function parseOrigRange(val: unknown): ParsedRange | null {
  if (!Array.isArray(val) || val.length !== 2) return null
  const s = Number(val[0])
  const e = Number(val[1])
  if (!Number.isFinite(s) || !Number.isFinite(e) || e <= s) return null
  return { start: s, end: e }
}

/**
 * 对单个顶层节点做分类判定。返回 "clean" 表示可以字节复用；
 * 返回 "dirty" 表示必须走重生成。
 */
function classifyOne(
  node: TiptapNode,
  partPath: string,
  reasons: ClassifierStats['dirtyReasons'],
): { clean: true; range: ParsedRange } | { clean: false } {
  if (WRAPPER_TYPES.has(node.type)) {
    return classifyWrapper(node, partPath, reasons)
  }

  const attrs = (node.attrs ?? {}) as Record<string, unknown>
  const range = parseOrigRange(attrs.__origRange)
  if (!range) {
    reasons.missingRange++
    return { clean: false }
  }
  const storedFp = attrs.__origContentFp
  if (typeof storedFp !== 'string') {
    reasons.missingFp++
    return { clean: false }
  }
  const origPart =
    typeof attrs.__origPart === 'string' ? attrs.__origPart : DEFAULT_PART
  if (origPart !== partPath) {
    reasons.partMismatch++
    return { clean: false }
  }

  const currentFp = computeContentFingerprint(node)
  if (currentFp !== storedFp) {
    reasons.fpMismatch++
    return { clean: false }
  }

  return { clean: true, range }
}

/**
 * 包装节点的 clean 判定：
 *   - 递归收集所有块级叶子后代的 origRange
 *   - 所有叶子都必须 clean（fp 相同、partPath 相同、range 存在）
 *   - 所有叶子的字节范围必须**严格连续递增**（range[i].start === range[i-1].end）
 *   - 满足则包装节点整体 reuse；否则 dirty（整块重生成）
 *
 * 连续性不满足通常意味着：原文档中这些段落被其他段落/表格打断，
 * 用户操作让它们聚拢成列表 —— 这种情况必须 regenerate 才能保真。
 */
function classifyWrapper(
  node: TiptapNode,
  partPath: string,
  reasons: ClassifierStats['dirtyReasons'],
): { clean: true; range: ParsedRange } | { clean: false } {
  const probe: LeafProbe = { ranges: [], dirty: false }
  collectWrapperLeaves(node, partPath, probe)

  if (probe.dirty || probe.ranges.length === 0) {
    reasons.wrapper++
    return { clean: false }
  }

  for (let i = 1; i < probe.ranges.length; i++) {
    if (probe.ranges[i].start !== probe.ranges[i - 1].end) {
      reasons.wrapper++
      return { clean: false }
    }
  }

  return {
    clean: true,
    range: {
      start: probe.ranges[0].start,
      end: probe.ranges[probe.ranges.length - 1].end,
    },
  }
}

interface LeafProbe {
  ranges: ParsedRange[]
  dirty: boolean
}

/**
 * 把包装节点的所有后代块级叶子按遍历顺序收入 probe.ranges；
 * 遇到任何 dirty 条件立即把 probe.dirty 置 true 并短路。
 */
function collectWrapperLeaves(
  node: TiptapNode,
  partPath: string,
  probe: LeafProbe,
): void {
  if (probe.dirty) return

  if (WRAPPER_TYPES.has(node.type)) {
    const children = Array.isArray(node.content) ? node.content : []
    for (const c of children) {
      collectWrapperLeaves(c, partPath, probe)
      if (probe.dirty) return
    }
    return
  }

  // 叶子节点：必须满足 origRange + fp + part + 内容未改
  const attrs = (node.attrs ?? {}) as Record<string, unknown>
  const range = parseOrigRange(attrs.__origRange)
  if (!range) {
    probe.dirty = true
    return
  }
  const storedFp = attrs.__origContentFp
  if (typeof storedFp !== 'string') {
    probe.dirty = true
    return
  }
  const origPart =
    typeof attrs.__origPart === 'string' ? attrs.__origPart : DEFAULT_PART
  if (origPart !== partPath) {
    probe.dirty = true
    return
  }
  if (computeContentFingerprint(node) !== storedFp) {
    probe.dirty = true
    return
  }

  probe.ranges.push(range)
}

// -------------------------------------------------------------------
// SDT/TOC 组识别
// -------------------------------------------------------------------

/** 取 tocEntry 的 __origSdtId；不存在返回 null */
function getTocSdtId(node: TiptapNode): string | null {
  const v = (node.attrs as Record<string, unknown> | undefined)?.__origSdtId
  return typeof v === 'string' ? v : null
}

/**
 * 从 i 开始扫描连续共享相同 __origSdtId 的 tocEntry；
 * 返回越过的最后一个下标 + 1（即切片 end）。
 *
 * 若 nodes[i] 没有 __origSdtId（例如前端新插入的纯 tocEntry），
 * 返回 i（表示不构成组），让主循环走单节点分类。
 */
function findTocGroupEnd(nodes: TiptapNode[], i: number): number {
  const sdtId = getTocSdtId(nodes[i])
  if (!sdtId) return i
  let j = i + 1
  while (
    j < nodes.length &&
    nodes[j].type === 'tocEntry' &&
    getTocSdtId(nodes[j]) === sdtId
  ) {
    j++
  }
  return j
}

/**
 * 整组 tocEntry（共享同一 SDT）判定：
 *   - 首项必须带完整 __origRange + __origPart + __origContentFp
 *   - 所有项的 contentFp 都与 __origContentFp 一致
 *   - 任何一条不满足 → 整组 dirty（导出时整块 SDT 走 regenerate 路径）
 */
function classifyTocGroup(
  group: TiptapNode[],
  partPath: string,
  reasons: ClassifierStats['dirtyReasons'],
): { clean: true; range: ParsedRange } | { clean: false } {
  if (group.length === 0) return { clean: false }
  const firstAttrs = (group[0].attrs ?? {}) as Record<string, unknown>
  const range = parseOrigRange(firstAttrs.__origRange)
  if (!range) {
    reasons.missingRange++
    return { clean: false }
  }
  const origPart =
    typeof firstAttrs.__origPart === 'string' ? firstAttrs.__origPart : DEFAULT_PART
  if (origPart !== partPath) {
    reasons.partMismatch++
    return { clean: false }
  }

  for (const n of group) {
    const a = (n.attrs ?? {}) as Record<string, unknown>
    const storedFp = a.__origContentFp
    if (typeof storedFp !== 'string') {
      reasons.missingFp++
      return { clean: false }
    }
    if (computeContentFingerprint(n) !== storedFp) {
      reasons.fpMismatch++
      return { clean: false }
    }
  }

  return { clean: true, range }
}

/**
 * 主入口：把节点数组划分为 segments[]。
 */
export function classifyNodes(
  nodes: TiptapNode[],
  options?: ClassifierOptions,
): ClassifierResult {
  const partPath = options?.partPath ?? DEFAULT_PART
  const segments: Segment[] = []
  const reasons: ClassifierStats['dirtyReasons'] = {
    missingRange: 0,
    missingFp: 0,
    partMismatch: 0,
    fpMismatch: 0,
    wrapper: 0,
  }

  let reuseNodes = 0
  let regenerateNodes = 0
  let reuseBytes = 0

  for (let i = 0; i < nodes.length; i++) {
    const node = nodes[i]

    // SDT/TOC 原子块：连续共享同一 __origSdtId 的 tocEntry 作为整组处理。
    // 导入时只给首项挂 __origRange + __origSdtXml，其余仅带 __origSdtId，
    // 必须识别为一个整体，否则后续 tocEntry 会因缺 range 被误判 dirty。
    if (node.type === 'tocEntry') {
      const groupEnd = findTocGroupEnd(nodes, i)
      if (groupEnd > i) {
        const group = nodes.slice(i, groupEnd)
        const groupVerdict = classifyTocGroup(group, partPath, reasons)
        if (groupVerdict.clean) {
          reuseNodes += group.length
          reuseBytes += groupVerdict.range.end - groupVerdict.range.start
          const last = segments[segments.length - 1]
          if (
            last &&
            last.kind === 'reuse' &&
            last.partPath === partPath &&
            last.end === groupVerdict.range.start
          ) {
            last.end = groupVerdict.range.end
            last.nodes.push(...group)
          } else {
            segments.push({
              kind: 'reuse',
              start: groupVerdict.range.start,
              end: groupVerdict.range.end,
              partPath,
              nodes: group,
            })
          }
        } else {
          regenerateNodes += group.length
          const last = segments[segments.length - 1]
          if (last && last.kind === 'regenerate') {
            last.nodes.push(...group)
          } else {
            segments.push({ kind: 'regenerate', nodes: [...group] })
          }
        }
        i = groupEnd - 1
        continue
      }
    }

    const verdict = classifyOne(node, partPath, reasons)

    if (verdict.clean) {
      reuseNodes++
      reuseBytes += verdict.range.end - verdict.range.start
      const last = segments[segments.length - 1]
      if (
        last &&
        last.kind === 'reuse' &&
        last.partPath === partPath &&
        last.end === verdict.range.start
      ) {
        // 字节邻接，合并
        last.end = verdict.range.end
        last.nodes.push(node)
      } else {
        segments.push({
          kind: 'reuse',
          start: verdict.range.start,
          end: verdict.range.end,
          partPath,
          nodes: [node],
        })
      }
    } else {
      regenerateNodes++
      const last = segments[segments.length - 1]
      if (last && last.kind === 'regenerate') {
        last.nodes.push(node)
      } else {
        segments.push({ kind: 'regenerate', nodes: [node] })
      }
    }
  }

  const reuseSegments = segments.filter((s) => s.kind === 'reuse').length
  const regenerateSegments = segments.filter((s) => s.kind === 'regenerate').length

  return {
    segments,
    stats: {
      totalNodes: nodes.length,
      reuseNodes,
      regenerateNodes,
      reuseBytes,
      reuseSegments,
      regenerateSegments,
      dirtyReasons: reasons,
    },
  }
}

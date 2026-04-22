/**
 * 章节编号计数器（Heading Numbering Counter）。
 *
 * 背景：
 *   Word 的标题样式（Heading 1/2/3）通常在 styles.xml 里挂了一个多级列表
 *   （w:numPr → numId=N, ilvl=0..8），用于自动生成 "1"、"1.2"、"1.2.1" 这类
 *   章节编号。章节编号不是文档正文的一部分，只是标题的装饰显示。
 *
 *   选择性保存方案对原 DOCX 的 numbering.xml 做字节复用，所以导出时不用
 *   重算；但前端编辑器需要直观显示这些编号，否则用户看不到 "1.2 产品应用"
 *   只看到 "产品应用"。
 *
 * 本模块：
 *   - 基于 numbering.xml（已由 numberingResolver 解析）维护一个全局的
 *     `counter[numId] = number[]`（每个 numId 对应一个 9 级计数器）。
 *   - 按文档顺序逐个 heading 调用 `advance(numId, ilvl)`：
 *       · counter[numId][ilvl] += 1（首次调用取 level.start 作为初值）
 *       · counter[numId][ilvl+1..8] = 0（清零更深层）
 *       · 根据该 level 的 lvlText（如 "%1.%2.%3"）把 "%N" 替换为对应计数值；
 *         每个 %N 按对应 level 的 numFmt 做数字→文本转换（目前支持
 *         decimal/chineseCounting/upperRoman/lowerRoman/upperLetter/lowerLetter）。
 *   - 返回拼装好的字符串，供 paragraph.ts 写入 heading attrs.numberingText。
 *
 * 注意：
 *   - Counter 是"有副作用"的状态机。必须按文档顺序调用 `advance`，否则
 *     计数错乱。importPipeline 的主遍历本来就按文档顺序，天然满足。
 *   - 若 numId/ilvl 找不到对应 level，返回空串（不崩，不污染 heading）。
 */

import { getNumberingLevel, type NumberingResult } from './numberingResolver.js'
import type { NumberingLevel } from '../types/ooxml.js'

const MAX_LEVELS = 9

export interface HeadingNumberingCounterOptions {
  numbering: NumberingResult
}

export class HeadingNumberingCounter {
  private readonly numbering: NumberingResult
  /** numId → [count0..count8]，未初始化时整数组为 null */
  private readonly state = new Map<number, (number | null)[]>()

  constructor(options: HeadingNumberingCounterOptions) {
    this.numbering = options.numbering
  }

  /**
   * 推进 (numId, ilvl) 的计数并生成章节编号文本。
   * 找不到对应 level 时返回空串。
   */
  advance(numId: number, ilvl: number): string {
    const level = getNumberingLevel(this.numbering, numId, ilvl)
    if (!level) return ''

    const counts = this.ensureCounts(numId)

    // 当前级别 +1；首次计数使用 level.start 作为基点
    const prev = counts[ilvl]
    if (prev == null) {
      counts[ilvl] = level.start
    } else {
      counts[ilvl] = prev + 1
    }

    // 清零更深层（Word 行为：更深 level 从 start 重新开始）
    for (let i = ilvl + 1; i < MAX_LEVELS; i++) {
      counts[i] = null
    }

    return this.renderLvlText(numId, ilvl, level)
  }

  /**
   * 把 level.lvlText 里的 "%N" 替换成对应 level 的当前计数值。
   *
   *   - N 从 1 开始（%1 对应 ilvl=0，%2 对应 ilvl=1，依此类推）
   *   - 若对应层 counts 为 null（父层未初始化，极少见），用 level.start 兜底
   */
  private renderLvlText(numId: number, ilvl: number, level: NumberingLevel): string {
    const lvlText = level.lvlText || ''
    if (!lvlText) return ''
    const counts = this.state.get(numId) ?? []

    return lvlText.replace(/%(\d)/g, (_, digit: string) => {
      const n = Number(digit)
      if (!Number.isFinite(n) || n < 1 || n > MAX_LEVELS) return ''
      const targetLvl = n - 1
      // 取目标层级在该 numId 下的 numFmt（不同 level 可能有不同 numFmt）
      const targetLevel =
        targetLvl === ilvl ? level : getNumberingLevel(this.numbering, numId, targetLvl)
      const count = counts[targetLvl] ?? targetLevel?.start ?? 1
      const fmt = targetLevel?.numFmt ?? 'decimal'
      return formatCount(count, fmt)
    })
  }

  private ensureCounts(numId: number): (number | null)[] {
    let arr = this.state.get(numId)
    if (!arr) {
      arr = new Array<number | null>(MAX_LEVELS).fill(null)
      this.state.set(numId, arr)
    }
    return arr
  }
}

/**
 * 把计数值按 numFmt 转成显示字符串。
 * 目前支持业务里最常见的几种；其它未识别的 numFmt 按 decimal 兜底。
 * bullet / none 不应出现在标题编号里，保底返回空串。
 */
function formatCount(count: number, numFmt: string): string {
  switch (numFmt) {
    case 'decimal':
    case 'decimalZero':
      return numFmt === 'decimalZero' && count < 10 ? `0${count}` : String(count)
    case 'upperLetter':
      return toAlpha(count, true)
    case 'lowerLetter':
      return toAlpha(count, false)
    case 'upperRoman':
      return toRoman(count).toUpperCase()
    case 'lowerRoman':
      return toRoman(count).toLowerCase()
    case 'chineseCounting':
    case 'chineseCountingThousand':
    case 'ideographDigital':
    case 'ideographTraditional':
      return toChineseCounting(count)
    case 'chineseLegalSimplified':
      return toChineseLegal(count)
    case 'bullet':
    case 'none':
      return ''
    default:
      return String(count)
  }
}

// A, B, ..., Z, AA, AB, ...
function toAlpha(n: number, upper: boolean): string {
  if (n <= 0) return ''
  const base = upper ? 'A'.charCodeAt(0) : 'a'.charCodeAt(0)
  let s = ''
  let x = n
  while (x > 0) {
    x--
    s = String.fromCharCode(base + (x % 26)) + s
    x = Math.floor(x / 26)
  }
  return s
}

function toRoman(n: number): string {
  if (n <= 0) return String(n)
  const table: [number, string][] = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'],
    [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'],
    [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
  ]
  let out = ''
  let x = n
  for (const [v, s] of table) {
    while (x >= v) {
      out += s
      x -= v
    }
  }
  return out
}

// 中文数字（一二三... 十一...）
function toChineseCounting(n: number): string {
  if (n <= 0) return String(n)
  const digits = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九']
  if (n < 10) return digits[n]
  if (n < 20) return n === 10 ? '十' : `十${digits[n - 10]}`
  if (n < 100) {
    const tens = digits[Math.floor(n / 10)]
    const ones = n % 10
    return ones === 0 ? `${tens}十` : `${tens}十${digits[ones]}`
  }
  // 简化：100 以上回退 decimal，避免实现过多
  return String(n)
}

// 大写中文数字（壹贰叁... 拾壹...）
function toChineseLegal(n: number): string {
  if (n <= 0) return String(n)
  const digits = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖']
  if (n < 10) return digits[n]
  if (n < 20) return n === 10 ? '拾' : `拾${digits[n - 10]}`
  if (n < 100) {
    const tens = digits[Math.floor(n / 10)]
    const ones = n % 10
    return ones === 0 ? `${tens}拾` : `${tens}拾${digits[ones]}`
  }
  return String(n)
}

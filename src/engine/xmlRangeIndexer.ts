/**
 * 选择性保存高保真方案：字节级 XML 顶层元素范围索引器。
 *
 * 目标：
 *   给定 word/document.xml（或 header/footer 等）原始字节，在**不做完整 XML 解析**
 *   的前提下，找到 `w:body`（或等价的顶层容器）内**每个顶层块**的字节区间 [start, end)。
 *
 * 为什么不直接用 fast-xml-parser？
 *   - 解析器会丢失字节偏移信息（它给出的是 JS 对象树，属性顺序/空白全部归一化）。
 *   - 解析成本是 O(n) 但常数大（构造对象），而我们真正要的只是顶层元素定位。
 *   - 一旦落到字符串层面，多字节字符（UTF-8 3/4 字节）与字节下标的对应关系需要
 *     额外记账；直接在字节数组上扫描可避免这个坑。
 *
 * 支持的顶层容器：w:body（document.xml）、w:hdr（headerN.xml）、w:ftr（footerN.xml）、
 *                  w:footnotes / w:endnotes（脚注/尾注）。
 * 支持的顶层元素：w:p、w:tbl、w:sdt、w:sectPr。
 *
 * 字节安全性说明：
 *   - `<` / `>` 在 UTF-8 里恒为单字节 0x3C / 0x3E，且在 XML 规范里不能出现在
 *     属性值或文本中（必须转义成 &lt; / &gt;），因此直接按字节扫描 `<` 是安全的。
 *   - XML 注释 `<!-- ... -->` 与 CDATA `<![CDATA[ ... ]]>` 内部可能包含 `<`/`>`，
 *     扫描器必须整块跳过。DOCX 几乎不用这两种结构，但为稳健起见仍做处理。
 *   - Processing Instruction `<? ... ?>` 与 DOCTYPE `<! ... >` 也按整体跳过。
 */

export type TopLevelTag = 'w:p' | 'w:tbl' | 'w:sdt' | 'w:sectPr'

export interface TopLevelRange {
  /** 节点的局部标签名，带命名空间前缀 */
  tag: TopLevelTag
  /** 起始字节偏移（含），指向 `<` 所在位置 */
  start: number
  /** 结束字节偏移（不含），指向 `>` 之后的下一个字节 */
  end: number
}

export interface RangeIndex {
  /** 容器标签，如 "w:body" / "w:hdr" / "w:ftr" / "w:footnotes" / "w:endnotes" */
  containerTag: string
  /** 容器开始标签的结束偏移（即容器内内容起点） */
  containerContentStart: number
  /** 容器结束标签的开始偏移（即容器内内容终点） */
  containerContentEnd: number
  /** 按文档顺序排列的顶层元素字节范围 */
  ranges: TopLevelRange[]
}

const CHAR_LT = 0x3c // '<'
const CHAR_GT = 0x3e // '>'
const CHAR_SLASH = 0x2f // '/'
const CHAR_BANG = 0x21 // '!'
const CHAR_QUESTION = 0x3f // '?'
const CHAR_DASH = 0x2d // '-'
const CHAR_LBRACK = 0x5b // '['

const TEXT_ENCODER = new TextEncoder()

/** 支持的容器标签集合 */
const SUPPORTED_CONTAINERS = ['w:body', 'w:hdr', 'w:ftr', 'w:footnotes', 'w:endnotes'] as const
type SupportedContainer = (typeof SUPPORTED_CONTAINERS)[number]

/** 支持的顶层元素集合 */
const SUPPORTED_TOP_LEVEL: TopLevelTag[] = ['w:p', 'w:tbl', 'w:sdt', 'w:sectPr']

/**
 * 扫描部件字节，产出顶层元素范围列表。
 *
 * @param bytes 部件原始字节（UTF-8 编码）
 * @param containerTag 容器标签，如 "w:body"；省略时自动识别（按 SUPPORTED_CONTAINERS 优先级）
 */
export function indexTopLevelRanges(
  bytes: Uint8Array,
  containerTag?: SupportedContainer,
): RangeIndex | null {
  const container = containerTag ?? autoDetectContainer(bytes)
  if (!container) return null

  const openMarker = TEXT_ENCODER.encode(`<${container}`)
  const closeMarker = TEXT_ENCODER.encode(`</${container}>`)

  const openStart = findSubsequence(bytes, openMarker, 0)
  if (openStart < 0) return null
  // 容器可能以 `<w:body>` 或 `<w:body xmlns:...>` 起始，找到对应 `>` 即可。
  const openTagEnd = findByteAfter(bytes, CHAR_GT, openStart)
  if (openTagEnd < 0) return null
  // 处理自闭合容器（理论上不该出现在 body/hdr/ftr，但脚注可能）：
  if (bytes[openTagEnd - 1] === CHAR_SLASH) {
    return {
      containerTag: container,
      containerContentStart: openTagEnd + 1,
      containerContentEnd: openTagEnd + 1,
      ranges: [],
    }
  }
  const contentStart = openTagEnd + 1

  const closeStart = findSubsequence(bytes, closeMarker, contentStart)
  if (closeStart < 0) return null
  const contentEnd = closeStart

  const ranges = scanTopLevelElements(bytes, contentStart, contentEnd)

  return {
    containerTag: container,
    containerContentStart: contentStart,
    containerContentEnd: contentEnd,
    ranges,
  }
}

/**
 * 读取 TopLevelRange 对应的字节子串。
 * 返回 subarray 视图；调用方不应修改内容。
 */
export function sliceRange(bytes: Uint8Array, range: TopLevelRange): Uint8Array {
  return bytes.subarray(range.start, range.end)
}

// -------------------------------------------------------------------
// 下方为内部扫描实现
// -------------------------------------------------------------------

function autoDetectContainer(bytes: Uint8Array): SupportedContainer | null {
  for (const tag of SUPPORTED_CONTAINERS) {
    const marker = TEXT_ENCODER.encode(`<${tag}`)
    const idx = findSubsequence(bytes, marker, 0)
    if (idx >= 0) {
      // 验证 `<tag` 后紧跟的是空格 / > / 换行，避免把 <w:bodyText> 之类误识别。
      const nextCharIdx = idx + marker.length
      if (nextCharIdx < bytes.length) {
        const next = bytes[nextCharIdx]
        if (next === CHAR_GT || next === 0x20 || next === 0x09 || next === 0x0a || next === 0x0d || next === CHAR_SLASH) {
          return tag
        }
      }
    }
  }
  return null
}

/**
 * 在 [contentStart, contentEnd) 范围内按层级扫描，挑出位于容器**直接子级**的
 * w:p / w:tbl / w:sdt / w:sectPr。遇到其他非支持的顶层元素（如 w:customXmlInsRangeStart）
 * 会跳过但不报错（这些元素当前不参与选择性保存，仍交给 handler 处理）。
 */
function scanTopLevelElements(
  bytes: Uint8Array,
  contentStart: number,
  contentEnd: number,
): TopLevelRange[] {
  const ranges: TopLevelRange[] = []
  let depth = 0
  let i = contentStart
  // 仅当 depth === 0 时遇到的开始标签才是顶层元素；
  // 在其打开后，记录 start，递归直到深度回到 0。
  let currentStart = -1
  let currentTag: TopLevelTag | null = null

  while (i < contentEnd) {
    const ch = bytes[i]
    if (ch !== CHAR_LT) {
      i++
      continue
    }
    // 快速跳过注释 / CDATA / PI / DOCTYPE
    if (i + 3 < contentEnd && bytes[i + 1] === CHAR_BANG && bytes[i + 2] === CHAR_DASH && bytes[i + 3] === CHAR_DASH) {
      const end = findSubsequence(bytes, CLOSE_COMMENT, i + 4)
      i = end < 0 ? contentEnd : end + CLOSE_COMMENT.length
      continue
    }
    if (i + 8 < contentEnd && bytes[i + 1] === CHAR_BANG && bytes[i + 2] === CHAR_LBRACK) {
      // <![CDATA[
      const end = findSubsequence(bytes, CLOSE_CDATA, i + 9)
      i = end < 0 ? contentEnd : end + CLOSE_CDATA.length
      continue
    }
    if (bytes[i + 1] === CHAR_QUESTION) {
      // <?xml ... ?>
      const end = findSubsequence(bytes, CLOSE_PI, i + 2)
      i = end < 0 ? contentEnd : end + CLOSE_PI.length
      continue
    }
    if (bytes[i + 1] === CHAR_BANG) {
      // <!DOCTYPE ... >
      const end = findByteAfter(bytes, CHAR_GT, i + 2)
      i = end < 0 ? contentEnd : end + 1
      continue
    }

    const isClose = bytes[i + 1] === CHAR_SLASH
    const tagEnd = findByteAfter(bytes, CHAR_GT, i)
    if (tagEnd < 0 || tagEnd >= contentEnd) break

    const isSelfClosing = bytes[tagEnd - 1] === CHAR_SLASH

    if (isClose) {
      depth--
      if (depth === 0 && currentStart >= 0 && currentTag) {
        ranges.push({ tag: currentTag, start: currentStart, end: tagEnd + 1 })
        currentStart = -1
        currentTag = null
      }
      i = tagEnd + 1
      continue
    }

    // 读取开始标签名
    const name = readTagName(bytes, i + 1, tagEnd)

    if (depth === 0) {
      const knownTag = matchTopLevel(name)
      if (knownTag) {
        currentStart = i
        currentTag = knownTag
        if (isSelfClosing) {
          ranges.push({ tag: knownTag, start: i, end: tagEnd + 1 })
          currentStart = -1
          currentTag = null
          i = tagEnd + 1
          continue
        }
      } else {
        // 未识别的顶层元素：不参与范围索引，但仍需维护 depth 以正确跳过它的内部。
        currentStart = -1
        currentTag = null
      }
    }

    if (!isSelfClosing) depth++
    i = tagEnd + 1
  }

  return ranges
}

/** 从 bytes[start..end) 中读取开始标签的名字（遇到空格 / `>` / `/` 停止） */
function readTagName(bytes: Uint8Array, start: number, end: number): string {
  let j = start
  while (j < end) {
    const c = bytes[j]
    if (c === CHAR_GT || c === CHAR_SLASH || c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d) {
      break
    }
    j++
  }
  return new TextDecoder('utf-8').decode(bytes.subarray(start, j))
}

function matchTopLevel(name: string): TopLevelTag | null {
  for (const tag of SUPPORTED_TOP_LEVEL) {
    if (name === tag) return tag
  }
  return null
}

// 预分配关闭标记，减少运行时字符串 → 字节的重复编码。
const CLOSE_COMMENT = Uint8Array.of(CHAR_DASH, CHAR_DASH, CHAR_GT) // -->
const CLOSE_CDATA = Uint8Array.of(0x5d, 0x5d, CHAR_GT) // ]]>
const CLOSE_PI = Uint8Array.of(CHAR_QUESTION, CHAR_GT) // ?>

/** 从 bytes 中找到 needle 首次出现的位置（≥ from），找不到返回 -1。 */
function findSubsequence(bytes: Uint8Array, needle: Uint8Array, from: number): number {
  const end = bytes.length - needle.length
  outer: for (let i = from; i <= end; i++) {
    for (let j = 0; j < needle.length; j++) {
      if (bytes[i + j] !== needle[j]) continue outer
    }
    return i
  }
  return -1
}

/** 从 bytes[from..] 中找到首个等于 target 的字节位置，找不到返回 -1。 */
function findByteAfter(bytes: Uint8Array, target: number, from: number): number {
  for (let i = from; i < bytes.length; i++) {
    if (bytes[i] === target) return i
  }
  return -1
}

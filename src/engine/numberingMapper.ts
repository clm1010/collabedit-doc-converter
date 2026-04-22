/**
 * 选择性保存高保真方案：numbering.xml 复用 / 兜底 patch。
 *
 * 背景：
 *   - 原 DOCX 的 word/numbering.xml 在选择性保存中保持字节不变，
 *     但 regenerate 段里由 docx 库新生成的 <w:numId w:val="N"/> 引用的
 *     是 docx 库自己分配的临时 numId（通常 1=bullet-list, 2=ordered-list），
 *     这些 numId 与原 numbering.xml 的语义完全不同，直接写回会导致编号错乱。
 *
 *   - 对策：在原 numbering.xml 中寻找一个"兼容的" numId（bullet 型或
 *     有序型），把 regen fragment 里的临时 numId 替换为原档真实的 numId。
 *     这样用户只是"改了列表里一段文字"时，新段落仍然用原档那个列表的编号
 *     体系，视觉保持一致。
 *
 *   - 若原档找不到兼容 numId（极端情况：原档根本没列表，用户新建了列表）：
 *     暂时保留 docx 库的临时 numId + abstractNum；阶段 5 兜底实现是
 *     直接把 temp numbering.xml 里相关 abstractNum/num 的 XML 片段
 *     追加到原 numbering.xml 的 `</w:numbering>` 前。
 *
 * 本模块职责：
 *   1. 构造时解析原 numbering.xml，产出 bulletNumIds / orderedNumIds。
 *   2. 接受一个 temp docx（localSerializer 临时打包产物）的 numbering.xml，
 *      识别 docx 库给 "bullet-list"/"ordered-list" reference 分配的 numId。
 *   3. 对外暴露 `getReplacementMap()` → Map<tempNumId, realNumId>。
 *   4. 兜底 `patchNumbering(originalBytes, ...)`：阶段 5 暂仅追加 abstractNum/num
 *      未实现，保留接口。
 */

const TEXT_DECODER = new TextDecoder('utf-8')

export class NumberingMapper {
  /** 原档 bullet 型 numId 候选（按优先级顺序） */
  private readonly bulletNumIds: number[] = []
  /** 原档有序型 numId 候选（decimal/upperLetter/...） */
  private readonly orderedNumIds: number[] = []
  /** 标志原 numbering.xml 是否可用 */
  readonly available: boolean

  constructor(originalNumberingBytes?: Uint8Array) {
    this.available = !!originalNumberingBytes
    if (!originalNumberingBytes) return
    const xml = TEXT_DECODER.decode(originalNumberingBytes)
    this.parseOriginal(xml)
  }

  /**
   * 根据一个 temp numbering.xml（docx 库临时生成），返回
   * Map<tempNumId, realNumId>，供 localSerializer 做字符串替换。
   */
  buildReplacement(tempNumberingBytes?: Uint8Array): Map<string, string> {
    const out = new Map<string, string>()
    if (!tempNumberingBytes || !this.available) return out
    const xml = TEXT_DECODER.decode(tempNumberingBytes)

    // docx 库把每个 reference 编号为一个 num + abstractNum。
    // 抽出 num 元素：<w:num w:numId="1"> ... <w:abstractNumId w:val="0"/> ... </w:num>
    const nums = parseNumElements(xml)
    const absClasses = parseAbstractNumClasses(xml)

    for (const n of nums) {
      const cls = absClasses.get(n.abstractNumId)
      if (cls === 'bullet' && this.bulletNumIds.length > 0) {
        out.set(String(n.numId), String(this.bulletNumIds[0]))
      } else if (cls === 'ordered' && this.orderedNumIds.length > 0) {
        out.set(String(n.numId), String(this.orderedNumIds[0]))
      }
    }
    return out
  }

  private parseOriginal(xml: string): void {
    const nums = parseNumElements(xml)
    const absClasses = parseAbstractNumClasses(xml)
    for (const n of nums) {
      const cls = absClasses.get(n.abstractNumId)
      if (cls === 'bullet') this.bulletNumIds.push(n.numId)
      else if (cls === 'ordered') this.orderedNumIds.push(n.numId)
    }
  }
}

interface NumMapping {
  numId: number
  abstractNumId: number
}

/**
 * 解析所有 <w:num w:numId="...">...<w:abstractNumId w:val="..."/>...</w:num>
 * 容忍属性顺序 / 自闭合等变种。
 */
function parseNumElements(xml: string): NumMapping[] {
  const out: NumMapping[] = []
  const rx = /<w:num\b[^>]*\bw:numId="(\d+)"[^>]*>([\s\S]*?)<\/w:num>/g
  let m: RegExpExecArray | null
  while ((m = rx.exec(xml)) !== null) {
    const numId = Number(m[1])
    const body = m[2]
    const absMatch = /<w:abstractNumId\b[^>]*\bw:val="(\d+)"/.exec(body)
    if (!absMatch) continue
    out.push({ numId, abstractNumId: Number(absMatch[1]) })
  }
  return out
}

/**
 * 解析每个 abstractNumId 的"大类"：bullet / ordered。
 *
 * 规则：取该 abstractNum 下首个 `<w:lvl>` 的 `<w:numFmt w:val="..."/>`：
 *   - "bullet" → bullet
 *   - 其它（decimal / decimalZero / upperLetter / lowerLetter / upperRoman / lowerRoman 等）→ ordered
 * 若 numFmt 缺失或 "none"，跳过不分类。
 */
function parseAbstractNumClasses(xml: string): Map<number, 'bullet' | 'ordered'> {
  const out = new Map<number, 'bullet' | 'ordered'>()
  const rx = /<w:abstractNum\b[^>]*\bw:abstractNumId="(\d+)"[^>]*>([\s\S]*?)<\/w:abstractNum>/g
  let m: RegExpExecArray | null
  while ((m = rx.exec(xml)) !== null) {
    const aid = Number(m[1])
    const body = m[2]
    const lvlMatch = /<w:lvl\b[^>]*>([\s\S]*?)<\/w:lvl>/.exec(body)
    const lvlBody = lvlMatch ? lvlMatch[1] : ''
    const fmtMatch = /<w:numFmt\b[^>]*\bw:val="([^"]+)"/.exec(lvlBody)
    const fmt = fmtMatch ? fmtMatch[1].toLowerCase() : ''
    if (!fmt || fmt === 'none') continue
    if (fmt === 'bullet') out.set(aid, 'bullet')
    else out.set(aid, 'ordered')
  }
  return out
}

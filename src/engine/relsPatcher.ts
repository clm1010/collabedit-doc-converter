/**
 * 选择性保存高保真方案：rels XML 字符串级 Patch。
 *
 * 策略：保留原 rels 文件中所有原始字节（含注释、属性顺序、自定义命名空间），
 * 仅在末尾 `</Relationships>` 之前追加新 `<Relationship .../>` 条目。
 *
 * 当原 rels 不存在（部件本身没有 rels 文件）时，构造一个最小化的合法 rels XML。
 *
 * 不变量：
 *   - 原条目字节零修改。Word 对此非常敏感（rId 顺序、Target 大小写都会影响渲染）。
 *   - 追加的条目放在最后，xmlns 用标准 OOXML rels 命名空间。
 */

import type { PendingRel } from './imageRefMapper.js'

const TEXT_DECODER = new TextDecoder('utf-8')
const TEXT_ENCODER = new TextEncoder()

const RELS_HEADER =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
const RELS_FOOTER = '</Relationships>'

export interface PatchRelsInput {
  /** 原 rels 文件字节，可为 undefined（部件无 rels） */
  originalBytes?: Uint8Array
  /** 该部件待追加的条目 */
  newRels: PendingRel[]
}

export interface PatchRelsOutput {
  bytes: Uint8Array
  appendedCount: number
}

export function patchRelsXml(input: PatchRelsInput): PatchRelsOutput {
  const { originalBytes, newRels } = input

  if (newRels.length === 0) {
    // 不变就把原字节原样返回（exportPipeline 不会调到这里，这里做防御）
    return {
      bytes: originalBytes ?? TEXT_ENCODER.encode(`${RELS_HEADER}${RELS_FOOTER}\n`),
      appendedCount: 0,
    }
  }

  let baseXml: string
  if (originalBytes && originalBytes.length > 0) {
    baseXml = TEXT_DECODER.decode(originalBytes)
  } else {
    baseXml = `${RELS_HEADER}${RELS_FOOTER}\n`
  }

  // 找最后一个 </Relationships> 的位置，在它之前插入
  const closeIdx = baseXml.lastIndexOf(RELS_FOOTER)
  if (closeIdx < 0) {
    throw new Error(
      `relsPatcher: original rels XML missing ${RELS_FOOTER}; ` +
        `len=${baseXml.length}`,
    )
  }

  const head = baseXml.slice(0, closeIdx)
  const tail = baseXml.slice(closeIdx)

  const additions = newRels
    .map(
      (r) =>
        `<Relationship Id="${escapeAttr(r.id)}" Type="${escapeAttr(r.type)}" Target="${escapeAttr(r.target)}"/>`,
    )
    .join('')

  const out = `${head}${additions}${tail}`
  return {
    bytes: TEXT_ENCODER.encode(out),
    appendedCount: newRels.length,
  }
}

function escapeAttr(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

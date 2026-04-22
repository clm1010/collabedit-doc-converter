/**
 * 选择性保存高保真方案：[Content_Types].xml 字符串级 Patch。
 *
 * 仅在新增图片扩展名（如新增了 .webp 而原文档无）时追加 `<Default ContentType=... Extension=.../>`。
 * 已有扩展名（jpeg/jpg/png/gif/bmp 等大多数 DOCX 已声明）不追加，避免破坏原顺序。
 *
 * 不变量：
 *   - 原条目字节零修改。
 *   - 追加的 Default 条目放在最后一个 Default / 第一个 Override 之前；
 *     若没有 Override 也放在 `</Types>` 之前。
 *   - 大小写：扩展名一律小写匹配（与 DOCX 标准一致）。
 */

import type { PendingContentType } from './imageRefMapper.js'

const TEXT_DECODER = new TextDecoder('utf-8')
const TEXT_ENCODER = new TextEncoder()

export interface PatchContentTypesInput {
  originalBytes: Uint8Array
  newDefaults: PendingContentType[]
}

export interface PatchContentTypesOutput {
  bytes: Uint8Array
  appendedCount: number
}

export function patchContentTypesXml(
  input: PatchContentTypesInput,
): PatchContentTypesOutput {
  const { originalBytes, newDefaults } = input
  if (newDefaults.length === 0) {
    return { bytes: originalBytes, appendedCount: 0 }
  }

  const xml = TEXT_DECODER.decode(originalBytes)

  // 解析现有 Default Extension 集（小写）
  const existingExts = new Set<string>()
  // ContentType 属性值里含 '/'（例如 image/png），不能用 [^/]；
  // 用 [^>]*? + 显式 /> 结尾。
  const rxDefault = /<Default\b[^>]*?\/>/g
  const rxExt = /\bExtension="([^"]*)"/
  let m: RegExpExecArray | null
  while ((m = rxDefault.exec(xml)) !== null) {
    const em = rxExt.exec(m[0])
    if (em) existingExts.add(em[1].toLowerCase())
  }

  const toAdd = newDefaults.filter((d) => !existingExts.has(d.extension.toLowerCase()))
  if (toAdd.length === 0) {
    return { bytes: originalBytes, appendedCount: 0 }
  }

  const additions = toAdd
    .map(
      (d) =>
        `<Default Extension="${escapeAttr(d.extension)}" ContentType="${escapeAttr(d.contentType)}"/>`,
    )
    .join('')

  // 优先插在第一个 <Override 之前，否则插在 </Types> 之前
  const overrideIdx = xml.indexOf('<Override')
  let insertIdx: number
  if (overrideIdx >= 0) {
    insertIdx = overrideIdx
  } else {
    insertIdx = xml.lastIndexOf('</Types>')
    if (insertIdx < 0) {
      throw new Error('contentTypesPatcher: missing </Types> in [Content_Types].xml')
    }
  }

  const out = xml.slice(0, insertIdx) + additions + xml.slice(insertIdx)
  return {
    bytes: TEXT_ENCODER.encode(out),
    appendedCount: toAdd.length,
  }
}

function escapeAttr(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

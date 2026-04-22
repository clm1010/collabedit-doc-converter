import { unzipSync, type Unzipped } from 'fflate'

/**
 * 选择性保存高保真方案：保留原始字节的 DOCX 解包。
 *
 * 与 `src/ooxml/zipExtractor.ts` 的区别：
 *   - 旧版直接把 fflate 的输出包成 { files, getText, ... }，使用时经常触发
 *     TextDecoder 把 Uint8Array 转成 string —— 一旦走到字符串层面，
 *     字节边界（对选择性保存至关重要）就无法可靠还原。
 *   - 本模块保留每个部件的原始 Uint8Array，并记录它们的插入顺序，
 *     用于导出阶段 rezip 时按原顺序回写（虽然 ZIP 不强依赖顺序，
 *     但 Word 对 [Content_Types].xml 在前的顺序敏感，需要保持）。
 *
 * 所有字节数组都返回 subarray 视图而非拷贝，调用方不应修改内容。
 */
export interface DocxPart {
  /** ZIP 内路径，如 "word/document.xml" */
  path: string
  /** 原始字节，保持 fflate 解压产出（已按 ZIP flag 处理 deflate） */
  bytes: Uint8Array
}

export interface RawDocxArchive {
  /** 按 ZIP 条目顺序排列的所有部件 */
  parts: DocxPart[]
  /** 路径 → 字节 的快速索引（与 parts 引用同一 Uint8Array） */
  partsByPath: Map<string, Uint8Array>
  /** 原始整体字节，导出时走 legacy 降级路径可整体回退 */
  originalZip: Uint8Array
}

/**
 * 解压 DOCX 到原始字节映射。
 *
 * 兼容 Buffer 与 Uint8Array 两种输入，外层 multer 通常给 Buffer。
 */
export function extractRawDocx(buffer: Buffer | Uint8Array): RawDocxArchive {
  const data = buffer instanceof Buffer ? new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength) : buffer
  const files: Unzipped = unzipSync(data)
  const parts: DocxPart[] = []
  const partsByPath = new Map<string, Uint8Array>()
  // Object.keys 对 fflate Unzipped 的遍历顺序等价于 ZIP 条目插入顺序
  for (const path of Object.keys(files)) {
    const bytes = files[path]
    parts.push({ path, bytes })
    partsByPath.set(path, bytes)
  }
  return {
    parts,
    partsByPath,
    originalZip: data instanceof Uint8Array ? data : new Uint8Array(data),
  }
}

/**
 * 读取部件的 UTF-8 文本。此方法**仅**供解析阶段使用；
 * 选择性保存依赖原始字节，任何需要"原样字节"的代码路径都应直接用 partsByPath。
 */
export function readPartAsText(archive: RawDocxArchive, path: string): string | null {
  const bytes = archive.partsByPath.get(path)
  if (!bytes) return null
  return new TextDecoder('utf-8').decode(bytes)
}

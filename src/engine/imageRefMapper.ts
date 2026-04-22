/**
 * 选择性保存高保真方案：图片 rId 复用与新增映射器（Image Ref Mapper）。
 *
 * 职责：
 *   - 解析原 DOCX 中各部件（document/headerN/footerN/footnotes/endnotes）对应的
 *     `_rels/<part>.rels`，建立 (partPath → rels Map) 索引。
 *   - 提供查询：
 *       findRidByTarget(partPath, target) → 复用已有 rId（用户没动过的图片）
 *       hasRid(partPath, rid)             → 校验复用的 rId 仍然存在
 *   - 提供新增：
 *       allocateRid(partPath)             → 在该部件 rels 中分配下一个未使用 rId
 *       addImageRel(partPath, target, rid)→ 把新 rels 条目记到 pendingRels
 *       addMediaFile(zipPath, bytes)      → 把新 media 字节记到 pendingMedia
 *       addContentTypeDefault(ext, mime)  → 把新扩展名条目记到 pendingTypes
 *
 * 阶段 4 仅处理 image 关系（type=...image）；超链接 / 字体 / oleObject
 * 等其它 rel 类型照原样保留（regenerate 段里若引用它们，复用其 rId）。
 *
 * 设计要点：
 *   - 不修改原 archive 字节；所有"追加"都收集到 pending* 字段，
 *     由 exportPipeline 在 patch 阶段合成新的 rels / Content_Types / media。
 *   - 部件 → rels 路径映射规则：
 *       word/document.xml             → word/_rels/document.xml.rels
 *       word/header1.xml              → word/_rels/header1.xml.rels
 *       word/footnotes.xml            → word/_rels/footnotes.xml.rels
 *     即在 part 所在目录下加 `_rels/<basename>.rels`。
 */

import type { RawDocxArchive } from './zipExtractor.js'

const TEXT_DECODER = new TextDecoder('utf-8')
const IMAGE_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

export interface RelEntry {
  id: string
  /** 例如 "media/image1.png"（相对当前 rels 所属部件目录） */
  target: string
  /** 完整 rel Type URI */
  type: string
  /** 整个 <Relationship .../> 元素原字节（rezip 阶段插入新 rels 时不会用到，这里仅供调试） */
  rawXml?: string
}

export interface PendingRel {
  partPath: string
  id: string
  target: string
  type: string
}

export interface PendingMedia {
  /** ZIP 内绝对路径，如 "word/media/image_xxx.png" */
  zipPath: string
  bytes: Uint8Array
}

export interface PendingContentType {
  /** 小写扩展名（不含点） */
  extension: string
  /** MIME 类型，如 "image/png" */
  contentType: string
}

export interface ImageRefMapperOptions {
  /** 默认 false。true 时 exportPipeline 会清理无引用的孤立 media 文件 */
  cleanOrphanMedia?: boolean
}

export class ImageRefMapper {
  /** partPath → rels Map<rId, RelEntry> */
  private readonly partRels = new Map<string, Map<string, RelEntry>>()
  /** partPath → 当前最大 rId 数字部分（用于 allocateRid） */
  private readonly partMaxRid = new Map<string, number>()
  /** 待追加的 rel 条目 */
  readonly pendingRels: PendingRel[] = []
  /** 待写入的新 media 文件字节 */
  readonly pendingMedia: PendingMedia[] = []
  /** 待追加的 Content_Types Default 条目（按 extension 去重） */
  private readonly pendingTypes = new Map<string, PendingContentType>()

  constructor(
    private readonly archive: RawDocxArchive,
    readonly options: ImageRefMapperOptions = {},
  ) {
    void this.options.cleanOrphanMedia // phase 4: 当前默认保留所有 media，env 控制留作扩展点
  }

  /** 懒加载并返回某个部件的 rels 索引；不存在返回空 Map */
  getPartRels(partPath: string): Map<string, RelEntry> {
    const cached = this.partRels.get(partPath)
    if (cached) return cached
    const relsPath = relsPathFor(partPath)
    const bytes = this.archive.partsByPath.get(relsPath)
    const map = new Map<string, RelEntry>()
    let max = 0
    if (bytes) {
      const xml = TEXT_DECODER.decode(bytes)
      // 用正则 + 字符串扫描，避免 fast-xml-parser 重序列化时把 attribute 顺序打乱
      // 属性值里可能出现 '/'（Type URL 含 http://...），不能用 [^/]；
      // 用 [^>]*? + 显式 /> 结尾，靠非贪婪从左到右匹配。
      const rxRel = /<Relationship\b([^>]*?)\/>/g
      let m: RegExpExecArray | null
      while ((m = rxRel.exec(xml)) !== null) {
        const attrs = m[1]
        const id = attrOf(attrs, 'Id')
        const target = attrOf(attrs, 'Target')
        const type = attrOf(attrs, 'Type') ?? ''
        if (id && target) {
          map.set(id, { id, target, type, rawXml: m[0] })
          const num = parseRidNumber(id)
          if (num > max) max = num
        }
      }
    }
    this.partRels.set(partPath, map)
    this.partMaxRid.set(partPath, max)
    return map
  }

  /** 在指定部件 rels 中按 target 查找已有 rId（图片 type 优先） */
  findRidByTarget(partPath: string, target: string): string | null {
    const rels = this.getPartRels(partPath)
    for (const [id, e] of rels) {
      if (e.target === target) return id
    }
    return null
  }

  /** 校验 rId 是否仍存在于该部件 rels 中（防止跨部件 rId 误用） */
  hasRid(partPath: string, rid: string): boolean {
    return this.getPartRels(partPath).has(rid)
  }

  /** 在该部件分配下一个未使用 rId（rIdN，N = 当前最大 + 1） */
  allocateRid(partPath: string): string {
    void this.getPartRels(partPath)
    const cur = this.partMaxRid.get(partPath) ?? 0
    const next = cur + 1
    this.partMaxRid.set(partPath, next)
    return `rId${next}`
  }

  /** 登记一个新的 image 类 Relationship */
  addImageRel(partPath: string, rid: string, target: string): void {
    this.pendingRels.push({
      partPath,
      id: rid,
      target,
      type: IMAGE_REL_TYPE,
    })
    // 同步到 partRels，确保后续 findRidByTarget 能感知本次新增
    this.getPartRels(partPath).set(rid, {
      id: rid,
      target,
      type: IMAGE_REL_TYPE,
    })
  }

  /** 登记一个新的 media 文件字节（zipPath 是 ZIP 内绝对路径） */
  addMediaFile(zipPath: string, bytes: Uint8Array): void {
    this.pendingMedia.push({ zipPath, bytes })
  }

  /** 登记一个新的 Content_Types Default 条目（按扩展名去重） */
  addContentTypeDefault(extension: string, contentType: string): void {
    const ext = extension.toLowerCase()
    if (this.pendingTypes.has(ext)) return
    this.pendingTypes.set(ext, { extension: ext, contentType })
  }

  /** 拿到所有待追加的 Content_Types 条目 */
  getPendingContentTypes(): PendingContentType[] {
    return Array.from(this.pendingTypes.values())
  }

  /**
   * 按部件分组返回所有 pendingRels，方便 relsPatcher 一次性 patch 一个 rels XML。
   */
  getPendingRelsByPart(): Map<string, PendingRel[]> {
    const grouped = new Map<string, PendingRel[]>()
    for (const r of this.pendingRels) {
      const list = grouped.get(r.partPath) ?? []
      list.push(r)
      grouped.set(r.partPath, list)
    }
    return grouped
  }
}

/** 计算 rels 文件的 ZIP 内路径 */
export function relsPathFor(partPath: string): string {
  const idx = partPath.lastIndexOf('/')
  const dir = idx >= 0 ? partPath.slice(0, idx) : ''
  const base = idx >= 0 ? partPath.slice(idx + 1) : partPath
  return dir ? `${dir}/_rels/${base}.rels` : `_rels/${base}.rels`
}

function attrOf(attrs: string, name: string): string | null {
  const re = new RegExp(`\\b${name}="([^"]*)"`)
  const m = re.exec(attrs)
  return m ? m[1] : null
}

function parseRidNumber(rid: string): number {
  const m = /^rId(\d+)$/i.exec(rid)
  return m ? Number(m[1]) : 0
}

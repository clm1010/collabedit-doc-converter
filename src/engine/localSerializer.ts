/**
 * 选择性保存高保真方案：本地片段序列化器（Local Serializer）。
 *
 * 把一个 regenerate segment 里的若干连续 Tiptap 节点转换成 OOXML 字节片段，
 * 供 xmlPatcher 插入到新 document.xml 的相应位置。
 *
 * 实现策略：
 *   1. 复用现有 `jsonToDocx`。把 segment.nodes 作为一个最小化 TiptapDoc 喂进去，
 *      让 docx 库生成一个完整 DOCX。
 *   2. 对该 DOCX 调 `extractRawDocx` 拿到 word/document.xml 字节。
 *   3. 用 `indexTopLevelRanges` 扫描 body，提取所有顶层元素字节（过滤 w:sectPr）。
 *   4. 顺序拼接成一个 Uint8Array，即 segment 的 XML 片段。
 *   5. （阶段 4）扫描片段里的 `r:embed="rIdN"`，按文档顺序对应到 nodes 中的
 *      image 节点：
 *        - 若节点带 `data-origin-rid` 且原 rels 仍有该 rId → 复用原 rId
 *        - 若仅带 `data-origin-target` → 在当前部件 rels 中按 target 找已有 rId
 *        - 否则 → 在 mapper 上分配新 rId，并登记新 media + 新 rels + 新 Content_Types
 *      最终替换片段里的 rId 字符串以匹配 mapper 决定的最终 rId。
 *
 * 性能：
 *   - 每个 regenerate segment 独立打包。典型编辑场景下 segment 数量很少（<10），
 *     打包本身是毫秒级，可以接受。
 *
 * 局限（由后续阶段补齐）：
 *   - 列表节点复用原 numbering 的 numId 在阶段 5。
 *   - SDT/TOC 原子级重用在阶段 5。
 */

import { jsonToDocx } from '../export/jsonToDocx.js'
import { extractRawDocx } from './zipExtractor.js'
import { indexTopLevelRanges } from './xmlRangeIndexer.js'
import type { ImageRefMapper } from './imageRefMapper.js'
import type { NumberingMapper } from './numberingMapper.js'
import type { TiptapDoc, TiptapNode } from '../types/tiptapJson.js'

const DOCUMENT_XML_PATH = 'word/document.xml'
const TEXT_DECODER = new TextDecoder('utf-8')
const TEXT_ENCODER = new TextEncoder()

export interface SerializeOptions {
  /** 该批节点所属部件（用于 image rId 复用 / 新增的 rels 归属） */
  partPath?: string
  /** 图片引用映射器；不传则走"全部新分配 rId"分支（兼容旧调用） */
  imageRefMapper?: ImageRefMapper
  /** 列表编号映射器；用于把临时 numId 换成原档兼容 numId（不传则不替换） */
  numberingMapper?: NumberingMapper
}

/**
 * 把一组连续节点序列化成 OOXML 片段字节。
 *
 * @returns 顶层 w:p / w:tbl 片段的字节连接结果（已过滤 w:sectPr，已重写图片 rId）
 * @throws 如果 jsonToDocx 生成失败或 document.xml 无法索引
 */
export async function serializeBlockNodesToXml(
  nodes: TiptapNode[],
  options: SerializeOptions = {},
): Promise<Uint8Array> {
  if (nodes.length === 0) return new Uint8Array(0)

  const doc: TiptapDoc = { type: 'doc', content: nodes }
  const { buffer, warnings } = await jsonToDocx(doc)
  void warnings

  const raw = extractRawDocx(buffer)
  const docXml = raw.partsByPath.get(DOCUMENT_XML_PATH)
  if (!docXml) {
    throw new Error(
      'localSerializer: generated docx missing word/document.xml',
    )
  }

  const index = indexTopLevelRanges(docXml, 'w:body')
  if (!index) {
    throw new Error('localSerializer: failed to index generated document.xml')
  }

  const pieces: Uint8Array[] = []
  let totalLen = 0
  for (const r of index.ranges) {
    if (r.tag === 'w:sectPr') continue
    const slice = docXml.subarray(r.start, r.end)
    pieces.push(slice)
    totalLen += slice.length
  }

  if (pieces.length === 0) return new Uint8Array(0)

  const concat = new Uint8Array(totalLen)
  let offset = 0
  for (const p of pieces) {
    concat.set(p, offset)
    offset += p.length
  }

  // 阶段 4：图片 rId 复用 / 新增
  // 显式声明为 Uint8Array（即 Uint8Array<ArrayBufferLike>），
  // 避免被推断为 Uint8Array<ArrayBuffer> 导致无法接收 rewriteXxx 的返回值（TS 5.7+）
  let out: Uint8Array = concat
  if (options.imageRefMapper && hasImageNode(nodes) && raw.partsByPath.has('word/_rels/document.xml.rels')) {
    out = rewriteImageRids(out, raw, nodes, options)
  }

  // 阶段 5：list 节点的 numId 复用
  if (options.numberingMapper && hasListOrListParagraph(nodes)) {
    out = rewriteListNumIds(out, raw, options.numberingMapper)
  }

  return out
}

/**
 * 批量入口：对一批 regenerate segments 并行生成 XML 片段。
 */
export async function serializeSegments(
  segments: { nodes: TiptapNode[] }[],
  options: SerializeOptions = {},
): Promise<Uint8Array[]> {
  return Promise.all(
    segments.map((s) => serializeBlockNodesToXml(s.nodes, options)),
  )
}

// -------------------------------------------------------------------
// 图片 rId 重写
// -------------------------------------------------------------------

function hasImageNode(nodes: TiptapNode[]): boolean {
  for (const n of nodes) {
    if (n.type === 'image') return true
    if (Array.isArray(n.content) && hasImageNode(n.content)) return true
  }
  return false
}

/** 按文档顺序收集 image 节点（与 docx 库生成顺序一致） */
function collectImageNodes(nodes: TiptapNode[]): TiptapNode[] {
  const out: TiptapNode[] = []
  const walk = (n: TiptapNode) => {
    if (n.type === 'image') out.push(n)
    if (Array.isArray(n.content)) for (const c of n.content) walk(c)
  }
  for (const n of nodes) walk(n)
  return out
}

/**
 * 在生成的临时 DOCX 中，docx 库给每张图片分配的临时 rId 出现在：
 *   - word/_rels/document.xml.rels （rId → media/imageN.xxx）
 *   - word/document.xml 里的 `<a:blip r:embed="rIdN"/>` / `<v:imagedata r:id="rIdN"/>`
 * 我们按 document.xml 中 r:embed 的出现顺序对应到 nodes 中的 image，
 * 然后决定每张图复用还是新增 rId，并把片段里的 rId 字符串替换为最终值。
 */
function rewriteImageRids(
  fragmentBytes: Uint8Array,
  rawTempDocx: ReturnType<typeof extractRawDocx>,
  nodes: TiptapNode[],
  options: SerializeOptions,
): Uint8Array {
  const partPath = options.partPath ?? DOCUMENT_XML_PATH
  const mapper = options.imageRefMapper!
  const tempRels = parseTempRels(rawTempDocx)
  const tempMedia = collectTempMedia(rawTempDocx)

  const fragmentText = TEXT_DECODER.decode(fragmentBytes)
  const tempRids = extractEmbedRids(fragmentText)
  const imageNodes = collectImageNodes(nodes)

  // 顺序对齐：tempRids[i] ↔ imageNodes[i]；如果不匹配则按"原样保留 rId"，
  // 这种情况一般不会发生，做防御性处理避免抛异常。
  const replacements = new Map<string, string>() // tempRid → finalRid
  const len = Math.min(tempRids.length, imageNodes.length)
  for (let i = 0; i < len; i++) {
    const tempRid = tempRids[i]
    const node = imageNodes[i]
    const finalRid = decideRid(tempRid, node, partPath, mapper, tempRels, tempMedia)
    if (finalRid && finalRid !== tempRid) {
      replacements.set(tempRid, finalRid)
    }
  }

  if (replacements.size === 0) return fragmentBytes

  const out = applyRidReplacements(fragmentText, replacements)
  return TEXT_ENCODER.encode(out)
}

interface TempRel {
  rid: string
  /** 在原 rels 里出现的 Target，如 "media/image1.png" */
  target: string
  type: string
}

function parseTempRels(raw: ReturnType<typeof extractRawDocx>): Map<string, TempRel> {
  const bytes = raw.partsByPath.get('word/_rels/document.xml.rels')
  const map = new Map<string, TempRel>()
  if (!bytes) return map
  const xml = TEXT_DECODER.decode(bytes)
  const rxRel = /<Relationship\b([^>]*?)\/>/g
  let m: RegExpExecArray | null
  while ((m = rxRel.exec(xml)) !== null) {
    const attrs = m[1]
    const id = matchAttr(attrs, 'Id')
    const target = matchAttr(attrs, 'Target')
    const type = matchAttr(attrs, 'Type') ?? ''
    if (id && target) map.set(id, { rid: id, target, type })
  }
  return map
}

function collectTempMedia(
  raw: ReturnType<typeof extractRawDocx>,
): Map<string, Uint8Array> {
  const out = new Map<string, Uint8Array>()
  for (const part of raw.parts) {
    if (part.path.startsWith('word/media/')) {
      out.set(part.path, part.bytes)
    }
  }
  return out
}

/** 提取 fragment 里所有 `r:embed="rIdN"` / `r:id="rIdN"`（去重，按出现顺序） */
function extractEmbedRids(xml: string): string[] {
  const seen = new Set<string>()
  const out: string[] = []
  // 既匹配 a:blip r:embed 也匹配 v:imagedata r:id；
  // 但只把 a:blip r:embed 视为"图像引用"主键，避免 v:imagedata 重复
  const rx = /<a:blip\b[^>]*?\br:embed="([^"]+)"/g
  let m: RegExpExecArray | null
  while ((m = rx.exec(xml)) !== null) {
    const rid = m[1]
    if (!seen.has(rid)) {
      seen.add(rid)
      out.push(rid)
    }
  }
  return out
}

function decideRid(
  tempRid: string,
  node: TiptapNode,
  partPath: string,
  mapper: ImageRefMapper,
  tempRels: Map<string, TempRel>,
  tempMedia: Map<string, Uint8Array>,
): string | null {
  const attrs = (node.attrs ?? {}) as Record<string, unknown>
  const originRid = (attrs['data-origin-rid'] as string | undefined) ?? null
  const originTarget = (attrs['data-origin-target'] as string | undefined) ?? null
  const originPart = (attrs['data-origin-part'] as string | undefined) ?? null

  // 优先 1：原 rId 存在且属于当前部件 → 复用
  if (originRid && (!originPart || originPart === partPath) && mapper.hasRid(partPath, originRid)) {
    return originRid
  }

  // 优先 2：按 origin target 在当前部件 rels 里反查
  if (originTarget) {
    const rid = mapper.findRidByTarget(partPath, originTarget)
    if (rid) return rid
  }

  // 优先 3：分配新 rId 并登记 media + rels + Content_Types
  const tempRel = tempRels.get(tempRid)
  if (!tempRel) {
    // docx 库未给该图片生成 rels（异常情况）；保留原 tempRid，避免抛错
    return tempRid
  }

  // tempRel.target 是 "media/imageN.xxx" 这类相对当前 rels 的路径
  const ext = inferExtension(tempRel.target)
  const newRid = mapper.allocateRid(partPath)
  // 新 media zipPath 用 tempMedia 中的字节，新文件名加 rId 前缀避免与原 media 重名
  const tempMediaPath = `word/${tempRel.target}` // word/media/imageN.xxx
  const tempBytes = tempMedia.get(tempMediaPath)
  if (!tempBytes) {
    // docx 库声明了 rels 但没把字节打包进 word/media/（理论上不会发生）
    return tempRid
  }
  const newRelTarget = `media/image_${newRid}.${ext}`
  const newZipPath = `word/${newRelTarget}`
  mapper.addMediaFile(newZipPath, tempBytes)
  mapper.addImageRel(partPath, newRid, newRelTarget)
  mapper.addContentTypeDefault(ext, mimeForExt(ext))
  return newRid
}

function applyRidReplacements(
  xml: string,
  replacements: Map<string, string>,
): string {
  // 替换 r:embed="rIdX" / r:id="rIdX"。
  // 用单次 scan + 字符串拼接，避免多次 replace 导致替换互相干扰。
  return xml.replace(
    /\b(r:embed|r:id)="([^"]+)"/g,
    (whole, attrName: string, rid: string) => {
      const next = replacements.get(rid)
      return next ? `${attrName}="${next}"` : whole
    },
  )
}

function matchAttr(attrs: string, name: string): string | null {
  const re = new RegExp(`\\b${name}="([^"]*)"`)
  const m = re.exec(attrs)
  return m ? m[1] : null
}

function inferExtension(target: string): string {
  const m = /\.([a-zA-Z0-9]+)(?:[?#].*)?$/.exec(target)
  return (m ? m[1] : 'png').toLowerCase()
}

// -------------------------------------------------------------------
// List numId 复用
// -------------------------------------------------------------------

function hasListOrListParagraph(nodes: TiptapNode[]): boolean {
  for (const n of nodes) {
    if (n.type === 'bulletList' || n.type === 'orderedList' || n.type === 'listItem') {
      return true
    }
    if (Array.isArray(n.content) && hasListOrListParagraph(n.content)) return true
  }
  return false
}

function rewriteListNumIds(
  fragmentBytes: Uint8Array,
  rawTempDocx: ReturnType<typeof extractRawDocx>,
  mapper: NumberingMapper,
): Uint8Array {
  const tempNumbering = rawTempDocx.partsByPath.get('word/numbering.xml')
  const replacement = mapper.buildReplacement(tempNumbering)
  if (replacement.size === 0) return fragmentBytes

  const xml = TEXT_DECODER.decode(fragmentBytes)
  const replaced = xml.replace(
    /<w:numId\b[^/]*\bw:val="(\d+)"\s*\/>/g,
    (whole, numId: string) => {
      const next = replacement.get(numId)
      if (!next) return whole
      return whole.replace(`w:val="${numId}"`, `w:val="${next}"`)
    },
  )
  if (replaced === xml) return fragmentBytes
  return TEXT_ENCODER.encode(replaced)
}

function mimeForExt(ext: string): string {
  const e = ext.toLowerCase()
  if (e === 'jpg' || e === 'jpeg') return 'image/jpeg'
  if (e === 'png') return 'image/png'
  if (e === 'gif') return 'image/gif'
  if (e === 'bmp') return 'image/bmp'
  if (e === 'webp') return 'image/webp'
  if (e === 'svg') return 'image/svg+xml'
  return `image/${e}`
}

import type { TiptapNode } from '../../types/tiptapJson.js'
import type { ParseContext } from '../../types/ooxml.js'
import { createNode } from '../../types/tiptapJson.js'
import { ensureArray, getAttr } from '../xmlParser.js'
import { getImageDataUrl } from '../imageExtractor.js'
import { resolveRelTarget } from '../relationships.js'
import { handleRun } from './run.js'

/** EMU → px (1 inch = 914400 EMU, 1 inch = 96px) */
function emuToPx(emu: number): number {
  return Math.round(emu / 914400 * 96)
}

/**
 * 展平文本框 w:txbxContent 中的段落为纯文本节点（行间用 hardBreak 分隔）。
 * 前端未注册 textBox 扩展，避免 ProseMirror 丢弃未知节点。
 *
 * 通过调用 handleRun() 保留 run 上的 marks（斜体、加粗、字号、颜色等），
 * 确保 logo 下方 slogan 这类文本框的视觉样式与原文档接近。
 */
function flattenTextBoxContent(
  txbxContent: Record<string, unknown>,
  ctx: ParseContext,
): TiptapNode[] {
  const out: TiptapNode[] = []
  const paragraphs = ensureArray(txbxContent['w:p'] as Record<string, unknown>[])
  for (let pi = 0; pi < paragraphs.length; pi++) {
    const p = paragraphs[pi]
    const runs = ensureArray(p['w:r'] as Record<string, unknown>[])
    for (const r of runs) {
      // 传空 RunProperties，避免继承外层段落样式（文本框样式是独立的）
      const runNodes = handleRun(r, ctx, {})
      out.push(...runNodes)
    }
    // 段落里可能还有 w:hyperlink，递归把超链接里的 run 抽出来
    const hyperlinks = ensureArray(p['w:hyperlink'] as Record<string, unknown>[])
    for (const hl of hyperlinks) {
      const hlRuns = ensureArray(hl['w:r'] as Record<string, unknown>[])
      for (const r of hlRuns) {
        out.push(...handleRun(r, ctx, {}))
      }
    }
    if (pi < paragraphs.length - 1) out.push({ type: 'hardBreak' })
  }
  return out
}

/**
 * 深度递归查找 w:txbxContent 节点。
 * Word 序列化文本框的路径并不固定，可能是：
 *   - wps:wsp > wps:txbx > w:txbxContent       （最常见：DrawingML 文本框）
 *   - wps:wsp > wps:linkedTxbx > w:txbxContent  （链接文本框）
 *   - 嵌在 mc:AlternateContent / wpg:wgp / wpc:wpc 子层里
 *   - v:shape > v:textbox > w:txbxContent       （VML 旧格式，走 handlePict 路径）
 * 这里把所有深度的 w:txbxContent 都挖出来，保证 slogan / 水印文字不被漏掉。
 */
function collectAllTxbxContent(
  node: unknown,
  acc: Record<string, unknown>[],
  visited: WeakSet<object>,
): void {
  if (!node || typeof node !== 'object') return
  if (Array.isArray(node)) {
    for (const item of node) collectAllTxbxContent(item, acc, visited)
    return
  }
  if (visited.has(node as object)) return
  visited.add(node as object)

  const obj = node as Record<string, unknown>

  // mc:AlternateContent 特殊处理：只走 Choice，不走 Fallback，避免同一文本框抽两份
  if (obj['mc:AlternateContent'] !== undefined) {
    const acs = ensureArray(obj['mc:AlternateContent'] as Record<string, unknown>[])
    for (const ac of acs) {
      const choices = ensureArray(ac['mc:Choice'] as Record<string, unknown>[])
      const fallback = ensureArray(ac['mc:Fallback'] as Record<string, unknown>[])
      const preferred = choices.length > 0 ? choices : fallback
      for (const c of preferred) collectAllTxbxContent(c, acc, visited)
    }
  }

  for (const [k, v] of Object.entries(obj)) {
    if (k.startsWith('@_') || k === '#text') continue
    if (k === 'mc:AlternateContent') continue // 已特殊处理
    if (k === 'w:txbxContent' && v && typeof v === 'object') {
      // w:txbxContent 可能是对象也可能是数组（极少见）
      if (Array.isArray(v)) {
        for (const item of v) if (item && typeof item === 'object') acc.push(item as Record<string, unknown>)
      } else {
        acc.push(v as Record<string, unknown>)
      }
      // 继续向下，防止 txbxContent 里又嵌文本框
    }
    collectAllTxbxContent(v, acc, visited)
  }
}

/**
 * 从 wps:wsp 数组中抽取文本框内容。采用深度查找策略以容忍 Word 的不同序列化。
 */
function extractTextFromWsps(
  wsps: Record<string, unknown>[],
  ctx: ParseContext,
): TiptapNode[] {
  const out: TiptapNode[] = []
  const txbxes: Record<string, unknown>[] = []
  const visited = new WeakSet<object>()
  for (const wsp of wsps) collectAllTxbxContent(wsp, txbxes, visited)
  for (const txbxContent of txbxes) {
    out.push(...flattenTextBoxContent(txbxContent, ctx))
  }
  return out
}

/**
 * 处理 wpg:wgp 分组图形：
 *   <a:graphicData uri=".../wordprocessingGroup">
 *     <wpg:wgp>
 *       <wpg:grpSpPr/>
 *       <pic:pic>...logo 图片...</pic:pic>
 *       <wps:wsp><wps:txbx><w:txbxContent>...slogan 文字...</w:txbxContent></wps:wsp>
 *     </wpg:wgp>
 *   </a:graphicData>
 *
 * 宇视 logo + "视无界 智以恒" 这种整块 "图片+slogan" 通常就是这个结构。
 * 我们在这里把分组里的 pic 转成 image，把 wsp 里的 txbxContent 展平成文本节点，
 * 让二者都出现在编辑器中。
 */
function extractFromGroupShape(
  wgp: Record<string, unknown>,
  display: 'inline' | 'block',
  ctx: ParseContext,
): TiptapNode[] {
  const out: TiptapNode[] = []

  // 组内可以有多个图片（极少见但合法）
  const pics = ensureArray(wgp['pic:pic'] as Record<string, unknown>[])
  for (const pic of pics) {
    const node = buildImageNodeFromPic(pic, display, undefined, undefined, ctx)
    if (node) out.push(node)
  }

  // 文本框
  const wsps = ensureArray(wgp['wps:wsp'] as Record<string, unknown>[])
  if (wsps.length > 0) {
    out.push(...extractTextFromWsps(wsps, ctx))
  }

  // 嵌套的分组（wpg:grpSp 或再次 wpg:wgp）
  const nestedGroups = [
    ...ensureArray(wgp['wpg:grpSp'] as Record<string, unknown>[]),
    ...ensureArray(wgp['wpg:wgp'] as Record<string, unknown>[]),
  ]
  for (const g of nestedGroups) {
    out.push(...extractFromGroupShape(g, display, ctx))
  }

  return out
}

/**
 * 从 pic:pic 节点直接构造 image 节点，供分组图形中复用。
 */
function buildImageNodeFromPic(
  pic: Record<string, unknown>,
  display: 'inline' | 'block',
  width: number | undefined,
  height: number | undefined,
  ctx: ParseContext,
): TiptapNode | null {
  const blipFill = pic['pic:blipFill'] as Record<string, unknown> | undefined
  const blip = blipFill?.['a:blip'] as Record<string, unknown> | undefined
  if (!blip) return null
  const rId = getAttr(blip, 'r:embed') ?? getAttr(blip, 'r:link')
  if (!rId) return null
  const dataUrl = getImageDataUrl(rId, ctx.relationships, ctx.images)
  const relTarget = resolveRelTarget(ctx.relationships, rId)
  const attrs: Record<string, unknown> = {
    src: dataUrl ?? '',
    display,
  }
  if (width) attrs.width = width
  if (height) attrs.height = height
  if (relTarget) attrs['data-origin-src'] = `word/${relTarget}`
  // 选择性保存高保真方案：记录原始 rels 追溯三元组，导出侧 localSerializer
  // 优先复用原 rId，避免重复嵌入 media 文件并导致 rels 冲突。
  attrs['data-origin-rid'] = rId
  if (relTarget) attrs['data-origin-target'] = relTarget
  if (ctx.partPath) attrs['data-origin-part'] = ctx.partPath
  if (!dataUrl) ctx.logs.warn.push(`Image not found for rId=${rId}, target=${relTarget}`)
  return createNode('image', attrs)
}

/** 从 w:drawing 元素中提取图片/文本框节点 */
export function handleDrawing(
  drawing: Record<string, unknown>,
  ctx: ParseContext,
): TiptapNode[] {
  const results: TiptapNode[] = []

  // 行内图片: wp:inline
  const inline = drawing['wp:inline'] as Record<string, unknown> | undefined
  if (inline) {
    const node = extractFromInlineOrAnchor(inline, 'inline', ctx)
    if (node) results.push(...node)
  }

  // 浮动图片: wp:anchor
  const anchor = drawing['wp:anchor'] as Record<string, unknown> | undefined
  if (anchor) {
    const node = extractFromInlineOrAnchor(anchor, 'block', ctx)
    if (node) results.push(...node)
  }

  return results
}

function extractFromInlineOrAnchor(
  node: Record<string, unknown>,
  display: 'inline' | 'block',
  ctx: ParseContext,
): TiptapNode[] | null {
  const extent = node['wp:extent'] as Record<string, unknown> | undefined
  let width: number | undefined
  let height: number | undefined
  if (extent) {
    const cx = getAttr(extent, 'cx')
    const cy = getAttr(extent, 'cy')
    if (cx) width = emuToPx(Number(cx))
    if (cy) height = emuToPx(Number(cy))
  }

  const graphic = node['a:graphic'] as Record<string, unknown> | undefined
  const graphicData = graphic?.['a:graphicData'] as Record<string, unknown> | undefined
  if (!graphicData) return null

  // 1) 图片 pic:pic
  const pic = graphicData['pic:pic'] as Record<string, unknown> | undefined
  if (pic) {
    const node = buildImageNodeFromPic(pic, display, width, height, ctx)
    if (!node) {
      ctx.logs.warn.push('Drawing pic:pic found without usable blip reference')
      return null
    }
    return [node]
  }

  // 2) 文本框 wps:wsp（uri 常为 wordprocessingShape）
  //    形状内可能根本没有文本框——只是装饰用线/矩形/连接符，此时静默返回 null，
  //    不再落到 "unsupported graphicData" 告警分支刷日志。
  const wsps = ensureArray(graphicData['wps:wsp'] as Record<string, unknown>[])
  if (wsps.length > 0) {
    const out = extractTextFromWsps(wsps, ctx)
    return out.length > 0 ? out : null
  }

  // 3) 分组图形 wpg:wgp（uri=wordprocessingGroup）
  //    里面通常同时包含 pic:pic（logo）和 wps:wsp（slogan）。
  const wgps = [
    ...ensureArray(graphicData['wpg:wgp'] as Record<string, unknown>[]),
    ...ensureArray(graphicData['wpg:grpSp'] as Record<string, unknown>[]),
  ]
  if (wgps.length > 0) {
    const out: TiptapNode[] = []
    for (const wgp of wgps) {
      out.push(...extractFromGroupShape(wgp, display, ctx))
    }
    return out.length > 0 ? out : null
  }

  // 4) 绘图画布 wpc:wpc（uri=wordprocessingCanvas）
  //    画布里的内容结构和分组图形一致：pic:pic + wps:wsp + 嵌套 wpg:wgp。
  const wpcs = ensureArray(graphicData['wpc:wpc'] as Record<string, unknown>[])
  if (wpcs.length > 0) {
    const out: TiptapNode[] = []
    for (const wpc of wpcs) {
      out.push(...extractFromGroupShape(wpc, display, ctx))
    }
    return out.length > 0 ? out : null
  }

  // 带上 uri 类型方便排查（chart / diagram / smartArt / inkml 等目前都不支持）。
  // 打印完整 uri，便于后续定位是 2006/chart、2014/chartex、2010/wordprocessingInk 等中的哪一种。
  const uri = getAttr(graphicData, 'uri')
  ctx.logs.warn.push(
    `Drawing with unsupported graphicData (uri=${uri ?? 'unknown'}); supported: pic:pic / wps:wsp / wpg:wgp / wpc:wpc`
  )
  return null
}

/**
 * 处理旧版 w:pict 格式。可能形态：
 *   1) v:shape 内仅有 v:imagedata  → 图片
 *   2) v:shape 内仅有 v:textbox    → 文本框（slogan / 水印）
 *   3) v:shape 同时包含二者        → 图+文
 *   4) v:group 里嵌套多个 v:shape  → 分组图形
 * 返回一组节点（而不是单个），让调用方统一 push。
 */
export function handlePict(
  pict: Record<string, unknown>,
  ctx: ParseContext,
): TiptapNode[] {
  const out: TiptapNode[] = []

  const shapes = ensureArray(pict['v:shape'] as Record<string, unknown>[])
  for (const shape of shapes) {
    out.push(...extractFromVmlShape(shape, ctx))
  }

  const groups = ensureArray(pict['v:group'] as Record<string, unknown>[])
  for (const g of groups) {
    const innerShapes = ensureArray(g['v:shape'] as Record<string, unknown>[])
    for (const shape of innerShapes) {
      out.push(...extractFromVmlShape(shape, ctx))
    }
  }

  return out
}

function extractFromVmlShape(
  shape: Record<string, unknown>,
  ctx: ParseContext,
): TiptapNode[] {
  const out: TiptapNode[] = []

  const style = getAttr(shape, 'style') ?? ''
  let width: number | undefined
  let height: number | undefined
  const widthMatch = style.match(/width:\s*([\d.]+)(?:pt|px)/)
  const heightMatch = style.match(/height:\s*([\d.]+)(?:pt|px)/)
  if (widthMatch) width = Math.round(parseFloat(widthMatch[1]) * (style.includes('pt') ? 1.333 : 1))
  if (heightMatch) height = Math.round(parseFloat(heightMatch[1]) * (style.includes('pt') ? 1.333 : 1))

  // 1) 图片
  const imagedata = shape['v:imagedata'] as Record<string, unknown> | undefined
  if (imagedata) {
    const rId = getAttr(imagedata, 'r:id')
    if (rId) {
      const dataUrl = getImageDataUrl(rId, ctx.relationships, ctx.images)
      const relTarget = resolveRelTarget(ctx.relationships, rId)
      const attrs: Record<string, unknown> = {
        src: dataUrl ?? '',
        display: 'inline',
      }
      if (width) attrs.width = width
      if (height) attrs.height = height
      if (relTarget) attrs['data-origin-src'] = `word/${relTarget}`
      // 选择性保存高保真方案：VML 老格式图片同样回填 rels 追溯三元组。
      attrs['data-origin-rid'] = rId
      if (relTarget) attrs['data-origin-target'] = relTarget
      if (ctx.partPath) attrs['data-origin-part'] = ctx.partPath
      out.push(createNode('image', attrs))
    }
  }

  // 2) VML 文本框：v:textbox > w:txbxContent（Word 老版文档的 slogan / 水印位置）
  const textboxes = ensureArray(shape['v:textbox'] as Record<string, unknown>[])
  for (const tb of textboxes) {
    const txbxContent = tb['w:txbxContent'] as Record<string, unknown> | undefined
    if (txbxContent) {
      out.push(...flattenTextBoxContent(txbxContent, ctx))
    }
  }

  return out
}

/** 在 run 中检查并提取图片 */
export function extractImagesFromRun(
  run: Record<string, unknown>,
  ctx: ParseContext,
): TiptapNode[] {
  const images: TiptapNode[] = []

  const drawings = ensureArray(run['w:drawing'] as Record<string, unknown>[])
  for (const drawing of drawings) {
    const out = handleDrawing(drawing, ctx)
    if (out.length > 0) images.push(...out)
  }

  const picts = ensureArray(run['w:pict'] as Record<string, unknown>[])
  for (const pict of picts) {
    const nodes = handlePict(pict, ctx)
    if (nodes.length > 0) images.push(...nodes)
  }

  return images
}

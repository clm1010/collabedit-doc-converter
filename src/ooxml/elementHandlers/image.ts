import type { TiptapNode } from '../../types/tiptapJson.js'
import type { ParseContext } from '../../types/ooxml.js'
import { createNode } from '../../types/tiptapJson.js'
import { ensureArray, getAttr } from '../xmlParser.js'
import { getImageDataUrl } from '../imageExtractor.js'
import { resolveRelTarget } from '../relationships.js'

/** EMU → px (1 inch = 914400 EMU, 1 inch = 96px) */
function emuToPx(emu: number): number {
  return Math.round(emu / 914400 * 96)
}

/** 从 w:drawing 元素中提取图片节点 */
export function handleDrawing(
  drawing: Record<string, unknown>,
  ctx: ParseContext,
): TiptapNode | null {
  // 行内图片: wp:inline
  const inline = drawing['wp:inline'] as Record<string, unknown> | undefined
  if (inline) {
    return extractImageFromInlineOrAnchor(inline, 'inline', ctx)
  }

  // 浮动图片: wp:anchor
  const anchor = drawing['wp:anchor'] as Record<string, unknown> | undefined
  if (anchor) {
    return extractImageFromInlineOrAnchor(anchor, 'block', ctx)
  }

  return null
}

function extractImageFromInlineOrAnchor(
  node: Record<string, unknown>,
  display: 'inline' | 'block',
  ctx: ParseContext,
): TiptapNode | null {
  // 获取尺寸 (extent)
  const extent = node['wp:extent'] as Record<string, unknown> | undefined
  let width: number | undefined
  let height: number | undefined
  if (extent) {
    const cx = getAttr(extent, 'cx')
    const cy = getAttr(extent, 'cy')
    if (cx) width = emuToPx(Number(cx))
    if (cy) height = emuToPx(Number(cy))
  }

  // 找到 blip 获取 rId
  const graphic = node['a:graphic'] as Record<string, unknown> | undefined
  const graphicData = graphic?.['a:graphicData'] as Record<string, unknown> | undefined
  const pic = graphicData?.['pic:pic'] as Record<string, unknown> | undefined
  const blipFill = pic?.['pic:blipFill'] as Record<string, unknown> | undefined
  const blip = blipFill?.['a:blip'] as Record<string, unknown> | undefined

  if (!blip) {
    ctx.logs.warn.push('Drawing found without blip reference')
    return null
  }

  const rId = getAttr(blip, 'r:embed') ?? getAttr(blip, 'r:link')
  if (!rId) {
    ctx.logs.warn.push('Drawing blip without r:embed or r:link')
    return null
  }

  const dataUrl = getImageDataUrl(rId, ctx.relationships, ctx.images)
  const relTarget = resolveRelTarget(ctx.relationships, rId)

  const attrs: Record<string, unknown> = {
    src: dataUrl ?? '',
    display,
  }
  if (width) attrs.width = width
  if (height) attrs.height = height
  if (relTarget) attrs['data-origin-src'] = `word/${relTarget}`

  if (!dataUrl) {
    ctx.logs.warn.push(`Image not found for rId=${rId}, target=${relTarget}`)
  }

  return createNode('image', attrs)
}

/** 处理旧版 w:pict / v:imagedata 格式 */
export function handlePict(
  pict: Record<string, unknown>,
  ctx: ParseContext,
): TiptapNode | null {
  // v:shape → v:imagedata
  const shape = pict['v:shape'] as Record<string, unknown> | undefined
  if (!shape) return null

  const imagedata = shape['v:imagedata'] as Record<string, unknown> | undefined
  if (!imagedata) return null

  const rId = getAttr(imagedata, 'r:id')
  if (!rId) return null

  const dataUrl = getImageDataUrl(rId, ctx.relationships, ctx.images)
  const relTarget = resolveRelTarget(ctx.relationships, rId)

  // 尝试从 v:shape 的 style 属性中获取尺寸
  const style = getAttr(shape, 'style') ?? ''
  let width: number | undefined
  let height: number | undefined

  const widthMatch = style.match(/width:\s*([\d.]+)(?:pt|px)/)
  const heightMatch = style.match(/height:\s*([\d.]+)(?:pt|px)/)
  if (widthMatch) width = Math.round(parseFloat(widthMatch[1]) * (style.includes('pt') ? 1.333 : 1))
  if (heightMatch) height = Math.round(parseFloat(heightMatch[1]) * (style.includes('pt') ? 1.333 : 1))

  const attrs: Record<string, unknown> = {
    src: dataUrl ?? '',
    display: 'inline',
  }
  if (width) attrs.width = width
  if (height) attrs.height = height
  if (relTarget) attrs['data-origin-src'] = `word/${relTarget}`

  return createNode('image', attrs)
}

/** 在 run 中检查并提取图片 */
export function extractImagesFromRun(
  run: Record<string, unknown>,
  ctx: ParseContext,
): TiptapNode[] {
  const images: TiptapNode[] = []

  const drawings = ensureArray(run['w:drawing'] as Record<string, unknown>[])
  for (const drawing of drawings) {
    const img = handleDrawing(drawing, ctx)
    if (img) images.push(img)
  }

  const picts = ensureArray(run['w:pict'] as Record<string, unknown>[])
  for (const pict of picts) {
    const img = handlePict(pict, ctx)
    if (img) images.push(img)
  }

  return images
}

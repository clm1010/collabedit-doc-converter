import type { TiptapNode } from '../../types/tiptapJson.js'
import type { ParseContext } from '../../types/ooxml.js'
import { createNode } from '../../types/tiptapJson.js'
import { ensureArray, getAttr, getWVal } from '../xmlParser.js'
import { handleParagraph } from './paragraph.js'

/** twips → px */
function twipsToPx(twips: number): number {
  return Math.round(twips / 15)
}

/** EMU → px (1 inch = 914400 EMU, 1 inch = 96px) */
function emuToPx(emu: number): number {
  return Math.round(emu / 914400 * 96)
}

export function handleTable(
  tbl: Record<string, unknown>,
  ctx: ParseContext,
): TiptapNode {
  // 表格属性
  const tblPr = tbl['w:tblPr'] as Record<string, unknown> | undefined
  const tableAttrs: Record<string, unknown> = {}

  if (tblPr) {
    const tblW = tblPr['w:tblW'] as Record<string, unknown> | undefined
    if (tblW) {
      const w = getAttr(tblW, 'w:w')
      const wType = getAttr(tblW, 'w:type')
      if (w && wType === 'dxa') {
        tableAttrs.tableWidth = twipsToPx(Number(w))
      } else if (w && wType === 'pct') {
        tableAttrs.tableWidth = `${Number(w) / 50}%`
      }
    }
  }

  // 表格网格（列宽）
  const tblGrid = tbl['w:tblGrid'] as Record<string, unknown> | undefined
  const gridCols: number[] = []
  if (tblGrid) {
    const cols = ensureArray(tblGrid['w:gridCol'] as Record<string, unknown>[])
    for (const col of cols) {
      const w = getAttr(col, 'w:w')
      if (w) gridCols.push(twipsToPx(Number(w)))
    }
  }

  // 处理行
  const rows = ensureArray(tbl['w:tr'] as Record<string, unknown>[])
  const rowNodes: TiptapNode[] = []

  // vMerge 状态追踪：colIndex → { rowspan, startRow, cellNode(起始单元格引用) }
  const vMergeState: Map<number, { rowspan: number; startRow: number; cellNode?: TiptapNode }> = new Map()

  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    const tr = rows[rowIdx]
    const rowNode = handleTableRow(tr, rowIdx, gridCols, vMergeState, ctx)
    if (rowNode) rowNodes.push(rowNode)
  }

  // 所有行处理完毕后，回写正确的 rowspan 到起始单元格
  for (const [, state] of vMergeState) {
    if (state.rowspan > 1 && state.cellNode) {
      if (!state.cellNode.attrs) state.cellNode.attrs = {}
      state.cellNode.attrs.rowspan = state.rowspan
    }
  }

  return createNode('table', Object.keys(tableAttrs).length > 0 ? tableAttrs : undefined, rowNodes)
}

function handleTableRow(
  tr: Record<string, unknown>,
  rowIdx: number,
  gridCols: number[],
  vMergeState: Map<number, { rowspan: number; startRow: number; cellNode?: TiptapNode }>,
  ctx: ParseContext,
): TiptapNode | null {
  const trPr = tr['w:trPr'] as Record<string, unknown> | undefined
  const rowAttrs: Record<string, unknown> = {}

  if (trPr) {
    const trHeight = trPr['w:trHeight'] as Record<string, unknown> | undefined
    if (trHeight) {
      const val = getAttr(trHeight, 'w:val')
      if (val) rowAttrs.height = twipsToPx(Number(val))
    }
  }

  const cells = ensureArray(tr['w:tc'] as Record<string, unknown>[])
  const cellNodes: TiptapNode[] = []
  let gridIdx = 0

  for (const tc of cells) {
    const cellResult = handleTableCell(tc, rowIdx, gridIdx, gridCols, vMergeState, ctx)
    if (cellResult) {
      cellNodes.push(cellResult.node)
      gridIdx += cellResult.colspan
    } else {
      gridIdx++
    }
  }

  return createNode(
    'tableRow',
    Object.keys(rowAttrs).length > 0 ? rowAttrs : undefined,
    cellNodes,
  )
}

function handleTableCell(
  tc: Record<string, unknown>,
  rowIdx: number,
  gridIdx: number,
  gridCols: number[],
  vMergeState: Map<number, { rowspan: number; startRow: number; cellNode?: TiptapNode }>,
  ctx: ParseContext,
): { node: TiptapNode; colspan: number } | null {
  const tcPr = tc['w:tcPr'] as Record<string, unknown> | undefined
  const attrs: Record<string, unknown> = {}
  let colspan = 1
  let isVMergeRestart = false

  if (tcPr) {
    // colspan
    const gridSpan = tcPr['w:gridSpan'] as Record<string, unknown> | undefined
    if (gridSpan) {
      const val = getWVal(gridSpan)
      if (val) {
        colspan = Number(val)
        if (colspan > 1) attrs.colspan = colspan
      }
    }

    // vMerge
    const vMerge = tcPr['w:vMerge'] as Record<string, unknown> | undefined | string
    if (vMerge !== undefined) {
      const val = typeof vMerge === 'string' ? vMerge : getWVal(vMerge as Record<string, unknown>)
      if (val === 'restart') {
        vMergeState.set(gridIdx, { rowspan: 1, startRow: rowIdx })
        isVMergeRestart = true
      } else {
        // continue merge
        const state = vMergeState.get(gridIdx)
        if (state) {
          state.rowspan++
          return null
        }
      }
    } else {
      vMergeState.delete(gridIdx)
    }

    // 列宽
    if (gridCols.length > 0) {
      const colwidths: number[] = []
      for (let i = gridIdx; i < gridIdx + colspan && i < gridCols.length; i++) {
        colwidths.push(gridCols[i])
      }
      if (colwidths.length > 0) attrs.colwidth = colwidths
    }

    // textAlign
    const jc = tcPr['w:jc'] as Record<string, unknown> | undefined
    if (jc) {
      const val = getWVal(jc)
      if (val) attrs.textAlign = val
    }

    // backgroundColor
    const shd = tcPr['w:shd'] as Record<string, unknown> | undefined
    if (shd) {
      const fill = getAttr(shd, 'w:fill')
      if (fill && fill !== 'auto') {
        attrs.backgroundColor = fill.startsWith('#') ? fill : `#${fill}`
      }
    }

    // verticalAlign
    const vAlign = tcPr['w:vAlign'] as Record<string, unknown> | undefined
    if (vAlign) {
      const val = getWVal(vAlign)
      if (val) attrs.verticalAlign = val
    }
  }

  // 处理单元格内容
  const content = processCellContent(tc, ctx)
  const cellContent = content.length > 0 ? content : [createNode('paragraph')]

  const cellNode = createNode('tableCell', Object.keys(attrs).length > 0 ? attrs : undefined, cellContent)

  // 保存起始单元格引用，rowspan 在所有行处理完毕后由 handleTable 回写
  if (isVMergeRestart) {
    const state = vMergeState.get(gridIdx)
    if (state) state.cellNode = cellNode
  }

  return { node: cellNode, colspan }
}

function processCellContent(
  tc: Record<string, unknown>,
  ctx: ParseContext,
): TiptapNode[] {
  const content: TiptapNode[] = []

  const paragraphs = ensureArray(tc['w:p'] as Record<string, unknown>[])
  for (const p of paragraphs) {
    const result = handleParagraph(p, ctx)
    if (result) {
      if (Array.isArray(result)) {
        content.push(...result)
      } else {
        content.push(result)
      }
    }
  }

  // 嵌套表格
  const tables = ensureArray(tc['w:tbl'] as Record<string, unknown>[])
  for (const tbl of tables) {
    content.push(handleTable(tbl, ctx))
  }

  return content
}

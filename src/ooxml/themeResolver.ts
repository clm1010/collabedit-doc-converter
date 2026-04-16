import type { ThemeColors, ThemeFonts } from '../types/ooxml.js'
import { parseXml, getAttr } from './xmlParser.js'
import type { DocxArchive } from './zipExtractor.js'

const DEFAULT_COLORS: ThemeColors = {
  dk1: '#000000', lt1: '#FFFFFF',
  dk2: '#44546A', lt2: '#E7E6E6',
  accent1: '#4472C4', accent2: '#ED7D31',
  accent3: '#A5A5A5', accent4: '#FFC000',
  accent5: '#5B9BD5', accent6: '#70AD47',
  hlink: '#0563C1', folHlink: '#954F72',
}

const DEFAULT_FONTS: ThemeFonts = {
  majorLatin: 'Calibri Light',
  majorEastAsia: '等线 Light',
  minorLatin: 'Calibri',
  minorEastAsia: '等线',
}

export function resolveTheme(archive: DocxArchive): { colors: ThemeColors; fonts: ThemeFonts } {
  const themeFiles = archive.listFiles('word/theme/')
  const themePath = themeFiles.find((f) => f.endsWith('.xml')) ?? 'word/theme/theme1.xml'
  const xml = archive.getText(themePath)

  if (!xml) return { colors: { ...DEFAULT_COLORS }, fonts: { ...DEFAULT_FONTS } }

  const parsed = parseXml(xml)
  const theme = parsed['a:theme'] as Record<string, unknown> | undefined
  if (!theme) return { colors: { ...DEFAULT_COLORS }, fonts: { ...DEFAULT_FONTS } }

  const themeElements = theme['a:themeElements'] as Record<string, unknown> | undefined
  if (!themeElements) return { colors: { ...DEFAULT_COLORS }, fonts: { ...DEFAULT_FONTS } }

  const colors = parseColorScheme(themeElements)
  const fonts = parseFontScheme(themeElements)

  return { colors, fonts }
}

function parseColorScheme(themeElements: Record<string, unknown>): ThemeColors {
  const clrScheme = themeElements['a:clrScheme'] as Record<string, unknown> | undefined
  if (!clrScheme) return { ...DEFAULT_COLORS }

  const colors = { ...DEFAULT_COLORS }
  const colorKeys: (keyof ThemeColors)[] = [
    'dk1', 'lt1', 'dk2', 'lt2',
    'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
    'hlink', 'folHlink',
  ]

  for (const key of colorKeys) {
    const entry = clrScheme[`a:${key}`] as Record<string, unknown> | undefined
    if (!entry) continue

    const srgb = entry['a:srgbClr'] as Record<string, unknown> | undefined
    if (srgb) {
      const val = getAttr(srgb, 'val')
      if (val) colors[key] = `#${val}`
      continue
    }

    const sys = entry['a:sysClr'] as Record<string, unknown> | undefined
    if (sys) {
      const lastClr = getAttr(sys, 'lastClr')
      if (lastClr) colors[key] = `#${lastClr}`
    }
  }

  return colors
}

function parseFontScheme(themeElements: Record<string, unknown>): ThemeFonts {
  const fontScheme = themeElements['a:fontScheme'] as Record<string, unknown> | undefined
  if (!fontScheme) return { ...DEFAULT_FONTS }

  const fonts = { ...DEFAULT_FONTS }

  const majorFont = fontScheme['a:majorFont'] as Record<string, unknown> | undefined
  if (majorFont) {
    const latin = majorFont['a:latin'] as Record<string, unknown> | undefined
    if (latin) fonts.majorLatin = getAttr(latin, 'typeface') ?? fonts.majorLatin
    const ea = majorFont['a:ea'] as Record<string, unknown> | undefined
    if (ea) fonts.majorEastAsia = getAttr(ea, 'typeface') ?? fonts.majorEastAsia
  }

  const minorFont = fontScheme['a:minorFont'] as Record<string, unknown> | undefined
  if (minorFont) {
    const latin = minorFont['a:latin'] as Record<string, unknown> | undefined
    if (latin) fonts.minorLatin = getAttr(latin, 'typeface') ?? fonts.minorLatin
    const ea = minorFont['a:ea'] as Record<string, unknown> | undefined
    if (ea) fonts.minorEastAsia = getAttr(ea, 'typeface') ?? fonts.minorEastAsia
  }

  return fonts
}

function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace(/^#/, '')
  return [
    parseInt(h.slice(0, 2), 16),
    parseInt(h.slice(2, 4), 16),
    parseInt(h.slice(4, 6), 16),
  ]
}

function rgbToHex(r: number, g: number, b: number): string {
  return '#' + [r, g, b].map(v => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, '0')).join('').toUpperCase()
}

function applyTint(rgb: [number, number, number], tint: number): [number, number, number] {
  return rgb.map(c => c + (255 - c) * tint) as [number, number, number]
}

function applyShade(rgb: [number, number, number], shade: number): [number, number, number] {
  return rgb.map(c => c * shade) as [number, number, number]
}

/** 解析主题颜色引用（如 accent1 + tint/shade 修改） */
export function resolveThemeColor(
  themeColors: ThemeColors,
  themeColorName: string,
  tint?: string,
  shade?: string,
): string | undefined {
  const mapped: Record<string, keyof ThemeColors> = {
    dark1: 'dk1', light1: 'lt1',
    dark2: 'dk2', light2: 'lt2',
    accent1: 'accent1', accent2: 'accent2',
    accent3: 'accent3', accent4: 'accent4',
    accent5: 'accent5', accent6: 'accent6',
    hyperlink: 'hlink', followedHyperlink: 'folHlink',
  }

  const key = mapped[themeColorName] ?? themeColorName
  const base = themeColors[key]
  if (!base) return undefined

  let rgb = hexToRgb(base)

  if (tint) {
    const t = parseInt(tint, 16) / 255
    if (!isNaN(t)) rgb = applyTint(rgb, t)
  }
  if (shade) {
    const s = parseInt(shade, 16) / 255
    if (!isNaN(s)) rgb = applyShade(rgb, s)
  }

  return rgbToHex(...rgb)
}

import type { DocMetadata } from '../types/docMetadata.js'
import type { DocxArchive } from '../ooxml/zipExtractor.js'
import { parseXml, getAttr } from '../ooxml/xmlParser.js'

export function extractMetadata(archive: DocxArchive): DocMetadata {
  const metadata: DocMetadata = {
    paperSize: { width: 210, height: 297 },
    margins: { top: 25.4, bottom: 25.4, left: 31.8, right: 31.8 },
    defaultFont: '宋体',
    defaultFontSize: 12,
    headers: {},
    footers: {},
    sections: [],
    hasFootnotes: false,
    hasEndnotes: false,
    numberingDefinitions: [],
    customStyles: [],
  }

  try {
    const documentXml = archive.getText('word/document.xml')
    if (documentXml) parseSectionProperties(documentXml, metadata)

    const stylesXml = archive.getText('word/styles.xml')
    if (stylesXml) parseDefaultStyles(stylesXml, metadata)

    const numberingXml = archive.getText('word/numbering.xml')
    if (numberingXml) metadata.numberingDefinitions = parseNumberingDefinitions(numberingXml)

    const headerFiles = archive.listFiles('word/').filter((f) => /^word\/header\d*\.xml$/.test(f))
    for (const hf of headerFiles) {
      const content = archive.getText(hf)
      if (content) {
        const key = hf.includes('1') ? 'default' : hf.includes('2') ? 'even' : 'first'
        metadata.headers[key as keyof typeof metadata.headers] = content
      }
    }

    const footerFiles = archive.listFiles('word/').filter((f) => /^word\/footer\d*\.xml$/.test(f))
    for (const ff of footerFiles) {
      const content = archive.getText(ff)
      if (content) {
        const key = ff.includes('1') ? 'default' : ff.includes('2') ? 'even' : 'first'
        metadata.footers[key as keyof typeof metadata.footers] = content
      }
    }

    metadata.hasFootnotes = archive.getBuffer('word/footnotes.xml') !== null
    metadata.hasEndnotes = archive.getBuffer('word/endnotes.xml') !== null
  } catch (err) {
    console.warn('Metadata extraction partially failed:', err)
  }

  return metadata
}

function parseSectionProperties(documentXml: string, metadata: DocMetadata) {
  const sectPrMatch = documentXml.match(/<w:sectPr[^>]*>([\s\S]*?)<\/w:sectPr>/g)
  if (!sectPrMatch) return

  for (const sectPr of sectPrMatch) {
    const pgSz = sectPr.match(/<w:pgSz\s+([^/]*)\/>/)
    if (pgSz) {
      const wMatch = pgSz[1].match(/w:w="(\d+)"/)
      const hMatch = pgSz[1].match(/w:h="(\d+)"/)
      if (wMatch) metadata.paperSize.width = twipsToMm(Number(wMatch[1]))
      if (hMatch) metadata.paperSize.height = twipsToMm(Number(hMatch[1]))
    }

    const pgMar = sectPr.match(/<w:pgMar\s+([^/]*)\/>/)
    if (pgMar) {
      const topMatch = pgMar[1].match(/w:top="(\d+)"/)
      const bottomMatch = pgMar[1].match(/w:bottom="(\d+)"/)
      const leftMatch = pgMar[1].match(/w:left="(\d+)"/)
      const rightMatch = pgMar[1].match(/w:right="(\d+)"/)
      if (topMatch) metadata.margins.top = twipsToMm(Number(topMatch[1]))
      if (bottomMatch) metadata.margins.bottom = twipsToMm(Number(bottomMatch[1]))
      if (leftMatch) metadata.margins.left = twipsToMm(Number(leftMatch[1]))
      if (rightMatch) metadata.margins.right = twipsToMm(Number(rightMatch[1]))
    }
  }
}

function parseDefaultStyles(stylesXml: string, metadata: DocMetadata) {
  const defaultRprMatch = stylesXml.match(
    /<w:docDefaults>[\s\S]*?<w:rPrDefault>[\s\S]*?<w:rPr>([\s\S]*?)<\/w:rPr>/,
  )
  if (defaultRprMatch) {
    const rPr = defaultRprMatch[1]
    const fontMatch = rPr.match(/<w:rFonts[^>]*w:ascii="([^"]*)"/)
    if (fontMatch) metadata.defaultFont = fontMatch[1]

    const sizeMatch = rPr.match(/<w:sz\s+w:val="(\d+)"/)
    if (sizeMatch) metadata.defaultFontSize = Number(sizeMatch[1]) / 2
  }
}

function parseNumberingDefinitions(numberingXml: string): object[] {
  const definitions: object[] = []
  const abstractNums = numberingXml.match(/<w:abstractNum[\s\S]*?<\/w:abstractNum>/g)
  if (abstractNums) {
    for (const an of abstractNums) {
      const idMatch = an.match(/w:abstractNumId="(\d+)"/)
      const levels: object[] = []
      const lvlMatches = an.match(/<w:lvl[\s\S]*?<\/w:lvl>/g)
      if (lvlMatches) {
        for (const lvl of lvlMatches) {
          const ilvlMatch = lvl.match(/w:ilvl="(\d+)"/)
          const numFmtMatch = lvl.match(/<w:numFmt\s+w:val="([^"]*)"/)
          const lvlTextMatch = lvl.match(/<w:lvlText\s+w:val="([^"]*)"/)
          levels.push({
            level: ilvlMatch ? Number(ilvlMatch[1]) : 0,
            numFmt: numFmtMatch?.[1] ?? 'decimal',
            lvlText: lvlTextMatch?.[1] ?? '',
          })
        }
      }
      definitions.push({
        abstractNumId: idMatch ? Number(idMatch[1]) : 0,
        levels,
      })
    }
  }
  return definitions
}

function twipsToMm(twips: number): number {
  return Math.round((twips / 1440) * 25.4 * 10) / 10
}

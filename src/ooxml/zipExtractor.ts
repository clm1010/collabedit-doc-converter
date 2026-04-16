import { unzipSync, type Unzipped } from 'fflate'

export interface DocxArchive {
  files: Unzipped
  getText(path: string): string | null
  getBuffer(path: string): Uint8Array | null
  listFiles(prefix?: string): string[]
}

export function extractDocx(buffer: Buffer | Uint8Array): DocxArchive {
  const data = buffer instanceof Buffer ? new Uint8Array(buffer) : buffer
  const files = unzipSync(data)

  return {
    files,

    getText(path: string): string | null {
      const entry = files[path]
      if (!entry) return null
      return new TextDecoder('utf-8').decode(entry)
    },

    getBuffer(path: string): Uint8Array | null {
      return files[path] ?? null
    },

    listFiles(prefix?: string): string[] {
      const keys = Object.keys(files)
      if (!prefix) return keys
      return keys.filter((k) => k.startsWith(prefix))
    },
  }
}

/**
 * Merge export service (P2 phase).
 * Uses original DOCX as template, replaces body content while preserving styles.
 * Not implemented in P0.
 */
export async function mergeExport(
  _originalDocx: Buffer,
  _editedHtml: string,
  _metadata?: object
): Promise<Buffer> {
  throw new Error('merge-export is not yet implemented (scheduled for P2)')
}

/** Parse .xlsx/.xls bytes and render one sheet as an HTML table via SheetJS, entirely in the browser. */
import { read as readWorkbook, utils } from 'xlsx'
import type { WorkBook } from 'xlsx'

/** .xlsx is a ZIP (OOXML) package; legacy .xls is an OLE2 Compound File. Checking both up front gives a
 * clear "invalid document" failure instead of SheetJS silently reading unrelated bytes as a one-cell sheet. */
const ZIP_SIGNATURE = [0x50, 0x4b]
const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]

function matchesSignature(data: Uint8Array, signature: readonly number[]): boolean {
  return data.length >= signature.length && signature.every((byte, index) => data[index] === byte)
}

function looksLikeSpreadsheet(data: Uint8Array): boolean {
  return matchesSignature(data, ZIP_SIGNATURE) || matchesSignature(data, CFB_SIGNATURE)
}

/**
 * Parse spreadsheet bytes into a workbook.
 * @param data - complete .xlsx or .xls file bytes; transient UI input, never persisted or shared.
 * @returns the parsed workbook, retaining every sheet name in file order.
 * @throws when the bytes are not a readable spreadsheet.
 */
export function parseSpreadsheet(data: Uint8Array<ArrayBuffer>): WorkBook {
  if (!looksLikeSpreadsheet(data)) throw new Error('not a spreadsheet (.xlsx or .xls) file')
  return readWorkbook(data, { type: 'array' })
}

/**
 * Render one workbook sheet as a complete HTML document.
 * @param workbook - workbook parsed by {@link parseSpreadsheet}.
 * @param sheetName - sheet to render; must be one of `workbook.SheetNames`.
 * @returns a standalone HTML document string, or undefined for an unknown sheet name.
 */
export function spreadsheetSheetDocument(workbook: WorkBook, sheetName: string): string | undefined {
  const sheet = workbook.Sheets[sheetName]
  if (sheet === undefined) return undefined
  return wrapSpreadsheetDocument(utils.sheet_to_html(sheet))
}

function wrapSpreadsheetDocument(table: string): string {
  return `<!doctype html><html><head><meta charset="utf-8"><style>${SPREADSHEET_DOCUMENT_STYLE}</style></head>`
    + `<body>${table}</body></html>`
}

const SPREADSHEET_DOCUMENT_STYLE = `
  :root { color-scheme: light dark; }
  body { margin: 0; padding: 16px; box-sizing: border-box;
    font: 13px/1.5 -apple-system, "Segoe UI", sans-serif; color: #1a1a1a; background: #fff; }
  table { border-collapse: collapse; }
  td, th { border: 1px solid #d0d0d5; padding: 4px 8px; white-space: pre; }
  tr:nth-child(even) td { background: rgba(0, 0, 0, .03); }
  @media (prefers-color-scheme: dark) {
    body { background: #1c1c1f; color: #e6e6e6; }
    td, th { border-color: #3a3a40; }
    tr:nth-child(even) td { background: rgba(255, 255, 255, .04); }
  }
`

/** Spreadsheet byte parsing: signature rejection, invalid packages, sheet listing, and per-sheet rendering. */
import { describe, expect, it } from 'vitest'
import { parseSpreadsheet, spreadsheetSheetDocument } from '../src/client/spreadsheet/convert.ts'
import { fixtureDocxBytes, fixtureXlsxBytes } from './office-local-fixture.ts'

describe('parseSpreadsheet', () => {
  it('rejects bytes with neither a ZIP nor a Compound File signature', () => {
    expect(() => parseSpreadsheet(new TextEncoder().encode('not a spreadsheet')))
      .toThrow('not a spreadsheet (.xlsx or .xls) file')
  })

  it('rejects empty bytes', () => {
    expect(() => parseSpreadsheet(new Uint8Array()))
      .toThrow('not a spreadsheet (.xlsx or .xls) file')
  })

  it('rejects a ZIP package that is not a spreadsheet', () => {
    const docxBytes = fixtureDocxBytes('not a spreadsheet') as Uint8Array<ArrayBuffer>
    expect(() => parseSpreadsheet(docxBytes)).toThrow()
  })

  it('parses a real multi-sheet .xlsx workbook and preserves sheet order', () => {
    const bytes = fixtureXlsxBytes(new Map([['First', 'A1 text'], ['Second', '第二 sheet']])) as Uint8Array<ArrayBuffer>
    const workbook = parseSpreadsheet(bytes)
    expect(workbook.SheetNames).toEqual(['First', 'Second'])
  })
})

describe('spreadsheetSheetDocument', () => {
  it('renders the requested sheet as a standalone HTML document', () => {
    const bytes = fixtureXlsxBytes(new Map([['First', 'A1 text'], ['Second', '第二 sheet']])) as Uint8Array<ArrayBuffer>
    const workbook = parseSpreadsheet(bytes)
    const first = spreadsheetSheetDocument(workbook, 'First')
    expect(first).toContain('<!doctype html>')
    expect(first).toContain('A1 text')
    expect(first).not.toContain('第二 sheet')
    const second = spreadsheetSheetDocument(workbook, 'Second')
    expect(second).toContain('第二 sheet')
  })

  it('returns undefined for a sheet name absent from the workbook', () => {
    const bytes = fixtureXlsxBytes(new Map([['Only', 'value']])) as Uint8Array<ArrayBuffer>
    const workbook = parseSpreadsheet(bytes)
    expect(spreadsheetSheetDocument(workbook, 'Missing')).toBeUndefined()
  })
})

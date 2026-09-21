/** Word-document byte conversion: signature rejection, invalid packages, and real conversion output. */
import { describe, expect, it } from 'vitest'
import { docxToDocument } from '../src/client/docx/convert.ts'
import { fixtureDocxBytes, fixtureXlsxBytes } from './office-local-fixture.ts'

describe('docxToDocument', () => {
  it('rejects bytes with no ZIP signature', async () => {
    await expect(docxToDocument(new TextEncoder().encode('not a docx')))
      .rejects.toThrow('not a Word document (.docx) package')
  })

  it('rejects empty bytes', async () => {
    await expect(docxToDocument(new Uint8Array()))
      .rejects.toThrow('not a Word document (.docx) package')
  })

  it('rejects a ZIP package that is not a Word document', async () => {
    const xlsxBytes = fixtureXlsxBytes(new Map([['Sheet1', 'not a docx']])) as Uint8Array<ArrayBuffer>
    await expect(docxToDocument(xlsxBytes)).rejects.toThrow()
  })

  it('converts a real .docx package to a standalone HTML document', async () => {
    const bytes = fixtureDocxBytes('Preview text 中文') as Uint8Array<ArrayBuffer>
    const html = await docxToDocument(bytes)
    expect(html).toContain('<!doctype html>')
    expect(html).toContain('<p>Preview text 中文</p>')
    expect(html).not.toContain('<script')
  })
})

/** Minimal, valid .docx and .xlsx byte fixtures for the local docx/spreadsheet preview unit tests. */
import { strToU8, zipSync } from 'fflate'

const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const CONTENT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'

function pack(files: Record<string, string>): Uint8Array {
  return zipSync(Object.fromEntries(Object.entries(files).map(([path, content]) => [path, strToU8(content)])))
}

/**
 * Build a minimal, valid one-paragraph .docx package.
 * @param text - paragraph text.
 * @returns compressed document bytes readable by mammoth.
 */
export function fixtureDocxBytes(text: string): Uint8Array {
  const main = 'word/document.xml'
  return pack({
    [main]: `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`,
    '_rels/.rels': `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${OFFICE_REL_NS}/officeDocument" Target="${main}"/></Relationships>`,
    '[Content_Types].xml': `<Types xmlns="${CONTENT_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/${main}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  })
}

/**
 * Build a minimal, valid multi-sheet .xlsx workbook, one A1 cell per sheet.
 * @param sheets - ordered sheet name to cell-text pairs; each becomes one worksheet.
 * @returns compressed workbook bytes readable by SheetJS.
 */
export function fixtureXlsxBytes(sheets: ReadonlyMap<string, string>): Uint8Array {
  const entries = [...sheets.entries()]
  const sheetXml = (text: string): string =>
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${text}</t></is></c></row></sheetData></worksheet>`
  const files: Record<string, string> = {
    'xl/workbook.xml': `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${OFFICE_REL_NS}">`
      + `<sheets>${entries.map(([name], index) => `<sheet name="${name}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join('')}</sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${REL_NS}">${entries.map((_, index) =>
      `<Relationship Id="rId${index + 1}" Type="${OFFICE_REL_NS}/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join('')}</Relationships>`,
    '_rels/.rels': `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${OFFICE_REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    '[Content_Types].xml': `<Types xmlns="${CONTENT_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + `${entries.map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`,
  }
  entries.forEach(([, text], index) => { files[`xl/worksheets/sheet${index + 1}.xml`] = sheetXml(text) })
  return pack(files)
}

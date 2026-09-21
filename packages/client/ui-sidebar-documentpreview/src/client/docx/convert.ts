/** Convert complete .docx bytes to a standalone HTML document via mammoth, entirely in the browser. */
import { convertToHtml } from 'mammoth'

/** Every .docx package is a ZIP archive; mammoth's own error otherwise names an unrelated failure. */
const ZIP_SIGNATURE = [0x50, 0x4b]

function looksLikeZip(data: Uint8Array): boolean {
  return data.length >= ZIP_SIGNATURE.length && ZIP_SIGNATURE.every((byte, index) => data[index] === byte)
}

/**
 * Convert Word document bytes to a complete HTML document.
 * @param data - complete .docx file bytes; transient UI input, never persisted or shared.
 * @returns a standalone HTML document string ready for an isolated iframe.
 * @throws when the bytes are not a readable Word document.
 */
export async function docxToDocument(data: Uint8Array<ArrayBuffer>): Promise<string> {
  if (!looksLikeZip(data)) throw new Error('not a Word document (.docx) package')
  // mammoth's Node entry point (the one this client bundle resolves regardless of
  // platform) recognizes `buffer`, not the `arrayBuffer` its browser entry declares;
  // it only forwards the value into JSZip.loadAsync, which accepts a plain Uint8Array,
  // so passing one here never reaches a real Buffer-only method. Default image
  // conversion reads through mammoth's base64 path, so no Node `Buffer` global is
  // needed either. mammoth's own type declares this field as `Buffer`, which this
  // browser-only package's tsconfig has no Node types to resolve, so the parameter
  // type is effectively unchecked here regardless.
  const result = await convertToHtml({ buffer: data })
  return wrapDocxDocument(result.value)
}

function wrapDocxDocument(body: string): string {
  return `<!doctype html><html><head><meta charset="utf-8"><style>${DOCX_DOCUMENT_STYLE}</style></head>`
    + `<body><article>${body}</article></body></html>`
}

const DOCX_DOCUMENT_STYLE = `
  :root { color-scheme: light dark; }
  body { margin: 0; padding: 32px; display: flex; justify-content: center; box-sizing: border-box;
    background: #f3f3f5; font: 15px/1.6 -apple-system, "Segoe UI", sans-serif; color: #1a1a1a; }
  article { max-width: 760px; width: 100%; background: #fff; padding: 48px 56px; box-sizing: border-box;
    box-shadow: 0 1px 6px rgba(0, 0, 0, .18); }
  article :first-child { margin-top: 0; }
  article :last-child { margin-bottom: 0; }
  table { border-collapse: collapse; }
  td, th { border: 1px solid currentColor; padding: 4px 8px; }
  img { max-width: 100%; }
  @media (prefers-color-scheme: dark) {
    body { background: #1c1c1f; color: #e6e6e6; }
    article { background: #26262b; box-shadow: none; }
  }
`

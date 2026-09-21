/** Locale-owned Word document implementation name and frame status text. */
export const zh = {
  title: 'Word 文档',
  frame: 'Word 文档预览',
  loading: '正在读取…',
  failed: '无法预览这份 Word 文档。文件可能已损坏、受密码保护，或扩展名有误。',
} satisfies Record<string, string>

/** Word document renderer dictionary keys. */
export type DocxPreviewKey = keyof typeof zh

/** English dictionary with the same keys as the Chinese dictionary. */
export const en = {
  title: 'Word document',
  frame: 'Word document preview',
  loading: 'Reading…',
  failed: 'This Word document could not be previewed. It may be damaged, password protected, or have the wrong extension.',
} satisfies Record<DocxPreviewKey, string>

declare module '@deepseek-ai/dsh-client-ui-slots' {
  interface LocaleNamespaceMap {
    /** Word document preview selection and status text. */
    documentDocx: DocxPreviewKey
  }
}

/** Locale-owned spreadsheet implementation name and frame status text. */
export const zh = {
  title: '表格',
  frame: '表格预览',
  loading: '正在读取…',
  failed: '无法预览这份表格文件。文件可能已损坏、受密码保护，或扩展名有误。',
  sheetTabs: '工作表',
} satisfies Record<string, string>

/** Spreadsheet renderer dictionary keys. */
export type SpreadsheetPreviewKey = keyof typeof zh

/** English dictionary with the same keys as the Chinese dictionary. */
export const en = {
  title: 'Spreadsheet',
  frame: 'Spreadsheet preview',
  loading: 'Reading…',
  failed: 'This spreadsheet could not be previewed. It may be damaged, password protected, or have the wrong extension.',
  sheetTabs: 'Sheets',
} satisfies Record<SpreadsheetPreviewKey, string>

declare module '@deepseek-ai/dsh-client-ui-slots' {
  interface LocaleNamespaceMap {
    /** Spreadsheet preview selection, sheet tabs, and status text. */
    documentSpreadsheet: SpreadsheetPreviewKey
  }
}

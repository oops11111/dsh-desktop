/** Builtin spreadsheet metadata and keyed body registration; parsing runs locally, no Host service. */
import type { Context } from '@deepseek-ai/cordis'
import type {} from '../index.ts'
import type { DocumentPreviewDefinition } from '../document/registry.ts'
import { LazySpreadsheetBody } from './LazySpreadsheetBody.tsx'
import { en, zh } from './locales.ts'

/** Spreadsheet implementation identity, shared by metadata and the keyed slot. */
export const SPREADSHEET_BODY_ID = '@deepseek-ai/dsh-client-ui-sidebar-documentpreview/spreadsheet'

/** Extensions this local implementation reads directly. Legacy binary .xls parses through the same SheetJS entry as .xlsx. */
const EXTENSIONS = ['xlsx', 'xls']

/**
 * Describe the local spreadsheet renderer's file types and loading mode.
 *
 * Priority defaults to `extension`, which outranks the Host-conversion Office
 * implementation's `builtin` band, so this local renderer is selected by
 * default wherever both are registered; the dropdown still offers the other
 * implementation when it is available.
 * @param title - locale-owned implementation name.
 * @returns metadata for complete .xlsx and .xls documents.
 */
export function spreadsheetBodyDefinition(title: () => string): DocumentPreviewDefinition {
  return { id: SPREADSHEET_BODY_ID, extensions: EXTENSIONS, binaryExtensions: EXTENSIONS, title, loading: 'bytes-complete', wrap: false }
}

/**
 * Register the spreadsheet dictionary, metadata and body with reversible effects.
 * @param ctx - owning plugin context.
 */
export function apply(ctx: Context): void {
  const t = ctx.locale.bind('documentSpreadsheet')
  ctx.effect(() => ctx.locale.register('documentSpreadsheet', { zh, en }))
  ctx.effect(() => ctx.documentPreviews.register(spreadsheetBodyDefinition(() => t('title'))))
  ctx.effect(() => ctx.slots.inject('sidebar.right.tab.document', () => ctx.slots.register(
    { name: 'sidebar.right.tab.document', key: SPREADSHEET_BODY_ID, locale: 'documentSpreadsheet' },
    LazySpreadsheetBody,
  )))
}

/** Builtin Word-document metadata and keyed body registration; conversion runs locally, no Host service. */
import type { Context } from '@deepseek-ai/cordis'
import type {} from '../index.ts'
import type { DocumentPreviewDefinition } from '../document/registry.ts'
import { LazyDocxBody } from './LazyDocxBody.tsx'
import { en, zh } from './locales.ts'

/** Word document implementation identity, shared by metadata and the keyed slot. */
export const DOCX_BODY_ID = '@deepseek-ai/dsh-client-ui-sidebar-documentpreview/docx'

/**
 * Describe the local Word-document renderer's file type and loading mode.
 *
 * Priority defaults to `extension`, which outranks the Host-conversion Office
 * implementation's `builtin` band, so this local renderer is selected by
 * default wherever both are registered; the dropdown still offers the other
 * implementation when it is available.
 * @param title - locale-owned implementation name.
 * @returns metadata for complete .docx documents.
 */
export function docxBodyDefinition(title: () => string): DocumentPreviewDefinition {
  return { id: DOCX_BODY_ID, extensions: ['docx'], binaryExtensions: ['docx'], title, loading: 'bytes-complete', wrap: false }
}

/**
 * Register the Word-document dictionary, metadata and body with reversible effects.
 * @param ctx - owning plugin context.
 */
export function apply(ctx: Context): void {
  const t = ctx.locale.bind('documentDocx')
  ctx.effect(() => ctx.locale.register('documentDocx', { zh, en }))
  ctx.effect(() => ctx.documentPreviews.register(docxBodyDefinition(() => t('title'))))
  ctx.effect(() => ctx.slots.inject('sidebar.right.tab.document', () => ctx.slots.register(
    { name: 'sidebar.right.tab.document', key: DOCX_BODY_ID, locale: 'documentDocx' },
    LazyDocxBody,
  )))
}

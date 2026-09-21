/** Word document preview: local mammoth conversion rendered in an isolated Blob iframe.
 *
 * The Blob-iframe lifecycle is deliberately not factored out into a module
 * shared with the spreadsheet renderer: this component is loaded only by
 * {@link LazyDocxBody}'s dynamic import, and a module shared between two
 * independently lazy-loaded chunks becomes a chunk of its own that this
 * package's eager bundle ends up requiring too, pulling this renderer's
 * weight back into the startup path it exists to stay out of.
 */
import { useEffect, useState } from 'react'
import type { ReactNode } from 'react'
import { IconLoadingOutline16 } from '@deepseek-ai/dsh-client-ui-primitives'
import type { PropsLocale } from '@deepseek-ai/dsh-client-ui-slots'
import type { DocumentPreviewProps } from '../document/contract.ts'
import { docxToDocument } from './convert.ts'
import type {} from './locales.ts'
import css from './DocxBody.module.css'

/** Standard document inputs plus this renderer's dictionary. */
export type DocxBodyProps = DocumentPreviewProps & PropsLocale<'documentDocx'>

interface ConversionState {
  readonly data: Uint8Array<ArrayBuffer>
  readonly html: string | undefined
  readonly url: string | undefined
  readonly failed: boolean
}

function DocxFrame({ data, t }: { readonly data: Uint8Array<ArrayBuffer>; readonly t: DocxBodyProps['t'] }): ReactNode {
  const [state, setState] = useState<ConversionState>()
  useEffect(() => {
    const controller = new AbortController()
    let url: string | undefined
    void (async () => {
      try {
        const html = await docxToDocument(data)
        if (controller.signal.aborted) return
        url = URL.createObjectURL(new Blob([html], { type: 'text/html' }))
        setState({ data, html, url, failed: false })
      } catch {
        if (!controller.signal.aborted) setState({ data, html: undefined, url: undefined, failed: true })
      }
    })()
    return () => {
      controller.abort()
      if (url !== undefined) URL.revokeObjectURL(url)
    }
  }, [data])

  const current = state?.data === data ? state : undefined
  if (current?.failed === true) return <p className={css.status} role="alert">{t('failed')}</p>
  if (current?.url === undefined) {
    return <span className={`${css.status} ${css.opening} ${css.loading}`} role="status" aria-label={t('loading')} data-document-loading>
      <span className={css.icon} aria-hidden="true"><IconLoadingOutline16 /></span>
    </span>
  }
  return <iframe key={current.url} className={css.frame} src={current.url} sandbox="" title={t('frame')} data-docx-preview />
}

/**
 * Render a complete Word document with the standard file and tab hooks.
 * @param props - document bytes and locale.
 * @returns the converted document in an isolated frame, or nothing for other content kinds.
 */
export function DocxBody({ content, resourceAddress, t }: DocxBodyProps): ReactNode {
  if (content.kind !== 'bytes') return null
  return <DocxFrame key={resourceAddress} data={content.data} t={t} />
}

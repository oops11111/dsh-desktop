/** Spreadsheet preview: local SheetJS parsing, one sheet at a time, rendered in an isolated Blob iframe.
 *
 * The Blob-iframe lifecycle is deliberately not factored out into a module
 * shared with the Word-document renderer: this component is loaded only by
 * {@link LazySpreadsheetBody}'s dynamic import, and a module shared between
 * two independently lazy-loaded chunks becomes a chunk of its own that this
 * package's eager bundle ends up requiring too, pulling this renderer's
 * weight back into the startup path it exists to stay out of.
 */
import { useEffect, useMemo, useState } from 'react'
import type { ReactNode } from 'react'
import type { WorkBook } from 'xlsx'
import type { PropsLocale } from '@deepseek-ai/dsh-client-ui-slots'
import type { DocumentPreviewProps } from '../document/contract.ts'
import { parseSpreadsheet, spreadsheetSheetDocument } from './convert.ts'
import type {} from './locales.ts'
import css from './SpreadsheetBody.module.css'

/** Standard document inputs plus this renderer's dictionary. */
export type SpreadsheetBodyProps = DocumentPreviewProps & PropsLocale<'documentSpreadsheet'>

interface ParsedWorkbook {
  readonly ok: true
  readonly workbook: WorkBook
}

function parse(data: Uint8Array<ArrayBuffer>): ParsedWorkbook | { readonly ok: false } {
  try {
    return { ok: true, workbook: parseSpreadsheet(data) }
  } catch {
    return { ok: false }
  }
}

interface SheetFrameProps {
  readonly html: string | undefined
  readonly failedLabel: string
  readonly frameLabel: string
}

function SheetFrame({ html, failedLabel, frameLabel }: SheetFrameProps): ReactNode {
  const [url, setUrl] = useState<string>()
  useEffect(() => {
    if (html === undefined) {
      setUrl(undefined)
      return
    }
    const objectUrl = URL.createObjectURL(new Blob([html], { type: 'text/html' }))
    setUrl(objectUrl)
    return () => { URL.revokeObjectURL(objectUrl) }
  }, [html])
  if (html === undefined) return <p className={css.status} role="alert">{failedLabel}</p>
  if (url === undefined) return null
  return <iframe key={url} className={css.frame} src={url} sandbox="" title={frameLabel} data-spreadsheet-preview />
}

function SpreadsheetFrame({ data, t }: { readonly data: Uint8Array<ArrayBuffer>; readonly t: SpreadsheetBodyProps['t'] }): ReactNode {
  const parsed = useMemo(() => parse(data), [data])
  const [selected, setSelected] = useState(0)
  const sheetNames = parsed.ok ? parsed.workbook.SheetNames : []
  const activeIndex = sheetNames.length === 0 ? 0 : Math.min(selected, sheetNames.length - 1)
  const activeSheet = sheetNames[activeIndex]
  const html = useMemo(() => {
    if (!parsed.ok || activeSheet === undefined) return undefined
    return spreadsheetSheetDocument(parsed.workbook, activeSheet)
  }, [parsed, activeSheet])

  if (!parsed.ok) return <p className={css.status} role="alert">{t('failed')}</p>
  return <div className={css.frameArea}>
    {sheetNames.length > 1 && <div className={css.tabs} role="tablist" aria-label={t('sheetTabs')}>
      {sheetNames.map((name, index) => (
        <button
          key={name} type="button" role="tab" aria-selected={index === activeIndex}
          className={index === activeIndex ? `${css.tab} ${css.tabActive}` : css.tab}
          onClick={() => { setSelected(index) }}
        >
          {name}
        </button>
      ))}
    </div>}
    <SheetFrame html={html} failedLabel={t('failed')} frameLabel={t('frame')} />
  </div>
}

/**
 * Render a spreadsheet's active sheet with the standard file and tab hooks.
 * @param props - document bytes and locale.
 * @returns the converted sheet in an isolated frame with sheet tabs, or nothing for other content kinds.
 */
export function SpreadsheetBody({ content, resourceAddress, t }: SpreadsheetBodyProps): ReactNode {
  if (content.kind !== 'bytes') return null
  return <SpreadsheetFrame key={resourceAddress} data={content.data} t={t} />
}

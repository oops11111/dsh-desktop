/** Load the spreadsheet renderer, and SheetJS, only after a spreadsheet body is mounted. */
import { lazy, Suspense, type ReactNode } from 'react'
import { LoadingIndicator } from '../LoadingIndicator.tsx'
import css from '../TextPreview.module.css'
import type { SpreadsheetBodyProps } from './SpreadsheetBody.tsx'

const LoadedSpreadsheetBody = lazy(async () => ({ default: (await import('./SpreadsheetBody.tsx')).SpreadsheetBody }))

/**
 * Suspend while the package-local spreadsheet chunk arrives.
 * @param props - spreadsheet body props supplied by the document slot.
 * @returns the deferred spreadsheet renderer.
 */
export function LazySpreadsheetBody(props: SpreadsheetBodyProps): ReactNode {
  return <Suspense fallback={<LoadingIndicator className={css.status} label={props.t('loading')} />}>
    <LoadedSpreadsheetBody {...props} />
  </Suspense>
}

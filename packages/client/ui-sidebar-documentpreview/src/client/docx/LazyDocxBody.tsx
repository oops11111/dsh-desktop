/** Load the Word-document renderer, and mammoth, only after a .docx body is mounted. */
import { lazy, Suspense, type ReactNode } from 'react'
import { LoadingIndicator } from '../LoadingIndicator.tsx'
import css from '../TextPreview.module.css'
import type { DocxBodyProps } from './DocxBody.tsx'

const LoadedDocxBody = lazy(async () => ({ default: (await import('./DocxBody.tsx')).DocxBody }))

/**
 * Suspend while the package-local Word-document chunk arrives.
 * @param props - Word-document body props supplied by the document slot.
 * @returns the deferred Word-document renderer.
 */
export function LazyDocxBody(props: DocxBodyProps): ReactNode {
  return <Suspense fallback={<LoadingIndicator className={css.status} label={props.t('loading')} />}>
    <LoadedDocxBody {...props} />
  </Suspense>
}

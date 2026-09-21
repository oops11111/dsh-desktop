// @vitest-environment jsdom
/** LazyDocxBody shows loading feedback until the package-local .docx chunk resolves. */
import { createElement } from 'react'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, render, screen, waitFor } from '@testing-library/react'
import { makeTranslate } from '@deepseek-ai/dsh-client-test-runtime'
import { en } from '../src/client/docx/locales.ts'

const renderedDocx = vi.hoisted(() => vi.fn(() => null))
vi.mock('../src/client/docx/DocxBody.tsx', () => ({ DocxBody: renderedDocx }))
import { LazyDocxBody } from '../src/client/docx/LazyDocxBody.tsx'

afterEach(() => { cleanup(); renderedDocx.mockClear() })

describe('LazyDocxBody', () => {
  it('shows localized loading feedback while the Word-document chunk resolves', async () => {
    render(createElement(LazyDocxBody, { t: makeTranslate(en) } as never))
    expect(screen.getByRole('status', { name: en.loading })).toBeDefined()
    await waitFor(() => { expect(renderedDocx).toHaveBeenCalledOnce() })
  })
})

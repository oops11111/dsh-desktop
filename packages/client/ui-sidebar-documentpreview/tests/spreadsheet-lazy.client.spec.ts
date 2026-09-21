// @vitest-environment jsdom
/** LazySpreadsheetBody shows loading feedback until the package-local spreadsheet chunk resolves. */
import { createElement } from 'react'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, render, screen, waitFor } from '@testing-library/react'
import { makeTranslate } from '@deepseek-ai/dsh-client-test-runtime'
import { en } from '../src/client/spreadsheet/locales.ts'

const renderedSpreadsheet = vi.hoisted(() => vi.fn(() => null))
vi.mock('../src/client/spreadsheet/SpreadsheetBody.tsx', () => ({ SpreadsheetBody: renderedSpreadsheet }))
import { LazySpreadsheetBody } from '../src/client/spreadsheet/LazySpreadsheetBody.tsx'

afterEach(() => { cleanup(); renderedSpreadsheet.mockClear() })

describe('LazySpreadsheetBody', () => {
  it('shows localized loading feedback while the spreadsheet chunk resolves', async () => {
    render(createElement(LazySpreadsheetBody, { t: makeTranslate(en) } as never))
    expect(screen.getByRole('status', { name: en.loading })).toBeDefined()
    await waitFor(() => { expect(renderedSpreadsheet).toHaveBeenCalledOnce() })
  })
})

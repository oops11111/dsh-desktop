// @vitest-environment jsdom
/** SpreadsheetBody parses real .xlsx bytes through SheetJS and renders the active sheet in an isolated Blob iframe. */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { cleanup, fireEvent, render, screen } from '@testing-library/react'
import type { SessionId } from '@deepseek-ai/dsh-session/types'
import { SpreadsheetBody } from '../src/client/spreadsheet/SpreadsheetBody.tsx'
import type { SpreadsheetBodyProps } from '../src/client/spreadsheet/SpreadsheetBody.tsx'
import { en } from '../src/client/spreadsheet/locales.ts'
import { fixtureXlsxBytes } from './office-local-fixture.ts'

const translations: ReadonlyMap<string, string> = new Map(Object.entries(en))
let createDescriptor: PropertyDescriptor | undefined
let revokeDescriptor: PropertyDescriptor | undefined
const create = vi.fn<(blob: Blob) => string>()
const revoke = vi.fn<(url: string) => void>()

beforeEach(() => {
  createDescriptor = Object.getOwnPropertyDescriptor(URL, 'createObjectURL')
  revokeDescriptor = Object.getOwnPropertyDescriptor(URL, 'revokeObjectURL')
  create.mockReset().mockImplementation(() => `blob:https://preview.invalid/${create.mock.calls.length}`)
  revoke.mockReset()
  Object.defineProperty(URL, 'createObjectURL', { configurable: true, value: create })
  Object.defineProperty(URL, 'revokeObjectURL', { configurable: true, value: revoke })
})

afterEach(() => {
  try { cleanup() } finally {
    if (createDescriptor === undefined) Reflect.deleteProperty(URL, 'createObjectURL')
    else Object.defineProperty(URL, 'createObjectURL', createDescriptor)
    if (revokeDescriptor === undefined) Reflect.deleteProperty(URL, 'revokeObjectURL')
    else Object.defineProperty(URL, 'revokeObjectURL', revokeDescriptor)
  }
})

function props(data: Uint8Array<ArrayBuffer>, address = 'dsh-resource://file/session/xlsx/book.xlsx'): SpreadsheetBodyProps {
  const signal = new AbortController().signal
  return {
    resourceAddress: address,
    content: { kind: 'bytes', data },
    wrap: false,
    sessionId: 'xlsx' as SessionId,
    useTabInfo: () => ({ tab: { signal } }),
    useResource: () => ({ value: undefined }),
    t: key => translations.get(key) ?? key,
  } as SpreadsheetBodyProps
}

describe('SpreadsheetBody', () => {
  it('renders a single-sheet workbook with no sheet tabs', async () => {
    const bytes = fixtureXlsxBytes(new Map([['Only', 'cell text']])) as Uint8Array<ArrayBuffer>
    const view = render(<SpreadsheetBody {...props(bytes)} />)
    const iframe = await screen.findByTitle(en.frame)
    expect(iframe.getAttribute('sandbox')).toBe('')
    expect(iframe.hasAttribute('data-spreadsheet-preview')).toBe(true)
    expect(screen.queryAllByRole('tab')).toHaveLength(0)
    view.unmount()
    expect(revoke).toHaveBeenCalledExactlyOnceWith('blob:https://preview.invalid/1')
  })

  it('shows sheet tabs for a multi-sheet workbook and switches the rendered sheet without a new file read', async () => {
    const bytes = fixtureXlsxBytes(new Map([['First', 'first value'], ['Second', 'second value']])) as Uint8Array<ArrayBuffer>
    render(<SpreadsheetBody {...props(bytes)} />)
    await screen.findByTitle(en.frame)
    const tabs = screen.getAllByRole('tab')
    expect(tabs.map(tab => tab.textContent)).toEqual(['First', 'Second'])
    expect(tabs[0]?.getAttribute('aria-selected')).toBe('true')
    expect(tabs[1]?.getAttribute('aria-selected')).toBe('false')
    expect(create).toHaveBeenCalledOnce()
    fireEvent.click(tabs[1]!)
    await screen.findByTitle(en.frame)
    expect(screen.getAllByRole('tab')[1]?.getAttribute('aria-selected')).toBe('true')
    expect(create).toHaveBeenCalledTimes(2)
  })

  it('reports a read-failure message for bytes that are not a spreadsheet, without leaving a frame', async () => {
    const invalid = new TextEncoder().encode('not a spreadsheet')
    render(<SpreadsheetBody {...props(invalid)} />)
    expect((await screen.findByRole('alert')).textContent).toBe(en.failed)
    expect(screen.queryByTitle(en.frame)).toBeNull()
    expect(create).not.toHaveBeenCalled()
  })

  it('reports a read-failure message for a workbook that parses with zero sheets', async () => {
    const empty = fixtureXlsxBytes(new Map()) as Uint8Array<ArrayBuffer>
    render(<SpreadsheetBody {...props(empty)} />)
    expect((await screen.findByRole('alert')).textContent).toBe(en.failed)
    expect(screen.queryByTitle(en.frame)).toBeNull()
    expect(create).not.toHaveBeenCalled()
  })

  it('renders nothing for a non-bytes content delivery', () => {
    const bytes = fixtureXlsxBytes(new Map([['Only', 'unused']])) as Uint8Array<ArrayBuffer>
    const base = props(bytes)
    const view = render(<SpreadsheetBody {...base} content={{ kind: 'text', text: 'plain', pages: [], eof: true }} />)
    expect(view.container.firstChild).toBeNull()
    expect(create).not.toHaveBeenCalled()
  })
})

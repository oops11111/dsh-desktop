/** Spreadsheet metadata and keyed slot contributions share one identity and unwind with their fiber. */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { Context } from '@deepseek-ai/cordis'
import { DocumentPreviewRegistry } from '../src/client/document/registry.ts'
import { apply, SPREADSHEET_BODY_ID, spreadsheetBodyDefinition } from '../src/client/spreadsheet/index.ts'
import { LazySpreadsheetBody } from '../src/client/spreadsheet/LazySpreadsheetBody.tsx'
import { en, zh } from '../src/client/spreadsheet/locales.ts'

type Registration = { name: string; key: string; locale: string }

let dispose: (() => Promise<void>) | undefined
afterEach(async () => { await dispose?.(); dispose = undefined })

describe('Spreadsheet registration', () => {
  it('claims the .xlsx and .xls suffixes as binary complete-byte renderers without wrap, at extension priority', () => {
    const title = vi.fn(() => 'localized spreadsheet')
    expect(spreadsheetBodyDefinition(title)).toMatchObject({
      id: SPREADSHEET_BODY_ID, extensions: ['xlsx', 'xls'], binaryExtensions: ['xlsx', 'xls'],
      title, loading: 'bytes-complete', wrap: false,
    })
    expect(spreadsheetBodyDefinition(title).priority).toBeUndefined()
    expect(title).not.toHaveBeenCalled()
    expect(spreadsheetBodyDefinition(title).title()).toBe('localized spreadsheet')
  })

  it('registers its dictionary and matching keyed body, and removes all contributions on disposal', async () => {
    const ctx = new Context()
    const registry = new DocumentPreviewRegistry()
    const dictionaries = new Map<string, unknown>()
    const bodies = new Map<string, unknown>()
    const register = vi.fn((options: Registration, body: unknown) => {
      bodies.set(options.key, body)
      return () => { bodies.delete(options.key) }
    })
    ctx.provide('documentPreviews', registry)
    ctx.provide('slots', { inject: (_key: string, callback: () => () => void) => callback(), register } as never)
    ctx.provide('locale', {
      bind: () => (key: keyof typeof en) => en[key],
      register: (name: string, value: unknown) => { dictionaries.set(name, value); return () => { dictionaries.delete(name) } },
    } as never)
    const fiber = ctx.plugin({ apply })
    dispose = async () => { await fiber.dispose() }
    await fiber.await()
    expect(registry.candidates('book.XLSX').map(entry => entry.id)).toEqual([SPREADSHEET_BODY_ID])
    expect(registry.candidates('book.xls').map(entry => entry.id)).toEqual([SPREADSHEET_BODY_ID])
    expect(registry.getSnapshot()[0]?.title()).toBe(en.title)
    expect(dictionaries.get('documentSpreadsheet')).toEqual({ zh, en })
    expect(register).toHaveBeenCalledOnce()
    const registration = register.mock.calls[0]?.[0]
    expect(registration).toMatchObject({ name: 'sidebar.right.tab.document', key: SPREADSHEET_BODY_ID, locale: 'documentSpreadsheet' })
    expect(register.mock.calls[0]?.[1]).toBe(LazySpreadsheetBody)
    await dispose()
    expect(registry.getSnapshot()).toEqual([])
    expect(bodies.size).toBe(0)
    expect(dictionaries.size).toBe(0)
  })
})

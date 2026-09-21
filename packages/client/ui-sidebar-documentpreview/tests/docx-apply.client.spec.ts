/** Word-document metadata and keyed slot contributions share one identity and unwind with their fiber. */
import { afterEach, describe, expect, it, vi } from 'vitest'
import { Context } from '@deepseek-ai/cordis'
import { DocumentPreviewRegistry } from '../src/client/document/registry.ts'
import { apply, DOCX_BODY_ID, docxBodyDefinition } from '../src/client/docx/index.ts'
import { LazyDocxBody } from '../src/client/docx/LazyDocxBody.tsx'
import { en, zh } from '../src/client/docx/locales.ts'

type Registration = { name: string; key: string; locale: string }

let dispose: (() => Promise<void>) | undefined
afterEach(async () => { await dispose?.(); dispose = undefined })

describe('Word document registration', () => {
  it('claims the .docx suffix as a binary complete-byte renderer without wrap, at extension priority', () => {
    const title = vi.fn(() => 'localized Word document')
    expect(docxBodyDefinition(title)).toMatchObject({
      id: DOCX_BODY_ID, extensions: ['docx'], binaryExtensions: ['docx'], title, loading: 'bytes-complete', wrap: false,
    })
    expect(docxBodyDefinition(title).priority).toBeUndefined()
    expect(title).not.toHaveBeenCalled()
    expect(docxBodyDefinition(title).title()).toBe('localized Word document')
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
    expect(registry.candidates('report.DOCX').map(entry => entry.id)).toEqual([DOCX_BODY_ID])
    expect(registry.getSnapshot()[0]?.title()).toBe(en.title)
    expect(dictionaries.get('documentDocx')).toEqual({ zh, en })
    expect(register).toHaveBeenCalledOnce()
    const registration = register.mock.calls[0]?.[0]
    expect(registration).toMatchObject({ name: 'sidebar.right.tab.document', key: DOCX_BODY_ID, locale: 'documentDocx' })
    expect(register.mock.calls[0]?.[1]).toBe(LazyDocxBody)
    await dispose()
    expect(registry.getSnapshot()).toEqual([])
    expect(bodies.size).toBe(0)
    expect(dictionaries.size).toBe(0)
  })
})

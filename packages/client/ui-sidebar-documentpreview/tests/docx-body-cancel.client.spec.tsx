// @vitest-environment jsdom
/** DocxFrame ignores a stale conversion result once the source bytes change before it settles. */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, cleanup, render, screen } from '@testing-library/react'
import type { SessionId } from '@deepseek-ai/dsh-session/types'

const convert = vi.hoisted(() => ({ docxToDocument: vi.fn<(data: Uint8Array<ArrayBuffer>) => Promise<string>>() }))
vi.mock('../src/client/docx/convert.ts', () => convert)
import { DocxBody } from '../src/client/docx/DocxBody.tsx'
import type { DocxBodyProps } from '../src/client/docx/DocxBody.tsx'
import { en } from '../src/client/docx/locales.ts'

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
  convert.docxToDocument.mockReset()
})

afterEach(() => {
  cleanup()
  if (createDescriptor === undefined) Reflect.deleteProperty(URL, 'createObjectURL')
  else Object.defineProperty(URL, 'createObjectURL', createDescriptor)
  if (revokeDescriptor === undefined) Reflect.deleteProperty(URL, 'revokeObjectURL')
  else Object.defineProperty(URL, 'revokeObjectURL', revokeDescriptor)
})

function props(data: Uint8Array<ArrayBuffer>): DocxBodyProps {
  const signal = new AbortController().signal
  return {
    resourceAddress: 'dsh-resource://file/session/docx/report.docx',
    content: { kind: 'bytes', data },
    wrap: false,
    sessionId: 'docx' as SessionId,
    useTabInfo: () => ({ tab: { signal } }),
    useResource: () => ({ value: undefined }),
    t: key => translations.get(key) ?? key,
  } as DocxBodyProps
}

describe('DocxBody cancellation', () => {
  it('ignores a stale resolved conversion once a newer file starts converting', async () => {
    const stale = Promise.withResolvers<string>()
    const fresh = Promise.withResolvers<string>()
    convert.docxToDocument.mockReturnValueOnce(stale.promise).mockReturnValueOnce(fresh.promise)
    const first = new Uint8Array([1])
    const second = new Uint8Array([2])
    const view = render(<DocxBody {...props(first)} />)
    view.rerender(<DocxBody {...props(second)} />)
    await act(async () => { stale.resolve('<p>stale</p>') })
    expect(create).not.toHaveBeenCalled()
    expect(screen.queryByTitle(en.frame)).toBeNull()
    await act(async () => { fresh.resolve('<p>fresh</p>') })
    expect(await screen.findByTitle(en.frame)).toBeDefined()
    expect(create).toHaveBeenCalledOnce()
  })

  it('ignores a stale rejected conversion once a newer file starts converting', async () => {
    const stale = Promise.withResolvers<string>()
    const fresh = Promise.withResolvers<string>()
    convert.docxToDocument.mockReturnValueOnce(stale.promise).mockReturnValueOnce(fresh.promise)
    const first = new Uint8Array([1])
    const second = new Uint8Array([2])
    const view = render(<DocxBody {...props(first)} />)
    view.rerender(<DocxBody {...props(second)} />)
    await act(async () => { stale.reject(new Error('stale failure')) })
    expect(screen.queryByRole('alert')).toBeNull()
    await act(async () => { fresh.resolve('<p>fresh</p>') })
    expect(await screen.findByTitle(en.frame)).toBeDefined()
  })
})

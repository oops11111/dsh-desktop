// @vitest-environment jsdom
/** DocxBody converts real .docx bytes through mammoth and renders the result in an isolated Blob iframe. */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { cleanup, render, screen } from '@testing-library/react'
import type { SessionId } from '@deepseek-ai/dsh-session/types'
import { DocxBody } from '../src/client/docx/DocxBody.tsx'
import type { DocxBodyProps } from '../src/client/docx/DocxBody.tsx'
import { en } from '../src/client/docx/locales.ts'
import { fixtureDocxBytes } from './office-local-fixture.ts'

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

function props(data: Uint8Array<ArrayBuffer>, address = 'dsh-resource://file/session/docx/report.docx'): DocxBodyProps {
  const signal = new AbortController().signal
  return {
    resourceAddress: address,
    content: { kind: 'bytes', data },
    wrap: false,
    sessionId: 'docx' as SessionId,
    useTabInfo: () => ({ tab: { signal } }),
    useResource: () => ({ value: undefined }),
    t: key => translations.get(key) ?? key,
  } as DocxBodyProps
}

describe('DocxBody', () => {
  it('converts a real .docx package and renders it in a script-denying Blob iframe', async () => {
    const view = render(<DocxBody {...props(fixtureDocxBytes('Preview text') as Uint8Array<ArrayBuffer>)} />)
    expect(screen.getByRole('status', { name: en.loading })).toBeDefined()
    const iframe = await screen.findByTitle(en.frame)
    expect(iframe.getAttribute('sandbox')).toBe('')
    expect(iframe.hasAttribute('data-docx-preview')).toBe(true)
    expect(create.mock.calls[0]?.[0].type).toBe('text/html')
    view.unmount()
    expect(revoke).toHaveBeenCalledExactlyOnceWith('blob:https://preview.invalid/1')
  })

  it('reports a read-failure message for bytes that are not a Word document, without leaving a frame', async () => {
    const invalid = new TextEncoder().encode('not a docx')
    render(<DocxBody {...props(invalid)} />)
    expect((await screen.findByRole('alert')).textContent).toBe(en.failed)
    expect(screen.queryByTitle(en.frame)).toBeNull()
    expect(create).not.toHaveBeenCalled()
  })

  it('replaces the frame and revokes the previous Blob when the file identity changes', async () => {
    const view = render(<DocxBody {...props(fixtureDocxBytes('first') as Uint8Array<ArrayBuffer>)} />)
    const first = await screen.findByTitle(en.frame)
    view.rerender(<DocxBody {...props(fixtureDocxBytes('second') as Uint8Array<ArrayBuffer>, 'dsh-resource://file/session/docx/other.docx')} />)
    expect(await screen.findByTitle(en.frame)).not.toBe(first)
    expect(revoke).toHaveBeenCalledWith('blob:https://preview.invalid/1')
  })

  it('renders nothing for a non-bytes content delivery', () => {
    const base = props(fixtureDocxBytes('unused') as Uint8Array<ArrayBuffer>)
    const view = render(<DocxBody {...base} content={{ kind: 'text', text: 'plain', pages: [], eof: true }} />)
    expect(view.container.firstChild).toBeNull()
    expect(create).not.toHaveBeenCalled()
  })
})

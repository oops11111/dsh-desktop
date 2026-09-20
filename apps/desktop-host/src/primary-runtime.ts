/** Offline installation and absolute paths for Desktop's bundled script dependencies. */

import { chmod, cp, lstat, mkdir, mkdtemp, readFile, rename, rm, stat, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'

/** Versions recorded by the Desktop build, independent of user-installed packages. */
export interface PrimaryRuntimeManifest {
  readonly desktopVersion: string
  readonly platform: string
  readonly arch: string
  /** Locked payload identity; absent only in installations made before payload hashing. */
  readonly payloadDigest?: string
  /** Installed wheel distribution versions; absent in older release manifests. */
  readonly pythonPackages?: Readonly<Record<string, string>>
  readonly components: {
    readonly python: string
    readonly pnpm: string
  }
}

/** Absolute entry points and bundled versions; pnpm runs through the returned Node executable. */
export interface WorkspaceDependencies {
  readonly python: string
  readonly node: string
  readonly pnpm: string
  readonly pythonPackages: string
  readonly nodePackages: string
  /** Locked distribution versions; excludes packages users add to the installed environment. */
  readonly pythonDistributions: Readonly<Record<string, string>>
}

/**
 * Read build metadata, rejecting duplicate normalized names and conflicting component/distribution versions.
 * @param root - Installed or bundled primary runtime directory.
 * @returns Validated component versions and target identifiers.
 */
export async function readPrimaryRuntime(root: string): Promise<PrimaryRuntimeManifest> {
  const value: unknown = JSON.parse(await readFile(join(root, 'runtime.json'), 'utf8'))
  if (typeof value !== 'object' || value === null) throw new Error('primary runtime: invalid metadata')
  const record = value as Record<string, unknown>
  const components = record.components
  const packages = record.pythonPackages
  if (typeof record.desktopVersion !== 'string' || record.desktopVersion.length === 0
    || !['win32', 'darwin'].includes(String(record.platform)) || !['x64', 'arm64'].includes(String(record.arch))
    || typeof components !== 'object' || components === null
    || !['python', 'pnpm'].every(key => /^\d+\.\d+\.\d+(?:[-+][\w.-]+)?$/u.test(String((components as Record<string, unknown>)[key])))
    || (record.payloadDigest !== undefined && (typeof record.payloadDigest !== 'string' || !/^[a-f0-9]{64}$/u.test(record.payloadDigest)))
    || (packages !== undefined && (typeof packages !== 'object' || packages === null || Array.isArray(packages)
      || !Object.entries(packages).every(([name, version]) => /^[A-Za-z0-9][A-Za-z0-9._-]*$/u.test(name)
        && typeof version === 'string' && /^\d[\w.!+-]*$/u.test(version))))) {
    throw new Error('primary runtime: invalid metadata')
  }
  const manifest = value as PrimaryRuntimeManifest
  const entries = Object.entries(manifest.pythonPackages ?? {})
  const distributions = new Set(entries.map(([name]) => name.toLowerCase().replace(/[-_.]+/gu, '-')))
  if (distributions.size !== entries.length) throw new Error('primary runtime: invalid metadata')
  return manifest
}

/**
 * Resolve platform-specific interpreter and library locations without changing the environment.
 * @param root - Absolute installation directory.
 * @param manifest - Validated runtime metadata.
 * @returns Absolute paths for explicit script execution and recorded bundled Python versions.
 */
export function workspaceDependencyPaths(root: string, manifest: PrimaryRuntimeManifest): WorkspaceDependencies {
  const dependencies = join(root, 'dependencies')
  const windows = manifest.platform === 'win32'
  return {
    python: join(dependencies, 'python', ...(windows ? ['python.exe'] : ['bin', 'python3'])),
    node: join(dependencies, 'node', 'bin', windows ? 'node.cmd' : 'node'),
    pnpm: join(dependencies, 'pnpm', 'bin', 'pnpm.mjs'),
    pythonPackages: join(dependencies, 'python', ...(windows ? ['Lib'] : ['lib', `python${manifest.components.python.split('.').slice(0, 2).join('.')}`]), 'site-packages'),
    nodePackages: join(dependencies, 'node', 'node_modules'),
    pythonDistributions: manifest.pythonPackages ?? {},
  }
}

/**
 * Write the Node launcher that forwards to the Electron executable running this Host in Node mode.
 * Rewritten on every installation because the executable path moves with the application, while a
 * matching payload skips the copy that would otherwise refresh it.
 * @param paths - Resolved entry points inside the staged or installed payload.
 * @param windows - Whether the payload targets Windows command files.
 */
async function writeNodeLauncher(paths: WorkspaceDependencies, windows: boolean): Promise<void> {
  await writeFile(paths.node, windows
    ? `@echo off\r\nsetlocal\r\nset "ELECTRON_RUN_AS_NODE=1"\r\n"${process.execPath}" %*\r\n`
    : `#!/bin/sh\nELECTRON_RUN_AS_NODE=1 exec "${process.execPath}" "$@"\n`)
  if (!windows) await chmod(paths.node, 0o755)
}

async function exists(path: string): Promise<boolean> {
  try {
    if ((await lstat(path)).isSymbolicLink()) throw new Error(`primary runtime: installation path is a filesystem link: ${path}`)
    return true
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return false
    throw error
  }
}

/**
 * Install the application-owned payload locally, retaining a complete previous tree on copy failure.
 * @param source - Payload carried by the current Desktop installation.
 * @param root - Fixed primary runtime directory under the Harness home.
 * @returns Paths into the installed payload; no PATH or package-manager configuration is changed.
 */
export async function installPrimaryRuntime(source: string, root: string): Promise<WorkspaceDependencies> {
  const manifest = await readPrimaryRuntime(source)
  if (manifest.platform !== process.platform || manifest.arch !== process.arch) throw new Error('primary runtime: incompatible platform or architecture')
  await mkdir(dirname(root), { recursive: true })
  const previous = `${root}.previous`
  await exists(previous)
  await exists(root)
  if (!await exists(root) && await exists(previous)) await rename(previous, root)
  if (await exists(join(root, 'runtime.json')) && JSON.stringify(await readPrimaryRuntime(root)) === JSON.stringify(manifest)) {
    const paths = workspaceDependencyPaths(root, manifest)
    await writeNodeLauncher(paths, manifest.platform === 'win32')
    for (const path of [paths.python, paths.node, paths.pnpm, paths.pythonPackages, paths.nodePackages]) await stat(path)
    return paths
  }
  const staging = await mkdtemp(join(dirname(root), '.primary-runtime-'))
  try {
    await cp(source, staging, { recursive: true, dereference: true })
    const paths = workspaceDependencyPaths(staging, manifest)
    await writeNodeLauncher(paths, manifest.platform === 'win32')
    for (const path of [paths.python, paths.node, paths.pnpm, paths.pythonPackages, paths.nodePackages]) await stat(path)
    await rm(previous, { recursive: true, force: true })
    const replacing = await exists(root)
    if (replacing) await rename(root, previous)
    try { await rename(staging, root) } catch (error) {
      if (replacing) await rename(previous, root)
      throw error
    }
    await rm(previous, { recursive: true, force: true })
  } finally {
    await rm(staging, { recursive: true, force: true })
  }
  return workspaceDependencyPaths(root, manifest)
}

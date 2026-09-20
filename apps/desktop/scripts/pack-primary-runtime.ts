/** Compress the primary runtime for packaging; run only after any required code-signing pass. */
import { join, resolve } from 'node:path'
import { packPrimaryRuntime } from './prepare-primary-runtime.ts'
import { resolveDesktopTargetBuildPaths } from './desktop-build-paths.mjs'

if (process.argv[1] !== undefined && resolve(process.argv[1]) === import.meta.filename) {
  await packPrimaryRuntime(join(resolveDesktopTargetBuildPaths().runtime, 'primary-runtime'))
}

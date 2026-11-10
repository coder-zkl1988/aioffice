/**
 * Compression feedback: a rebuild that does not make the file smaller leaves the
 * user with the same bytes under a new name, so the UI tells them that the run
 * happened but the file cannot be compressed further.
 */

/**
 * True when compression actually reduced the file. Equal sizes count as no
 * shrink: the output is not smaller, so the notice applies.
 */
export function compressionShrank(before: number, after: number): boolean {
  return after < before
}

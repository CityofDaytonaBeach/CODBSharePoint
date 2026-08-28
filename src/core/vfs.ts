// ============================================================================
// VFS - Virtual File System
// Browser-native file system for SPFx project generation
// ============================================================================

import type { VFSFile, VFS } from '../types/index.js';
import { zipSync, strToU8 } from 'fflate';

export type { VFS };

// Simple glob-to-regex converter
function globToRegex(pattern: string): RegExp {
  const escaped = pattern
    .replace(/[.+^${}()|[\]\\]/g, '\\$&')
    .replace(/\*\*/g, '{{GLOBSTAR}}')
    .replace(/\*/g, '[^/]*')
    .replace(/\?/g, '[^/]')
    .replace(/\{\{GLOBSTAR\}\}/g, '.*');
  return new RegExp(`^${escaped}$`);
}

export function createVFS(): VFS {
  const files = new Map<string, VFSFile>();

  function normalizePath(path: string): string {
    return path
      .replace(/\\/g, '/')
      .replace(/\/+/g, '/')
      .replace(/^\.\//, '')
      .replace(/\/$/, '');
  }

  function addFile(path: string, content: string | Uint8Array, encoding: 'utf-8' | 'binary' = 'utf-8'): void {
    const normalized = normalizePath(path);
    files.set(normalized, {
      path: normalized,
      content,
      encoding,
      timestamp: Date.now()
    });
  }

  function getFile(path: string): VFSFile | undefined {
    return files.get(normalizePath(path));
  }

  function removeFile(path: string): void {
    files.delete(normalizePath(path));
  }

  function getFiles(): VFSFile[] {
    return Array.from(files.values()).sort((a, b) => a.path.localeCompare(b.path));
  }

  function getFilesByPattern(pattern: string | RegExp): VFSFile[] {
    const regex = typeof pattern === 'string' ? globToRegex(pattern) : pattern;
    return getFiles().filter(f => regex.test(f.path));
  }

  function hasFile(path: string): boolean {
    return files.has(normalizePath(path));
  }

  function readAsString(path: string): string | undefined {
    const file = getFile(path);
    if (!file) return undefined;
    if (typeof file.content === 'string') return file.content;
    return new TextDecoder().decode(file.content);
  }

  function toZip(): Uint8Array {
    const zipData: Record<string, Uint8Array> = {};

    for (const file of files) {
      const [, vfsFile] = file;
      if (typeof vfsFile.content === 'string') {
        zipData[vfsFile.path] = strToU8(vfsFile.content);
      } else {
        zipData[vfsFile.path] = vfsFile.content;
      }
    }

    return zipSync(zipData, { level: 6 });
  }

  return {
    files,
    addFile,
    getFile,
    removeFile,
    getFiles,
    getFilesByPattern,
    hasFile,
    readAsString,
    toZip
  };
}

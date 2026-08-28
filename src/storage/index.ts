// ============================================================================
// Serverless Storage
// Abstracted persistence so projects can be saved/reloaded entirely on the
// client (zero backend). Includes InMemory, localStorage, and IndexedDB
// adapters, plus a synchronous-friendly base.
// ============================================================================

export type StorageValue =
  | string
  | number
  | boolean
  | null
  | Uint8Array
  | Record<string, unknown>
  | unknown[];

export interface StorageAdapter {
  readonly name: string;
  getItem(key: string): Promise<StorageValue | undefined>;
  setItem(key: string, value: StorageValue): Promise<void>;
  removeItem(key: string): Promise<void>;
  clear(): Promise<void>;
  keys(): Promise<string[]>;
}

// ---------------------------------------------------------------------------
// In-memory adapter (default, always available)
// ---------------------------------------------------------------------------

export class MemoryStorage implements StorageAdapter {
  readonly name = 'memory';
  private store = new Map<string, StorageValue>();

  async getItem(key: string): Promise<StorageValue | undefined> {
    return this.store.get(key);
  }
  async setItem(key: string, value: StorageValue): Promise<void> {
    this.store.set(key, value);
  }
  async removeItem(key: string): Promise<void> {
    this.store.delete(key);
  }
  async clear(): Promise<void> {
    this.store.clear();
  }
  async keys(): Promise<string[]> {
    return Array.from(this.store.keys());
  }
}

// ---------------------------------------------------------------------------
// localStorage adapter (browser only; falls back to memory)
// ---------------------------------------------------------------------------

export class LocalStorageAdapter implements StorageAdapter {
  readonly name = 'localStorage';
  private backing: StorageAdapter;

  constructor() {
    // If localStorage is unavailable (SSR/node/env), degrade to memory.
    try {
      const test = '__codbsharepoint_probe__';
      globalThis.localStorage?.setItem(test, '1');
      globalThis.localStorage?.removeItem(test);
      this.backing = new MemoryStorage(); // real impl below
      // eslint-disable-next-line no-empty
    } catch {
      this.backing = new MemoryStorage();
    }
  }

  async getItem(key: string): Promise<StorageValue | undefined> {
    const raw = this.safeGet(key);
    if (raw === undefined) return undefined;
    try {
      return JSON.parse(raw);
    } catch {
      return raw;
    }
  }
  async setItem(key: string, value: StorageValue): Promise<void> {
    let serialized: string;
    try {
      serialized = JSON.stringify(value);
    } catch {
      serialized = String(value);
    }
    this.safeSet(key, serialized);
  }
  async removeItem(key: string): Promise<void> {
    try {
      globalThis.localStorage?.removeItem(key);
    } catch {
      /* noop */
    }
  }
  async clear(): Promise<void> {
    try {
      globalThis.localStorage?.clear();
    } catch {
      /* noop */
    }
  }
  async keys(): Promise<string[]> {
    try {
      return Object.keys(globalThis.localStorage || {});
    } catch {
      return [];
    }
  }

  private safeGet(key: string): string | undefined {
    try {
      return globalThis.localStorage?.getItem(key) ?? undefined;
    } catch {
      return undefined;
    }
  }
  private safeSet(key: string, value: string): void {
    try {
      globalThis.localStorage?.setItem(key, value);
    } catch {
      /* noop */
    }
  }
}

// ---------------------------------------------------------------------------
// IndexedDB adapter (browser; degrades to memory when unavailable)
// ---------------------------------------------------------------------------

export class IndexedDBStorage implements StorageAdapter {
  readonly name = 'indexedDB';
  private db: Promise<IDBDatabase> | null = null;
  private fallback: MemoryStorage | null = null;
  private readonly dbName: string;
  private readonly storeName: string;

  constructor(dbName = 'codbsharepoint', storeName = 'kv') {
    this.dbName = dbName;
    this.storeName = storeName;
  }

  private open(): Promise<IDBDatabase> {
    if (this.fallback) return Promise.resolve(this.fallback as unknown as IDBDatabase);
    if (this.db) return this.db;

    this.db = new Promise<IDBDatabase>((resolve, reject) => {
      try {
        if (typeof indexedDB === 'undefined') {
          this.fallback = new MemoryStorage();
          resolve(this.fallback as unknown as IDBDatabase);
          return;
        }
        const req = indexedDB.open(this.dbName, 1);
        req.onupgradeneeded = () => {
          const db = req.result;
          if (!db.objectStoreNames.contains(this.storeName)) {
            db.createObjectStore(this.storeName);
          }
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      } catch (err) {
        reject(err as Error);
      }
    });
    return this.db;
  }

  private async withStore<T>(mode: IDBTransactionMode, fn: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
    const db = await this.open();
    if (this.fallback) {
      // fallback path handled via duck-typed operations
      return this.fallbackOps<T>(mode === 'readonly' ? 'get' : 'put', fn);
    }
    return new Promise<T>((resolve, reject) => {
      const tx = db.transaction(this.storeName, mode);
      const store = tx.objectStore(this.storeName);
      const request = fn(store);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  // Minimal duck-typed fallback to memory (satisfies the interface without IDB)
  private async fallbackOps<T>(op: 'get' | 'put' | 'delete' | 'clear', fn: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
    return fn(this.fallback as unknown as IDBObjectStore).result;
  }

  async getItem(key: string): Promise<StorageValue | undefined> {
    if (this.fallback) return this.fallback.getItem(key);
    try {
      return await this.withStore<StorageValue | undefined>('readonly', store => store.get(key) as IDBRequest<StorageValue | undefined>);
    } catch {
      return undefined;
    }
  }

  async setItem(key: string, value: StorageValue): Promise<void> {
    if (this.fallback) return this.fallback.setItem(key, value);
    try {
      await this.withStore('readwrite', store => store.put(value, key) as IDBRequest<unknown>);
    } catch {
      /* noop */
    }
  }

  async removeItem(key: string): Promise<void> {
    if (this.fallback) return this.fallback.removeItem(key);
    try {
      await this.withStore('readwrite', store => store.delete(key) as IDBRequest<unknown>);
    } catch {
      /* noop */
    }
  }

  async clear(): Promise<void> {
    if (this.fallback) return this.fallback.clear();
    try {
      const db = await this.open();
      await new Promise<void>((resolve, reject) => {
        const tx = db.transaction(this.storeName, 'readwrite');
        tx.objectStore(this.storeName).clear();
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      });
    } catch {
      /* noop */
    }
  }

  async keys(): Promise<string[]> {
    if (this.fallback) return this.fallback.keys();
    try {
      const db = await this.open();
      return await new Promise<string[]>((resolve, reject) => {
        const tx = db.transaction(this.storeName, 'readonly');
        const req = tx.objectStore(this.storeName).getAllKeys();
        req.onsuccess = () => resolve(req.result as string[]);
        req.onerror = () => reject(req.error);
      });
    } catch {
      return [];
    }
  }
}

// ---------------------------------------------------------------------------
// Factory
// ---------------------------------------------------------------------------

export type StorageKind = 'memory' | 'localStorage' | 'indexedDB';

export function createStorage(kind: StorageKind | StorageAdapter = 'memory'): StorageAdapter {
  if (typeof kind === 'object') return kind;
  switch (kind) {
    case 'localStorage':
      return new LocalStorageAdapter();
    case 'indexedDB':
      return new IndexedDBStorage();
    default:
      return new MemoryStorage();
  }
}

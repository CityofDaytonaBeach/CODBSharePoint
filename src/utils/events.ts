// ============================================================================
// Event Emitter for CODBSharePoint pipeline events
// ============================================================================

import type { CODBEvent, CODBEventHandler } from '../types/index.js';

export class EventEmitter {
  private handlers: Map<string, Set<CODBEventHandler>> = new Map();

  on(handler: CODBEventHandler): () => void {
    const type = '*' as string;
    if (!this.handlers.has(type)) {
      this.handlers.set(type, new Set());
    }
    this.handlers.get(type)!.add(handler);

    // Return unsubscribe function
    return () => {
      this.handlers.get(type)?.delete(handler);
    };
  }

  emit(event: CODBEvent): void {
    // Notify all handlers
    const wildcardHandlers = this.handlers.get('*');
    if (wildcardHandlers) {
      for (const handler of wildcardHandlers) {
        try {
          handler(event);
        } catch (e) {
          console.error('Event handler error:', e);
        }
      }
    }
  }

  removeAllListeners(): void {
    this.handlers.clear();
  }
}

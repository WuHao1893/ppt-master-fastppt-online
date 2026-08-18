import type { EventEnvelope } from '../shared/models.js';
import type { StateStore } from './store.js';
import { nowIso } from './utils.js';

type Listener = (event: EventEnvelope) => void;

export class EventBus {
  private readonly listeners = new Map<string, Set<Listener>>();

  constructor(private readonly store: StateStore) {}

  async publish(input: Omit<EventEnvelope, 'seq' | 'createdAt'>): Promise<EventEnvelope> {
    let event!: EventEnvelope;
    await this.store.update((state) => {
      event = { ...input, seq: ++state.seq, createdAt: nowIso() };
      state.events.push(event);
      if (state.events.length > 5000) state.events.splice(0, state.events.length - 5000);
    });
    const projectListeners = this.listeners.get(input.projectId);
    projectListeners?.forEach((listener) => listener(event));
    return event;
  }

  subscribe(projectId: string, listener: Listener): () => void {
    const listeners = this.listeners.get(projectId) || new Set<Listener>();
    listeners.add(listener);
    this.listeners.set(projectId, listeners);
    return () => {
      listeners.delete(listener);
      if (!listeners.size) this.listeners.delete(projectId);
    };
  }

  history(projectId: string, afterSeq = 0): EventEnvelope[] {
    return this.store.state.events.filter((event) => event.projectId === projectId && event.seq > afterSeq);
  }
}

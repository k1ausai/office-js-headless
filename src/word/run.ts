import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { Body } from "./body";

export type QueuedOperation = () => void;

export class RequestContext {
  readonly document: { body: Body };
  private readonly queue: QueuedOperation[] = [];

  constructor(doc: FlatOpcDocument) {
    this.document = { body: new Body(doc, (op) => this.enqueue(op)) };
  }

  private enqueue(op: QueuedOperation): void {
    this.queue.push(op);
  }

  async sync(): Promise<void> {
    // Drain the queue up front, not iterate-while-mutating — a queued op
    // failing partway through must not silently apply the ops after it, and
    // must not leave the failed op replayed on the next sync() call.
    const pending = this.queue.splice(0, this.queue.length);
    for (const op of pending) {
      op();
    }
  }
}

export async function wordRun<T>(
  doc: FlatOpcDocument,
  callback: (context: RequestContext) => Promise<T>
): Promise<T> {
  const context = new RequestContext(doc);
  return callback(context);
}

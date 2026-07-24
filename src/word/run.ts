import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { SupportedPlatform } from "../office/context";
import { Body } from "./body";
import { Style } from "./style";

export type QueuedOperation = () => void;

// Every shim proxy object that tracks loaded properties (Body, and later
// Range/Paragraph/Table) implements this so RequestContext can snapshot all
// of them after a sync()'s mutation queue runs — not just a single hardcoded
// Body reference, since later receiver types create their own instances
// dynamically (e.g. every body.getRange() call makes a new Range).
export interface Syncable {
  sync(): void;
}

export class RequestContext {
  readonly document: { body: Body; getStyles(): Style[] };
  private readonly queue: QueuedOperation[] = [];
  private readonly syncables: Syncable[] = [];

  constructor(doc: FlatOpcDocument, platform: SupportedPlatform) {
    const body = new Body(
      doc,
      platform,
      (op) => this.enqueue(op),
      (obj) => this.registerSyncable(obj)
    );
    this.document = {
      body,
      // Real Word.Document.getStyles() — not Body.getStyles(), which
      // doesn't exist on the real Word.Body class. Synchronous, like
      // Body.getComments()/.search() — a query, not a mutation.
      getStyles: () =>
        doc.getStyleElements().map((s) => {
          const style = new Style(doc, s);
          this.registerSyncable(style);
          return style;
        }),
    };
    this.registerSyncable(body);
  }

  private enqueue(op: QueuedOperation): void {
    this.queue.push(op);
  }

  registerSyncable(obj: Syncable): void {
    this.syncables.push(obj);
  }

  async sync(): Promise<void> {
    // Drain the queue up front, not iterate-while-mutating — a queued op
    // failing partway through must not silently apply the ops after it, and
    // must not leave the failed op replayed on the next sync() call.
    const pending = this.queue.splice(0, this.queue.length);
    for (const op of pending) {
      op();
    }
    // Loaded properties snapshot AFTER mutations apply — real add-in code
    // relies on load()+sync() in the same batch as a mutation seeing the
    // post-mutation value.
    for (const syncable of this.syncables) {
      syncable.sync();
    }
  }
}

export async function wordRun<T>(
  doc: FlatOpcDocument,
  platform: SupportedPlatform,
  callback: (context: RequestContext) => Promise<T>
): Promise<T> {
  const context = new RequestContext(doc, platform);
  return callback(context);
}

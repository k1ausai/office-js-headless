import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { richApiError } from "./errors";
import { Table } from "./table";
import type { QueuedOperation, Syncable } from "./run";

// Mirrors ParagraphCollection (#14): getFirst() is a synchronous proxy
// factory, not enqueued, since it doesn't mutate the document, and the
// created Table is registered as syncable so a later context.sync() can
// snapshot whatever properties get .load()ed on it. Narrower than
// ParagraphCollection — no getLast() — since nothing in issue #15's scope
// or the driving consumer's call sites asks for it; add it if a real need
// shows up, per the "extend on demand" non-goal.
export class TableCollection {
  constructor(
    private readonly doc: FlatOpcDocument,
    private readonly enqueue: (op: QueuedOperation) => void,
    private readonly registerSyncable: (obj: Syncable) => void
  ) {}

  getFirst(): Table {
    const first = this.doc.getTableElements()[0];
    if (!first) {
      throw richApiError("ItemNotFound", "The requested item was not found in the collection.");
    }
    const table = new Table(this.doc, first, this.enqueue, this.registerSyncable);
    this.registerSyncable(table);
    return table;
  }
}

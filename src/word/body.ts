import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { InsertLocation, InsertLocationValue } from "./insertLocation";
import type { QueuedOperation } from "./run";

export class Body {
  constructor(
    private readonly doc: FlatOpcDocument,
    private readonly enqueue: (op: QueuedOperation) => void
  ) {}

  insertText(text: string, insertLocation: InsertLocationValue): void {
    // Validation is deferred into the queued op, not thrown here at the call
    // site — matches real Office.js, which validates queued operations when
    // the batch is processed at sync(), not synchronously at the call site.
    this.enqueue(() => {
      if (insertLocation !== InsertLocation.end) {
        throw new Error(
          `Body.insertText: InsertLocation "${insertLocation}" is not yet implemented (only "End" is supported so far)`
        );
      }
      this.doc.appendParagraph(text);
    });
  }

  getOoxml(): string {
    return this.doc.getOoxml();
  }
}

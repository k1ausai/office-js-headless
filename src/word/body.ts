import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { InsertLocation, InsertLocationValue } from "./insertLocation";
import { TrackedProperties } from "./proxy";
import type { QueuedOperation } from "./run";

type BodyProperty = "text";

export class Body {
  private readonly tracked = new TrackedProperties();

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

  load(propertyNames: BodyProperty | BodyProperty[]): void {
    this.tracked.load(propertyNames);
  }

  get text(): string {
    return this.tracked.read<string>("text");
  }

  /** Called by RequestContext.sync() after the mutation queue has run. */
  sync(): void {
    this.tracked.sync((name) => this.computeProperty(name));
  }

  private computeProperty(name: string): unknown {
    switch (name) {
      case "text":
        return this.doc.getBodyText();
      default:
        throw new Error(`Body: unknown property "${name}"`);
    }
  }
}

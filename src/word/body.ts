import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { bodyInsertEnd } from "../operations/bodyInsertEnd";
import { bodyInsertStart } from "../operations/bodyInsertStart";
import { bodyReplace } from "../operations/bodyReplace";
import { InsertLocation, InsertLocationValue } from "./insertLocation";
import { TrackedProperties } from "./proxy";
import type { QueuedOperation } from "./run";

type BodyProperty = "text";

// Body has no "existing content" for insertText's Start/End to prepend/append
// into the way Range/Paragraph do — both insertText and insertParagraph
// always create a whole new paragraph, so they share one dispatch table.
const BODY_LOCATION_HANDLERS: Partial<
  Record<InsertLocationValue, (doc: FlatOpcDocument, text: string) => void>
> = {
  [InsertLocation.start]: bodyInsertStart,
  [InsertLocation.end]: bodyInsertEnd,
  [InsertLocation.replace]: bodyReplace,
};

export class Body {
  private readonly tracked = new TrackedProperties();

  constructor(
    private readonly doc: FlatOpcDocument,
    private readonly enqueue: (op: QueuedOperation) => void
  ) {}

  insertText(text: string, insertLocation: InsertLocationValue): void {
    this.insert(text, insertLocation);
  }

  insertParagraph(text: string, insertLocation: InsertLocationValue): void {
    this.insert(text, insertLocation);
  }

  private insert(text: string, insertLocation: InsertLocationValue): void {
    // Validation and dispatch are deferred into the queued op, not resolved
    // here at the call site — matches real Office.js, which validates
    // queued operations when the batch is processed at sync().
    this.enqueue(() => {
      const handler = BODY_LOCATION_HANDLERS[insertLocation];
      if (!handler) {
        throw new Error(
          `Body: InsertLocation "${insertLocation}" is not applicable to Body (only "Start"/"End"/"Replace" apply)`
        );
      }
      handler(this.doc, text);
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

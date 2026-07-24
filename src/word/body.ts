import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { SupportedPlatform } from "../office/context";
import { bodyInsertEnd } from "../operations/bodyInsertEnd";
import { bodyInsertOoxmlEnd } from "../operations/bodyInsertOoxmlEnd";
import { bodyInsertOoxmlReplace } from "../operations/bodyInsertOoxmlReplace";
import { bodyInsertOoxmlStart } from "../operations/bodyInsertOoxmlStart";
import { bodyInsertStart } from "../operations/bodyInsertStart";
import { bodyReplace } from "../operations/bodyReplace";
import { assertOoxmlSupported } from "./errors";
import { InsertLocation, InsertLocationValue } from "./insertLocation";
import { TrackedProperties } from "./proxy";
import type { QueuedOperation } from "./run";

type BodyProperty = "text";
type LocationHandler = (doc: FlatOpcDocument, text: string) => void;

// Body has no "existing content" for insertText's Start/End to prepend/append
// into the way Range/Paragraph do — both insertText and insertParagraph
// always create a whole new paragraph, so they share one dispatch table.
const TEXT_LOCATION_HANDLERS: Partial<Record<InsertLocationValue, LocationHandler>> = {
  [InsertLocation.start]: bodyInsertStart,
  [InsertLocation.end]: bodyInsertEnd,
  [InsertLocation.replace]: bodyReplace,
};

const OOXML_LOCATION_HANDLERS: Partial<Record<InsertLocationValue, LocationHandler>> = {
  [InsertLocation.start]: bodyInsertOoxmlStart,
  [InsertLocation.end]: bodyInsertOoxmlEnd,
  [InsertLocation.replace]: bodyInsertOoxmlReplace,
};

export class Body {
  private readonly tracked = new TrackedProperties();

  constructor(
    private readonly doc: FlatOpcDocument,
    private readonly platform: SupportedPlatform,
    private readonly enqueue: (op: QueuedOperation) => void
  ) {}

  insertText(text: string, insertLocation: InsertLocationValue): void {
    this.dispatch(TEXT_LOCATION_HANDLERS, text, insertLocation, "Body.insertText");
  }

  insertParagraph(text: string, insertLocation: InsertLocationValue): void {
    this.dispatch(TEXT_LOCATION_HANDLERS, text, insertLocation, "Body.insertParagraph");
  }

  insertOoxml(ooxml: string, insertLocation: InsertLocationValue): void {
    // Checked at sync() time, not the call site, matching every other
    // deferred validation in this shim.
    this.enqueue(() => {
      assertOoxmlSupported(this.platform);
      const handler = OOXML_LOCATION_HANDLERS[insertLocation];
      if (!handler) {
        throw new Error(
          `Body: InsertLocation "${insertLocation}" is not applicable to Body (only "Start"/"End"/"Replace" apply)`
        );
      }
      handler(this.doc, ooxml);
    });
  }

  private dispatch(
    handlers: Partial<Record<InsertLocationValue, LocationHandler>>,
    text: string,
    insertLocation: InsertLocationValue,
    methodName: string
  ): void {
    // Validation and dispatch are deferred into the queued op, not resolved
    // here at the call site — matches real Office.js, which validates
    // queued operations when the batch is processed at sync().
    this.enqueue(() => {
      const handler = handlers[insertLocation];
      if (!handler) {
        throw new Error(
          `${methodName}: InsertLocation "${insertLocation}" is not applicable to Body (only "Start"/"End"/"Replace" apply)`
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

import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { paragraphInsertAfter } from "../operations/paragraphInsertAfter";
import { paragraphInsertBefore } from "../operations/paragraphInsertBefore";
import { paragraphInsertEnd } from "../operations/paragraphInsertEnd";
import { paragraphInsertStart } from "../operations/paragraphInsertStart";
import { paragraphReplace } from "../operations/paragraphReplace";
import { paragraphTextInsertEnd } from "../operations/paragraphTextInsertEnd";
import { paragraphTextInsertStart } from "../operations/paragraphTextInsertStart";
import { SupportedPlatform } from "../office/context";
import { InsertLocation, InsertLocationValue } from "./insertLocation";
import { TrackedProperties } from "./proxy";
import { Range } from "./range";
import { RangeLocation, RangeLocationValue } from "./rangeLocation";
import type { QueuedOperation, Syncable } from "./run";

type LocationHandler = (doc: FlatOpcDocument, target: Element, text: string) => void;

// insertText's Start/End insert text WITHIN the paragraph's own content;
// insertParagraph's Start/End always create a whole new paragraph — genuinely
// different operations at the same (receiver, location) pair. Before/After/
// Replace behave identically for both methods.
const TEXT_LOCATION_HANDLERS: Record<InsertLocationValue, LocationHandler> = {
  [InsertLocation.before]: paragraphInsertBefore,
  [InsertLocation.after]: paragraphInsertAfter,
  [InsertLocation.start]: paragraphTextInsertStart,
  [InsertLocation.end]: paragraphTextInsertEnd,
  [InsertLocation.replace]: paragraphReplace,
};

const PARAGRAPH_LOCATION_HANDLERS: Record<InsertLocationValue, LocationHandler> = {
  [InsertLocation.before]: paragraphInsertBefore,
  [InsertLocation.after]: paragraphInsertAfter,
  [InsertLocation.start]: paragraphInsertStart,
  [InsertLocation.end]: paragraphInsertEnd,
  [InsertLocation.replace]: paragraphReplace,
};

type ParagraphProperty = "text";

// Whole-paragraph granularity (design spec, established alongside Range in
// #12): wraps exactly one <w:p>. Style and other Paragraph-only properties
// remain out of scope for v1.
export class Paragraph implements Syncable {
  private readonly tracked = new TrackedProperties();

  constructor(
    private readonly doc: FlatOpcDocument,
    private readonly target: Element,
    private readonly platform: SupportedPlatform,
    private readonly enqueue: (op: QueuedOperation) => void
  ) {}

  insertText(text: string, insertLocation: InsertLocationValue): void {
    this.insert(TEXT_LOCATION_HANDLERS, text, insertLocation);
  }

  insertParagraph(text: string, insertLocation: InsertLocationValue): void {
    this.insert(PARAGRAPH_LOCATION_HANDLERS, text, insertLocation);
  }

  // Whole-paragraph granularity means every rangeLocation returns a Range
  // over this same paragraph — same reasoning as Range.getRange() (#12).
  getRange(_rangeLocation: RangeLocationValue = RangeLocation.whole): Range {
    return new Range(this.doc, this.target, this.platform, this.enqueue);
  }

  load(propertyNames: ParagraphProperty | ParagraphProperty[]): void {
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
        return this.doc.getParagraphText(this.target);
      default:
        throw new Error(`Paragraph: unknown property "${name}"`);
    }
  }

  private insert(
    handlers: Record<InsertLocationValue, LocationHandler>,
    text: string,
    insertLocation: InsertLocationValue
  ): void {
    this.enqueue(() => {
      handlers[insertLocation](this.doc, this.target, text);
    });
  }
}

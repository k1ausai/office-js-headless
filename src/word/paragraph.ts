import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { paragraphInsertAfter } from "../operations/paragraphInsertAfter";
import { paragraphInsertBefore } from "../operations/paragraphInsertBefore";
import { paragraphInsertEnd } from "../operations/paragraphInsertEnd";
import { paragraphInsertStart } from "../operations/paragraphInsertStart";
import { paragraphReplace } from "../operations/paragraphReplace";
import { paragraphTextInsertEnd } from "../operations/paragraphTextInsertEnd";
import { paragraphTextInsertStart } from "../operations/paragraphTextInsertStart";
import { InsertLocation, InsertLocationValue } from "./insertLocation";
import type { QueuedOperation } from "./run";

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

// MVP scope (issue #9): wraps exactly one <w:p>. Full Paragraph mechanics
// (getRange, style, and the paragraphs.getFirst()/getLast() factory that
// constructs one through the real API path) are issue #14's job; this class
// only proves InsertLocation splicing is correct per receiver type.
export class Paragraph {
  constructor(
    private readonly doc: FlatOpcDocument,
    private readonly target: Element,
    private readonly enqueue: (op: QueuedOperation) => void
  ) {}

  insertText(text: string, insertLocation: InsertLocationValue): void {
    this.insert(TEXT_LOCATION_HANDLERS, text, insertLocation);
  }

  insertParagraph(text: string, insertLocation: InsertLocationValue): void {
    this.insert(PARAGRAPH_LOCATION_HANDLERS, text, insertLocation);
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

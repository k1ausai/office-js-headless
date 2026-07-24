import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { rangeInsertAfter } from "../operations/rangeInsertAfter";
import { rangeInsertBefore } from "../operations/rangeInsertBefore";
import { rangeInsertEnd } from "../operations/rangeInsertEnd";
import { rangeInsertStart } from "../operations/rangeInsertStart";
import { rangeParagraphInsertEnd } from "../operations/rangeParagraphInsertEnd";
import { rangeParagraphInsertStart } from "../operations/rangeParagraphInsertStart";
import { rangeReplace } from "../operations/rangeReplace";
import { InsertLocation, InsertLocationValue } from "./insertLocation";
import type { QueuedOperation } from "./run";

type LocationHandler = (doc: FlatOpcDocument, target: Element, text: string) => void;

// insertText's Start/End insert text WITHIN the range's own content;
// insertParagraph's Start/End always create a whole new paragraph — genuinely
// different operations at the same (receiver, location) pair. Before/After/
// Replace behave identically for both methods.
const TEXT_LOCATION_HANDLERS: Record<InsertLocationValue, LocationHandler> = {
  [InsertLocation.before]: rangeInsertBefore,
  [InsertLocation.after]: rangeInsertAfter,
  [InsertLocation.start]: rangeInsertStart,
  [InsertLocation.end]: rangeInsertEnd,
  [InsertLocation.replace]: rangeReplace,
};

const PARAGRAPH_LOCATION_HANDLERS: Record<InsertLocationValue, LocationHandler> = {
  [InsertLocation.before]: rangeInsertBefore,
  [InsertLocation.after]: rangeInsertAfter,
  [InsertLocation.start]: rangeParagraphInsertStart,
  [InsertLocation.end]: rangeParagraphInsertEnd,
  [InsertLocation.replace]: rangeReplace,
};

// MVP scope (issue #9): wraps exactly one whole <w:p> — no sub-paragraph
// offsets yet. Full Range mechanics (getOoxml, delete, select, search,
// arbitrary sub-paragraph spans) are issue #12's job; this class only proves
// InsertLocation splicing is correct per receiver type. Not yet reachable
// from Body (no body.getRange() factory yet — also #12).
export class Range {
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

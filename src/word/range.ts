import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { SupportedPlatform } from "../office/context";
import { rangeInsertAfter } from "../operations/rangeInsertAfter";
import { rangeInsertBefore } from "../operations/rangeInsertBefore";
import { rangeInsertEnd } from "../operations/rangeInsertEnd";
import { rangeInsertHtmlAfter } from "../operations/rangeInsertHtmlAfter";
import { rangeInsertHtmlBefore } from "../operations/rangeInsertHtmlBefore";
import { rangeInsertHtmlEnd } from "../operations/rangeInsertHtmlEnd";
import { rangeInsertHtmlReplace } from "../operations/rangeInsertHtmlReplace";
import { rangeInsertHtmlStart } from "../operations/rangeInsertHtmlStart";
import { rangeInsertOoxmlAfter } from "../operations/rangeInsertOoxmlAfter";
import { rangeInsertOoxmlBefore } from "../operations/rangeInsertOoxmlBefore";
import { rangeInsertOoxmlEnd } from "../operations/rangeInsertOoxmlEnd";
import { rangeInsertOoxmlReplace } from "../operations/rangeInsertOoxmlReplace";
import { rangeInsertOoxmlStart } from "../operations/rangeInsertOoxmlStart";
import { rangeInsertStart } from "../operations/rangeInsertStart";
import { rangeParagraphInsertEnd } from "../operations/rangeParagraphInsertEnd";
import { rangeParagraphInsertStart } from "../operations/rangeParagraphInsertStart";
import { rangeReplace } from "../operations/rangeReplace";
import { assertOoxmlSupported } from "./errors";
import { InsertLocation, InsertLocationValue } from "./insertLocation";
import { RangeLocation, RangeLocationValue } from "./rangeLocation";
import type { QueuedOperation } from "./run";
import { SearchOptions } from "./searchOptions";
import { SelectionMode, SelectionModeValue } from "./selectionMode";

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

const OOXML_LOCATION_HANDLERS: Record<InsertLocationValue, LocationHandler> = {
  [InsertLocation.before]: rangeInsertOoxmlBefore,
  [InsertLocation.after]: rangeInsertOoxmlAfter,
  [InsertLocation.start]: rangeInsertOoxmlStart,
  [InsertLocation.end]: rangeInsertOoxmlEnd,
  [InsertLocation.replace]: rangeInsertOoxmlReplace,
};

const HTML_LOCATION_HANDLERS: Record<InsertLocationValue, LocationHandler> = {
  [InsertLocation.before]: rangeInsertHtmlBefore,
  [InsertLocation.after]: rangeInsertHtmlAfter,
  [InsertLocation.start]: rangeInsertHtmlStart,
  [InsertLocation.end]: rangeInsertHtmlEnd,
  [InsertLocation.replace]: rangeInsertHtmlReplace,
};

// Whole-paragraph granularity, permanently: wraps exactly one whole <w:p>,
// no sub-paragraph character offsets (design spec's Range model doesn't
// track cursor position within a paragraph — that's out of scope for v1,
// not deferred to a later issue). search() (#13) is the other source of
// Range instances; getRange() below returns more of the same shape. Not yet
// reachable from Word.run's context — Body has no getRange() factory in the
// design spec's v1 API coverage; Paragraph.getRange() (#14) and search()
// (#13) are the actual entry points, still to come.
export class Range {
  // Shared by Body.search() and Range.search() — both wrap a set of matched
  // paragraph elements as Range instances the same way.
  static fromParagraphs(
    doc: FlatOpcDocument,
    paragraphs: Element[],
    platform: SupportedPlatform,
    enqueue: (op: QueuedOperation) => void
  ): Range[] {
    return paragraphs.map((p) => new Range(doc, p, platform, enqueue));
  }

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

  insertOoxml(ooxml: string, insertLocation: InsertLocationValue): void {
    this.enqueue(() => {
      assertOoxmlSupported(this.platform);
      OOXML_LOCATION_HANDLERS[insertLocation](this.doc, this.target, ooxml);
    });
  }

  // Never platform-gated, unlike insertOoxml — design spec's "Platform
  // selection": insertHtml is fully supported on every platform, no
  // divergence to model.
  insertHtml(html: string, insertLocation: InsertLocationValue): void {
    this.insert(HTML_LOCATION_HANDLERS, html, insertLocation);
  }

  getOoxml(): string {
    return this.doc.getRangeOoxml(this.target);
  }

  delete(): void {
    this.enqueue(() => {
      this.doc.deleteNode(this.target);
    });
  }

  // Real select() changes UI selection — inherently unobservable in a
  // headless shim with no UI. Still enqueued, matching every other Range
  // action method's deferred-until-sync() timing, but applies no document
  // mutation.
  select(_selectionMode: SelectionModeValue = SelectionMode.select): void {
    this.enqueue(() => {});
  }

  // Whole-paragraph granularity means Start/End/After/Content/Whole can't be
  // distinguished — there's no sub-paragraph position to collapse to, so
  // every rangeLocation returns a new Range over the same paragraph. Never
  // wrong, just not as precise as real Word's character-level cursor.
  getRange(_rangeLocation: RangeLocationValue = RangeLocation.whole): Range {
    return new Range(this.doc, this.target, this.platform, this.enqueue);
  }

  // Synchronous, like getRange() — a query, not a mutation, matching real
  // Office.js's proxy-factory methods. Whole-paragraph granularity means
  // this range can match at most once (itself, if its own paragraph
  // contains `text`) — a real sub-paragraph Range could return several
  // matches within one paragraph, which isn't representable here.
  search(text: string, options?: SearchOptions): Range[] {
    const matches = this.doc
      .search(text, options?.matchCase ?? false)
      .filter((p) => p === this.target);
    return Range.fromParagraphs(this.doc, matches, this.platform, this.enqueue);
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

import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { SupportedPlatform } from "../office/context";
import { bodyInsertEnd } from "../operations/bodyInsertEnd";
import { bodyInsertHtmlEnd } from "../operations/bodyInsertHtmlEnd";
import { bodyInsertHtmlReplace } from "../operations/bodyInsertHtmlReplace";
import { bodyInsertHtmlStart } from "../operations/bodyInsertHtmlStart";
import { bodyInsertOoxmlEnd } from "../operations/bodyInsertOoxmlEnd";
import { bodyInsertOoxmlReplace } from "../operations/bodyInsertOoxmlReplace";
import { bodyInsertOoxmlStart } from "../operations/bodyInsertOoxmlStart";
import { bodyInsertStart } from "../operations/bodyInsertStart";
import { bodyReplace } from "../operations/bodyReplace";
import { Comment } from "./comment";
import { assertOoxmlSupported } from "./errors";
import { InsertLocation, InsertLocationValue } from "./insertLocation";
import { ParagraphCollection } from "./paragraphCollection";
import { TrackedProperties } from "./proxy";
import { Range } from "./range";
import type { QueuedOperation, Syncable } from "./run";
import { SearchOptions } from "./searchOptions";
import { TableCollection } from "./tableCollection";

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

const HTML_LOCATION_HANDLERS: Partial<Record<InsertLocationValue, LocationHandler>> = {
  [InsertLocation.start]: bodyInsertHtmlStart,
  [InsertLocation.end]: bodyInsertHtmlEnd,
  [InsertLocation.replace]: bodyInsertHtmlReplace,
};

export class Body {
  private readonly tracked = new TrackedProperties();

  constructor(
    private readonly doc: FlatOpcDocument,
    private readonly platform: SupportedPlatform,
    private readonly enqueue: (op: QueuedOperation) => void,
    private readonly registerSyncable: (obj: Syncable) => void
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

  // Never platform-gated, unlike insertOoxml — design spec's "Platform
  // selection": insertHtml is fully supported on every platform, no
  // divergence to model (it's real add-in code's actual fallback route
  // around insertOoxml's OfficeOnline gap).
  insertHtml(html: string, insertLocation: InsertLocationValue): void {
    this.dispatch(HTML_LOCATION_HANDLERS, html, insertLocation, "Body.insertHtml");
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

  // Synchronous, like Range.getRange() — a query, not a mutation.
  search(text: string, options?: SearchOptions): Range[] {
    const matches = this.doc.search(text, options?.matchCase ?? false);
    return Range.fromParagraphs(this.doc, matches, this.platform, this.enqueue);
  }

  get paragraphs(): ParagraphCollection {
    return new ParagraphCollection(this.doc, this.platform, this.enqueue, this.registerSyncable);
  }

  get tables(): TableCollection {
    return new TableCollection(this.doc, this.enqueue, this.registerSyncable);
  }

  // Synchronous, like .search()/.paragraphs — a query, not a mutation.
  // Top-level comments only, each carrying its own threaded replies via
  // .replies — matches real Word.CommentCollection, which doesn't flatten
  // replies in as peers.
  getComments(): Comment[] {
    return this.doc.getTopLevelCommentElements().map((c) => {
      const comment = new Comment(this.doc, c);
      this.registerSyncable(comment);
      return comment;
    });
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

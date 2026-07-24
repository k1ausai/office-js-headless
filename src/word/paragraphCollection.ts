import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { SupportedPlatform } from "../office/context";
import { richApiError } from "./errors";
import { Paragraph } from "./paragraph";
import type { QueuedOperation, Syncable } from "./run";

// getFirst()/getLast() are synchronous proxy factories, like Range.getRange()
// (#12) — not enqueued, since they don't mutate the document. Each created
// Paragraph is registered as syncable so a later context.sync() can snapshot
// whatever properties get .load()ed on it (see run.ts's Syncable doc comment,
// which anticipated exactly this: "every body.getRange() call makes a new
// Range").
export class ParagraphCollection {
  constructor(
    private readonly doc: FlatOpcDocument,
    private readonly platform: SupportedPlatform,
    private readonly enqueue: (op: QueuedOperation) => void,
    private readonly registerSyncable: (obj: Syncable) => void
  ) {}

  getFirst(): Paragraph {
    const paragraphs = this.doc.getParagraphElements();
    return this.wrap(this.requireItem(paragraphs[0]));
  }

  getLast(): Paragraph {
    const paragraphs = this.doc.getParagraphElements();
    return this.wrap(this.requireItem(paragraphs[paragraphs.length - 1]));
  }

  // Real Word's documented behavior: "Throws an ItemNotFound error if the
  // collection is empty." The exact message text isn't confirmed against a
  // real runtime capture (unlike the wayfinder-traced errors in errors.ts) —
  // best-effort wording, not a hard-coded assumption.
  private requireItem(paragraph: Element | undefined): Element {
    if (!paragraph) {
      throw richApiError("ItemNotFound", "The requested item was not found in the collection.");
    }
    return paragraph;
  }

  private wrap(target: Element): Paragraph {
    const paragraph = new Paragraph(this.doc, target, this.platform, this.enqueue);
    this.registerSyncable(paragraph);
    return paragraph;
  }
}

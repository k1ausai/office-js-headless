import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

// insertOoxml always inserts whole paragraph(s) (unlike insertText's Start,
// which inserts plain text within existing content) — for this MVP Range
// (wraps one whole <w:p>, no sub-paragraph split support), "insert ooxml at
// the start of my content" collapses to "insert immediately before me",
// mirroring the same collapse insertParagraph's Start/End already uses
// (see rangeParagraphInsertStart.ts). Revisit once Range gains real
// sub-paragraph splitting (issue #12).
export function rangeInsertOoxmlStart(doc: FlatOpcDocument, target: Element, ooxml: string) {
  return doc.insertOoxmlBefore(target, ooxml);
}

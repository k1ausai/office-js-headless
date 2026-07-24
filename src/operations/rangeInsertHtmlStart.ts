import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

// insertHtml always inserts whole paragraph(s) (unlike insertText's Start,
// which inserts plain text within existing content) — for this MVP Range
// (wraps one whole <w:p>, no sub-paragraph split support), "insert html at
// the start of my content" collapses to "insert immediately before me",
// mirroring the same collapse insertOoxml's Start already uses (see
// rangeInsertOoxmlStart.ts).
export function rangeInsertHtmlStart(doc: FlatOpcDocument, target: Element, html: string) {
  return doc.insertHtmlBefore(target, html);
}

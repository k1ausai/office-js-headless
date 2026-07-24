import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

// insertParagraph always creates a whole new paragraph (unlike insertText's
// Start, which inserts text within existing content via rangeInsertStart.ts)
// — so for a Range, "insert a new paragraph at the start of my content"
// collapses to "insert immediately before me" while this MVP Range wraps a
// whole <w:p> with no sub-paragraph split support. Revisit once Range gains
// real sub-paragraph splitting (issue #12).
export function rangeParagraphInsertStart(doc: FlatOpcDocument, target: Element, text: string) {
  return doc.insertParagraphBefore(target, text);
}

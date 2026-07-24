import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

// insertParagraph always creates a whole new paragraph (unlike insertText's
// Start/End, which insert text within existing content) — so for a
// Paragraph, "insert a new paragraph at the start of my content" collapses
// to "insert immediately before me" while this MVP Paragraph wraps a whole
// <w:p> with no sub-paragraph split support. Revisit once Range/Paragraph
// gain real sub-paragraph splitting (issue #12).
export function paragraphInsertStart(doc: FlatOpcDocument, target: Element, text: string) {
  return doc.insertParagraphBefore(target, text);
}

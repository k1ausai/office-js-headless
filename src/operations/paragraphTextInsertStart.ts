import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

// insertText's Start prepends text WITHIN the paragraph's own content —
// distinct from insertParagraph's Start (paragraphInsertStart.ts), which
// always creates a whole new paragraph.
export function paragraphTextInsertStart(doc: FlatOpcDocument, target: Element, text: string) {
  doc.prependTextInParagraph(target, text);
}

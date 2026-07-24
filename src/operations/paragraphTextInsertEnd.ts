import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

// See paragraphTextInsertStart.ts — insertText's End appends WITHIN the
// paragraph's own content, distinct from insertParagraph's End.
export function paragraphTextInsertEnd(doc: FlatOpcDocument, target: Element, text: string) {
  doc.appendTextInParagraph(target, text);
}

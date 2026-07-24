import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

// insertText's Start prepends text WITHIN the range's own paragraph content —
// genuinely distinct from Before (which splices a new sibling paragraph).
export function rangeInsertStart(doc: FlatOpcDocument, target: Element, text: string) {
  doc.prependTextInParagraph(target, text);
}

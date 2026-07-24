import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

// See paragraphInsertStart.ts — the same MVP collapse applies symmetrically:
// "insert a new paragraph at the end of my content" collapses to "insert
// immediately after me" until Paragraph gains real sub-paragraph splitting.
export function paragraphInsertEnd(doc: FlatOpcDocument, target: Element, text: string) {
  return doc.insertParagraphAfter(target, text);
}

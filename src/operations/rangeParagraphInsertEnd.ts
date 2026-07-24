import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

// See rangeParagraphInsertStart.ts — the same MVP collapse applies
// symmetrically for insertParagraph's End.
export function rangeParagraphInsertEnd(doc: FlatOpcDocument, target: Element, text: string) {
  return doc.insertParagraphAfter(target, text);
}

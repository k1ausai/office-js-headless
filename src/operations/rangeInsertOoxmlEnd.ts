import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

// See rangeInsertOoxmlStart.ts — the same MVP collapse applies
// symmetrically for End.
export function rangeInsertOoxmlEnd(doc: FlatOpcDocument, target: Element, ooxml: string) {
  return doc.insertOoxmlAfter(target, ooxml);
}

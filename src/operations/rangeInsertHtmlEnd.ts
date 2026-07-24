import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

// See rangeInsertHtmlStart.ts — the same MVP collapse applies
// symmetrically for End.
export function rangeInsertHtmlEnd(doc: FlatOpcDocument, target: Element, html: string) {
  return doc.insertHtmlAfter(target, html);
}

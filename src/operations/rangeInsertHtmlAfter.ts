import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function rangeInsertHtmlAfter(doc: FlatOpcDocument, target: Element, html: string) {
  return doc.insertHtmlAfter(target, html);
}

import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function rangeInsertHtmlBefore(doc: FlatOpcDocument, target: Element, html: string) {
  return doc.insertHtmlBefore(target, html);
}

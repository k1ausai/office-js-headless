import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function rangeInsertHtmlReplace(doc: FlatOpcDocument, target: Element, html: string) {
  return doc.replaceHtmlAtTarget(target, html);
}

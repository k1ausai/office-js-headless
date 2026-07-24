import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function rangeInsertOoxmlReplace(doc: FlatOpcDocument, target: Element, ooxml: string) {
  return doc.replaceOoxmlAtTarget(target, ooxml);
}

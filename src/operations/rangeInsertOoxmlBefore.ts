import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function rangeInsertOoxmlBefore(doc: FlatOpcDocument, target: Element, ooxml: string) {
  return doc.insertOoxmlBefore(target, ooxml);
}

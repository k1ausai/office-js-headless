import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function rangeInsertOoxmlAfter(doc: FlatOpcDocument, target: Element, ooxml: string) {
  return doc.insertOoxmlAfter(target, ooxml);
}

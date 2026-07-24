import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function rangeInsertAfter(doc: FlatOpcDocument, target: Element, text: string) {
  return doc.insertParagraphAfter(target, text);
}

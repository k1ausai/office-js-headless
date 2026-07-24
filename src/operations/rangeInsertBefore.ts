import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function rangeInsertBefore(doc: FlatOpcDocument, target: Element, text: string) {
  return doc.insertParagraphBefore(target, text);
}

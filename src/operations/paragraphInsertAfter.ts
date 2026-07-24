import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function paragraphInsertAfter(doc: FlatOpcDocument, target: Element, text: string) {
  return doc.insertParagraphAfter(target, text);
}

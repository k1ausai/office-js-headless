import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function rangeInsertEnd(doc: FlatOpcDocument, target: Element, text: string) {
  doc.appendTextInParagraph(target, text);
}

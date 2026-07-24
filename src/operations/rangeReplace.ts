import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function rangeReplace(doc: FlatOpcDocument, target: Element, text: string) {
  doc.replaceParagraphContent(target, text);
}

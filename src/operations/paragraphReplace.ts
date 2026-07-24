import { type Element } from "@xmldom/xmldom";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function paragraphReplace(doc: FlatOpcDocument, target: Element, text: string) {
  doc.replaceParagraphContent(target, text);
}

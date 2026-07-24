import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function bodyInsertEnd(doc: FlatOpcDocument, text: string) {
  return doc.appendParagraph(text);
}

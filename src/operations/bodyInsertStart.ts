import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function bodyInsertStart(doc: FlatOpcDocument, text: string) {
  return doc.insertParagraphAsFirstChild(text);
}

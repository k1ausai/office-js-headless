import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function bodyInsertOoxmlStart(doc: FlatOpcDocument, ooxml: string) {
  return doc.insertOoxmlAsFirstChild(ooxml);
}

import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function bodyInsertOoxmlEnd(doc: FlatOpcDocument, ooxml: string) {
  return doc.insertOoxmlAsLastChild(ooxml);
}

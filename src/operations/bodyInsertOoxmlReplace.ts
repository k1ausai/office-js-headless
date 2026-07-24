import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function bodyInsertOoxmlReplace(doc: FlatOpcDocument, ooxml: string) {
  return doc.replaceOoxmlBodyContent(ooxml);
}

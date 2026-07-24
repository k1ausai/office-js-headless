import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function bodyReplace(doc: FlatOpcDocument, text: string) {
  return doc.replaceBodyContent(text);
}

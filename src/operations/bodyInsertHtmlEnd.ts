import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function bodyInsertHtmlEnd(doc: FlatOpcDocument, html: string) {
  return doc.insertHtmlAsLastChild(html);
}

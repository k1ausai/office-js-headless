import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function bodyInsertHtmlStart(doc: FlatOpcDocument, html: string) {
  return doc.insertHtmlAsFirstChild(html);
}

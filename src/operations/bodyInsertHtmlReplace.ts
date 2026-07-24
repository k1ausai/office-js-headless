import { FlatOpcDocument } from "../document/FlatOpcDocument";

export function bodyInsertHtmlReplace(doc: FlatOpcDocument, html: string) {
  return doc.replaceHtmlBodyContent(html);
}

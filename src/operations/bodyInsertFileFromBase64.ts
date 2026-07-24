import JSZip from "jszip";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

// Design spec's "insertFileFromBase64 — deliberate fidelity gap": unzip the
// .docx (the one place this package touches real zip/binary I/O), copy
// styles.xml/numbering.xml entries in without real conflict resolution.
// Doesn't also insert the imported file's body content — real Word's full
// implementation does, but the stub's scope (issue #19) is limited to the
// styles/numbering merge, per the design spec and PRD's explicit framing
// ("exercise the template-import code path," not full-fidelity import).
export async function bodyInsertFileFromBase64(
  doc: FlatOpcDocument,
  base64File: string
): Promise<void> {
  const zip = await JSZip.loadAsync(base64File, { base64: true });
  const stylesXml = await zip.file("word/styles.xml")?.async("string");
  const numberingXml = await zip.file("word/numbering.xml")?.async("string");
  if (stylesXml) doc.mergeStylesXml(stylesXml);
  if (numberingXml) doc.mergeNumberingXml(numberingXml);
}

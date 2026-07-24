import JSZip from "jszip";
import { FlatOpcDocument } from "../document/FlatOpcDocument";

// Design spec's "insertFileFromBase64 — deliberate fidelity gap": unzip the
// .docx (the one place this package touches real zip/binary I/O), copy
// styles.xml/numbering.xml entries in without real conflict resolution.
// Doesn't also insert the imported file's body content — real Word's full
// implementation does, but the stub's scope (issue #19) is limited to the
// styles/numbering merge, per the design spec and PRD's explicit framing
// ("exercise the template-import code path," not full-fidelity import).
//
// One file for all InsertLocation values, unlike every other operation in
// this directory (AGENTS.md: one file per receiver x InsertLocation) —
// this stub has no per-location behavior at all (location is validated by
// the caller for signature parity, then never consulted here), so
// splitting into three identical-body files would be pure ceremony.
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

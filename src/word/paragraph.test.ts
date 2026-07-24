import { type Element } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { MINIMAL_SEED_OOXML } from "../document/__fixtures__/minimalSeed";
import { InsertLocation } from "./insertLocation";
import { Paragraph } from "./paragraph";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function firstParagraph(doc: FlatOpcDocument): Element {
  const paragraph = doc.bodyElement.getElementsByTagNameNS(W_NS, "p")[0];
  if (!paragraph) throw new Error("test fixture has no <w:p>");
  return paragraph;
}

function immediateEnqueue(op: () => void): void {
  op();
}

describe("Paragraph InsertLocation dispatch", () => {
  it("insertText Before/After/Start/End/Replace match Range's behavior (same underlying primitives)", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    const paragraph = new Paragraph(doc, target, immediateEnqueue);

    paragraph.insertText("PREPENDED ", InsertLocation.start);
    expect(doc.getParagraphText(target)).toBe("PREPENDED Seed paragraph.");

    paragraph.insertText(" APPENDED", InsertLocation.end);
    expect(doc.getParagraphText(target)).toBe("PREPENDED Seed paragraph. APPENDED");

    paragraph.insertText("Replaced.", InsertLocation.replace);
    expect(doc.getParagraphText(target)).toBe("Replaced.");
  });

  it("insertText Before/After create a new sibling paragraph", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    const paragraph = new Paragraph(doc, target, immediateEnqueue);
    paragraph.insertText("Before.", InsertLocation.before);
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("Before.")).toBeLessThan(ooxml.indexOf("Seed paragraph."));
  });

  it("insertParagraph Start/End collapse to Before/After — MVP limitation, documented in operations/paragraphInsertStart.ts", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    const paragraph = new Paragraph(doc, target, immediateEnqueue);

    paragraph.insertParagraph("Start-as-before.", InsertLocation.start);
    const ooxml = doc.getOoxml();
    // Lands as a new sibling paragraph before the target, same as Before.
    expect(ooxml.indexOf("Start-as-before.")).toBeLessThan(ooxml.indexOf("Seed paragraph."));
  });

  it("insertParagraph Replace swaps the paragraph's content entirely", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    const paragraph = new Paragraph(doc, target, immediateEnqueue);
    paragraph.insertParagraph("Replaced via insertParagraph.", InsertLocation.replace);
    expect(doc.getParagraphText(target)).toBe("Replaced via insertParagraph.");
  });
});

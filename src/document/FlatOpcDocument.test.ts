import { describe, expect, it } from "vitest";
import { FlatOpcDocument } from "./FlatOpcDocument";
import { MINIMAL_SEED_OOXML } from "./__fixtures__/minimalSeed";

describe("FlatOpcDocument", () => {
  it("parses the seed and round-trips it unchanged through getOoxml()", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const ooxml = doc.getOoxml();
    expect(ooxml).toContain("Seed paragraph.");
    expect(ooxml).toContain("<w:sectPr>");
  });

  it("appendParagraph adds a new paragraph before the trailing sectPr", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.appendParagraph("New paragraph.");
    const ooxml = doc.getOoxml();

    const sectPrIndex = ooxml.indexOf("<w:sectPr>");
    const newParaIndex = ooxml.indexOf("New paragraph.");
    expect(newParaIndex).toBeGreaterThan(-1);
    expect(newParaIndex).toBeLessThan(sectPrIndex);
  });

  it("appendParagraph keeps existing content and appends after it, in order", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.appendParagraph("First append.");
    doc.appendParagraph("Second append.");
    const ooxml = doc.getOoxml();

    expect(ooxml.indexOf("Seed paragraph.")).toBeLessThan(ooxml.indexOf("First append."));
    expect(ooxml.indexOf("First append.")).toBeLessThan(ooxml.indexOf("Second append."));
    expect(ooxml.indexOf("Second append.")).toBeLessThan(ooxml.indexOf("<w:sectPr>"));
  });

  it("appendParagraph works on a document with no sectPr (appends as last child)", () => {
    const seedWithoutSectPr = MINIMAL_SEED_OOXML.replace(/<w:sectPr>[\s\S]*?<\/w:sectPr>/, "");
    const doc = new FlatOpcDocument(seedWithoutSectPr);
    doc.appendParagraph("Appended.");
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("Seed paragraph.")).toBeLessThan(ooxml.indexOf("Appended."));
  });

  it("appendParagraph anchors on the body's own trailing sectPr, not a nested one from a mid-document section break", () => {
    // A <w:sectPr> nested inside a paragraph's <w:pPr> marks a mid-document
    // section break — a real, legal OOXML construct, and a DIFFERENT element
    // from <w:body>'s own trailing <w:sectPr>. Anchoring on it instead (a
    // subtree search would do this) means body.insertBefore(paragraph,
    // sectPr) throws, since that sectPr isn't body's direct child.
    const seedWithSectionBreak = MINIMAL_SEED_OOXML.replace(
      "</w:p>",
      `</w:p>
          <w:p>
            <w:pPr>
              <w:sectPr>
                <w:pgSz w:w="12240" w:h="15840"/>
              </w:sectPr>
            </w:pPr>
          </w:p>`
    );
    const doc = new FlatOpcDocument(seedWithSectionBreak);

    expect(() => doc.appendParagraph("After the section break.")).not.toThrow();
    const ooxml = doc.getOoxml();

    // Lands before the body's own trailing sectPr — the last sectPr in the
    // document — not before the mid-document section-break sectPr.
    const bodyTrailingSectPrIndex = ooxml.lastIndexOf("<w:sectPr>");
    expect(ooxml.indexOf("After the section break.")).toBeLessThan(bodyTrailingSectPrIndex);
  });

  it("reset() restores the document to its originally-seeded state", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.appendParagraph("Will be reset away.");
    expect(doc.getOoxml()).toContain("Will be reset away.");

    doc.reset();
    expect(doc.getOoxml()).not.toContain("Will be reset away.");
    expect(doc.getOoxml()).toContain("Seed paragraph.");
  });
});

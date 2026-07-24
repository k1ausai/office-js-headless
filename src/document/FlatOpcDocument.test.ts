import { type Element } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { FlatOpcDocument } from "./FlatOpcDocument";
import { MINIMAL_SEED_OOXML } from "./__fixtures__/minimalSeed";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function firstParagraph(doc: FlatOpcDocument): Element {
  const paragraph = doc.bodyElement.getElementsByTagNameNS(W_NS, "p")[0];
  if (!paragraph) throw new Error("test fixture has no <w:p>");
  return paragraph;
}

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

  it("getBodyText() joins each paragraph's text, in document order, with newlines", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.appendParagraph("Second paragraph.");
    doc.appendParagraph("Third paragraph.");
    expect(doc.getBodyText()).toBe("Seed paragraph.\nSecond paragraph.\nThird paragraph.");
  });

  it("insertParagraphBefore splices a new paragraph immediately before the target", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.insertParagraphBefore(firstParagraph(doc), "Before.");
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("Before.")).toBeLessThan(ooxml.indexOf("Seed paragraph."));
  });

  it("insertParagraphAfter splices a new paragraph immediately after the target", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const seedParagraph = firstParagraph(doc);
    doc.appendParagraph("Gamma.");
    doc.insertParagraphAfter(seedParagraph, "After.");
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("Seed paragraph.")).toBeLessThan(ooxml.indexOf("After."));
    expect(ooxml.indexOf("After.")).toBeLessThan(ooxml.indexOf("Gamma."));
  });

  it("replaceParagraphContent swaps the target paragraph's text entirely", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.replaceParagraphContent(firstParagraph(doc), "Replaced.");
    const ooxml = doc.getOoxml();
    expect(ooxml).toContain("Replaced.");
    expect(ooxml).not.toContain("Seed paragraph.");
  });

  it("prependTextInParagraph inserts text at the start of the target paragraph's content", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const seedParagraph = firstParagraph(doc);
    doc.prependTextInParagraph(seedParagraph, "PREPENDED ");
    expect(doc.getParagraphText(seedParagraph)).toBe("PREPENDED Seed paragraph.");
  });

  it("appendTextInParagraph inserts text at the end of the target paragraph's content", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const seedParagraph = firstParagraph(doc);
    doc.appendTextInParagraph(seedParagraph, " APPENDED");
    expect(doc.getParagraphText(seedParagraph)).toBe("Seed paragraph. APPENDED");
  });

  it("insertParagraphAsFirstChild inserts before all existing content", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.insertParagraphAsFirstChild("First now.");
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("First now.")).toBeLessThan(ooxml.indexOf("Seed paragraph."));
  });

  it("insertParagraphAsFirstChild works on a body with no existing paragraphs", () => {
    const emptyBodySeed = MINIMAL_SEED_OOXML.replace(
      /<w:p w14:paraId="00000001">[\s\S]*?<\/w:p>/,
      ""
    );
    const doc = new FlatOpcDocument(emptyBodySeed);
    expect(() => doc.insertParagraphAsFirstChild("Only paragraph.")).not.toThrow();
    expect(doc.getOoxml()).toContain("Only paragraph.");
  });

  it("replaceBodyContent clears the whole body and inserts one paragraph", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.appendParagraph("Second paragraph.");
    doc.replaceBodyContent("Only this remains.");
    const ooxml = doc.getOoxml();
    expect(ooxml).toContain("Only this remains.");
    expect(ooxml).not.toContain("Seed paragraph.");
    expect(ooxml).not.toContain("Second paragraph.");
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

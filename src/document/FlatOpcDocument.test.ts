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

  it("reset() restores the document to its originally-seeded state", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.appendParagraph("Will be reset away.");
    expect(doc.getOoxml()).toContain("Will be reset away.");

    doc.reset();
    expect(doc.getOoxml()).not.toContain("Will be reset away.");
    expect(doc.getOoxml()).toContain("Seed paragraph.");
  });
});

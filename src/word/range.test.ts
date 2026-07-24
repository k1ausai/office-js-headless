import { type Element } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { FlatOpcDocument } from "../document/FlatOpcDocument";
import { MINIMAL_SEED_OOXML } from "../document/__fixtures__/minimalSeed";
import { InsertLocation } from "./insertLocation";
import { Range } from "./range";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function firstParagraph(doc: FlatOpcDocument): Element {
  const paragraph = doc.bodyElement.getElementsByTagNameNS(W_NS, "p")[0];
  if (!paragraph) throw new Error("test fixture has no <w:p>");
  return paragraph;
}

// Executes queued ops immediately — these tests target splicing correctness
// per location, not deferred-batching (already covered in run.test.ts).
function immediateEnqueue(op: () => void): void {
  op();
}

describe("Range InsertLocation dispatch", () => {
  it("Before splices a new sibling paragraph immediately before the range", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const range = new Range(doc, firstParagraph(doc), "PC", immediateEnqueue);
    range.insertText("Before.", InsertLocation.before);
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("Before.")).toBeLessThan(ooxml.indexOf("Seed paragraph."));
  });

  it("After splices a new sibling paragraph immediately after the range", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    doc.appendParagraph("Gamma.");
    const range = new Range(doc, target, "PC", immediateEnqueue);
    range.insertText("After.", InsertLocation.after);
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("Seed paragraph.")).toBeLessThan(ooxml.indexOf("After."));
    expect(ooxml.indexOf("After.")).toBeLessThan(ooxml.indexOf("Gamma."));
  });

  it("Start prepends text within the range's own paragraph — no new paragraph", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    const range = new Range(doc, target, "PC", immediateEnqueue);
    range.insertText("PREPENDED ", InsertLocation.start);
    expect(doc.getParagraphText(target)).toBe("PREPENDED Seed paragraph.");
    expect(doc.getBodyText().split("\n")).toHaveLength(1);
  });

  it("End appends text within the range's own paragraph — no new paragraph", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    const range = new Range(doc, target, "PC", immediateEnqueue);
    range.insertText(" APPENDED", InsertLocation.end);
    expect(doc.getParagraphText(target)).toBe("Seed paragraph. APPENDED");
    expect(doc.getBodyText().split("\n")).toHaveLength(1);
  });

  it("Replace swaps the range's own content entirely", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    const range = new Range(doc, target, "PC", immediateEnqueue);
    range.insertText("Replaced.", InsertLocation.replace);
    expect(doc.getParagraphText(target)).toBe("Replaced.");
  });

  it("insertParagraph(Before/After) creates a new sibling paragraph, same as insertText", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    const range = new Range(doc, target, "PC", immediateEnqueue);
    range.insertParagraph("New para.", InsertLocation.after);
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("Seed paragraph.")).toBeLessThan(ooxml.indexOf("New para."));
  });

  it("insertParagraph(Start/End) creates a new sibling paragraph too — NOT text-within-paragraph like insertText(Start/End)", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    const range = new Range(doc, target, "PC", immediateEnqueue);

    range.insertParagraph("Start-as-before.", InsertLocation.start);
    expect(doc.getParagraphText(target)).toBe("Seed paragraph.");
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("Start-as-before.")).toBeLessThan(ooxml.indexOf("Seed paragraph."));
  });

  it("insertParagraph(End) collapses to after — new sibling paragraph, target unchanged", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    const range = new Range(doc, target, "PC", immediateEnqueue);

    range.insertParagraph("End-as-after.", InsertLocation.end);
    expect(doc.getParagraphText(target)).toBe("Seed paragraph.");
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("Seed paragraph.")).toBeLessThan(ooxml.indexOf("End-as-after."));
  });

  it("insertParagraph(Replace) swaps the range's own content, same as insertText(Replace)", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    const range = new Range(doc, target, "PC", immediateEnqueue);
    range.insertParagraph("Replaced via insertParagraph.", InsertLocation.replace);
    expect(doc.getParagraphText(target)).toBe("Replaced via insertParagraph.");
  });

  it("insertOoxml applies the fragment on PC, splicing it after the target", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    const range = new Range(doc, target, "PC", immediateEnqueue);
    const fragment = `<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}">
        <w:body>
          <w:p><w:r><w:t>Fragment content.</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
    range.insertOoxml(fragment, InsertLocation.after);
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("Seed paragraph.")).toBeLessThan(ooxml.indexOf("Fragment content."));
  });

  it("insertOoxml throws on OfficeOnline (no client-side merge engine there)", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    const range = new Range(doc, target, "OfficeOnline", immediateEnqueue);
    expect(() => range.insertOoxml("<pkg:package/>", InsertLocation.after)).toThrow();
  });
});

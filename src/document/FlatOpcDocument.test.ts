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

  it("insertOoxmlBefore imports and splices the fragment's paragraphs before the target, preserving order", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    const fragment = `<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}">
        <w:body>
          <w:p><w:r><w:t>Fragment one.</w:t></w:r></w:p>
          <w:p><w:r><w:t>Fragment two.</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
    doc.insertOoxmlBefore(target, fragment);
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("Fragment one.")).toBeLessThan(ooxml.indexOf("Fragment two."));
    expect(ooxml.indexOf("Fragment two.")).toBeLessThan(ooxml.indexOf("Seed paragraph."));
  });

  it("insertOoxmlAfter imports and splices the fragment's paragraphs after the target, preserving order", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    doc.appendParagraph("Gamma.");
    const fragment = `<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}">
        <w:body>
          <w:p><w:r><w:t>Fragment one.</w:t></w:r></w:p>
          <w:p><w:r><w:t>Fragment two.</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
    doc.insertOoxmlAfter(target, fragment);
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("Seed paragraph.")).toBeLessThan(ooxml.indexOf("Fragment one."));
    expect(ooxml.indexOf("Fragment one.")).toBeLessThan(ooxml.indexOf("Fragment two."));
    expect(ooxml.indexOf("Fragment two.")).toBeLessThan(ooxml.indexOf("Gamma."));
  });

  it("insertOoxmlAsFirstChild inserts the fragment's paragraphs before all existing content, preserving order", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const fragment = `<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}">
        <w:body>
          <w:p><w:r><w:t>Frag one.</w:t></w:r></w:p>
          <w:p><w:r><w:t>Frag two.</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
    doc.insertOoxmlAsFirstChild(fragment);
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("Frag one.")).toBeLessThan(ooxml.indexOf("Frag two."));
    expect(ooxml.indexOf("Frag two.")).toBeLessThan(ooxml.indexOf("Seed paragraph."));
  });

  it("insertOoxmlAsLastChild inserts the fragment's paragraphs after all existing content, before sectPr", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const fragment = `<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}">
        <w:body>
          <w:p><w:r><w:t>Frag one.</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
    doc.insertOoxmlAsLastChild(fragment);
    const ooxml = doc.getOoxml();
    expect(ooxml.indexOf("Seed paragraph.")).toBeLessThan(ooxml.indexOf("Frag one."));
    expect(ooxml.indexOf("Frag one.")).toBeLessThan(ooxml.indexOf("<w:sectPr>"));
  });

  it("replaceOoxmlBodyContent clears the whole body and inserts the fragment's paragraphs", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const fragment = `<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}">
        <w:body>
          <w:p><w:r><w:t>Only this.</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
    doc.replaceOoxmlBodyContent(fragment);
    const ooxml = doc.getOoxml();
    expect(ooxml).toContain("Only this.");
    expect(ooxml).not.toContain("Seed paragraph.");
  });

  it("replaceOoxmlAtTarget removes the target and inserts the fragment's paragraphs in its place", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    const fragment = `<?xml version="1.0" encoding="UTF-8"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}">
        <w:body>
          <w:p><w:r><w:t>Replacement one.</w:t></w:r></w:p>
          <w:p><w:r><w:t>Replacement two.</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
    doc.replaceOoxmlAtTarget(target, fragment);
    const ooxml = doc.getOoxml();
    expect(ooxml).not.toContain("Seed paragraph.");
    expect(ooxml.indexOf("Replacement one.")).toBeLessThan(ooxml.indexOf("Replacement two."));
  });

  it("reset() restores the document to its originally-seeded state", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.appendParagraph("Will be reset away.");
    expect(doc.getOoxml()).toContain("Will be reset away.");

    doc.reset();
    expect(doc.getOoxml()).not.toContain("Will be reset away.");
    expect(doc.getOoxml()).toContain("Seed paragraph.");
  });

  it("deleteNode removes the target paragraph, leaving the rest of the body intact", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    const target = firstParagraph(doc);
    doc.appendParagraph("Survivor.");

    doc.deleteNode(target);

    expect(doc.getOoxml()).not.toContain("Seed paragraph.");
    expect(doc.getOoxml()).toContain("Survivor.");
  });

  it("getRangeOoxml serializes only the target paragraph, wrapped as its own Flat-OPC package", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
    doc.appendParagraph("Other paragraph.");
    const target = firstParagraph(doc);

    const rangeOoxml = doc.getRangeOoxml(target);

    expect(rangeOoxml).toContain("<pkg:package");
    expect(rangeOoxml).toContain("Seed paragraph.");
    expect(rangeOoxml).not.toContain("Other paragraph.");
  });

  it("getRangeOoxml on OfficeOnline churns the target's ids just like a whole-document read would", () => {
    const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML, "OfficeOnline");
    const target = firstParagraph(doc);
    const paraIdPattern = /w14:paraId="([0-9A-F]{8})"/;

    const firstRead = doc.getRangeOoxml(target);
    const secondRead = doc.getRangeOoxml(target);

    expect(firstRead.match(paraIdPattern)?.[1]).not.toBe(secondRead.match(paraIdPattern)?.[1]);
  });

  describe("search", () => {
    it("returns the paragraphs whose text contains the search string, case-insensitively by default", () => {
      const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
      doc.appendParagraph("Another one.");

      const matches = doc.search("SEED", false);

      expect(matches).toHaveLength(1);
      expect(doc.getParagraphText(matches[0]!)).toBe("Seed paragraph.");
    });

    it("matchCase: true only matches on exact case", () => {
      const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);

      expect(doc.search("SEED", true)).toHaveLength(0);
      expect(doc.search("Seed", true)).toHaveLength(1);
    });

    it("returns no matches when the text isn't found anywhere", () => {
      const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
      expect(doc.search("nonexistent", false)).toHaveLength(0);
    });

    it("treats wildcard-special characters in the search text literally, never as a pattern", () => {
      const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
      doc.replaceParagraphContent(firstParagraph(doc), "Contains * and ? and .* literally.");

      expect(doc.search("* and ?", false)).toHaveLength(1);
      // A regex-as-pattern interpretation of ".*" would match everything;
      // plain substring matching must not.
      expect(doc.search(".*", false)).toHaveLength(1);
      expect(doc.search("nomatch.*", false)).toHaveLength(0);
    });

    it("excludes the trailing paragraph mark from search results — it's not real content", () => {
      const doc = new FlatOpcDocument(MINIMAL_SEED_OOXML);
      expect(doc.search("", false).length).toBe(1);
    });
  });
});

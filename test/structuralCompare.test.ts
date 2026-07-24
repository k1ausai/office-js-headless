import { describe, expect, it } from "vitest";
import { expectStructuralMatch, structuralSummary } from "./structuralCompare";

const SEED = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="00000001"><w:r><w:t>First.</w:t></w:r></w:p>
          <w:p w14:paraId="00000002"><w:r><w:t>Second.</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

describe("structuralSummary", () => {
  it("extracts paragraph text in document order, excluding the trailing mark", () => {
    expect(structuralSummary(SEED).paragraphTexts).toEqual(["First.", "Second."]);
  });

  it("reports paraIdsValid: true when every real paragraph has a well-formed 8-hex-uppercase paraId", () => {
    expect(structuralSummary(SEED).paraIdsValid).toBe(true);
  });

  it("reports paraIdsValid: false when a real paragraph is missing a paraId", () => {
    const noId = SEED.replace(' w14:paraId="00000001"', "");
    expect(structuralSummary(noId).paraIdsValid).toBe(false);
  });

  it("extracts style names when the document has a styles.xml part, empty array otherwise", () => {
    expect(structuralSummary(SEED).styleNames).toEqual([]);

    const withStyles = SEED.replace(
      "</pkg:package>",
      `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml">
        <pkg:xmlData>
          <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
          </w:styles>
        </pkg:xmlData>
      </pkg:part></pkg:package>`
    );
    expect(structuralSummary(withStyles).styleNames).toEqual(["Normal"]);
  });
});

const RESULT_WITH_TRAILING_MARK = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="00000001"><w:r><w:t>Real content.</w:t></w:r></w:p>
          <w:p w14:paraId="00000002"><w:r><w:t></w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

describe("structuralSummary — getOoxml()-shaped input", () => {
  it("excludes a trailing empty paragraph — every getOoxml() dump (this shim's or real Word's) carries one as its actual last <w:p>", () => {
    expect(structuralSummary(RESULT_WITH_TRAILING_MARK).paragraphTexts).toEqual(["Real content."]);
  });

  it("two getOoxml()-shaped strings differing only in the trailing mark's id still compare equal — re-parsing must not double-count it", () => {
    const other = RESULT_WITH_TRAILING_MARK.replace(
      'w14:paraId="00000002"',
      'w14:paraId="ABCDEF02"'
    );
    expect(() => expectStructuralMatch(RESULT_WITH_TRAILING_MARK, other)).not.toThrow();
  });
});

const WITH_TABLE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:tbl>
            <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
            <w:tr>
              <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>R1C1</w:t></w:r></w:p></w:tc>
              <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>R1C2</w:t></w:r></w:p></w:tc>
            </w:tr>
          </w:tbl>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

describe("structuralSummary — tables", () => {
  it("extracts table cell text, in document (row-major) order — table-cell paragraphs live outside body's direct paragraph list", () => {
    expect(structuralSummary(WITH_TABLE).tableCellTexts).toEqual(["R1C1", "R1C2"]);
  });

  it("expectStructuralMatch catches a table-cell content difference that paragraphTexts alone would miss", () => {
    const different = WITH_TABLE.replace("R1C2", "Different.");
    expect(() => expectStructuralMatch(WITH_TABLE, different)).toThrow();
  });
});

describe("expectStructuralMatch", () => {
  it("passes when paragraph text, style names, and paraId validity all match", () => {
    expect(() => expectStructuralMatch(SEED, SEED)).not.toThrow();
  });

  it("throws when paragraph text differs", () => {
    const different = SEED.replace("First.", "Different.");
    expect(() => expectStructuralMatch(SEED, different)).toThrow();
  });

  it("does not require exact paraId value equality — only format validity on both sides", () => {
    const differentIds = SEED.replace('w14:paraId="00000001"', 'w14:paraId="ABCDEF01"');
    expect(() => expectStructuralMatch(SEED, differentIds)).not.toThrow();
  });
});

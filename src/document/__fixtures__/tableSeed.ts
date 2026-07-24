// A minimal Flat-OPC seed with one paragraph followed by a plain, unmerged
// 2-row x 3-column table — no w:tblPr/w:trPr (issue #15: omitted entirely
// for v1), one w:p per cell (the minimum valid OOXML cell content).
export const TABLE_SEED_OOXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="00000001">
            <w:r>
              <w:t>Seed paragraph.</w:t>
            </w:r>
          </w:p>
          <w:tbl>
            <w:tblGrid>
              <w:gridCol w:w="2000"/>
              <w:gridCol w:w="2000"/>
              <w:gridCol w:w="2000"/>
            </w:tblGrid>
            <w:tr>
              <w:tc>
                <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                <w:p><w:r><w:t>R1C1</w:t></w:r></w:p>
              </w:tc>
              <w:tc>
                <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                <w:p><w:r><w:t>R1C2</w:t></w:r></w:p>
              </w:tc>
              <w:tc>
                <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                <w:p><w:r><w:t>R1C3</w:t></w:r></w:p>
              </w:tc>
            </w:tr>
            <w:tr>
              <w:tc>
                <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                <w:p><w:r><w:t>R2C1</w:t></w:r></w:p>
              </w:tc>
              <w:tc>
                <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                <w:p><w:r><w:t>R2C2</w:t></w:r></w:p>
              </w:tc>
              <w:tc>
                <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                <w:p><w:r><w:t>R2C3</w:t></w:r></w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>
`;

// A 2x2-grid table where row 1 has a single horizontally-merged cell
// (gridSpan="2") spanning both grid columns, and column 1 has a vertically
// merged pair (vMerge w:val="restart" on row 1, bare <w:vMerge/> — schema
// default "continue" — on row 2). Exercises issue #15's "table with a
// pre-existing merge doesn't silently miscount cells" requirement: row 1
// has only 1 <w:tc> and row 2 has 2 <w:tc>, but w:tblGrid still reports 2
// columns.
export const MERGED_TABLE_SEED_OOXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:tbl>
            <w:tblGrid>
              <w:gridCol w:w="2000"/>
              <w:gridCol w:w="2000"/>
            </w:tblGrid>
            <w:tr>
              <w:tc>
                <w:tcPr>
                  <w:tcW w:w="4000" w:type="dxa"/>
                  <w:gridSpan w:val="2"/>
                  <w:vMerge w:val="restart"/>
                </w:tcPr>
                <w:p><w:r><w:t>Merged top.</w:t></w:r></w:p>
              </w:tc>
            </w:tr>
            <w:tr>
              <w:tc>
                <w:tcPr>
                  <w:tcW w:w="2000" w:type="dxa"/>
                  <w:vMerge/>
                </w:tcPr>
                <w:p><w:r><w:t></w:t></w:r></w:p>
              </w:tc>
              <w:tc>
                <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
                <w:p><w:r><w:t>R2C2</w:t></w:r></w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>
`;

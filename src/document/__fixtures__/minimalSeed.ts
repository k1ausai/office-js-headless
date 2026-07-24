// Minimal Flat-OPC seed document — just the `/word/document.xml` part, since
// that's all FlatOpcDocument needs for body-level operations. Real Word's own
// Flat-OPC dump includes many more parts (styles.xml, relationships, content
// types, ...); those aren't modeled here because nothing in this shim reads
// them yet.
export const MINIMAL_SEED_OOXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
          <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
          </w:sectPr>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>
`;

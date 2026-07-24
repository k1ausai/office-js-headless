// A Flat-OPC seed with document.xml + styles.xml parts. styles.xml has four
// distinct paragraph styles: three built-in (Normal, Heading1, Title — no
// w:customStyle attribute, matching real Word's convention for its built-in
// styles, whose w:styleId values are exactly the Word.BuiltInStyleName
// enum's string values) and one custom (w:customStyle="1"). Only w:type,
// w:styleId, and w:name are modeled — w:basedOn/w:next/w:link/w:uiPriority/
// w:pPr/w:rPr aren't read by this shim's style primitives, so they're
// omitted rather than included unused.
export const STYLE_SEED_OOXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml">
    <pkg:xmlData>
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:type="paragraph" w:styleId="Normal">
          <w:name w:val="Normal"/>
        </w:style>
        <w:style w:type="paragraph" w:styleId="Heading1">
          <w:name w:val="heading 1"/>
        </w:style>
        <w:style w:type="paragraph" w:styleId="Title">
          <w:name w:val="Title"/>
        </w:style>
        <w:style w:type="paragraph" w:customStyle="1" w:styleId="MyCustomStyle">
          <w:name w:val="My Custom Style"/>
        </w:style>
      </w:styles>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>
`;

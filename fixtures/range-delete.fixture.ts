import type { Fixture } from "../test/types";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const fixture: Fixture = {
  description: "Range.delete() removes the range's paragraph from the document",
  seedOoxml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="00000001"><w:r><w:t>Keep this.</w:t></w:r></w:p>
          <w:p w14:paraId="00000002"><w:r><w:t>Delete this.</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`,
  apply: (context) => {
    const lastParagraph = context.document.body.paragraphs.getLast();
    const range = lastParagraph.getRange();
    range.delete();
  },
  resultOoxml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="00000001"><w:r><w:t>Keep this.</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`,
};

export default fixture;

import { InsertLocation } from "../src/word/insertLocation";
import type { Fixture } from "../test/types";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const fixture: Fixture = {
  description: "Body.insertHtml(End) converts <b>/<i> formatting into w:b/w:i runs",
  seedOoxml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="00000001"><w:r><w:t>Seed paragraph.</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`,
  apply: (context) => {
    context.document.body.insertHtml("<p><b>Bold</b> and <i>italic</i>.</p>", InsertLocation.end);
  },
  resultOoxml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
        <w:body>
          <w:p w14:paraId="00000001"><w:r><w:t>Seed paragraph.</w:t></w:r></w:p>
          <w:p w14:paraId="00000002">
            <w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>
            <w:r><w:t> and </w:t></w:r>
            <w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>
            <w:r><w:t>.</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`,
};

export default fixture;
